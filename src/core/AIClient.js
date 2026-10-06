/* AIClient — BYOK multi-provider, lazy-loaded. Daftar provider dari config/providers.json
   (preset, ikut git) + localStorage (custom per-perangkat). Gaya API: openai|gemini|anthropic.
   Streaming + AbortController agar UI tetap responsif. */
import { SYSTEM, buildUserPrompt } from "./PromptPack.js";

const REG_URL = "../../config/providers.json"; // relatif thd taskpane.html (src/taskpane/ → root/config/)
const LS_CUSTOM = "docassist:customProviders";
const LS_REG_CACHE = "docassist:registryCache";
const FALLBACK = [
  { id: "openai", label: "OpenAI", style: "openai", baseUrl: "https://api.openai.com/v1", models: ["gpt-4o-mini", "gpt-4o"] },
  { id: "gemini", label: "Google Gemini", style: "gemini", baseUrl: "https://generativelanguage.googleapis.com/v1beta", models: ["gemini-2.0-flash"] },
];

let regMem = null;
export function getCustomProviders() { try { return JSON.parse(localStorage.getItem(LS_CUSTOM) || "[]"); } catch (e) { return []; } }
export function saveCustomProvider(p) {
  const all = getCustomProviders().filter((x) => x.id !== p.id); all.push(p);
  localStorage.setItem(LS_CUSTOM, JSON.stringify(all)); regMem = null; return p;
}
export function deleteCustomProvider(id) {
  localStorage.setItem(LS_CUSTOM, JSON.stringify(getCustomProviders().filter((x) => x.id !== id))); regMem = null;
}

// Preset file + custom runtime. Cache file ke localStorage agar tetap jalan offline/CDN gagal.
export async function getProviders() {
  if (regMem) return regMem;
  let preset = null;
  try { const r = await fetch(REG_URL, { cache: "no-cache" }); if (r.ok) preset = (await r.json()).providers; } catch (e) {}
  if (preset && preset.length) { try { localStorage.setItem(LS_REG_CACHE, JSON.stringify(preset)); } catch (e) {} }
  else { try { preset = JSON.parse(localStorage.getItem(LS_REG_CACHE) || "null") || FALLBACK; } catch (e) { preset = FALLBACK; } }
  regMem = [...preset, ...getCustomProviders()];
  return regMem;
}
export async function getProvider(id) {
  const all = await getProviders();
  return all.find((p) => p.id === id) || all[0];
}

const cfg = () => ({ p: localStorage.getItem("docassist:provider") || "openai",
  k: localStorage.getItem("docassist:key") || "", m: localStorage.getItem("docassist:model") || "",
  t: parseFloat(localStorage.getItem("docassist:temp") || "0.3") });

let abort = null;
export function cancel() { abort && abort.abort(); }
const base = (prov) => String(prov.baseUrl || "").replace(/\/+$/, "");

// Gaya OpenAI-compatible: OpenAI, Groq, OpenRouter, DeepSeek, Mistral, Cerebras, HF, Ollama, dsb. SSE streaming.
async function callOpenAIStyle(prov, q, ctx, onTok) {
  const { k, m, t } = cfg();
  abort = new AbortController();
  const headers = { "Content-Type": "application/json" };
  if (!prov.noKey) headers.Authorization = `Bearer ${k}`;
  const r = await fetch(`${base(prov)}${prov.chatPath || "/chat/completions"}`, { method: "POST", signal: abort.signal, headers,
    body: JSON.stringify({ model: m || prov.models?.[0], temperature: t, stream: true,
      messages: [{ role: "system", content: SYSTEM }, { role: "user", content: buildUserPrompt(q, ctx) }] }) });
  if (!r.ok) throw new Error(prov.label + " " + r.status);
  const rd = r.body.getReader(), dec = new TextDecoder(); let buf = "", out = "";
  for (;;) { const { done, value } = await rd.read(); if (done) break;
    buf += dec.decode(value, { stream: true });
    for (const line of buf.split("\n")) { const s = line.trim(); if (!s.startsWith("data:")) continue;
      const d = s.slice(5).trim(); if (d === "[DONE]") break;
      try { const tok = JSON.parse(d).choices?.[0]?.delta?.content || ""; out += tok; onTok && onTok(tok); } catch (e) {} }
    buf = buf.slice(-2000); }
  return out;
}

async function callGeminiStyle(prov, q, ctx, onTok) {
  const { k, m, t } = cfg();
  abort = new AbortController();
  const r = await fetch(`${base(prov)}/models/${m || prov.models?.[0]}:generateContent?key=${encodeURIComponent(k)}`,
    { method: "POST", signal: abort.signal, headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ systemInstruction: { parts: [{ text: SYSTEM }] },
        generationConfig: { temperature: t, responseMimeType: "application/json" },
        contents: [{ parts: [{ text: buildUserPrompt(q, ctx) }] }] }) });
  if (!r.ok) throw new Error(prov.label + " " + r.status);
  const j = await r.json(); const txt = j.candidates?.[0]?.content?.parts?.[0]?.text || "{}";
  onTok && onTok(txt); return txt;
}

async function callAnthropicStyle(prov, q, ctx, onTok) {
  const { k, m, t } = cfg();
  abort = new AbortController();
  const r = await fetch(`${base(prov)}/v1/messages`,
    { method: "POST", signal: abort.signal,
      headers: { "Content-Type": "application/json", "x-api-key": k, "anthropic-version": "2023-06-01", "anthropic-dangerous-direct-browser-access": "true" },
      body: JSON.stringify({ model: m || prov.models?.[0], max_tokens: 2048, temperature: t,
        system: SYSTEM + " Balas HANYA JSON.", messages: [{ role: "user", content: buildUserPrompt(q, ctx) }] }) });
  if (!r.ok) throw new Error(prov.label + " " + r.status);
  const j = await r.json();
  const txt = (j.content || []).filter((b) => b.type === "text").map((b) => b.text).join("") || "{}";
  onTok && onTok(txt); return txt;
}

export async function ask(q, ctx, onTok) {
  const { p, k } = cfg();
  const prov = await getProvider(p);
  if (!prov.noKey && !k) throw new Error("Isi API key dulu di ⚙ (BYOK, tersimpan lokal).");
  const raw = prov.style === "gemini" ? await callGeminiStyle(prov, q, ctx, onTok)
    : prov.style === "anthropic" ? await callAnthropicStyle(prov, q, ctx, onTok)
    : await callOpenAIStyle(prov, q, ctx, onTok);
  try { const start = raw.indexOf("{"), end = raw.lastIndexOf("}");
    return normOps(JSON.parse(raw.slice(start, end + 1)));
  } catch (e) { return { ops: [], alasan_singkat: "Gagal parse JSON AI.", preview: raw.slice(0, 500) }; }
}

// Normalisasi toleran: AI kadang mengarang key (action/target/content) atau
// mengembalikan DATA mentah tanpa ops. Tanpa ini, op tak dikenal dilewati diam-diam.
export function normOps(res) {
  if (!res || typeof res !== "object") return { ops: [], alasan_singkat: "Respons AI kosong." };
  let ops = Array.isArray(res.ops) ? res.ops : [];
  ops = ops.map((o) => {
    if (!o || typeof o !== "object") return null;
    const op = { ...o };
    op.tool = op.tool || op.action || op.name || "";
    delete op.action; delete op.name;
    if (op.teks == null) op.teks = op.content ?? op.text ?? op.value ?? "";
    if (typeof op.teks !== "string") op.teks = String(op.teks ?? "");
    return op.tool ? op : null;
  }).filter(Boolean);
  if (!ops.length) { // heuristik bungkus-data: selamatkan respons berisi data, bukan ops
    if (res.nama || res.ringkasan || res.pengalaman) ops = [{ tool: "buat_cv", template: "cv-ats", data: res }];
    else if (Array.isArray(res.headers) && Array.isArray(res.rows)) ops = [{ tool: "buat_chart", chartType: "bar", title: res.title || "Chart", headers: res.headers, rows: res.rows }];
    else if (typeof res.mermaid === "string") ops = [{ tool: "buat_flowchart", mermaid: res.mermaid, caption: res.caption || "" }];
    else if (Array.isArray(res.nodes)) ops = [{ tool: "buat_flowchart_native", nodes: res.nodes, edges: res.edges || [], caption: res.caption || "" }];
    else if (Number.isInteger(res.rows) && Number.isInteger(res.cols)) ops = [{ tool: "buat_tabel", rows: res.rows, cols: res.cols, data: res.data }];
  }
  res.ops = ops; return res;
}

export async function testConn() { // payload terkecil per gaya
  const { p, k, m } = cfg();
  const prov = await getProvider(p);
  if (!prov.noKey && !k) throw new Error("Key kosong.");
  let r;
  if (prov.style === "gemini") {
    r = await fetch(`${base(prov)}/models/${m || prov.models?.[0]}:generateContent?key=${encodeURIComponent(k)}`,
      { method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ contents: [{ parts: [{ text: "ping" }] }] }) });
  } else if (prov.style === "anthropic") {
    r = await fetch(`${base(prov)}/v1/messages`,
      { method: "POST", headers: { "Content-Type": "application/json", "x-api-key": k, "anthropic-version": "2023-06-01", "anthropic-dangerous-direct-browser-access": "true" },
        body: JSON.stringify({ model: m || prov.models?.[0], max_tokens: 1, messages: [{ role: "user", content: "ping" }] }) });
  } else if (prov.noKey) { // tanpa key (Ollama/Pollinations): ping chat minimal, bukan GET /models
    r = await fetch(`${base(prov)}${prov.chatPath || "/chat/completions"}`,
      { method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ model: m || prov.models?.[0], messages: [{ role: "user", content: "ping" }], max_tokens: 1, stream: false }) });
  } else {
    r = await fetch(`${base(prov)}/models`, { headers: { Authorization: `Bearer ${k}` } });
  }
  if (!r.ok) throw new Error("Koneksi gagal (" + r.status + ")." + (r.status === 404 ? " Cek base URL provider." : "")); return true;
}

// Daftar model dari key user (browse available models). Cache 24 jam per provider.
const LS_MODELS = "docassist:models:";
const NON_CHAT = ["embed", "tts", "whisper", "dall-e", "moderation", "transcrib", "text-to-image", "audio"];
export function filterChatModels(ids) {
  return [...new Set(ids || [])].filter((id) => !NON_CHAT.some((b) => String(id).toLowerCase().includes(b)));
}
export async function listModels(force) {
  const { p, k } = cfg();
  const prov = await getProvider(p);
  if (!prov.noKey && !k) throw new Error("Isi key dulu.");
  const ck = LS_MODELS + prov.id;
  if (!force) {
    try { const c = JSON.parse(localStorage.getItem(ck) || "null");
      if (c && Date.now() - c.at < 24 * 3600e3) return { models: c.models, cached: true };
    } catch (e) {}
  }
  let ids = [];
  if (prov.style === "gemini") {
    const r = await fetch(`${base(prov)}/models?key=${encodeURIComponent(k)}`);
    if (!r.ok) throw new Error("Key ditolak (" + r.status + ").");
    const j = await r.json();
    ids = (j.models || []).filter((m) => (m.supportedGenerationMethods || []).includes("generateContent"))
      .map((m) => String(m.name || "").replace(/^models\//, ""));
  } else if (prov.style === "anthropic") {
    const r = await fetch(`${base(prov)}/v1/models`,
      { headers: { "x-api-key": k, "anthropic-version": "2023-06-01" } });
    if (!r.ok) throw new Error("Key ditolak (" + r.status + ").");
    ids = ((await r.json()).data || []).map((m) => m.id);
  } else { // openai-style: OpenAI, Groq, OpenRouter, HF, NVIDIA, Ollama, dsb.
    const headers = { "Content-Type": "application/json" };
    if (!prov.noKey) headers.Authorization = `Bearer ${k}`;
    const r = await fetch(`${base(prov)}/models`, { headers });
    if (!r.ok) throw new Error(r.status === 401 || r.status === 403 ? "Key ditolak (" + r.status + ")." : "Daftar model gagal (" + r.status + ").");
    ids = ((await r.json()).data || []).map((m) => m.id);
  }
  ids = filterChatModels(ids);
  try { localStorage.setItem(ck, JSON.stringify({ at: Date.now(), models: ids })); } catch (e) {}
  return { models: ids, cached: false };
}
