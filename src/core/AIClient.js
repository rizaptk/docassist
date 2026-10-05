/* AIClient — BYOK, lazy-loaded (diimport dinamis dari taskpane.js). Streaming + AbortController agar UI tetap responsif. */
import { SYSTEM, buildUserPrompt } from "./PromptPack.js";

const cfg = () => ({ p: localStorage.getItem("docassist:provider") || "openai",
  k: localStorage.getItem("docassist:key") || "", m: localStorage.getItem("docassist:model") || "gpt-4o-mini",
  t: parseFloat(localStorage.getItem("docassist:temp") || "0.3") });

let abort = null;
export function cancel() { abort && abort.abort(); }

async function streamOpenAI(q, ctx, onTok) {
  const { k, m, t } = cfg();
  abort = new AbortController();
  const r = await fetch("https://api.openai.com/v1/chat/completions", { method: "POST", signal: abort.signal,
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${k}` },
    body: JSON.stringify({ model: m, temperature: t, stream: true,
      messages: [{ role: "system", content: SYSTEM }, { role: "user", content: buildUserPrompt(q, ctx) }] }) });
  if (!r.ok) throw new Error("OpenAI " + r.status);
  const rd = r.body.getReader(), dec = new TextDecoder(); let buf = "", out = "";
  for (;;) { const { done, value } = await rd.read(); if (done) break;
    buf += dec.decode(value, { stream: true });
    for (const line of buf.split("\n")) { const s = line.trim(); if (!s.startsWith("data:")) continue;
      const d = s.slice(5).trim(); if (d === "[DONE]") break;
      try { const tok = JSON.parse(d).choices?.[0]?.delta?.content || ""; out += tok; onTok && onTok(tok); } catch (e) {} }
    buf = buf.slice(-2000); }
  return out;
}

async function callGemini(q, ctx, onTok) {
  const { k, m, t } = cfg();
  abort = new AbortController();
  const r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${m}:generateContent?key=${encodeURIComponent(k)}`,
    { method: "POST", signal: abort.signal, headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ systemInstruction: { parts: [{ text: SYSTEM }] },
        generationConfig: { temperature: t, responseMimeType: "application/json" },
        contents: [{ parts: [{ text: buildUserPrompt(q, ctx) }] }] }) });
  if (!r.ok) throw new Error("Gemini " + r.status);
  const j = await r.json(); const txt = j.candidates?.[0]?.content?.parts?.[0]?.text || "{}";
  onTok && onTok(txt); return txt;
}

export async function ask(q, ctx, onTok) {
  const { p, k } = cfg();
  if (!k) throw new Error("Isi API key dulu di ⚙ (BYOK, tersimpan lokal).");
  const raw = p === "gemini" ? await callGemini(q, ctx, onTok) : await streamOpenAI(q, ctx, onTok);
  try { const start = raw.indexOf("{"), end = raw.lastIndexOf("}");
    return JSON.parse(raw.slice(start, end + 1));
  } catch (e) { return { ops: [], alasan_singkat: "Gagal parse JSON AI.", preview: raw.slice(0, 500) }; }
}

export async function testConn() { // payload terkecil
  const { p, k, m } = cfg(); if (!k) throw new Error("Key kosong.");
  const url = p === "gemini"
    ? `https://generativelanguage.googleapis.com/v1beta/models/${m}:generateContent?key=${encodeURIComponent(k)}`
    : "https://api.openai.com/v1/models";
  const r = await fetch(url, p === "gemini"
    ? { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ contents: [{ parts: [{ text: "ping" }] }] }) }
    : { headers: { Authorization: `Bearer ${k}` } });
  if (!r.ok) throw new Error("Koneksi gagal (" + r.status + ")."); return true;
}
