/* Smoke test logika murni (tanpa Word/DOM): layout, escaping, OOXML, guard.
   Jalankan: node scripts/smoke.mjs */
import { strict as a } from "node:assert";
import { layoutFlow, buildFlowchartOoxml, flowToMermaid } from "../src/core/FlowchartOoxml.js";
import { parseDataTable } from "../src/core/Visuals.js";
import { readFileSync } from "node:fs";

const lin = { nodes: [{ id: "a", label: "Mulai", shape: "start" }, { id: "b", label: "Proses A&B <x>" }, { id: "c", label: "Selesai", shape: "end" }], edges: [{ from: "a", to: "b" }, { from: "b", to: "c", label: "Ya" }] };
const L = layoutFlow(lin);
a.equal(L.nodes[0].rank, 0); a.equal(L.nodes[2].rank, 2);
a.ok(L.nodes[2].y > L.nodes[0].y, "rank bawah = y lebih besar");

const br = { nodes: [{ id: "a", label: "A" }, { id: "b", label: "B" }, { id: "c", label: "C" }], edges: [{ from: "a", to: "b" }, { from: "a", to: "c" }] };
const B = layoutFlow(br);
a.equal(B.at.b.rank, B.at.c.rank, "cabang se-rank");
a.notEqual(B.at.b.x, B.at.c.x, "cabang beda x");

const cyc = { nodes: [{ id: "a", label: "A" }, { id: "b", label: "B" }], edges: [{ from: "a", to: "b" }, { from: "b", to: "a" }] };
layoutFlow(cyc); // tidak hang

const xml = buildFlowchartOoxml(lin, 123456);
a.ok(xml.includes("wpg:wgp"), "ada grup");
a.equal((xml.match(/<wps:wsp>/g) || []).length, 4, "3 node + 1 label edge");
a.equal((xml.match(/<wps:cxnSp>/g) || []).length, 2, "2 panah");
a.ok(xml.includes("Proses A&amp;B &lt;x&gt;"), "escaping XML");
a.ok(xml.includes('prst="diamond"') === false, "tanpa decision = tanpa diamond");
a.ok(xml.includes('prst="ellipse"'), "start/end = ellipse");
a.ok(xml.includes('tailEnd type="triangle"'), "panah arrowhead");
a.ok(xml.includes('id="123456"'), "docPr id terpakai");

const dec = buildFlowchartOoxml({ nodes: [{ id: "d", label: "Setuju?", shape: "decision" }], edges: [] });
a.ok(dec.includes('prst="diamond"'), "decision = diamond");

const m = flowToMermaid(lin);
a.ok(m.includes("-->") && m.includes('|"Ya"|'), "mermaid edge + label");

a.throws(() => layoutFlow({ nodes: [], edges: [] }), "node kosong ditolak");
a.throws(() => layoutFlow({ nodes: Array.from({ length: 25 }, (_, i) => ({ id: "n" + i, label: "x" })), edges: [] }), "di atas 20 ditolak");

const t = parseDataTable("Apel: 10\nJeruk: 20");
a.equal(t.rows.length, 2, "parse gaya Label: angka");

console.log("smoke OK: layout, escaping, ooxml, guard, parse — semua lolos");

// Registry provider: preset file valid + custom bisa ditambah/dihapus (mock localStorage).
const reg = JSON.parse(readFileSync("config/providers.json", "utf8"));
a.ok(reg.providers.length >= 2, "registry punya preset");
for (const p of reg.providers) a.ok(p.id && p.label && p.style && p.baseUrl, "provider lengkap: " + p.id);
a.ok(["openai", "gemini", "anthropic"].includes(reg.providers[0].style) || ["openai", "gemini", "anthropic"].includes(reg.providers.find((p) => p.id === "openai").style), "style dikenal");
for (const p of reg.providers) a.ok(["openai", "gemini", "anthropic"].includes(p.style), "style valid: " + p.id);
a.ok(reg.providers.some((p) => p.noKey), "ada opsi tanpa key");
const store = {};
globalThis.localStorage = { getItem: (k) => store[k] ?? null, setItem: (k, v) => { store[k] = String(v); }, removeItem: (k) => { delete store[k]; } };
globalThis.fetch = async () => { throw new Error("offline"); }; // paksa jalur fallback + custom
const AI = await import("../src/core/AIClient.js");
AI.saveCustomProvider({ id: "custom-x", label: "X", style: "openai", baseUrl: "https://x.example/v1", models: ["m1"], custom: true });
const all = await AI.getProviders();
a.ok(all.some((p) => p.id === "custom-x"), "custom tergabung");
a.ok(all.some((p) => p.id === "openai"), "fallback preset saat offline");
AI.deleteCustomProvider("custom-x");
a.ok(!(await AI.getProviders()).some((p) => p.id === "custom-x"), "custom terhapus");
console.log("smoke OK: registry provider — semua lolos");

// listModels: normalisasi + filter non-chat + cache 24 jam (mock fetch).
a.deepEqual(AI.filterChatModels(["chat-a", "text-embedding-x", "whisper-1", "chat-a"]), ["chat-a"]);
let calls = 0;
globalThis.fetch = async (url) => { calls++;
  if (String(url).includes("/models?key="))
    return { ok: true, json: async () => ({ models: [{ name: "models/gemini-x", supportedGenerationMethods: ["generateContent"] }, { name: "models/embed-y", supportedGenerationMethods: ["embedContent"] }] }) };
  return { ok: true, json: async () => ({ data: [{ id: "chat-a" }, { id: "text-embedding-x" }] }) };
};
store["docassist:provider"] = "openai"; store["docassist:key"] = "k";
const r1 = await AI.listModels();
a.deepEqual(r1.models, ["chat-a"], "non-chat terfilter");
a.equal(calls, 1);
const r2 = await AI.listModels(); // cache: tanpa fetch ulang
a.equal(r2.cached, true); a.equal(calls, 1);
store["docassist:provider"] = "gemini";
const r3 = await AI.listModels();
a.deepEqual(r3.models, ["gemini-x"], "gemini filter generateContent");
console.log("smoke OK: browse model dari key — semua lolos");
