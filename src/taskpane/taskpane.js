/* taskpane.js — bootstrap ringan: render instan, lazy-load AI/Word, pause saat hidden (hemat CPU/Word). */
import { buildContext } from "../core/ContextBuilder.js";
import { loadManifest, ensureManifest, saveManifestLocal, cheapHash, isStale } from "../core/ManifestManager.js";
import { eksekusiOps, bacaMeta, bacaJendela, auditStruktur } from "../core/WordOps.js";

const $ = (id) => document.getElementById(id);
const msgs = $("msgs"), badge = $("cacheBadge");
let manifest = null, pendingOps = null, busy = false, aiMod = null, visMod = null;
const lazyAI = () => aiMod || import("../core/AIClient.js").then((m) => (aiMod = m));
const lazyVis = () => visMod || import("../core/Visuals.js").then((m) => (visMod = m));

// Render: cap 50 bubble terakhir (hemat DOM), throttle streaming via rAF.
let rafQ = "", rafOn = false;
function addMsg(cls, txt) {
  const d = document.createElement("div"); d.className = "msg " + cls; d.textContent = txt;
  msgs.appendChild(d); while (msgs.children.length > 50) msgs.firstChild.remove();
  msgs.scrollTop = msgs.scrollHeight; return d;
}
function streamTo(el) { return (tok) => { rafQ += tok; if (rafOn) return; rafOn = true;
  requestAnimationFrame(() => { el.textContent += rafQ; rafQ = ""; rafOn = false; msgs.scrollTop = msgs.scrollHeight; }); }; }

// Collapse non-intrusif + sembunyi total.
$("btnCollapse").onclick = () => $("app").classList.add("collapsed");
$("btnExpand").onclick = () => $("app").classList.remove("collapsed");
$("btnHide").onclick = async () => { try { await Office.addin.hide(); } catch (e) { $("app").classList.add("collapsed"); } };
function openSettings() { const s = $("settings"); s.classList.add("open"); s.hidden = false; const b = $("backdrop"); b.classList.add("open"); b.hidden = false; }
function closeSettings() { const s = $("settings"); s.classList.remove("open"); s.hidden = true; const b = $("backdrop"); b.classList.remove("open"); b.hidden = true; }
$("btnSettings").onclick = openSettings;
$("btnCloseSettings").onclick = closeSettings;
$("backdrop").onclick = closeSettings;
document.addEventListener("keydown", (e) => { if (e.key === "Escape") { closeSettings(); hideSlash(); } });

// Slash commands ala Hermes/opencode: "/" → autocomplete, "/tools" → daftar klik di chat.
const SLASH = [
  { cmd: "rapikan", ic: "✨", desc: "Rapikan paragraf, pertahankan istilah", template: "Rapikan paragraf ini, pertahankan istilah teknis." },
  { cmd: "lanjutkan", ic: "📝", desc: "Lanjutkan tulisan dengan nada sama", template: "Lanjutkan 2 paragraf dengan nada yang sama." },
  { cmd: "tabel", ic: "▦", desc: "Buatkan tabel rapi", template: "Buatkan tabel 3x3 rapi dari teks ini." },
  { cmd: "formal", ic: "🎓", desc: "Formal-kan tanpa ubah makna", template: "Formal-kan tanpa ubah makna." },
  { cmd: "flowchart", ic: "🔀", desc: "Flowchart editable dari alur", template: "Buatkan flowchart yang bisa diedit dari alur berikut: " },
  { cmd: "chart", ic: "📊", desc: "Chart dari data paste", template: "Buatkan chart dari data berikut (paste tabel/angka): " },
  { cmd: "cv", ic: "📄", desc: "CV ATS-friendly dari data", template: "Buatkan CV ATS-friendly dari data berikut (nama, kontak, pengalaman, pendidikan, skill): " },
  { cmd: "cek-ats", ic: "✅", desc: "Audit kelolosan ATS dokumen", template: "Cek apakah dokumen ini lolos ATS, beri saran perbaikan." },
  { cmd: "tools", ic: "🧰", desc: "Tampilkan daftar tools", template: "" },
  { cmd: "help", ic: "❓", desc: "Bantuan slash command", template: "" },
  { cmd: "clear", ic: "🧹", desc: "Bersihkan chat", template: "" },
];
let slashIdx = -1;
function slashMatches() {
  const v = $("input").value;
  if (!v.startsWith("/") || /\s/.test(v)) return null;
  const q = v.slice(1).toLowerCase();
  return SLASH.filter((s) => s.cmd.startsWith(q));
}
function renderSlash() {
  const box = $("slashBox"), m = slashMatches();
  if (!m) { box.hidden = true; slashIdx = -1; return; }
  box.innerHTML = "";
  m.forEach((s, i) => {
    const b = document.createElement("button");
    b.className = "slash-item" + (i === slashIdx ? " active" : ""); b.setAttribute("role", "option");
    b.innerHTML = "";
    const ic = document.createElement("span"); ic.className = "ic"; ic.textContent = s.ic;
    const cd = document.createElement("code"); cd.textContent = "/" + s.cmd;
    const ds = document.createElement("span"); ds.className = "ds"; ds.textContent = s.desc;
    b.append(ic, cd, ds);
    b.onclick = () => pickSlash(i);
    box.appendChild(b);
  });
  box.hidden = false;
}
function hideSlash() { $("slashBox").hidden = true; slashIdx = -1; autoGrow(); }
function pickSlash(i) {
  const m = slashMatches(); if (!m || !m[i]) return;
  const s = m[i];
  if (s.cmd === "tools") { hideSlash(); $("input").value = ""; renderToolsList(); $("input").focus(); return; }
  if (s.cmd === "help") { hideSlash(); $("input").value = ""; showSlashHelp(); $("input").focus(); return; }
  if (s.cmd === "clear") { hideSlash(); $("input").value = ""; msgs.innerHTML = ""; return; }
  $("input").value = "/" + s.cmd + " ";
  if (s.template && !s.template.endsWith(" ")) { $("input").value = s.template + " "; }
  else if (s.template) { $("input").value = s.template; }
  hideSlash(); $("input").focus();
}
function renderToolsList() {
  const d = addMsg("a", "Tools tersedia — klik untuk mengisi prompt:");
  const w = document.createElement("div"); w.className = "tool-list";
  for (const s of SLASH.filter((x) => x.template)) {
    const b = document.createElement("button"); b.className = "tool-pick"; b.textContent = `${s.ic} /${s.cmd}`;
    b.title = s.desc; b.dataset.tpl = s.template; w.appendChild(b);
  }
  d.appendChild(w); msgs.scrollTop = msgs.scrollHeight;
}
function showSlashHelp() {
  addMsg("a", "Ketik / untuk autocomplete. Contoh: /cv Nama… + data, lalu Enter. Konteks dokumen aktif otomatis terbaca. /tools = daftar klik, /clear = bersihkan chat.");
}
msgs.addEventListener("click", (e) => {
  const b = e.target.closest && e.target.closest(".tool-pick");
  if (!b) return;
  $("input").value = b.dataset.tpl + " "; hideSlash(); $("input").focus();
});
$("input").addEventListener("input", () => { slashIdx = -1; renderSlash(); autoGrow(); });
function autoGrow() { const t = $("input"); t.style.height = "auto"; t.style.height = Math.min(t.scrollHeight, 76) + "px"; }
$("input").addEventListener("keydown", (e) => {
  const box = $("slashBox");
  if (!box.hidden) {
    const n = box.children.length;
    if (e.key === "ArrowDown") { e.preventDefault(); slashIdx = (slashIdx + 1) % n; renderSlash(); return; }
    else if (e.key === "ArrowUp") { e.preventDefault(); slashIdx = (slashIdx - 1 + n) % n; renderSlash(); return; }
    else if ((e.key === "Enter" || e.key === "Tab") && slashIdx >= 0) { e.preventDefault(); pickSlash(slashIdx); return; }
    else if (e.key === "Escape") { hideSlash(); return; }
  }
  if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); $("form").requestSubmit(); }
});
function expandSlash(raw) {
  if (!raw.startsWith("/")) return { q: raw };
  const m = raw.match(/^\/(\S+)\s*([\s\S]*)$/);
  const cmd = (m[1] || "").toLowerCase(), rest = (m[2] || "").trim();
  if (cmd === "tools") { renderToolsList(); return { handled: true }; }
  if (cmd === "help") { showSlashHelp(); return { handled: true }; }
  if (cmd === "clear") { msgs.innerHTML = ""; return { handled: true }; }
  const hit = SLASH.find((s) => s.cmd === cmd);
  if (hit && hit.template) return { q: (hit.template + " " + rest).trim() };
  return { q: raw }; // slash tak dikenal → kirim mentah ke AI
}

// Settings BYOK + registry provider dinamis (config/providers.json + custom lokal).
const S = (k, v) => v === undefined ? localStorage.getItem(k) : localStorage.setItem(k, v);
let provList = [];
async function refreshProviders() {
  const ai = await lazyAI();
  provList = await ai.getProviders();
  const sel = $("sProvider"), cur = S("docassist:provider") || "openai";
  sel.innerHTML = "";
  for (const p of provList) { const o = document.createElement("option"); o.value = p.id; o.textContent = p.label + (p.custom ? " ✎" : ""); sel.appendChild(o); }
  sel.value = provList.some((p) => p.id === cur) ? cur : provList[0].id;
  paintProvider();
}
function paintProvider() {
  const p = provList.find((x) => x.id === $("sProvider").value) || provList[0];
  if (!p) return;
  const dl = $("modelList"); dl.innerHTML = "";
  for (const m of p.models || []) { const o = document.createElement("option"); o.value = m; dl.appendChild(o); }
  if (!$("sModel").value) $("sModel").value = (p.models || [""])[0] || "";
  $("sKey").placeholder = p.keyPlaceholder || "tempel key…";
  $("keyRow").style.display = p.noKey ? "none" : "";
  const h = $("sProvHint");
  const parts = [];
  if (p.keyUrl && !p.noKey) parts.push(`<a href="${p.keyUrl}" target="_blank">ambil key →</a>`);
  if (p.notes) parts.push(p.notes);
  h.innerHTML = parts.join(" · "); h.hidden = !parts.length;
  refreshModelList(); // gabung daftar bawaan + hasil browse key user
}

// Browse model dari key user → isi datalist. Dijaga dari balapan + gagal diam-diam.
let mdlSeq = 0;
async function refreshModelList(force) {
  const my = ++mdlSeq;
  try {
    const pid = $("sProvider").value;
    const prov = provList.find((x) => x.id === pid);
    if (!prov) return;
    if (!prov.noKey && !S("docassist:key")) return; // belum ada key → jangan tembak 401
    const ai = await lazyAI();
    const { models, cached } = await ai.listModels(force);
    if (my !== mdlSeq) return;
    const seen = new Set(), dl = $("modelList"); dl.innerHTML = "";
    for (const m of [...(prov.models || []), ...models]) {
      if (!m || seen.has(m)) continue; seen.add(m);
      const o = document.createElement("option"); o.value = m; dl.appendChild(o);
    }
    const h = $("sProvHint");
    h.innerHTML += (h.innerHTML ? " · " : "") + `${models.length} model dari key${cached ? " (cache)" : ""} — klik kolom Model`;
    h.hidden = false;
    if (!$("sModel").value && prov.models?.[0]) $("sModel").value = prov.models[0];
  } catch (e) {
    if (/401|403|ditolak/i.test(e.message || "")) { const h = $("sProvHint"); h.textContent = "Key ditolak — periksa kembali."; h.hidden = false; }
    // gagal lain (CORS/endpoint): diam, daftar bawaan tetap ada
  }
}
function loadSettings() { // sinkron dulu (cat cepat), registry menyusul async
  $("sKey").value = S("docassist:key") || ""; $("sModel").value = S("docassist:model") || ""; $("sTemp").value = S("docassist:temp") || "0.3";
  refreshProviders().catch(() => {});
}
$("sProvider").onchange = () => { S("docassist:provider", $("sProvider").value); $("sModel").value = ""; paintProvider(); };
$("sKey").onchange = () => { S("docassist:key", $("sKey").value); refreshModelList(); };
$("sModel").onfocus = (e) => e.target.select(); // klik → pilih semua, dropdown tampil penuh
$("sModel").onchange = () => S("docassist:model", $("sModel").value);
$("sTemp").onchange = () => S("docassist:temp", $("sTemp").value);
$("btnAddProv").onclick = async () => {
  const name = $("cName").value.trim(), base = $("cBase").value.trim().replace(/\/+$/, "");
  if (!name || !/^https?:\/\/.+\..+/.test(base)) { addMsg("sys", "Nama + base URL https tidak valid."); return; }
  const id = "custom-" + name.toLowerCase().replace(/[^a-z0-9]+/g, "-").slice(0, 24);
  (await lazyAI()).saveCustomProvider({ id, label: name, style: $("cStyle").value, baseUrl: base,
    models: [$("cModel").value.trim() || "default"], custom: true });
  S("docassist:provider", id); $("sModel").value = ""; $("cName").value = $("cBase").value = $("cModel").value = "";
  await refreshProviders(); addMsg("sys", "Provider '" + name + "' tersimpan di perangkat ini.");
};
$("btnDelProv").onclick = async () => {
  const id = $("sProvider").value, ai = await lazyAI();
  if (!provList.find((p) => p.id === id)?.custom) { addMsg("sys", "Hanya provider custom (✎) yang bisa dihapus."); return; }
  ai.deleteCustomProvider(id); S("docassist:provider", "openai"); $("sModel").value = "";
  await refreshProviders(); addMsg("sys", "Provider custom dihapus.");
};
$("btnWipe").onclick = () => { localStorage.removeItem("docassist:key"); $("sKey").value = ""; };
$("btnTest").onclick = async () => { try { await (await lazyAI()).testConn(); addMsg("sys", "Koneksi OK."); refreshModelList(true); } catch (e) { addMsg("sys", String(e.message)); } };
$("btnRefreshCtx").onclick = () => refreshCtx(true);

// Badge cache.
function paintBadge(stale) { badge.textContent = !manifest ? "Konteks: baru" : stale ? "Konteks: stale [Refresh]" : "Konteks: cached ✓"; }
async function refreshCtx(force) {
  const [meta, win] = await Promise.all([bacaMeta(), bacaJendela(2000)]);
  const h = cheapHash(meta, (win.konteks || "").length);
  manifest = (await loadManifest()) || ensureManifest(h);
  const stale = force ? true : isStale(manifest.docHash, h);
  if (force || !manifest.docHash) { manifest.docHash = h; manifest.outline = meta.headings || manifest.outline; saveManifestLocal(manifest); }
  paintBadge(stale);
}

// Submit: 1 Word.run (konteks) + 1 fetch (AI) + preview sebelum tulis.
$("form").onsubmit = async (e) => {
  e.preventDefault(); if (busy) return;
  const q = $("input").value.trim(); if (!q) return;
  $("input").value = ""; hideSlash(); addMsg("u", q); busy = true;
  const ex = expandSlash(q); if (ex.handled) { busy = false; return; } // /tools /help /clear
  const q2 = ex.q; // slash command → template + konteks user
  const thinking = addMsg("a", "Menyusun operasi…");
  try {
    await refreshCtx(false);
    const ctx = await buildContext(manifest);
    const ai = await lazyAI();
    const res = await ai.ask(q2, ctx, null); // tanpa streaming mentah: output AI selalu JSON mesin
    thinking.textContent = res.alasan_singkat || "Siap.";
    if (res.preview) { $("diffBefore").textContent = (ctx.l1.seleksi || "(kosong)").slice(0, 600); $("diffAfter").textContent = res.preview.slice(0, 600); $("diffBox").hidden = false; }
    pendingOps = res.ops || [];
    // Audit ATS: jalan duluan, hasilnya bubble chat (bukan diff), lalu dibuang dari antrean.
    if (pendingOps.some((o) => o.tool === "cek_ats")) {
      thinking.textContent = "Memeriksa struktur dokumen…";
      try {
        const B = await import("../core/Builders/CvBuilder.js");
        const v = B.auditVerdict(await auditStruktur());
        addMsg("a", (v.lolos ? "✅ " : "⚠️ ") + v.pesan);
      } catch (err) { addMsg("sys", "Audit gagal: " + err.message); }
      pendingOps = pendingOps.filter((o) => o.tool !== "cek_ats");
      if (!pendingOps.length && !res.preview) { thinking.textContent = res.alasan_singkat || "Audit selesai."; busy = false; return; }
    }
    // Visual (flowchart/chart): render dulu untuk pratinjau, sisip saat Terapkan.
    // Flowchart native: pratinjau = render mermaid dari nodes/edges; yang disisip = shapes editable.
    const vis = pendingOps.filter((o) => o.tool === "buat_flowchart" || o.tool === "buat_chart" || o.tool === "buat_flowchart_native");
    if (vis.length) {
      thinking.textContent = "Merender visual…";
      const V = await lazyVis();
      for (const op of vis) {
        try {
          if (op.tool === "buat_flowchart") {
            op._png = (await V.renderMermaidPNG(op.mermaid || "graph TD;A-->B")).split(",")[1];
            op.caption = op.caption || "Flowchart";
          } else if (op.tool === "buat_flowchart_native") {
            const F = await import("../core/FlowchartOoxml.js");
            const flow = { nodes: op.nodes || [], edges: op.edges || [] };
            op._ooxml = F.buildFlowchartOoxml(flow); // siap sisip saat Terapkan
            op._png = (await V.renderMermaidPNG(F.flowToMermaid(flow))).split(",")[1]; // pratinjau
            op.caption = op.caption || "Flowchart (editable di Word)";
          } else {
            const t = V.parseDataTable(op.data || ctx.l1.seleksi || "");
            const png = V.renderChartPNG({ title: op.title || "Chart", type: op.chartType || "bar",
              headers: op.headers || t?.headers || ["", "Nilai"], rows: op.rows || t?.rows || [] });
            op._png = png.split(",")[1]; op.caption = op.title || "Chart";
          }
        } catch (err) { addMsg("sys", "Visual gagal: " + err.message + " — dilewati."); op._png = null; }
      }
      const first = vis.find((v) => v._png);
      if (first) {
        const img = $("diffImg"); img.src = "data:image/png;base64," + first._png; img.hidden = false;
        $("diffAfter").textContent = (first.caption || "visual") + " — cek gambar di atas, Terapkan untuk sisip.";
        $("diffBox").hidden = false;
      }
      thinking.textContent = res.alasan_singkat || "Visual siap dipratinjau.";
    }
    if (!pendingOps.length && res.preview) pendingOps = [{ tool: "tulis_ganti", teks: res.preview }];
    if (pendingOps.length && !res.preview) {
      const n = await eksekusiOps(pendingOps); pendingOps = null;
      addMsg("sys", n ? "Diterapkan langsung (aksi kecil). Ctrl+Z untuk batal." : "Tidak ada aksi valid yang bisa diterapkan.");
    }
  } catch (err) { thinking.textContent = "Gagal: " + err.message; }
  busy = false;
};
$("btnApply").onclick = async () => {
  const n = pendingOps ? await eksekusiOps(pendingOps) : 0;
  pendingOps = null; $("diffBox").hidden = true; $("diffImg").hidden = true;
  addMsg("sys", n ? `Diterapkan (${n} aksi). Ctrl+Z untuk batal.` : "Tidak ada aksi valid yang diterapkan.");
};
$("btnDiscard").onclick = () => { pendingOps = null; $("diffBox").hidden = true; $("diffImg").hidden = true; };

// Pause background saat pane hidden (jangan ganggu Word).
let idleT = null;
function armIdle() { clearInterval(idleT); idleT = setInterval(() => { if (!$("app").classList.contains("collapsed")) refreshCtx(false); }, 15000); }

// Theme-aware: ikuti tema Word (OfficeTheme) bila ada, fallback ke system theme.
function normHex(c) { if (!c) return null; c = String(c).trim(); if (/^[0-9a-fA-F]{6}$/.test(c)) return "#" + c; return /^#[0-9a-fA-F]{6}$/.test(c) ? c : null; }
function lum(hex) { const n = parseInt(hex.slice(1), 16); return (0.299 * ((n >> 16) & 255) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255)) / 255; }
function applyTheme() {
  const root = document.documentElement;
  let bg = null, fg = null;
  try { const t = Office.context.officeTheme || {}; bg = normHex(t.bodyBackgroundColor); fg = normHex(t.bodyForegroundColor); } catch (e) {}
  if (!bg && window.matchMedia && matchMedia("(prefers-color-scheme: dark)").matches) { bg = "#1e1e1e"; fg = "#e8e8e8"; }
  if (bg) {
    root.style.setProperty("--bg", bg);
    const dark = lum(bg) < 0.5;
    document.getElementById("app").dataset.theme = dark ? "dark" : "light";
    root.style.setProperty("--fg", fg || (dark ? "#e8e8e8" : "#1b1b1b"));
  }
  try { if (window.matchMedia) matchMedia("(prefers-color-scheme: dark)").addEventListener("change", applyTheme); } catch (e) {}
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) { addMsg("sys", "Buka di Word untuk memakai DocAssist."); return; }
  applyTheme(); loadSettings(); addMsg("a", "Halo! Saya asisten operasional dokumen. Ketik / untuk melihat perintah.");
  try { Office.addin.onVisibilityModeChanged(async (args) => { if (args.visibilityMode === "hidden") clearInterval(idleT); else armIdle(); }); } catch (e) {}
  // Render dulu, kerja berat di idle (TTI cepat).
  const boot = () => { refreshCtx(false); armIdle(); };
  if ("requestIdleCallback" in window) requestIdleCallback(boot, { timeout: 1500 }); else setTimeout(boot, 600);
});
