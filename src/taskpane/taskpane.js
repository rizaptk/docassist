/* taskpane.js — bootstrap ringan: render instan, lazy-load AI/Word, pause saat hidden (hemat CPU/Word). */
import { buildContext } from "../core/ContextBuilder.js";
import { loadManifest, ensureManifest, saveManifestLocal, cheapHash, isStale } from "../core/ManifestManager.js";
import { eksekusiOps, bacaMeta, bacaJendela } from "../core/WordOps.js";

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
$("btnSettings").onclick = () => { const s = $("settings"); s.hidden = !s.hidden; };
document.querySelectorAll("#chips button").forEach((b) => (b.onclick = () => {
  if (b.dataset.prefill) { $("input").value = b.dataset.prefill; $("input").focus(); } // visual: user lanjutkan ketik/paste
  else { $("input").value = b.dataset.q; $("form").requestSubmit(); }
}));

// Settings BYOK (sinkron lokal, tanpa network).
const S = (k, v) => v === undefined ? localStorage.getItem(k) : localStorage.setItem(k, v);
function loadSettings() { $("sProvider").value = S("docassist:provider") || "openai"; $("sKey").value = S("docassist:key") || "";
  $("sModel").value = S("docassist:model") || "gpt-4o-mini"; $("sTemp").value = S("docassist:temp") || "0.3"; }
["sProvider", "sKey", "sModel", "sTemp"].forEach((id) => $(id).onchange = () =>
  { S("docassist:provider", $("sProvider").value); S("docassist:key", $("sKey").value);
    S("docassist:model", $("sModel").value); S("docassist:temp", $("sTemp").value); });
$("btnWipe").onclick = () => { localStorage.removeItem("docassist:key"); $("sKey").value = ""; };
$("btnTest").onclick = async () => { try { await (await lazyAI()).testConn(); addMsg("sys", "Koneksi OK."); } catch (e) { addMsg("sys", String(e.message)); } };
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
  $("input").value = ""; addMsg("u", q); busy = true;
  const thinking = addMsg("a", "…");
  try {
    await refreshCtx(false);
    const ctx = await buildContext(manifest);
    const ai = await lazyAI();
    const res = await ai.ask(q, ctx, streamTo(thinking));
    thinking.textContent = res.alasan_singkat || "Siap.";
    if (res.preview) { $("diffBefore").textContent = (ctx.l1.seleksi || "(kosong)").slice(0, 600); $("diffAfter").textContent = res.preview.slice(0, 600); $("diffBox").hidden = false; }
    pendingOps = res.ops || [];
    // Visual (flowchart/chart): render dulu untuk pratinjau, sisip saat Terapkan.
    const vis = pendingOps.filter((o) => o.tool === "buat_flowchart" || o.tool === "buat_chart");
    if (vis.length) {
      thinking.textContent = "Merender visual…";
      const V = await lazyVis();
      for (const op of vis) {
        try {
          if (op.tool === "buat_flowchart") {
            op._png = (await V.renderMermaidPNG(op.mermaid || "graph TD;A-->B")).split(",")[1];
            op.caption = op.caption || "Flowchart";
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
    if (pendingOps.length && !res.preview) { await eksekusiOps(pendingOps); pendingOps = null; addMsg("sys", "Diterapkan langsung (aksi kecil)."); }
  } catch (err) { thinking.textContent = "Gagal: " + err.message; }
  busy = false;
};
$("btnApply").onclick = async () => { if (pendingOps) await eksekusiOps(pendingOps); pendingOps = null; $("diffBox").hidden = true; $("diffImg").hidden = true; addMsg("sys", "Diterapkan. Ctrl+Z untuk batal."); };
$("btnDiscard").onclick = () => { pendingOps = null; $("diffBox").hidden = true; $("diffImg").hidden = true; };

// Pause background saat pane hidden (jangan ganggu Word).
let idleT = null;
function armIdle() { clearInterval(idleT); idleT = setInterval(() => { if (!$("app").classList.contains("collapsed")) refreshCtx(false); }, 15000); }

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) { addMsg("sys", "Buka di Word untuk memakai DocAssist."); return; }
  loadSettings(); addMsg("a", "Halo! Saya asisten operasional dokumen. Pilih chip atau ketik perintah.");
  try { Office.addin.onVisibilityModeChanged(async (args) => { if (args.visibilityMode === "hidden") clearInterval(idleT); else armIdle(); }); } catch (e) {}
  // Render dulu, kerja berat di idle (TTI cepat).
  const boot = () => { refreshCtx(false); armIdle(); };
  if ("requestIdleCallback" in window) requestIdleCallback(boot, { timeout: 1500 }); else setTimeout(boot, 600);
});
