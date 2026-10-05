/* Visuals — flowchart & chart → gambar → Word. Cross-platform: render jadi PNG,
   sisip via insertInlinePictureFromBase64 (API yang ada di semua platform).
   Mermaid lazy-load dari CDN hanya saat dibutuhkan. Chart digambar via canvas
   murni (tanpa dependensi) agar ringan. */

// ---------- 1. Parse data user (CSV / tab / "Label: 10") ----------
function parseNum(s) {
  s = String(s).trim().replace(/\s/g, "");
  if (/[.].*,/.test(s)) s = s.replace(/[.]/g, "").replace(",", "."); // 1.234,5 (id)
  else if (/[,]/.test(s) && !/[.]/.test(s)) s = s.replace(",", ".");
  const n = Number(s); return Number.isFinite(n) ? n : null;
}

export function parseDataTable(text) {
  const lines = String(text || "").split(/\r?\n/).map((l) => l.trim()).filter(Boolean).slice(0, 25);
  if (lines.length < 2) return null;
  // Gaya "Apel: 10" per baris.
  if (lines.every((l) => /^[^:;,\t|]+[:=]\s*-?[\d.,\s]+$/.test(l))) {
    const rows = [];
    for (const l of lines) { const m = l.match(/^([^:=]+)[:=]\s*(-?[\d.,\s]+)/); const n = parseNum(m[2]); if (n != null) rows.push([m[1].trim(), n]); }
    return rows.length ? { headers: ["", "Nilai"], rows } : null;
  }
  const d = ["\t", ";", "|", ","].find((x) => lines[0].includes(x)) || null;
  const cells = lines.map((l) => (d ? l.split(d) : l.split(/\s{2,}|\s/)).map((c) => c.trim()).filter((c) => c !== ""));
  if (cells.some((c) => c.length < 2)) return null;
  let headers = cells[0], start = 1;
  if (headers.slice(1).every((c) => parseNum(c) != null) && parseNum(headers[0]) == null) { start = 1; } // baris 1 = header
  else { headers = cells[0].map((_, i) => (i === 0 ? "" : `Seri ${i}`)); start = 0; } // tanpa header
  const rows = [];
  for (let i = start; i < cells.length; i++) {
    const nums = cells[i].slice(1).map(parseNum);
    if (nums.some((n) => n == null)) continue;
    rows.push([cells[i][0], ...nums]);
  }
  return rows.length ? { headers, rows } : null;
}

// ---------- 2. Chart renderer: canvas murni, tanpa dep ----------
const PALET = ["#2563eb", "#16a34a", "#f59e0b", "#ef4444", "#8b5cf6", "#06b6d4"];

export function renderChartPNG({ title = "Chart", type = "bar", headers = [], rows = [] }, W = 960, H = 540) {
  if (!rows.length) throw new Error("tidak ada data angka untuk chart");
  const c = document.createElement("canvas"); c.width = W; c.height = H;
  const x = c.getContext("2d");
  x.fillStyle = "#fff"; x.fillRect(0, 0, W, H);
  x.fillStyle = "#111"; x.font = "bold 26px system-ui,sans-serif"; x.textAlign = "center";
  x.fillText(String(title).slice(0, 60), W / 2, 44);
  const L = 80, R = 30, T = 70, B = 70, pw = W - L - R, ph = H - T - B;
  const labels = rows.map((r) => String(r[0]).slice(0, 14));
  const nSer = Math.max(headers.length - 1, 1);
  const vals = rows.map((r) => r.slice(1));
  const max = Math.max(...vals.flat(), 0) || 1;
  const nice = Math.pow(10, Math.floor(Math.log10(max)));
  const yMax = Math.ceil(max / nice) * nice;
  // grid + sumbu Y
  x.font = "16px system-ui,sans-serif"; x.fillStyle = "#6b7280"; x.textAlign = "right";
  for (let i = 0; i <= 4; i++) {
    const v = (yMax * i) / 4, y = T + ph - (ph * i) / 4;
    x.strokeStyle = "#e5e7eb"; x.beginPath(); x.moveTo(L, y); x.lineTo(W - R, y); x.stroke();
    x.fillText(String(Math.round(v)), L - 10, y + 6);
  }
  const gw = pw / labels.length;
  labels.forEach((lb, i) => {
    const cx = L + gw * i + gw / 2;
    if (type === "line") {
      for (let s = 0; s < nSer; s++) {
        x.strokeStyle = PALET[s % PALET.length]; x.lineWidth = 3; x.beginPath();
        labels.forEach((_, j) => { const px = L + gw * j + gw / 2, py = T + ph - (ph * vals[j][s]) / yMax; j ? x.lineTo(px, py) : x.moveTo(px, py); });
        x.stroke();
        labels.forEach((_, j) => { const px = L + gw * j + gw / 2, py = T + ph - (ph * vals[j][s]) / yMax;
          x.fillStyle = PALET[s % PALET.length]; x.beginPath(); x.arc(px, py, 5, 0, 7); x.fill(); });
      }
    } else {
      const bw = Math.min(60, (gw * 0.6) / nSer);
      for (let s = 0; s < nSer; s++) {
        const v = vals[i][s], bh = (ph * v) / yMax, bx = cx - (bw * nSer) / 2 + s * bw;
        x.fillStyle = PALET[s % PALET.length]; x.fillRect(bx, T + ph - bh, bw - 3, bh);
        x.fillStyle = "#111"; x.textAlign = "center"; x.fillText(String(v), bx + (bw - 3) / 2, T + ph - bh - 8);
      }
    }
    x.fillStyle = "#374151"; x.textAlign = "center"; x.fillText(lb, cx, H - 30);
  });
  // legenda
  if (nSer > 1) headers.slice(1).forEach((h, s) => {
    const lx = L + s * 150; x.fillStyle = PALET[s % PALET.length]; x.fillRect(lx, H - 14, 18, 12);
    x.fillStyle = "#374151"; x.textAlign = "left"; x.fillText(String(h).slice(0, 18), lx + 24, H - 4);
  });
  return c.toDataURL("image/png");
}

// ---------- 3. Mermaid flowchart → PNG (lazy CDN) ----------
// ponytail: CDN jauh > bundel mermaid 1MB+ ke hosting gratis; gagal load → fallback pesan jujur.
let mInit = null;
function mermaid() {
  if (!mInit) mInit = import("https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.esm.min.mjs").then((m) => {
    m.default.initialize({ startOnLoad: false, theme: "neutral" }); return m.default;
  });
  return mInit;
}

export async function renderMermaidPNG(code, outW = 1400) {
  const mm = await mermaid();
  const { svg } = await mm.render("da" + Date.now(), code);
  let sized = svg;
  if (!/width=/.test(sized)) sized = sized.replace("<svg", '<svg width="1200" height="800"');
  const url = URL.createObjectURL(new Blob([sized], { type: "image/svg+xml;charset=utf-8" }));
  try {
    const img = new Image();
    await new Promise((res, rej) => { img.onload = res; img.onerror = () => rej(new Error("gagal raster diagram")); img.src = url; });
    const k = outW / (img.naturalWidth || outW);
    const c = document.createElement("canvas");
    c.width = outW; c.height = Math.round((img.naturalHeight || 800) * k);
    const x = c.getContext("2d");
    x.fillStyle = "#fff"; x.fillRect(0, 0, c.width, c.height);
    x.drawImage(img, 0, 0, c.width, c.height);
    return c.toDataURL("image/png");
  } finally { URL.revokeObjectURL(url); }
}

// demo mandiri: parseDataTable("Apel: 10\nJeruk: 20") → 2 baris.
