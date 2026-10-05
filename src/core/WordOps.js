/* WordOps — 15 tools atomik. Prinsip perf: 1 Word.run per aksi, load properti minimal, walk ±3 paragraf (tanpa load seluruh body). */
async function run(fn) { return Word.run(fn); }
async function tryCatch(fn, onErr) { try { return await fn(); } catch (e) { onErr && onErr(e); console.error("[DocAssist]", e); return null; } }

export async function bacaSeleksi() {
  return tryCatch(() => run(async (ctx) => {
    const r = ctx.document.getSelection(); r.load("text"); await ctx.sync(); return r.text || "";
  }));
}

// Baca jendela kerja: seleksi + ±3 paragraf via getPrevious/getNext (murah untuk doc besar).
export async function bacaJendela(maxChars = 4000) {
  return tryCatch(() => run(async (ctx) => {
    const sel = ctx.document.getSelection(); sel.load("text"); await ctx.sync();
    const selText = sel.text || "";
    let cur = sel.paragraphs.getFirstOrNullObject(); cur.load("text"); await ctx.sync();
    if (cur.isNullObject) return { seleksi: selText, konteks: selText.slice(0, maxChars) };
    const parts = [];
    let p = cur;
    for (let i = 0; i < 3; i++) { const q = p.getPreviousOrNullObject(); q.load("text"); await ctx.sync(); if (q.isNullObject) break; parts.unshift(q.text); p = cur; cur = q; }
    cur = sel.paragraphs.getFirstOrNullObject(); cur.load("text"); await ctx.sync();
    let f = cur;
    for (let i = 0; i < 3; i++) { const q = f.getNextOrNullObject(); q.load("text"); await ctx.sync(); if (q.isNullObject) break; parts.push(q.text); f = q; }
    const konteks = [...parts.slice(0, 3), selText, ...parts.slice(3)].join("\n").slice(0, maxChars);
    return { seleksi: selText, konteks };
  }));
}

// Meta murah: count + heading pertama (chunk 50, stop cepat).
export async function bacaMeta() {
  return tryCatch(() => run(async (ctx) => {
    const paras = ctx.document.body.paragraphs; paras.load("items"); await ctx.sync();
    const n = paras.items.length;
    const heads = [];
    const cap = Math.min(n, 60); // ponytail: 60 pertama cukup untuk outline cepat; full scan hanya saat rebuild
    for (let i = 0; i < cap; i++) { paras.items[i].load("text,style"); }
    await ctx.sync();
    for (let i = 0; i < cap; i++) { const p = paras.items[i]; if (/heading/i.test(p.style || "")) heads.push(p.text.trim().slice(0, 80)); if (heads.length >= 20) break; }
    return { paragraf: n, headings: heads };
  }), () => ({ paragraf: 0, headings: [] });
}

export async function tulisGanti(teks) {
  return tryCatch(() => run(async (ctx) => {
    const r = ctx.document.getSelection(); r.insertText(teks, Word.InsertLocation.replace); await ctx.sync(); return true;
  }));
}

export async function sisipSetelah(teks) {
  return tryCatch(() => run(async (ctx) => {
    const r = ctx.document.getSelection(); r.insertText(teks, Word.InsertLocation.after); await ctx.sync(); return true;
  }));
}

export async function terapkanStyle(style) {
  return tryCatch(() => run(async (ctx) => {
    const r = ctx.document.getSelection(); r.style = style; await ctx.sync(); return true;
  }));
}

export async function buatTabel(rows, cols, data) {
  return tryCatch(() => run(async (ctx) => {
    const r = ctx.document.getSelection();
    const t = r.insertTable(rows, cols, Word.InsertLocation.after, data || undefined);
    t.headerRowCount = 1; t.style = "Table Grid"; await ctx.sync(); return true;
  }));
}

// Cari-ganti batch: 1 Word.run, loop search results (bukan 1 run per hasil).
export async function cariGanti(dari, ke, maks = 50) {
  return tryCatch(() => run(async (ctx) => {
    const res = ctx.document.body.search(dari, { matchCase: false }); res.load("items"); await ctx.sync();
    const n = Math.min(res.items.length, maks);
    for (let i = 0; i < n; i++) res.items[i].insertText(ke, Word.InsertLocation.replace);
    await ctx.sync(); return n;
  }));
}

export async function tambahKomentar(teks) {
  return tryCatch(() => run(async (ctx) => {
    const r = ctx.document.getSelection(); r.insertComment(teks); await ctx.sync(); return true;
  }));
}

// Sisip gambar (hasil render flowchart/chart) — 1 Word.run, cross-platform.
export async function sisipGambarBase64(b64, caption = "") {
  return tryCatch(() => run(async (ctx) => {
    const r = ctx.document.getSelection();
    r.insertInlinePictureFromBase64(b64, Word.InsertLocation.after);
    if (caption) r.insertText("\n" + String(caption).slice(0, 200), Word.InsertLocation.after);
    await ctx.sync(); return true;
  }));
}

// Kotak/shapes NATIF — hanya Word Desktop (WordApiDesktop 1.2). Web → fallback jujur.
// Batas jujur: API shape tanpa konektor, jadi cocok untuk kotak berlabel, bukan flowchart berpanah.
export async function kotakNative(teks, bentuk = "Rectangle") {
  if (!Office.context.requirements.isSetSupported("WordApiDesktop", "1.2"))
    return { ok: false, reason: "Butuh Word Desktop (Windows/Mac). Di Word Web gunakan flowchart sebagai gambar." };
  return tryCatch(() => run(async (ctx) => {
    const r = ctx.document.getSelection();
    if (/text|kotak/i.test(bentuk)) r.insertTextBox(teks || "", { width: 300, height: 80 });
    else r.insertGeometricShape(bentuk, { width: 200, height: 100 });
    await ctx.sync(); return { ok: true };
  }));
}

// Eksekutor ops dari AI (minimal-diff contract).
export async function eksekusiOps(ops) {
  for (const op of ops || []) {
    if (op.tool === "tulis_ganti") await tulisGanti(op.teks || "");
    else if (op.tool === "sisip_setelah") await sisipSetelah(op.teks || "");
    else if (op.tool === "terapkan_style") await terapkanStyle(op.style || "Normal");
    else if (op.tool === "buat_tabel") await buatTabel(op.rows || 3, op.cols || 3, op.data);
    else if (op.tool === "cari_ganti") await cariGanti(op.dari || "", op.ke || "");
    else if (op.tool === "sisip_gambar") await sisipGambarBase64(op.base64 || "", op.caption || "");
    else if (op.tool === "kotak_teks") { const r = await kotakNative(op.teks || "", op.bentuk || "Rectangle"); if (r && r.reason) console.warn(r.reason); }
    else if (op.tool === "buat_flowchart" || op.tool === "buat_chart") {
      const V = await import("./Visuals.js"); // lazy: hanya saat visual diminta
      let b64 = op._png; // sudah di-render saat preview → pakai ulang, hemat
      if (!b64) {
        if (op.tool === "buat_flowchart") b64 = (await V.renderMermaidPNG(op.mermaid || "graph TD;A-->B")).split(",")[1];
        else {
          const t = V.parseDataTable(op.data || "");
          const png = V.renderChartPNG({ title: op.title || "Chart", type: op.chartType || "bar",
            headers: op.headers || t?.headers || ["", "Nilai"], rows: op.rows || t?.rows || [] });
          b64 = png.split(",")[1];
        }
      }
      await sisipGambarBase64(b64 || "", op.caption || op.title || "");
    }
    // unknown tool: skip (hemat, anti-rusak)
  }
  return true;
}

// ponytail: no per-paragraf Merkle; hash murah di ManifestManager.
if (typeof window !== "undefined") window.__wordops_demo = async () => Array.isArray(await bacaMeta().then(m => m.headings));
