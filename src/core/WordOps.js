/* WordOps — 16 tools atomik. Prinsip perf: 1 Word.run per aksi, load properti minimal, walk ±3 paragraf (tanpa load seluruh body). */
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
  }), () => ({ paragraf: 0, headings: [] }));
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

export async function buatTabel(rows, cols, data, head) {
  return tryCatch(() => run(async (ctx) => {
    const r = ctx.document.getSelection();
    const t = r.insertTable(rows, cols, Word.InsertLocation.after, data || undefined);
    t.headerRowCount = 1; t.style = "Table Grid"; await ctx.sync();
    try { // header premium (bold + aksen): best-effort, gagal diam-diam di build lama
      t.rows.load("items"); await ctx.sync();
      const cells = t.rows.items[0].cells; cells.load("items"); await ctx.sync();
      for (const c of cells.items) c.body.font.set({ bold: true, color: (head && head.color) || "1F4E79" });
      await ctx.sync();
    } catch (e) {}
    return true;
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

// Sisip 1 paragraf blok bergaya di AKHIR dokumen (untuk hasil generate: CV, artikel, surat).
// font opsional {bold,color,size} untuk aksen premium; gagal di build lama → ditelan tryCatch.
export async function sisipBlok(teks, style = "Normal", font) {
  return tryCatch(() => run(async (ctx) => {
    const p = ctx.document.body.insertParagraph(teks || "", Word.InsertLocation.end);
    if (style && style !== "Normal") p.style = style;
    if (font) p.font.set(font);
    await ctx.sync(); return true;
  }));
}

// Audit struktur via 1x getOoxml + scan string (anti-API-uncertainty, murah untuk CV 1-2 hlm).
// stat: {tables, images, textboxes, columns}
export async function auditStruktur() {
  return tryCatch(() => run(async (ctx) => {
    const xml = ctx.document.body.getOoxml();
    await ctx.sync();
    const n = (re) => (xml.match(re) || []).length;
    return {
      tables: n(/<w:tbl[\s>]/g),
      images: n(/<pic:pic[\s>]/g) + n(/<wp:inline[\s>]/g), // termasuk grup shapes/flowchart (ATS juga tak membacanya)
      textboxes: n(/<wps:wsp[\s>]/g) + n(/<w:txbxContent[\s>]/g),
      columns: /<w:cols[^>]*w:num="[^1"]/.test(xml),
    };
  }), () => ({ tables: 0, images: 0, textboxes: 0, columns: false }));
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

// Sisip OOXML (flowchart native editable). Gagal → {ok:false} agar caller fallback ke gambar.
export async function sisipOoxml(ooxml) {
  return new Promise((resolve) => {
    try {
      Office.context.document.setSelectedDataAsync(ooxml, { coercionType: Office.CoercionType.Ooxml }, (r) => {
        if (r.status === Office.AsyncResultStatus.Succeeded) resolve(true);
        else resolve({ ok: false, reason: (r.error && r.error.message) || "ditolak Word" });
      });
    } catch (e) { resolve({ ok: false, reason: String((e && e.message) || e) }); }
  });
}

// Eksekutor ops dari AI (minimal-diff contract). execSatu agar builder bisa reuse.
// Kembali: true bila tool dikenali & dieksekusi, false bila dilewati → pesan jujur di UI.
export async function execSatu(op) {
  if (!op || !op.tool) return false;
  if (op.tool === "tulis_ganti") return !!(await tulisGanti(op.teks || ""));
  else if (op.tool === "sisip_setelah") return !!(await sisipSetelah(op.teks || ""));
  else   if (op.tool === "sisip_blok") return !!(await sisipBlok(op.teks || "", op.style || "Normal", op.font));
  else if (op.tool === "terapkan_style") return !!(await terapkanStyle(op.style || "Normal"));
  else if (op.tool === "buat_tabel") return !!(await buatTabel(op.rows || 3, op.cols || 3, op.data, op.head));
  else if (op.tool === "cari_ganti") return !!(await cariGanti(op.dari || "", op.ke || ""));
  else if (op.tool === "sisip_gambar") return !!(await sisipGambarBase64(op.base64 || "", op.caption || ""));
  else if (op.tool === "kotak_teks") { const r = await kotakNative(op.teks || "", op.bentuk || "Rectangle"); if (r && r.reason) console.warn(r.reason); return r === true || !!(r && r.ok); }
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
    return !!(await sisipGambarBase64(b64 || "", op.caption || op.title || ""));
  }
  else if (op.tool === "buat_flowchart_native") {
    const F = await import("./FlowchartOoxml.js");
    let ooxml = op._ooxml;
    if (!ooxml) { try { ooxml = F.buildFlowchartOoxml({ nodes: op.nodes, edges: op.edges }); } catch (e) { ooxml = null; } }
    const r = ooxml ? await sisipOoxml(ooxml) : { ok: false };
    if (r !== true) { // fallback jujur: gambar (selalu bisa di semua platform)
      console.warn("[DocAssist] native gagal, fallback gambar:", r && r.reason);
      const V = await import("./Visuals.js");
      const png = await V.renderMermaidPNG(F.flowToMermaid({ nodes: op.nodes, edges: op.edges }));
      return !!(await sisipGambarBase64(png.split(",")[1], (op.caption || "Flowchart") + " (gambar — mode editable gagal)"));
    }
    return true;
  }
  else if (op.tool === "buat_cv" || op.tool === "buat_cover_letter") {
    const B = await import("./Builders/CvBuilder.js"); // lazy: hanya saat kit karir diminta
    const tpl = String(op.template || op.style || "");
    const S = (tpl.startsWith("cv-style") || (op.tool === "buat_cover_letter" && op.style))
      ? await import("./Builders/CvStyles.js") : null;
    if (S) { // premium OOXML editable; gagal → fallback teks ATS di bawah
      const xml = tpl.startsWith("cv-style") ? S.buildCvStyled(tpl, op.data) : S.buildCoverStyled(op.data, op.style);
      if (await sisipOoxml(xml) === true) return true;
      console.warn("[DocAssist] styled gagal, fallback teks");
    }
    const subs = op.tool === "buat_cv"
      ? (tpl === "cv-modern" ? B.buildCvModern(op.data) : B.buildCvAts(op.data))
      : B.buildCoverLetter(op.data);
    let n = 0;
    for (const sub of subs) if (await execSatu(sub)) n++;
    return n > 0;
  }
  return false; // unknown tool: skip (hemat, anti-rusak)
}
export async function eksekusiOps(ops) {
  let n = 0;
  for (const op of ops || []) if (await execSatu(op)) n++;
  return n;
}

// ponytail: no per-paragraf Merkle; hash murah di ManifestManager.
if (typeof window !== "undefined") window.__wordops_demo = async () => Array.isArray(await bacaMeta().then(m => m.headings));
