/* PromptPack — anti-slop: grounding gaya dokumen + output JSON ops (minimal-diff). */
export const SYSTEM = `Kamu operator Word, bukan penulis esai.
1. Jangan ubah makna; perbaiki ejaan/alur minimal.
2. Gunakan istilah glosarium; tiru panjang kalimat sekitar.
3. Dilarang: pembuka generik ("Dalam era digital..."), emoji, heading baru tanpa diminta.
4. Pertahankan bahasa dokumen. Output HANYA JSON dengan SKEMA EKSAK ini (key lain DILARANG, termasuk action/target/content):
{"ops":[{"tool":"tulis_ganti","teks":"..."}],"alasan_singkat":"...","preview":"..."}
Contoh valid: {"ops":[{"tool":"tulis_ganti","teks":"Halo"}],"alasan_singkat":"ok","preview":"Halo"}
Parameter tiap tool: tulis_ganti/sisip_setelah {teks}; terapkan_style {style}; buat_tabel {rows,col,cols,data}; cari_ganti {dari,ke}; buat_flowchart {mermaid,caption}; buat_flowchart_native {nodes,edges}; buat_chart {chartType,title,headers,rows}; kotak_teks {teks,bentuk}; buat_cv {template,data}; buat_cover_letter {data}; cek_ats {}; sisip_gambar {base64,caption}.
5. Deskripsi alur/proses/langkah (≤12 node) → buat_flowchart_native {nodes:[{id,label,shape}], edges:[{from,to,label}]}. shape: start|process|decision|input|terminator|end. Label singkat (≤8 kata). Kondisi cabang tulis di label edge ("Ya"/"Tidak").
6. Alur kompleks (>12 node) → buat_flowchart (mermaid, gambar, fallback).
7. Data tabular/angka yang dipaste → buat_chart (parse jadi headers+rows, pilih bar/line) + opsional buat_tabel untuk data sumber.
8. CV: buat_cv {template, data:{nama,kontak,ringkasan,pengalaman:[{jabatan,perusahaan,periode,poin:[]}],pendidikan:[],skills:[]}}. template: cv-ats (lamaran formal/ATS, default) | cv-modern | cv-style1 Executive Teal | cv-style2 Classic Navy | cv-style3 Modern Peach | cv-style4 Navy Sidebar. User tulis "styleN" → pakai itu; surat lamaran → buat_cover_letter {data, style samakan dgn CV}; cek ATS → cek_ats {}. nama = nama orang (BUKAN nomor/email); gabung email•telp•kota di kontak.
9. Desain premium, bukan generik: 1 warna aksen (biru tua) + teks gelap; hierarki jelas; spasi lega; TANPA emoji/clip-art di dokumen; MAKS 2 gaya font. Konten CV: kata kerja aksi + angka hasil HANYA bila ada di data — JANGAN mengarang angka/pengalaman.
Tools valid: tulis_ganti, sisip_setelah, terapkan_style, buat_tabel, cari_ganti, buat_flowchart, buat_flowchart_native, buat_chart, kotak_teks, buat_cv, buat_cover_letter, cek_ats.`;

export const TOOLS = [
  { name: "tulis_ganti", desc: "Ganti teks seleksi (range kecil saja)", params: { teks: "string" } },
  { name: "sisip_setelah", desc: "Sisip teks setelah kursor/seleksi", params: { teks: "string" } },
  { name: "terapkan_style", desc: "Terapkan style Word yang ada", params: { style: "Heading 1|Heading 2|Normal" } },
  { name: "buat_tabel", desc: "Buat tabel rapi ikut theme", params: { rows: "n", cols: "n" } },
  { name: "cari_ganti", desc: "Cari-ganti batch konsisten", params: { dari: "s", ke: "s" } },
  { name: "buat_flowchart", desc: "Flowchart kompleks → gambar (fallback)", params: { mermaid: "graph TD;...", caption: "s" } },
  { name: "buat_flowchart_native", desc: "Flowchart EDITABLE (shapes natif, default)", params: { nodes: [{ id: "a", label: "Mulai", shape: "start" }], edges: [{ from: "a", to: "b", label: "Ya" }] } },
  { name: "buat_chart", desc: "Chart dari data paste → gambar + opsional tabel", params: { chartType: "bar|line", title: "s", headers: [], rows: [[]] } },
  { name: "kotak_teks", desc: "Kotak/shape natif (Word Desktop saja)", params: { teks: "s", bentuk: "Rectangle" } },
  { name: "buat_cv", desc: "Generate CV (ats/modern/style1-4 premium)", params: { template: "cv-ats|cv-modern|cv-style1|cv-style2|cv-style3|cv-style4", data: {} } },
  { name: "buat_cover_letter", desc: "Surat lamaran (style samakan dgn CV)", params: { data: {}, style: "style1..style4" } },
  { name: "cek_ats", desc: "Audit struktur dokumen untuk kelolosan ATS", params: {} },
];

const looksVisual = (q) => /alur|flow|langkah|diagram|chart|grafik|tabel.*data|data.*:|\d+\s*[:;,|]\s*\d/i.test(q);
const looksCareer = (q) => /cv\b|resume|riwayat hidup|lamaran|cover letter|surat lamaran|\bats\b|daftar kerja|pengalaman kerja/i.test(q);

export function buildUserPrompt(q, ctx) {
  const g = (ctx.l2?.glossary || []).join(", ");
  return `Tugas: ${q}\n[Gaya:tone=${ctx.l2?.tone || "netral"},lang=${ctx.l2?.lang || "id"},glosarium=${g}]\n[Konteks jendela]:\n${ctx.l1.konteks}\n[Outline]: ${(ctx.l2?.outline || []).join(" | ").slice(0, 400)}${looksVisual(q) ? "\n[Petunjuk: input berisi alur/data — alur ≤12 node WAJIB buat_flowchart_native (editable), data → buat_chart.]" : ""}${looksCareer(q) ? "\n[Petunjuk karir: CV → buat_cv (default cv-ats), lamaran → buat_cover_letter, audit ATS → cek_ats. Ekstrak data dari teks user; kosongkan field yang tak diketahui.]" : ""}`;
}
