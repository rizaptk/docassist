/* PromptPack — anti-slop: grounding gaya dokumen + output JSON ops (minimal-diff). */
export const SYSTEM = `Kamu operator Word, bukan penulis esai.
1. Jangan ubah makna; perbaiki ejaan/alur minimal.
2. Gunakan istilah glosarium; tiru panjang kalimat sekitar.
3. Dilarang: pembuka generik ("Dalam era digital..."), emoji, heading baru tanpa diminta.
4. Pertahankan bahasa dokumen. Output HANYA JSON {ops:[...], alasan_singkat:"...", preview:"..."}.
5. Deskripsi alur/proses/langkah (≤12 node) → buat_flowchart_native {nodes:[{id,label,shape}], edges:[{from,to,label}]}. shape: start|process|decision|input|terminator|end. Label singkat (≤8 kata). Kondisi cabang tulis di label edge ("Ya"/"Tidak").
6. Alur kompleks (>12 node) → buat_flowchart (mermaid, gambar, fallback).
7. Data tabular/angka yang dipaste → buat_chart (parse jadi headers+rows, pilih bar/line) + opsional buat_tabel untuk data sumber.
Tools valid: tulis_ganti, sisip_setelah, terapkan_style, buat_tabel, cari_ganti, buat_flowchart, buat_flowchart_native, buat_chart, kotak_teks.`;

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
];

const looksVisual = (q) => /alur|flow|langkah|diagram|chart|grafik|tabel.*data|data.*:|\d+\s*[:;,|]\s*\d/i.test(q);

export function buildUserPrompt(q, ctx) {
  const g = (ctx.l2?.glossary || []).join(", ");
  return `Tugas: ${q}\n[Gaya:tone=${ctx.l2?.tone || "netral"},lang=${ctx.l2?.lang || "id"},glosarium=${g}]\n[Konteks jendela]:\n${ctx.l1.konteks}\n[Outline]: ${(ctx.l2?.outline || []).join(" | ").slice(0, 400)}${looksVisual(q) ? "\n[Petunjuk: input berisi alur/data — alur ≤12 node WAJIB buat_flowchart_native (editable), data → buat_chart.]" : ""}`;
}
