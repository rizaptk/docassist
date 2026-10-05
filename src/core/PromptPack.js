/* PromptPack — anti-slop: grounding gaya dokumen + output JSON ops (minimal-diff). */
export const SYSTEM = `Kamu operator Word, bukan penulis esai.
1. Jangan ubah makna; perbaiki ejaan/alur minimal.
2. Gunakan istilah glosarium; tiru panjang kalimat sekitar.
3. Dilarang: pembuka generik ("Dalam era digital..."), emoji, heading baru tanpa diminta.
4. Pertahankan bahasa dokumen. Output HANYA JSON {ops:[...], alasan_singkat:"...", preview:"..."}.
Tools valid: tulis_ganti, sisip_setelah, terapkan_style, buat_tabel, cari_ganti.`;

export const TOOLS = [
  { name: "tulis_ganti", desc: "Ganti teks seleksi (range kecil saja)", params: { teks: "string" } },
  { name: "sisip_setelah", desc: "Sisip teks setelah kursor/seleksi", params: { teks: "string" } },
  { name: "terapkan_style", desc: "Terapkan style Word yang ada", params: { style: "Heading 1|Heading 2|Normal" } },
  { name: "buat_tabel", desc: "Buat tabel rapi ikut theme", params: { rows: "n", cols: "n" } },
  { name: "cari_ganti", desc: "Cari-ganti batch konsisten", params: { dari: "s", ke: "s" } },
];

export function buildUserPrompt(q, ctx) {
  const g = (ctx.l2?.glossary || []).join(", ");
  return `Tugas: ${q}\n[Gaya:tone=${ctx.l2?.tone || "netral"},lang=${ctx.l2?.lang || "id"},glosarium=${g}]\n[Konteks jendela]:\n${ctx.l1.konteks}\n[Outline]: ${(ctx.l2?.outline || []).join(" | ").slice(0, 400)}`;
}
