/* CvBuilder — pure builder kit karir: data terstruktur → ops primitif WordOps.
   Tanpa DOM/Word di sini (bisa unit-test di Node). Angka: tidak ada hitungan —
   builder hanya menyusun teks; tanggung jawab AI = isi, bukan aritmetika. */

const s = (v) => String(v ?? "").trim();
const BUL = "• ";
// flat: objek {jenjang,sekolah,tahun} → "jenjang — sekolah — tahun". Anti "[object Object]".
const flat = (v) => Array.isArray(v) ? v.map(flat).filter(Boolean).join(", ")
  : (v && typeof v === "object" ? Object.values(v).map(flat).filter(Boolean).join(" — ") : s(v));

// Normalisasi data CV dari AI (toleran field hilang).
export function normCv(d = {}) {
  const exp = (d.pengalaman || []).map((e) => ({
    jabatan: flat(e.jabatan), perusahaan: flat(e.perusahaan), periode: flat(e.periode),
    poin: (e.poin || []).map(flat).filter(Boolean).slice(0, 6),
  })).filter((e) => e.jabatan || e.perusahaan).slice(0, 8);
  return {
    nama: flat(d.nama), kontak: flat(d.kontak), ringkasan: flat(d.ringkasan),
    pengalaman: exp,
    pendidikan: (d.pendidikan || []).map(flat).filter(Boolean).slice(0, 5),
    skills: (d.skills || []).map(flat).filter(Boolean).slice(0, 20),
  };
}

function sek(ops, judul) { ops.push({ tool: "sisip_blok", teks: judul, style: "Heading 2", font: { bold: true, color: "2E74B5" } }); }

export function buildCvAts(d) {
  const v = normCv(d), ops = [];
  ops.push({ tool: "sisip_blok", teks: v.nama || "Nama Lengkap", style: "Heading 1", font: { bold: true, color: "1F3864", size: 26 } });
  if (v.kontak) ops.push({ tool: "sisip_blok", teks: v.kontak, style: "Normal" });
  if (v.ringkasan) { sek(ops, "Ringkasan"); ops.push({ tool: "sisip_blok", teks: v.ringkasan, style: "Normal" }); }
  if (v.pengalaman.length) {
    sek(ops, "Pengalaman Kerja");
    for (const e of v.pengalaman) {
      ops.push({ tool: "sisip_blok", teks: [e.jabatan, e.perusahaan, e.periode].filter(Boolean).join(" — "), style: "Heading 3", font: { bold: true } });
      for (const p of e.poin) ops.push({ tool: "sisip_blok", teks: BUL + p, style: "Normal" });
    }
  }
  if (v.pendidikan.length) { sek(ops, "Pendidikan"); for (const p of v.pendidikan) ops.push({ tool: "sisip_blok", teks: BUL + p, style: "Normal" }); }
  if (v.skills.length) { sek(ops, "Keahlian"); ops.push({ tool: "sisip_blok", teks: v.skills.join(", "), style: "Normal" }); }
  return ops;
}

export function buildCvModern(d) {
  const ops = buildCvAts(d), v = normCv(d);
  if (v.skills.length) {
    // Tabel skill 2 kolom di akhir sebagai penegas visual (tetap ATS-aman: tabel sederhana).
    const rows = [["Kategori", "Detail"]];
    const half = Math.ceil(v.skills.length / 2);
    rows.push(["Teknis", v.skills.slice(0, half).join(", ")], ["Pendukung", v.skills.slice(half).join(", ") || "-"]);
    ops.push({ tool: "sisip_blok", teks: "Matriks Keahlian", style: "Heading 2", font: { bold: true, color: "2E74B5" } });
    ops.push({ tool: "buat_tabel", rows: 3, cols: 2, data: rows, head: { color: "1F4E79" } });
  }
  return ops;
}

export function buildCoverLetter(d = {}) {
  const ops = [], v = {
    nama: flat(d.nama), kontak: flat(d.kontak), perusahaan: flat(d.perusahaan),
    posisi: flat(d.posisi), sumber: flat(d.sumber), isi: flat(d.isi), kota: flat(d.kota), tanggal: flat(d.tanggal),
  };
  if (v.kota || v.tanggal) ops.push({ tool: "sisip_blok", teks: [v.kota, v.tanggal].filter(Boolean).join(", "), style: "Normal" });
  ops.push({ tool: "sisip_blok", teks: "Yth. HRD " + (v.perusahaan || "[Nama Perusahaan]"), style: "Normal" });
  ops.push({ tool: "sisip_blok", teks: `Dengan hormat, saya ${v.nama || "[Nama]"} bermaksud melamar posisi ${v.posisi || "[Posisi]"}${v.sumber ? " berdasarkan informasi " + v.sumber : ""}.`, style: "Normal" });
  ops.push({ tool: "sisip_blok", teks: v.isi || "[Ceritakan pengalaman dan motivasi Anda dalam 1 paragraf.]", style: "Normal" });
  ops.push({ tool: "sisip_blok", teks: "Besar harapan saya untuk dapat mengikuti tahap seleksi berikutnya. Atas perhatiannya saya ucapkan terima kasih.", style: "Normal" });
  ops.push({ tool: "sisip_blok", teks: `Hormat saya,\n${v.nama || "[Nama]"}\n${v.kontak}`, style: "Normal" });
  return ops;
}

// Verdict ATS dari hasil auditStruktur(): aturan berbasis struktur, tanpa AI (gratis + deterministik).
export function auditVerdict(st = {}) {
  const masalah = [];
  if (st.images > 0) masalah.push(`ditemukan ${st.images} gambar/foto — ATS sering gagal membacanya; hapus untuk versi ATS`);
  if (st.textboxes > 0) masalah.push(`ditemukan ${st.textboxes} textbox/shape — teks di dalamnya tak terbaca ATS`);
  if (st.tables > 3) masalah.push(`ditemukan ${st.tables} tabel — layout tabel kompleks berisiko teracak saat parsing`);
  if (st.columns) masalah.push("dokumen memakai kolom — ATS membaca per kolom secara berurutan (berantakan)");
  if (!masalah.length) return { lolos: true, pesan: "Struktur ATS-aman ✓ (tanpa gambar, textbox, tabel kompleks, kolom). Tips: font standar, 1–2 halaman, kirim sebagai .docx/.pdf teks." };
  return { lolos: false, pesan: "Belum ATS-aman:\n• " + masalah.join("\n• ") };
}

// demo: buildCvAts({nama:"A", pengalaman:[{jabatan:"X"}]}).length > 0
