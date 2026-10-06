/* CvStyles — 4 layout CV premium ala contoh (native OOXML, editable).
   Token dari ui-ux-pro-max: navy #1E3A5F + putih; serif Cambria (ganti EB Garamond)
   + sans Calibri (ganti Lato) karena font Word-native. Foto = placeholder yang
   user ganti manual (Insert > Pictures). Style non-ATS (shading/kolom) — untuk
   lamaran formal/ATS tetap pakai cv-ats. Murni string builder (testable Node). */
import { P, CELL, ROW, TBL, EMPTY, wrapPkg, DXA_PAGE } from "./OoxmlKit.js";
import { normCv } from "./CvBuilder.js";

export const ACCENTS = { style1: "0E5F5B", style2: "1E3A5F", style3: "1A1A1A", style4: "1E3A5F" };
const INK = "262626", MUT = "475569", WHITE = "FFFFFF", PAPER = "F2F2F2", PEACH = "FBE9DC";

const sec = (t, accent) => P(t.toUpperCase(), { bold: true, color: accent, size: 11, after: 60, rule: accent });
const job = (e) => [
  P([e.jabatan, e.perusahaan].filter(Boolean).join(", "), { bold: true, after: 20 }),
  ...(e.periode ? [P(e.periode, { color: MUT, size: 9, after: 20 })] : []),
  ...e.poin.map((p) => P("•  " + p, { after: 40 })),
  EMPTY(60),
];
const fotoBox = (w, fill, tc) => CELL([P("[ FOTO ]", { align: "center", color: tc, size: 10, after: 40 }), P("ganti via Insert > Pictures", { align: "center", color: tc, size: 7 })], { fill, width: w });
const kontakLines = (v) => v.kontak.split(/[•\n|]/).map((s) => s.trim()).filter(Boolean);

// style1 EXECUTIVE TEAL: band header + body 2 kolom (sidebar abu).
export function style1(v) {
  const A = ACCENTS.style1;
  const band = TBL([ROW([
    CELL([P(v.nama || "Nama Lengkap", { bold: true, color: WHITE, size: 26, after: 40 }), P(v.kontak, { color: WHITE, size: 10 })], { fill: A, width: 7400 }),
    fotoBox(1960, "0B4A46", WHITE),
  ])], [7400, 1960]);
  const main = [
    ...(v.ringkasan ? [sec("Profil", A), P(v.ringkasan, { after: 120 })] : []),
    ...(v.pengalaman.length ? [sec("Pengalaman Kerja", A), ...v.pengalaman.flatMap(job)] : []),
    ...(v.pendidikan.length ? [sec("Pendidikan", A), ...v.pendidikan.map((p) => P("•  " + p, { after: 40 }))] : []),
  ];
  const side = [
    ...(v.kontak ? [sec("Kontak", A), ...kontakLines(v).map((c) => P(c, { size: 9, after: 30 })), EMPTY(80)] : []),
    ...(v.skills.length ? [sec("Keahlian", A), ...v.skills.map((s) => P("•  " + s, { size: 9, after: 30 }))] : []),
  ];
  const body = TBL([ROW([CELL(main, { width: 6100 }), CELL(side, { fill: PAPER, width: 3000 })])], [6100, 3000]);
  return wrapPkg(band + EMPTY(120) + body);
}

// style2 CLASSIC NAVY: foto + nama serif + divider + sidebar.
export function style2(v) {
  const A = ACCENTS.style2;
  const head = TBL([ROW([
    fotoBox(1800, "E9EEF5", MUT),
    CELL([P(v.nama || "Nama Lengkap", { bold: true, color: A, size: 28, after: 40, font: "Cambria" }), P(v.kontak, { color: MUT, size: 10 })], { width: 7560 }),
  ])], [1800, 7560]);
  const side = [
    ...(v.kontak ? [sec("Kontak", A), ...kontakLines(v).map((c) => P(c, { size: 9, after: 30 })), EMPTY(80)] : []),
    ...(v.pendidikan.length ? [sec("Pendidikan", A), ...v.pendidikan.map((p) => P(p, { size: 9, after: 30 })), EMPTY(80)] : []),
    ...(v.skills.length ? [sec("Keahlian", A), ...v.skills.map((s) => P("•  " + s, { size: 9, after: 30 }))] : []),
  ];
  const main = [
    ...(v.ringkasan ? [sec("Profil Singkat", A), P(v.ringkasan, { after: 120 })] : []),
    ...(v.pengalaman.length ? [sec("Pengalaman Kerja", A), ...v.pengalaman.flatMap(job)] : []),
  ];
  const body = TBL([ROW([CELL(side, { width: 3000 }), CELL(main, { width: 6100 })])], [3000, 6100]);
  return wrapPkg(head + P("", { rule: A, after: 120 }) + body);
}

// style3 MODERN PEACH: nama tengah + bar kontak hitam + sidebar peach.
export function style3(v) {
  const A = ACCENTS.style3, RULE = "C97B4A";
  const head = P(v.nama || "Nama Lengkap", { align: "center", bold: true, size: 26, after: 20 })
    + P("Curriculum Vitae", { align: "center", color: MUT, size: 11, after: 100 });
  const bar = TBL([ROW([CELL([P(v.kontak || "kontak • email • kota", { align: "center", color: WHITE, size: 9 })], { fill: A, width: DXA_PAGE })])], [DXA_PAGE]);
  const side = [
    ...(v.pendidikan.length ? [sec("Pendidikan", A), ...v.pendidikan.map((p) => P(p, { size: 9, after: 30 })), EMPTY(80)] : []),
    fotoBox(0, PEACH, MUT),
    ...(v.skills.length ? [EMPTY(60), sec("Keahlian", A), ...v.skills.map((s) => P("•  " + s, { size: 9, after: 30 }))] : []),
  ];
  const main = [
    ...(v.ringkasan ? [sec("Tentang Saya", RULE), P(v.ringkasan, { after: 120 })] : []),
    ...(v.pengalaman.length ? [sec("Pengalaman Kerja", RULE), ...v.pengalaman.flatMap(job)] : []),
  ];
  const body = TBL([ROW([CELL(side, { fill: PEACH, width: 3000 }), CELL(main, { width: 6100 })])], [3000, 6100]);
  return wrapPkg(head + bar + EMPTY(120) + body);
}

// style4 NAVY SIDEBAR: sidebar navy penuh + konten terang.
export function style4(v) {
  const A = ACCENTS.style4, W = WHITE;
  const side = [
    fotoBox(0, "2E5987", W), EMPTY(80),
    ...(v.kontak ? [sec("Kontak", W), ...kontakLines(v).map((c) => P(c, { color: W, size: 9, after: 30 })), EMPTY(80)] : []),
    ...(v.pendidikan.length ? [sec("Pendidikan", W), ...v.pendidikan.map((p) => P(p, { color: W, size: 9, after: 30 })), EMPTY(80)] : []),
    ...(v.skills.length ? [sec("Keahlian", W), ...v.skills.map((s) => P("•  " + s, { color: W, size: 9, after: 30 }))] : []),
  ];
  const title = (v.pengalaman[0] && v.pengalaman[0].jabatan) || "Curriculum Vitae";
  const main = [
    P(v.nama || "Nama Lengkap", { bold: true, color: A, size: 26, after: 20 }),
    P(title, { color: MUT, size: 11, after: 40, rule: A }),
    ...(v.ringkasan ? [sec("Tentang Saya", A), P(v.ringkasan, { after: 120 })] : []),
    ...(v.pengalaman.length ? [sec("Pengalaman Kerja", A), ...v.pengalaman.flatMap(job)] : []),
  ];
  const body = TBL([ROW([CELL(side, { fill: A, width: 3000 }), CELL(main, { width: 6100 })])], [3000, 6100]);
  return wrapPkg(body);
}

const STYLES = { "cv-style1": style1, "cv-style2": style2, "cv-style3": style3, "cv-style4": style4 };
export const STYLE_IDS = Object.keys(STYLES);
export function buildCvStyled(template, data) {
  const fn = STYLES[template] || style1;
  return fn(normCv(data));
}

// Cover letter ber-aksen: kop band + isi (pola style sama, reusable kit lain).
const clean = (v, fb) => {
  const f = Array.isArray(v) ? v.join(", ") : (v && typeof v === "object" ? Object.values(v).join(" — ") : String(v ?? ""));
  return f.trim() || fb;
};
export function buildCoverStyled(d = {}, style = "style1") {
  const A = ACCENTS[style] || ACCENTS.style1;
  const v = {
    nama: clean(d.nama, "[Nama]").slice(0, 80),
    kontak: clean(d.kontak, "").slice(0, 160),
    perusahaan: clean(d.perusahaan, "[Nama Perusahaan]").slice(0, 80),
    posisi: clean(d.posisi, "[Posisi]").slice(0, 80),
    sumber: clean(d.sumber, "").slice(0, 160),
    isi: clean(d.isi, "[Ceritakan pengalaman dan motivasi Anda dalam 1 paragraf.]").slice(0, 2000),
    kota: clean(d.kota, "").slice(0, 60),
    tanggal: clean(d.tanggal, "").slice(0, 60),
  };
  const kop = TBL([ROW([CELL([
    P(v.nama, { bold: true, color: WHITE, size: 18, after: 20 }),
    P(v.kontak, { color: WHITE, size: 9 }),
  ], { fill: A, width: DXA_PAGE })])], [DXA_PAGE]);
  const body = [
    ...((v.kota || v.tanggal) ? [P([v.kota, v.tanggal].filter(Boolean).join(", "))] : []),
    P("Yth. HRD " + v.perusahaan),
    P(`Dengan hormat, saya ${v.nama} bermaksud melamar posisi ${v.posisi}${v.sumber ? " berdasarkan informasi " + v.sumber : ""}.`),
    P(v.isi),
    P("Besar harapan saya untuk dapat mengikuti tahap seleksi berikutnya. Atas perhatiannya saya ucapkan terima kasih."),
    P(`Hormat saya,\n${v.nama}\n${v.kontak}`),
  ];
  return wrapPkg(kop + EMPTY(120) + body.join(""));
}
