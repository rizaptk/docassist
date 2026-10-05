# DocAssist — Riset Tools Baru (Artikel, Promo, CV, Presentasi-ala-Word)

> Tanggal: Okt 2026 | Sumber: pola pemakaian Word populer + kapabilitas Word JS API terverifikasi.

## Temuan kunci (menentukan desain)

1. **Permintaan user Word paling sering**: CV/resume + cover letter, flyer/poster/brosur promosi, newsletter, laporan/proposal, invoice, agenda/undangan. Kompetitor template (30rb+ template di app sejenis) membuktikan pasarnya.
2. **API Word mendukung**: content controls (cross-platform, ideal untuk template isi-form), `insertFileFromBase64` (sisip blok dokumen), tabel/gambar/shapes, TOC/daftar isi.
3. **BATAS PENTING — `PageSetup` & header/footer = `WordApiDesktop 1.3` (Desktop Windows/Mac saja, preview di Web)**. Artinya: margin, orientasi landscape, ukuran kertas, kolom teks, header/footer kustom **tidak bisa diandalkan di Word Web**. Semua tool desain halaman wajib punya **fallback graceful** (bangun konten dulu, page-setup hanya bila didukung via `isSetSupported`).
4. **Strategi template pemenang** (dipakai industri, mis. Textkernel): template = dokumen biasa berisi **tag `{{placeholder}}`**; AI/user isi data → tag diganti dengan format tag dipertahankan; blok berulang (pengalaman kerja) via pola foreach sederhana. Tanpa server, tanpa dep.

## Backlog prioritas (dampak × usaha × kelayakan API)

### P1 — CV Generator + Cover Letter (bangun duluan, full cross-platform)
- **Job user**: "buatkan CV dari data saya", "rapikan CV ini agar lolos ATS", "buatkan cover letter".
- **Tools**: `buat_cv {template, data}` (isi tag `{{nama}}`… + tambah/hapus baris pengalaman), `cek_ats` (pindai: tabel/kolom/gambar/fancy-font yang merusak parsing ATS → sarankan versi polos), `buat_cover_letter {data, lowongan}`.
- **Template**: `config/templates.json` (mirip `providers.json`!) berisi 3 starter: ATS-friendly, Modern 2-kolom (tabel), Kreatif (shapes, desktop). Template disimpan sebagai **fungsi builder JS** (bukan file .dotx — tanpa backend, builder = kode).
- **Usaha**: kecil. Semua API sudah ada (tabel, style, insertText). ATS-check = 1 `Word.run` baca struktur.

### P1 — Article Writer & Enhancer (bangun duluan, full cross-platform)
- **Job user**: "tulis artikel 800 kata tentang X", "perbaiki alur + EYD", "buatkan abstrak/kesimpulan/daftar pustaka", "ubah gaya bahasa akademik".
- **Tools**: `tulis_artikel {topik, panjang, gaya, outline?}` (generate per-section agar hemat token + bisa stop), `enhance_artikel {mode: eyd|alur|gaya|ringkas}`, `buat_abstrak`, `format_heading_otomatis` (terapkan Heading 1-3 dari struktur).
- **Anti-slop khusus artikel**: grounding glosarium + larangan klise sudah ada; tambah `cek_fakta_tandai` (AI tandai klaim angka/tahun dengan komentar "verifikasi manual" — jujur soal halusinasi).
- **Usaha**: kecil-menengah. Reuse ContextBuilder + PromptPack (tambah 2 template prompt).

### P2 — Promo Page Builder: flyer/poster/brosur 1-halaman
- **Job user**: "buatkan flyer promo produk X", "brosur lipat-3", "poster acara".
- **Tools**: `buat_flyer {judul, promo, kontak, gaya}` (cover shapes + headline + tabel harga + gambar + CTA box), `tata_brosur` (3 kolom via textColumns — desktop; fallback tabel 3-kolom di Web!).
- **Kelayakan**: konten 100% bisa; page-setup (landscape/margin sempit) desktop-only → pola `coba pageSetup, gagal → lanjut konten + pesan jujur`.
- **Usaha**: menengah (butuh builder OOXML/DrawingML seperti flowchart + 2-3 preset gaya).

### P2 — Pitch Deck ala Word ("presentasi tanpa PowerPoint")
- **Job user**: UMKM/student yang hanya punya Word tapi butuh materi pitching.
- **Tools**: `buat_pitch_doc {topik}` = cover page + 6-8 section 1-topik (masalah, solusi, pasar, model bisnis, traction, tim, ajakan) + tiap section diawali heading + visual (chart/flowchart reuse Visuals.js!) + daftar isi otomatis.
- **Kelayakan**: full cross-platform (tidak butuh pageSetup). Ini diferensiasi unik vs aplikasi template lain.
- **Usaha**: menengah-kecil (orkestrasi tools yang sudah ada + 1 builder cover).

### P3 — Business kit (laporan, proposal, invoice, agenda, newsletter)
- Pola sama: 1 builder + 1 prompt pack per jenis. Newsletter butuh kolom (fallback tabel). Invoice butuh tabel + total otomatis (hitung di JS, bukan AI — anti salah hitung!).
- **Prinsip UploadAble**: tiap kit = entri di `config/templates.json` + 1 file builder. Tidak membengkakkan core.

## Keputusan arsitektur (agar tidak jadi spaghetti)

1. **`config/templates.json`** — registry template (nama, jenis, builder, deskripsi, desktopOnly flag). Cerminan `providers.json`.
2. **`src/core/Builders/`** — 1 file per kit (`CvBuilder.js`, `FlyerBuilder.js`, …), pure function data → Word ops. Core (`WordOps`, `ContextBuilder`) tidak diutak-atik.
3. **3. **Aturan angka**: harga/total/diskon SELALU dihitung JS, AI hanya format teks. (AI + aritmetika = sumber komplain.)
4. **Desktop-only selalu feature-detect** (`isSetSupported`) + fallback Web + pesan jujur di chat.

## Rekomendasi eksekusi

Gelombang 1 (nilai tercepat, risiko nol): **P1 CV + P1 Artikel** — 2 builder + 3 prompt pack + registry template, tanpa API baru.
Gelombang 2: **Pitch-doc**, lalu **Flyer** (butuh DrawingML preset seperti flowchart).
