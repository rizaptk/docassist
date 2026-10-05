# DocAssist — Perencanaan Add-on AI Asisten Operasional MS Word

> Status: DRAFT perencanaan (MVP free-usage) | Target: Word M365 Modern | Distribusi: manual sideload | AI: BYOK OpenAI + Gemini

## 1. Ringkasan Eksekutif

**DocAssist** adalah add-on MS Word yang berperan sebagai **asisten operasional dokumen**: membantu user memformat, membuat, mengubah, mengedit, menambahkan, dan mendesain dokumen yang sedang dibuka — lewat **chat widget kecil** (task pane sempit yang terasa seperti floating widget).

**Keputusan kunci (dari diskusi):**

- Platform: **Office Web Add-in (Office.js + Word.js)** — cross-platform (Windows/Mac/Web), hosting statis gratis.
- Bukan VSTO/COM/VBA — agar tetap gratis, ringan, dan bisa manual install.
- Distribusi: **manual install (sideload manifest)** + panduan. Bukan Marketplace dulu (hindari batasan `showAsTaskpane`/`auto-open` 2026 untuk listing publik).
- AI: **BYOK (Bring Your Own Key)** per user. Tanpa backend sendiri di MVP → biaya server Rp0, user bayar token masing-masing.
- Model default murah: `gpt-4o-mini` / `gemini-2.0-flash`. Model premium opsional via dropdown.

**Verdict peluang: 8/10 — feasible.** Pesaing (M365 Copilot) butuh lisensi berbayar; ceruk DocAssist = Bahasa Indonesia, BYOK murah, workflow operasional custom, enterprise yang mau sideload internal.

---

## 2. Tujuan & Ruang Lingkup

### 2.1 Tujuan

User dapat, lewat bahasa natural Indonesia:

- Memformat (style, font, heading, spacing, list, tabel).
- Membuat / melanjutkan / menulis ulang teks **menyatu dengan gaya dokumen** (anti-slop).
- Mengubah / mencari-ganti / merapikan konsisten istilah.
- Menambahkan (paragraf, tabel, gambar, daftar isi, komentar, footnote).
- Mendesain (tabel rapi ikut theme, struktur heading untuk TOC, layout sederhana).

### 2.2 Non-tujuan MVP (sengaja di-skip)

- True-floating overlay bebas di atas kanvas (tidak diizinkan sandbox Office.js).
- Backend proxy, login/OAuth, billing, analitik server.
- Publish ke Microsoft Marketplace / AppSource.
- Integrasi Copilot Agent (preview, butuh lisensi Copilot).
- Vector DB / embeddings server. Diganti cache ringan (lihat §5).
- 100% parity VBA/COM (dialog print, mail-merge penuh, ActiveX, akses file lokal).

> Prinsip: kirim MVP lazy yang bekerja, tambah yang berat saat ada user berbayar. Tandai simplifikasi dengan komentar `ponytail:`.

---

## 3. Arsitektur Sistem (MVP)

```
┌─────────────────────────────────────────────┐
│  MS Word (Windows / Mac / Web)              │
│  ┌───────────────────────────────────────┐  │
│  │ Task Pane (330px, docked chat widget) │  │
│  │  React/Vanilla JS + CSS               │  │
│  │  - ChatBubble, QuickChips             │  │
│  │  - Settings (provider/key/model)      │  │
│  │  - Badge cache + Preview Diff         │  │
│  └──────────────┬────────────────────────┘  │
│                 │ Office.js (Word.js)        │
│  ┌──────────────▼────────────────────────┐  │
│  │ WordBridge.js (12 tools atomik)       │  │
│  │ ContextBuilder.js (layered context)   │  │
│  │ ManifestManager.js (cache AI)         │  │
│  │ AIClient.js (fetch langsung)          │  │
│  └──────────────┬────────────────────────┘  │
└─────────────────┼───────────────────────────┘
                  │ HTTPS langsung (BYOK)
        ┌─────────▼──────────┐   ┌──────────────────────┐
        │ api.openai.com     │   │ generativelanguage.  │
        │ (user key)         │   │ googleapis.com (key) │
        └────────────────────┘   └──────────────────────┘

Hosting add-in: GitHub Pages / Netlify / Vercel (statis, gratis).
Penyimpanan: localStorage (key, preferensi) + CustomXmlParts (manifest ikut .docx).
Tanpa server aplikasi di MVP.
```

### 3.1 Struktur file yang diusulkan

```
docassist/
├── manifest.xml              # sideload manual (ReadWriteDocument)
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html     # shell chat widget
│   │   ├── taskpane.css      # compact, 330px
│   │   └── taskpane.ts       # bootstrap Office.onReady
│   ├── core/
│   │   ├── WordOps.ts        # 12 tools -> Word.run()
│   │   ├── ContextBuilder.ts # Layer 0/1/2/3
│   │   ├── ManifestManager.ts# cache CustomXmlParts + localStorage
│   │   ├── AIClient.ts       # OpenAI + Gemini, satu interface
│   │   └── PromptPack.ts     # system prompt + template anti-slop
│   └── ui/
│       ├── Chat.tsx          # bubble + chips + diff preview
│       └── Settings.tsx      # provider/key/model/suhu/test conn
└── docs/
    └── PANDUAN_INSTALL.md    # sideload step-by-step + video
```

### 3.2 Alur operasi (happy path)

1. User seleksi teks / letakkan kursor → klik Ribbon `🤖 DocAssist` (atau pane auto-open jika sideload).
2. Task pane tampil sebagai chat kecil. Badge tampil: `Konteks: cached 2 jam lalu [Refresh]`.
3. User ketik: "rapikan paragraf ini jadi formal, pertahankan istilah teknis".
4. `ContextBuilder` susun konteks berlapis (lihat §4) + `ManifestManager` suntik ringkasan cache.
5. `AIClient` kirim ke OpenAI/Gemini dengan `tools` (function-calling). AI balas **JSON ops**, bukan blob teks.
6. UI tampilkan **preview diff** (sebelum → sesudah) untuk aksi >1 paragraf.
7. User klik `Terapkan` → `WordOps` eksekusi via `Word.run()` + `context.sync()`. `Batal` = buang.
8. Perubahan kecil (<10% char) → patch manifest inkremental; besar → tandai cache stale.

---

## 4. Strategi Cerdas Efisiensi Token

### 4.1 Prinsip: layered context, bukan full-doc dump

| Layer | Isi | Estimasi | Kapan dipakai |
|-------|-----|----------|---------------|
| L0 meta | intent, jml paragraf/tabel, daftar style, posisi kursor, bahasa | ~300 tk | selalu |
| L1 jendela kerja | seleksi + ±3 paragraf atas/bawah + heading section + format range itu | 500–2000 tk | selalu untuk operasi lokal |
| L2 cache | outline + glosarium + profil gaya + ringkasan per-section (dari manifest) | <800 tk | saat butuh gaya/konsistensi global |
| L3 penuh | chunk per-heading via `body.search()` / `getFileAsync` slice | sesuai ukuran | hanya: ringkas full-doc, audit konsistensi |

**Routing:**

- `format seleksi ini` → L0+L1 saja.
- `lanjutkan / tulis ulang bab ini` → L1 + profil gaya (L2).
- `rapikan seluruh dokumen / buatkan TOC / cek istilah` → L2 + scan bertahap per-heading (loop, bukan sekali kirim).
- `desain tabel/gambar` → kirim struktur (rows/cols/caption), bukan narasi.

Hasil: hemat 70–90% vs kirim full OOXML tiap turn. Model murah sudah cukup karena tugas operasional, bukan reasoning berat.

### 4.2 Aturan praktis hemat token

- Jangan kirim OOXML mentah ke LLM. Konversi ke teks + atribut minimal (`style, bold, listLevel`) di client.
- Potong jendela kerja: maks 4000 karakter sekitar kursor untuk edit lokal.
- Scan global = iterasi per-heading, agregasi di client, 1 panggilan sintesis akhir.
- `temperature` rendah (0.2–0.4) untuk edit; 0.7 hanya ideasi.
- Cache prompt: system prompt + profil gaya disimpan, tidak dikirim ulang sebagai narasi panjang — cukup delta per turn.
- Default model termurah; naikkan hanya jika user pilih manual.

---

## 5. Compat Cached File Generated by AI (DocManifest)

### 5.1 Konsep

**DocManifest AI** = JSON kecil (<800 token) yang dibuat AI sekali, dipakai ulang sebagai pengganti baca ulang full-doc. Memberi ilusi "AI ingat seluruh dokumen".

### 5.2 Skema (v1)

```json
{
  "schema": "docassist/manifest/v1",
  "docHash": "a3f9c1e2",
  "updatedAt": "2026-10-05T00:00:00Z",
  "outline": ["1. Pendahuluan", "2. Metode", "3. Hasil"],
  "styleProfile": {
    "lang": "id",
    "tone": "formal-akademik",
    "headingFont": "Calibri 14 Bold",
    "bodyFont": "Calibri 11",
    "avgSentenceLen": 18,
    "forbiddenOpeners": ["Dalam era digital yang serba cepat"]
  },
  "glossary": ["efisiensi token", "CustomXmlParts", "task pane"],
  "sectionSummary": {
    "1": "latar belakang + rumusan masalah, 2 kalimat",
    "2": "metode layered context + cache"
  },
  "styleInventory": ["Heading 1", "Heading 2", "Normal", "Caption"],
  "doNotTouch": ["halaman sampul", "tabel lampiran B"],
  "contextBudget": { "l1Chars": 4000, "l2Tokens": 800 }
}
```

### 5.3 Penyimpanan ganda (gratis, tanpa backend)

1. **`CustomXmlParts`** (namespace `http://docassist/ai-manifest/v1`) — ikut tersimpan di dalam `.docx`, pindah laptop tetap ada.
2. **`localStorage["docassist:<docId>:manifest"]`** — akses instan tanpa panggil Word API.

### 5.4 Fingerprint & invalidasi

- `docHash = hash(jumlahParagraf + totalChars + daftarHeading + versiStyle)`.
- Hitung tiap idle 10 detik atau saat AI dipanggil.
- Hash sama → pakai cache (0 token rebuild).
- Delta kecil (<10% char) → patch inkremental: ringkas ulang section berubah saja.
- Delta besar → tandai stale, tampilkan `[Refresh]`, rebuild di latar atau saat user klik.
- Badge UI wajib: `Konteks: cached 2 jam lalu [Refresh]` — transparan ke user.

### 5.5 API ManifestManager (kontrak)

```ts
loadManifest(): Promise<Manifest | null>
saveManifest(m: Manifest): Promise<void>   // tulis ke CustomXmlParts + localStorage
computeHash(): Promise<string>
isStale(cached: Manifest): Promise<boolean>
patchSection(heading: string, newSummary: string): Promise<void>
rebuildFull(onProgress): Promise<Manifest>  // loop per-heading, agregasi client
```

> `ponytail: checksum sederhana (bukan Merkle-tree per-paragraf); upgrade ke diff per-paragraf bila dokumen >100 hlm bermasalah.`

---

## 6. Anti-Slop: Edit/Generate Menyatu dengan Konteks

Slop = gaya generik + menimpa formatting dengan plain-text. Dicegah via 4 lapis:

### 6.1 Style-grounding

Setiap request generate sertakan otomatis:

- `styleProfile` (tone, avgSentenceLen, font) + 2–3 kalimat asli dokumen sebagai few-shot.
- `glossary` sebagai istilah wajib dipakai konsisten.
- Instruksi larangan klise generik + larangan emoji/heading baru tanpa diminta.

### 6.2 Minimal-diff (AI output operasi, bukan blob)

AI **dilarang** mengembalikan 1 halaman teks untuk `Replace` penuh. AI mengembalikan:

```json
{
  "ops": [
    { "tool": "ganti_kata", "paragraf": 4, "dari": "efisien", "ke": "hemat-token" },
    { "tool": "terapkan_style", "range": "seleksi", "style": "Heading 2" },
    { "tool": "sisip_setelah", "anchor": "bab2-akhir", "teks": "...2 paragraf...", "pertahankanStyle": true }
  ],
  "alasan_singkat": "rapikan tanpa ubah makna"
}
```

Eksekutor (`WordOps`) yang menerjemahkan ke `Word.run()` agar style, comment, dan track-changes asli tidak hancur.

### 6.3 Preserve-formatting pipeline

1. `range.load("text,font,style")` sebelum edit.
2. AI hanya kirim teks + delta style.
3. Terapkan delta ke **range yang sama** (bukan buat paragraf baru).
4. Tabel/gambar: AI kirim spec → builder buat via Word API native (ikut theme dokumen).

### 6.4 Preview + Undo contract

- Aksi >1 paragraf: tampilkan diff sebelum/sesudah di widget + tombol `Terapkan / Batal`.
- Simpan `backupRange` untuk undo manual.
- Operasi destruktif (hapus section, replace massal) wajib konfirmasi eksplisit.

### 6.5 Prompt pack (inti)

**System prompt operator (disimpan sekali):**

```
Kamu operator Word, bukan penulis esai.
1. Jangan ubah makna; perbaiki ejaan/alur minimal.
2. Gunakan istilah glosarium; tiru panjang kalimat sekitar (±18 kata).
3. Dilarang: pembuka generik ("Dalam era digital..."), emoji, heading baru tanpa diminta.
4. Pertahankan bahasa dokumen (id). Suhu rendah, presisi tinggi.
5. Output HANYA JSON {ops:[...], alasan_singkat:"..."}.
```

**Template tugas:**

- `edit`: L1 + profil gaya → ops ganti/sisip.
- `lanjutkan`: L1 + 2 kalimat terakhir + tone → sisip setelah kursor.
- `format/desain`: L1 struktur + styleInventory → terapkan style/tabel.

---

## 7. Daftar 12 Tools Atomik (WordOps)

| # | Tool (nama fungsi AI) | Pemetaan Word.js | Catatan token |
|---|----------------------|------------------|---------------|
| 1 | `baca_seleksi` | `getSelection().load("text")` | input L1 |
| 2 | `baca_jendela` | seleksi ±3 paragraf + heading | maks 4000 char |
| 3 | `baca_outline` | iterasi `body.paragraphs` filter Heading | murah, cacheable |
| 4 | `tulis_ganti` | `range.insertText(t, "Replace")` | hanya range kecil |
| 5 | `sisip_setelah/sebelum` | `insertParagraph / insertText("After"/"Before")` | pertahankan style |
| 6 | `terapkan_style` | `range.style = "Heading 2"` / font ops | dari inventory |
| 7 | `buat_tabel` | `insertTable(r,c)` + isi + header bold | spec, bukan narasi |
| 8 | `cari_ganti` | `body.search(term)` loop replace | batch, konfirmasi |
| 9 | `tambah_komentar` | `insertComment` / content control note | non-destruktif |
| 10 | `ringkas_section` | baca section → 2 kalimat → tulis ke manifest | untuk rebuild cache |
| 11 | `cek_konsistensi` | scan istilah glosarium vs body | agregasi client |
| 12 | `sisip_gambar` | `insertInlinePictureFromBase64` | dari URL/base64 user |

Setiap tool = 1 fungsi `Word.run()` kecil yang bisa diuji mandiri. Satu `demo()` assert-based per modul non-trivial.

---

## 8. Pengaturan Kelola AI (Settings)

Minimal untuk MVP (jangan over-build):

- `provider`: [OpenAI | Gemini]
- `apiKey`: password field, simpan `localStorage`, tombol `Test Connection` + `Hapus`
- `model`: dropdown (`gpt-4o-mini`, `gpt-4o`, `gemini-2.0-flash`, `gemini-1.5-pro`) + input custom
- `temperature`: slider 0–1 (default 0.3)
- `bahasa output`: [ikut dokumen | id | en]
- `budget`: batas char L1 + toggle "selalu preview sebelum terapkan"

Peringatan keamanan di UI: *"Key tersimpan lokal di perangkat ini (BYOK). Untuk enterprise, gunakan backend proxy agar key tidak terekspos via DevTools."*

---

## 9. UX Chat Widget (bukan floating bebas)

Keterbatasan jujur: sandbox hanya izinkan Ribbon + Task Pane + Dialog. Desain agar terasa seperti widget:

- Lebar default 330px (min Web Word), collapsible ke 86px strip ikon di Windows.
- Ribbon `🤖 DocAssist` → `showAsTaskpane()`; sideload bisa auto-open.
- Komponen: header mini (status cache + settings gear), area bubble, chips cepat (`Rapikan`, `Lanjutkan`, `Buat tabel`, `Formal-kan`), input + tombol kirim, panel preview diff.
- Jangan tiru Copilot penuh; fokus 4 chips operasional tercepat.

---

## 10. Risiko & Mitigasi

| Risiko | Dampak | Mitigasi |
|--------|--------|----------|
| AI halusinasi / rusak format | Tinggi | minimal-diff + preview + backupRange + suhu rendah |
| Token jebol (doc 100 hlm) | Biaya/lambat | layered context + manifest + scan per-heading |
| Key BYOK bocor via DevTools | Sedang | edukasi + peringatan; roadmap backend proxy enterprise |
| Beda perilaku Word Web vs Desktop | Sedang | uji keduanya; hindari API desktop-only; feature-detect requirement set |
| Penolakan Marketplace kelak | Rendah (MVP sideload) | patuhi `showAsTaskpane` manual-click path sejak awal |
| Privasi dokumen via API publik | Sedang | kirim potongan minimal; nyatakan di panduan; opsi Gemini/OpenAI sesuai pilihan user |

---

## 11. Roadmap MVP → v1

**Fase 0 — Scaffold (0.5 hari):** `yo office` Word TaskPane TS, manifest `ReadWriteDocument`, hosting statis, sideload lokal berhasil.

**Fase 1 — Chat + BYOK (1–2 hari):** UI chat mini, Settings, `AIClient` OpenAI+Gemini, Test Connection.

**Fase 2 — WordOps inti (2–3 hari):** 8 tools pertama (baca/tulis/style/tabel/cari-ganti/komentar), tiap tool ada `demo()` assert.

**Fase 3 — Konteks cerdas (2 hari):** ContextBuilder L0/L1, ManifestManager + CustomXmlParts, badge cache + refresh.

**Fase 4 — Anti-slop (1–2 hari):** PromptPack, preview diff, glosarium + styleProfile grounding.

**Fase 5 — Panduan & uji (1 hari):** `PANDUAN_INSTALL.md` sideload, uji Word Windows + Web, dokumen sampel id 10 hlm.

**v1 (nanti):** backend proxy (enterprise), login, usage meter, publish internal catalog, opsi embeddings lokal.

---

## 12. Kriteria Selesai MVP (Definition of Done)

- [ ] Sideload manual berhasil di Word Windows + Word Web dari hosting gratis.
- [ ] Chat bisa: rapikan seleksi, lanjutkan 2 paragraf, buat tabel 3×3, terapkan Heading — tanpa merusak style sekitar.
- [ ] Badge cache tampil benar; edit kecil tidak rebuild penuh; refresh manual bekerja.
- [ ] Preview diff muncul untuk aksi >1 paragraf; Batal tidak mengubah dokumen.
- [ ] Key BYOK tersimpan lokal, Test Connection sukses untuk OpenAI dan Gemini.
- [ ] Panduan install manual bisa diikuti user awam dalam <10 menit.

---

## Lampiran A — Contoh Manifest Sideload (cuplikan)

```xml
<Permissions>ReadWriteDocument</Permissions>
<!-- Ribbon button DocAssist -> showAsTaskpane (shared runtime) -->
```

```json
// unified manifest (alternatif modern) — permissions:
"authorization": { "permissions": { "resourceSpecific": [{ "name": "Document.ReadWrite.User", "type": "Delegated" }] } }
```

## Lampiran B — Contoh Pertukaran Hemat Token

```
User: "rapikan paragraf ini"
Kirim: L0(300) + L1(1200) + styleProfile(150) + glossary(100) = ~1750 tk
Hemat vs full-doc 20 hlm (~15rb tk): ~88%.
AI balas: {"ops":[{"tool":"tulis_ganti","rangeId":"sel","teks":"..."}]}
Eksekusi: 1x Word.run(). Tanpa rebuild cache (delta <10%).
```

---

*Dokumen ini adalah blueprint. Implementasi mulai dari Fase 0. Setiap simplifikasi yang memotong corner nyata wajib diberi komentar `ponytail:` + batas + jalur upgrade.*
