# 🤖 DocAssist — AI Asisten Operasional MS Word

Add-on MS Word ringan yang menjadi **asisten operasional dokumen**: format, buat, ubah, edit, tambah, dan desain dokumen langsung dari **chat widget kecil** yang tidak mengganggu kerja.

- 🆓 **Gratis**: tanpa server, hosting statis. Live di `https://rizaptk.github.io/docassist/`.
- 🔑 **BYOK multi-provider**: Gemini, Groq, NVIDIA NIM, OpenRouter :free, Cerebras, Mistral, Hugging Face, Pollinations (tanpa daftar), Ollama lokal — plus OpenAI/DeepSeek/Claude berbayar. Daftar di `config/providers.json`, user tinggal pilih + isi key. Daftar model bisa di-browse langsung dari key.
- 🪶 **Ringan**: tanpa framework, tanpa build step — HTML/CSS/JS murni + Office.js.
- 🧠 **Hemat token**: konteks berlapis + cache ringkasan AI (DocManifest).
- 🛡️ **Anti-slop**: AI mengembalikan operasi minimal-diff + pratinjau sebelum diterapkan.

## ✨ Fitur

| Fitur | Keterangan |
|-------|-----------|
| Chat widget 330px | Docked task pane, bisa dilipat jadi strip ikon 48px |
| Quick chips | Rapikan, Lanjutkan, Tabel, Formal, Flowchart ✎, Chart, CV 📄, Cek ATS — satu klik |
| BYOK multi-provider | 12 provider + custom, browse model dari key, Test koneksi |
| 19 tools Word atomik | Teks, tabel, visual, flowchart editable, CV, cover letter, cek ATS |
| Template kits | Registry `config/templates.json` + builder (CV ATS/modern, cover letter) |
| Flowchart editable | Shapes DrawingML natif (teks bisa diedit di Word), fallback gambar otomatis |
| Chart dari data | Paste tabel/angka → grafik batang/garis + tabel sumber |
| DocManifest cache | Ringkasan AI ikut tersimpan di `.docx` (CustomXmlParts) |
| Preview diff | Setiap aksi besar tampil sebelum → sesudah + Terapkan/Batal |
| Streaming responsif | Token mengalir bertahap, bisa dibatalkan, Word tetap bisa diketik |

## 🚀 Cara Menggunakan (tahap demi tahap)

### Tahap 0 — Persiapan (sekali saja)

- Word versi Microsoft 365 (Desktop Windows/Mac) atau Word Web.
- Satu API key gratis, misal dari [Google AI Studio](https://aistudio.google.com/apikey) (Gemini, ±1500 req/hari) atau [Groq Console](https://console.groq.com/keys) (cepat, ±1K req/hari). Tanpa kartu kredit.

### Tahap 1 — Install add-in ke Word (sekali per perangkat)

1. Download file [`manifest.xml`](https://github.com/rizaptk/docassist/blob/master/manifest.xml) dari repo (tombol `···` → Download, atau `git clone`).
2. Buka Word → **Insert → Add-ins → Advanced → Upload My Add-in** (Word baru: tab **Developer → Add-ins → Upload**).
3. Pilih `manifest.xml` → **Upload**.
4. Di Ribbon **Home** muncul grup **DocAssist** → klik **AI Asisten** untuk membuka panel chat. Klik **–** untuk melipat jadi strip ikon, **×** untuk menyembunyikan total.

> Jika panel kosong: pastikan perangkat online (add-in dimuat dari GitHub Pages) dan ulangi Upload.

### Tahap 2 — Sambungkan AI (sekali per perangkat)

1. Di panel klik **⚙** → pilih **Provider** (mis. `Google Gemini`, `Groq`, atau `OpenRouter :free`).
2. Tempel **API key** → kolom **Model** otomatis terisi daftar model dari akunmu (klik untuk pilih, atau ketik manual).
3. Klik **Test** → "Koneksi OK." berarti siap.
4. (Opsional) Provider lain: pilih dari daftar, atau tambah sendiri via **+ Provider lain** (nama + base URL + gaya API) — tersimpan di perangkat ini.

> Key hanya tersimpan di `localStorage` perangkat ini, tidak dikirim ke mana pun selain provider yang kamu pilih.

### Tahap 3 — Pakai sehari-hari

**Edit/format dokumen aktif** — seleksi teks (atau letakkan kursor), lalu:
- Chip **Rapikan / Formal / Lanjutkan**, atau ketik bebas: `rapikan paragraf ini`, `lanjutkan 2 paragraf dengan nada yang sama`, `ganti kata X jadi Y di seluruh dokumen`.
- Aksi besar (>1 paragraf) selalu muncul di **Pratinjau** → **Terapkan** / **Batal**.

**Flowchart** — klik `Flowchart ✎`, deskripsikan alur (`Buatkan flowchart yang bisa diedit dari alur berikut: daftar → verifikasi → ...`). Hasil = shapes asli Word (teks tiap kotak bisa diklik-edit, dipindah, di-resize). Pratinjau gambar dulu, Terapkan untuk sisip.

**Chart dari data** — klik `Chart`, paste tabel/angka setelah titik dua. Hasil = gambar grafik + tabel sumbernya ikut disisip.

**CV & lamaran** — klik `CV 📄`, paste data diri (nama, kontak, pengalaman, skill) → jadi CV ATS-friendly. Chip `Cek ATS` mengaudit dokumen aktif (gambar/textbox/tabel/kolom) dan memberi saran — gratis, tanpa token AI.

**Ganti gaya panel** — badge `Konteks: cached ✓` artinya AI ingat ringkasan dokumen; klik **Refresh konteks** (di ⚙) setelah edit besar.

### Tahap 4 — Jika ada masalah

| Gejala | Solusi |
|---|---|
| Panel blank | Cek internet; Upload ulang manifest; reload panel (tombol panah bawah → Reload) |
| `Key ditolak` | Salah provider vs key (`sk-…` OpenAI, `AIza…` Gemini, `gsk_…` Groq, `nvapi-…` NVIDIA); Test ulang |
| `Koneksi gagal (429)` | Limit gratis tercapai — ganti provider/model atau tunggu reset harian |
| Flowchart jadi gambar | Mode editable gagal di dokumen itu (lihat console) — tetap bisa dipakai, atau regenerate |
| Word Web beda perilaku | Page-setup/shapes natif = Desktop saja; konten teks/tabel/chart selalu jalan di semua platform |
| Model tidak ada di daftar | Ketik manual ID-nya di kolom Model (daftar `:free` OpenRouter berputar) |

Panduan instal detail: [`docs/PANDUAN_INSTALL.md`](./docs/PANDUAN_INSTALL.md)

## 🏗️ Arsitektur

```
Word Task Pane (statis, GitHub Pages)
 ├── Chat UI + Settings (BYOK, registry provider dinamis)
 ├── WordOps: 19 tools → Word.js API (+ OOXML untuk shapes)
 ├── Builders/: CvBuilder, (kit lain menyusul) — data → ops primitif
 ├── FlowchartOoxml: layout + DrawingML editable
 ├── Visuals: Mermaid→PNG, chart canvas murni
 ├── ContextBuilder: konteks berlapis L0–L2 (hemat 70–90% token)
 ├── ManifestManager: cache CustomXmlParts + localStorage
 └── AIClient: registry config/providers.json, 3 gaya API, browse model dari key
```

Dokumen perencanaan: [`DOCASSIST_PLAN.md`](./DOCASSIST_PLAN.md) · Riset tools: [`docs/TOOLS_RESEARCH.md`](./docs/TOOLS_RESEARCH.md)

## 📁 Struktur Proyek

```
├── manifest.xml              # sideload (ReadWriteDocument), host = GitHub Pages
├── config/
│   ├── providers.json        # registry provider AI (preset ikut git)
│   └── templates.json        # registry kit dokumen
├── src/
│   ├── taskpane/             # taskpane.html/.css/.js, commands.html
│   └── core/                 # WordOps, AIClient, ContextBuilder, ManifestManager,
│                             # PromptPack, Visuals, FlowchartOoxml, Builders/
├── scripts/smoke.mjs         # uji logika murni (node scripts/smoke.mjs)
├── docs/                     # panduan install + riset tools
├── DOCASSIST_PLAN.md         # blueprint
└── package.json              # type:module (untuk smoke test; tanpa build step)
```

## ⚡ Catatan Performance

- 1× `Word.run()` per aksi; `load` properti minimal; walk ±3 paragraf (tanpa load seluruh body).
- Kerja berat di `requestIdleCallback`; background refresh pause saat panel collapsed/hidden.
- AI client lazy-import; streaming via rAF-throttle; DOM chat dibatasi 50 bubble.
- Tulis cache debounced di idle — mengetik di Word tidak pernah diblokir.

## 🗺️ Roadmap

- [x] MVP: chat + BYOK + cache + anti-slop
- [x] Visual: flowchart/chart + flowchart editable + CV generator + ATS-check
- [x] Provider registry + free tiers + browse model dari key
- [ ] P1 Article Enhancer (tulis/enhance artikel, abstrak, heading otomatis)
- [ ] P2 Pitch-doc Word + Promo builder (flyer/brosur)
- [ ] Backend proxy enterprise (key aman, usage meter)
- [ ] Publish Marketplace / katalog internal

## 📄 Lisensi

MIT — bebas dipakai dan dimodifikasi.
