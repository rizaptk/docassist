# 🤖 DocAssist — AI Asisten Operasional MS Word

Add-on MS Word ringan yang menjadi **asisten operasional dokumen**: format, buat, ubah, edit, tambah, dan desain dokumen langsung dari **chat widget kecil** yang tidak mengganggu kerja.

- 🆓 **Gratis**: tanpa server, hosting statis (GitHub Pages/Netlify/Vercel).
- 🔑 **BYOK multi-provider**: Gemini, Groq, NVIDIA NIM, OpenRouter :free, Cerebras, Mistral, Hugging Face, Pollinations (tanpa daftar), Ollama lokal — plus OpenAI/DeepSeek/Claude berbayar. Daftar di `config/providers.json`, user tinggal pilih + isi key.
- 🪶 **Ringan**: tanpa framework, tanpa build step — HTML/CSS/JS murni + Office.js.
- 🧠 **Hemat token**: konteks berlapis + cache ringkasan AI (DocManifest).
- 🛡️ **Anti-slop**: AI mengembalikan operasi minimal-diff + pratinjau sebelum diterapkan.

## ✨ Fitur MVP

| Fitur | Keterangan |
|-------|-----------|
| Chat widget 330px | Docked task pane, bisa dilipat jadi strip ikon 48px |
| Quick chips | Rapikan, Lanjutkan, Tabel, Formal — satu klik |
| BYOK OpenAI + Gemini | Dropdown provider/model, slider suhu, Test Connection |
| 12 tools Word atomik | Baca/tulis seleksi, style, tabel, cari-ganti, komentar |
| DocManifest cache | Ringkasan AI ikut tersimpan di `.docx` (CustomXmlParts) |
| Preview diff | Setiap aksi besar tampil sebelum → sesudah + Terapkan/Batal |
| Streaming responsif | Token mengalir bertahap, bisa dibatalkan, Word tetap bisa diketik |

## 🏗️ Arsitektur

```
Word Task Pane (statis, gratis)
 ├── Chat UI + Settings (BYOK)
 ├── WordBridge: 12 tools → Word.js API
 ├── ContextBuilder: konteks berlapis L0–L2 (hemat 70–90% token)
 ├── ManifestManager: cache CustomXmlParts + localStorage
 └── AIClient: fetch langsung ke OpenAI/Gemini (tanpa backend)
```

Detail lengkap: [`DOCASSIST_PLAN.md`](./DOCASSIST_PLAN.md)

## 🚀 Mulai Cepat (10 menit)

### 1. Host file statis

Upload repo ini ke GitHub Pages / Netlify / Vercel (HTTPS wajib), misal
`https://<user>.github.io/docassist/`, lalu ganti semua `~HOST~` di
[`manifest.xml`](./manifest.xml) dengan domain tersebut.

### 2. Sideload ke Word

- **Word Desktop**: Insert → Add-ins → Advanced → Upload My Add-in → pilih `manifest.xml`.
- **Word Web**: Insert → Add-ins → Advanced → Upload My Add-in → pilih `manifest.xml`.
- Grup **DocAssist** muncul di Ribbon Home → klik **AI Asisten**.

### 3. Isi API key

Klik **⚙** di panel → pilih provider → tempel key → pilih model murah dulu
(`gpt-4o-mini` / `gemini-2.0-flash`) → **Test**.

> Key hanya tersimpan di `localStorage` perangkat ini. Untuk enterprise,
> gunakan backend proxy (roadmap v1).

Panduan lengkap: [`docs/PANDUAN_INSTALL.md`](./docs/PANDUAN_INSTALL.md)

## 📁 Struktur Proyek

```
├── manifest.xml              # sideload manual (ReadWriteDocument)
├── src/
│   ├── taskpane/             # taskpane.html/.css/.js, commands.html
│   └── core/                 # WordOps, ContextBuilder, ManifestManager, AIClient, PromptPack
├── docs/                     # panduan install
└── DOCASSIST_PLAN.md         # blueprint perencanaan
```

## ⚡ Catatan Performance

- 1× `Word.run()` per aksi; `load` properti minimal; walk ±3 paragraf (tanpa load seluruh body).
- Kerja berat di `requestIdleCallback`; background refresh pause saat panel collapsed/hidden.
- AI client lazy-import; streaming via rAF-throttle; DOM chat dibatasi 50 bubble.
- Tulis cache debounced di idle — mengetik di Word tidak pernah diblokir.

## 🗺️ Roadmap

- [x] MVP: chat + BYOK + 12 tools + cache + anti-slop
- [ ] Backend proxy enterprise (key aman, usage meter)
- [ ] Katalog deploy internal / publish Marketplace
- [ ] Embeddings lokal untuk search semantik full-doc

## 📄 Lisensi

MIT — bebas dipakai dan dimodifikasi.
