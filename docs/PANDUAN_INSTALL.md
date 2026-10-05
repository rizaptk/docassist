# Panduan Install Manual DocAssist (Sideload) — <10 menit

Karena distribusi manual (bukan Marketplace), ikuti sekali saja per perangkat.

## 1. Hosting (gratis, sekali saja oleh admin)

1. Upload isi repo ini ke GitHub Pages / Netlify / Vercel dengan HTTPS. Misal `https://kamu.github.io/docassist/`.
2. Buka `manifest.xml`, ganti semua `~HOST~` dengan domain itu. Contoh: `https://kamu.github.io/docassist`.
3. (Opsional) Taruh 3 ikon di `assets/icon-16.png, icon-32.png, icon-80.png`. Bisa 1 file PNG disalin 3 nama.

## 2. Sideload di Word Windows

1. Word → **File > Options > Trust Center > Trusted Add-in Catalogs** → pastikan sideload diizinkan; atau langsung:
2. **Insert > Add-ins > Advanced > Upload My Add-in** (pada Word baru: **Developer > Add-ins > Upload**).
3. Pilih `manifest.xml` hasil edit tadi → **Upload**.
4. Ribbon **Home** muncul grup **DocAssist → AI Asisten**. Klik untuk buka panel 330px.
5. Klik `–` untuk lipat jadi strip 🤖 (tidak intrusif), `×` untuk sembunyikan total.

## 3. Sideload di Word Web (Microsoft 365)

1. Buka Word Web → **Insert > Add-ins > Advanced > Upload My Add-in** → pilih `manifest.xml`.
2. Panel kanan tampil. Sematkan via **... > Pin** bila tersedia agar tidak hilang.

## 4. Isi API key (BYOK)

1. Di panel klik **⚙** → Provider **OpenAI/Gemini** → tempel key → pilih model murah dulu (`gpt-4o-mini` / `gemini-2.0-flash`) → **Test**.
2. Key hanya di `localStorage` perangkat ini. Jangan bagikan manifest berisi key.

## 5. Uji cepat (verifikasi tidak mengganggu Word)

1. Ketik 2 paragraf Bahasa Indonesia → seleksi satu paragraf.
2. Chip **Rapikan** → harus muncul pratinjau diff → **Terapkan**. Dokumen tetap bisa diketik selama AI berpikir (streaming, abortable).
3. Lipat panel (`–`), ketik bebas — Word tidak lag (background hash pause saat collapsed/hidden).
4. Badge `Konteks: cached ✓` tampil; edit kecil tidak rebuild penuh.

## Troubleshooting

- Panel blank → pastikan URL `SourceLocation` HTTPS publik, bukan `localhost`, dan `AppDomain` mencakup host itu.
- `showAsTaskpane` tidak bereaksi di Marketplace build → wajar; sideload/internal saja mendukung auto-open.
- Error key 401 → cek provider vs key (sk-… untuk OpenAI, AIza… untuk Gemini) + model tersedia di akunmu.
- Word lemot → pastikan panel di-collapse saat tidak dipakai; background refresh tiap 15 dtk otomatis pause saat hidden.
