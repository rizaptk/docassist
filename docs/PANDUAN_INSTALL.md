# Panduan Install Manual DocAssist (Sideload) — <10 menit

Karena distribusi manual (bukan Marketplace), ikuti sekali saja per perangkat.

## 1. Hosting (gratis, sekali saja oleh admin)

1. Upload isi repo ini ke GitHub Pages / Netlify / Vercel dengan HTTPS. Misal `https://kamu.github.io/docassist/`.
2. Buka `manifest.xml`, ganti semua `~HOST~` dengan domain itu. Contoh: `https://kamu.github.io/docassist`.
3. (Opsional) Taruh 3 ikon di `assets/icon-16.png, icon-32.png, icon-80.png`. Bisa 1 file PNG disalin 3 nama.

## 2. Sideload di Word Windows

**Metode A — script sekali-klik (disarankan, tanpa admin, tanpa menu Upload):**

1. Download repo ini sebagai ZIP (GitHub → Code → Download ZIP) → ekstrak.
2. Klik kanan `scripts\install.ps1` → **Run with PowerShell** (bila ditolak: `powershell -ExecutionPolicy Bypass -File scripts\install.ps1`).
3. Tutup SEMUA Word → buka lagi → grup **DocAssist** di Ribbon Home.
4. Melepas: jalankan `scripts\uninstall.ps1`.

Script ini melakukan hal yang sama dengan Metode CLI di bawah (tulis `HKCU\...\WEF\Developer\<id>`,
registry per-user) — cocok dibagikan ke user lain beserta file repo.

**Metode B — CLI resmi Microsoft (setara script, untuk yang nyaman terminal):**

```powershell
npm install -g office-addin-dev-settings
office-addin-dev-settings register manifest.xml
```

Perintah ini menulis `HKCU\...\WEF\Developer\<id>` = path manifest (registry per-user, tanpa admin).
Tutup SEMUA jendela Word → buka lagi → grup **DocAssist** muncul di Ribbon Home
(atau Insert → My Add-ins). Untuk melepas: `office-addin-dev-settings unregister manifest.xml`.
Validasi manifest dulu bila ragu: `npx -y office-addin-manifest validate manifest.xml`.

**Metode C — dialog Upload (bila ada):**

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

## Lampiran: Word 2016/2019 atau tanpa menu Advanced/Upload

**Catatan konsep (dari docs Microsoft): "katalog" = FOLDER berisi file manifest, bukan file manifest itu sendiri.**
Satu folder katalog bisa menampung banyak `manifest.xml` (tiap file = 1 add-in di tab SHARED FOLDER).
Word memindai folder itu; aplikasi web-nya sendiri tetap diambil dari `<SourceLocation>` di tiap
manifest (punya kita sudah menunjuk ke GitHub Pages ✅). Format path wajib UNC (`\\NAMAPC\Folder`),
bukan `C:\...`. Cara ini hanya untuk testing di Windows (bukan produksi; manifest unified tidak didukung;
tiap ubah ribbon harus reinstall).

1. Di dialog Office Add-ins, klik tab **MY ADD-INS** (bukan STORE) → cari link **Upload My Add-in** di kanan atas → pilih `manifest.xml`. Ini cara tercepat bila link-nya ada.
2. Bila tidak ada (kasus Office Home & Student 2016): siapkan shared folder:
   - Buat folder mis. `C:\Addins`, salin **hanya** `manifest.xml` ke sana → klik kanan folder → Properties → Sharing → Share (setujui UAC) → catat alamat `\\NAMAPC\Addins`.
   - Word → **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs** → isi **Catalog Url** dengan `\\NAMAPC\Addins` → **Add catalog** → centang **Show in Menu** → OK → OK → **tutup SEMUA Word dan buka lagi**.
   - **Insert → My Add-ins → SHARED FOLDER** → pilih DocAssist → Add. (Alternatif tanpa klik: file `.reg` ke `HKCU\...\WEF\TrustedCatalogs\{GUID}` dengan `Url` + `Flags=1` — tanpa admin.)
3. Alternatif termudah: **Word Web** (office.com) → dokumen baru → Insert → Add-ins → Upload → pilih `manifest.xml`.
4. Keluar dari **Compatibility Mode** dulu: **File → Info → Convert** (atau Save As → `.docx`). Add-in paling stabil di format `.docx` modern.
5. **Word lawas menolak ribbon** (manifest terdaftar tapi tidak muncul di SHARED FOLDER, padahal manifest minimal muncul): gunakan **`manifest-legacy.xml`** — sama persis tanpa blok VersionOverrides/ribbon. Taskpane dibuka via Insert → My Add-ins → Add (bukan tombol ribbon). Terbukti jalan di Office Home & Student 2016 (Okt 2026).
