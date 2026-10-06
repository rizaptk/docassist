<#
  DocAssist installer — Windows, TANPA admin.
  Cara pakai: klik kanan > Run with PowerShell (atau: powershell -ExecutionPolicy Bypass -File install.ps1)
  Kerja: daftarkan manifest.xml ke HKCU\...\WEF\Developer (registry per-user),
  lalu minta restart Word. Untuk melepas: jalankan uninstall.ps1.
#>
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$manifest = Join-Path $root "manifest.xml"
if (-not (Test-Path -LiteralPath $manifest)) { throw "manifest.xml tidak ditemukan di: $manifest" }

[xml]$m = Get-Content -LiteralPath $manifest -Encoding UTF8
$id = $m.OfficeApp.Id
if (-not $id) { throw "Id tidak terbaca dari manifest.xml" }
$name = $m.OfficeApp.DisplayName.DefaultValue

$dev = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
if (-not (Test-Path -LiteralPath $dev)) { New-Item -Path $dev -Force | Out-Null }
New-ItemProperty -Path $dev -Name $id -Value $manifest -PropertyType String -Force | Out-Null

Write-Output "OK: '$name' terdaftar untuk Word di perangkat ini."
Write-Output "ID: $id"
Write-Output ""
Write-Output "Langkah terakhir (wajib):"
Write-Output "  1. Tutup SEMUA jendela Word."
Write-Output "  2. Buka Word > dokumen baru (.docx, bukan Compatibility Mode)."
Write-Output "  3. Cari grup DocAssist di tab Home (atau Insert > Add-ins)."
Write-Output ""
Write-Output "Melepas: jalankan uninstall.ps1"
