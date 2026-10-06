<#
  DocAssist uninstaller — menghapus registrasi add-in (tanpa admin).
  Data chat/key user di Word TIDAK dihapus (tersimpan di dokumen & localStorage terpisah).
#>
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$manifest = Join-Path $root "manifest.xml"
$dev = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"

$id = $null
if (Test-Path -LiteralPath $manifest) {
  [xml]$m = Get-Content -LiteralPath $manifest -Encoding UTF8
  $id = $m.OfficeApp.Id
}
if ($id -and (Get-ItemProperty -Path $dev -Name $id -ErrorAction SilentlyContinue)) {
  Remove-ItemProperty -Path $dev -Name $id -Force
  Write-Output "OK: registrasi DocAssist dihapus. Restart Word untuk menerapkan."
} else {
  Write-Output "Tidak ada registrasi DocAssist di perangkat ini."
}
