$path = ".\scripts.html"
$txt = Get-Content $path -Raw

$needle = "safeLog_('fetchEntityList '+entity, res);"

if ($txt.IndexOf($needle) -lt 0) {
  Write-Host "ERROR: No encuentro el needle exacto. No hice cambios."
  exit 1
}

$txt2 = $txt.Replace($needle, "safeLog_('fetchEntityList proformas', res);")

Set-Content -Path $path -Value $txt2 -Encoding UTF8
Write-Host "OK: Reemplacé safeLog_ con entity undefined (proformas fijo)."

# Verificación
$after = Get-Content $path -Raw
if ($after.IndexOf($needle) -ge 0) {
  Write-Host "ERROR: todavía existe el needle. Algo falló."
  exit 1
}
Write-Host "OK: ya no existe el safeLog_ roto."
