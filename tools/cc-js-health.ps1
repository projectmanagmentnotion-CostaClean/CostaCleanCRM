param(
  [string]$Path = ".\scripts.html",
  [string]$Tmp  = ".\_tmp_scripts_check.js",
  [string]$Out  = ".\_diagnostics\js-health.txt",
  [int]$Context = 12
)

function Ensure-Dir([string]$filePath){
  $dir = Split-Path -Parent $filePath
  if($dir -and !(Test-Path $dir)){
    New-Item -ItemType Directory -Force $dir | Out-Null
  }
}

function Write-Report([string]$text){
  Ensure-Dir $Out
  Set-Content -Encoding UTF8 $Out $text
  Write-Host "REPORT => $Out" -ForegroundColor Cyan
  try { notepad $Out } catch {}
}

function Extract-JS([string]$srcPath, [string]$tmpPath){
  $all = Get-Content -Encoding UTF8 $srcPath
  $js  = $all | Where-Object { $_ -notmatch "^\s*<script>\s*$" -and $_ -notmatch "^\s*</script>\s*$" }
  Set-Content -Encoding UTF8 $tmpPath $js
}

function Find-StrayText([string[]]$lines){
  $bad = @()
  for($i=0; $i -lt $lines.Count; $i++){
    $s = $lines[$i]
    if($s -match "^\s*$")  { continue }
    if($s -match "^\s*//") { continue }
    if($s -match "^\s*/\*"){ continue }
    if($s -match "^\s*\*") { continue }
    if($s -match "^\s*<")  { continue } # HTML

    $t = $s.Trim()

    # Ignorar líneas válidas típicas dentro de objetos/arrays:
    # - properties: "key: value"
    # - items con coma final "debug," / "'x'," / "123,"
    if($t -match "^[A-Za-z_$][\w$]*\s*:"){ continue }
    if($t -match "^(debug|mock)\s*,\s*$"){ continue }
    if($t -match "^[^;]*,\s*$"){ 
      # si termina en coma y NO empieza con '#/' ni '-' es probablemente item/property
      if($t -notmatch "^(#/|-)") { continue }
    }

    # Reportar patrones peligrosos que YA vimos que rompen JS:
    if($t -match "^\s*#\/"){ $bad += [pscustomobject]@{ Line=($i+1); Text=$t }; continue }
    if($t -match "^\s*-\s"){ $bad += [pscustomobject]@{ Line=($i+1); Text=$t }; continue }

    # Heurística general: texto titulo con letras y sin tokens JS
    if(
      $t -match "^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ].*$" -and
      $t -notmatch "[;=(){}\[\]]" -and
      $t -notmatch "^\s*(function|const|let|var|if|for|while|return|try|catch|switch|case|break|default|throw|class)\b"
    ){
      $bad += [pscustomobject]@{ Line=($i+1); Text=$t }
    }
  }
  return $bad
}

try{
  if(!(Test-Path $Path)){ throw "No existe: $Path" }

  $srcLines = Get-Content -Encoding UTF8 $Path
  $stray = Find-StrayText $srcLines

  Extract-JS $Path $Tmp
  $outText = (& node --check $Tmp 2>&1 | Out-String)

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine("JS HEALTH REPORT")
  [void]$sb.AppendLine("Src:  $Path")
  [void]$sb.AppendLine("Tmp:  $Tmp")
  [void]$sb.AppendLine("")

  if($stray.Count -gt 0){
    [void]$sb.AppendLine("WARN: Texto suelto potencialmente peligroso (no JS) => " + $stray.Count)
    [void]$sb.AppendLine("----- first 25 -----")
    $stray | Select-Object -First 25 | ForEach-Object {
      [void]$sb.AppendLine(("{0,5}: {1}" -f $_.Line, $_.Text))
    }
    [void]$sb.AppendLine("")
  } else {
    [void]$sb.AppendLine("OK: No se detectó texto suelto peligroso.")
    [void]$sb.AppendLine("")
  }

  if($outText -notmatch "SyntaxError"){
    [void]$sb.AppendLine("OK: node --check no reporta errores.")
    Write-Report $sb.ToString()
    exit 0
  }

  $m = [regex]::Match($outText, "^(?<file>[A-Za-z]:\\.*?):(?<line>\d+)\s*$", "Multiline")
  $file = if($m.Success){ $m.Groups["file"].Value } else { $Tmp }
  $line = if($m.Success){ [int]$m.Groups["line"].Value } else { 1 }

  $lines = Get-Content -Encoding UTF8 $file
  $from = [Math]::Max(1, $line - $Context)
  $to   = [Math]::Min($lines.Count, $line + $Context)

  [void]$sb.AppendLine("ERROR: node --check encontró SyntaxError")
  [void]$sb.AppendLine("File: $file")
  [void]$sb.AppendLine("Line: $line")
  [void]$sb.AppendLine("")
  [void]$sb.AppendLine("---- node output ----")
  [void]$sb.AppendLine($outText.TrimEnd())
  [void]$sb.AppendLine("")
  [void]$sb.AppendLine("---- context ----")
  for($i=$from; $i -le $to; $i++){
    $prefix = if($i -eq $line) { ">>" } else { "  " }
    [void]$sb.AppendLine(("{0}{1,5}: {2}" -f $prefix, $i, $lines[$i-1]))
  }

  Write-Report $sb.ToString()
  exit 1
}
catch{
  Write-Report ("FATAL `r`n" + $_.Exception.Message)
  exit 2
}
