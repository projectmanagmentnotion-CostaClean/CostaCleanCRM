$L = Get-Content .\WEBAPP_API.js

function Dump($pattern, $before=5, $after=120){
  $hit = (rg -n $pattern .\WEBAPP_API.js | Select-Object -First 1)
  if(-not $hit){ "NO ENCONTRÉ: $pattern"; return }
  $line = [int](($hit -split ":")[0])
  $start = [Math]::Max(1, $line - $before)
  $end = $line + $after
  "HIT: $hit"
  for($i=$start; $i -le $end; $i++){
    "{0,4}: {1}" -f $i, $L[$i-1]
  }
  ""
}

Dump "function _mapListResult_"
Dump "function _listFromView_"
