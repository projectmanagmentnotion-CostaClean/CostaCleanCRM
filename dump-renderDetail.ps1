$L = Get-Content .\scripts.html
$start = 760
$end = 980
for($i=$start; $i -le $end; $i++){
  "{0,4}: {1}" -f $i, $L[$i-1]
}
