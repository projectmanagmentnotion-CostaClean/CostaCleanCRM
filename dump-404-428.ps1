$L = Get-Content .\scripts.html
$start = 404
$end = 428
for($i=$start; $i -le $end; $i++){
  "{0,4}: {1}" -f $i, $L[$i-1]
}
