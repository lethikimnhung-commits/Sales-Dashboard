<#
  update.ps1 - Rebuild dashboard data from the latest Excel file.
  Usage:
    1) Save the new Excel file (keep the "Raw data" sheet structure).
    2) Open PowerShell in this folder and run:  ./update.ps1
       (or right-click > Run with PowerShell)
  It reads Raw data -> rebuilds data.js + cust.js -> commits & pushes to GitHub.
  GitHub Pages auto-rebuilds in about 1 minute.

  Options:
    -Excel "D:\path\file.xlsb"   use a different Excel file
    -NoPush                      rebuild data only, do not commit/push
#>
param(
  [string]$Excel = "C:\Users\mgr-q\OneDrive\2026\Sales Update\Sales Update Jun 2026.xlsb",
  [switch]$NoPush
)
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
Write-Host "Reading Excel: $Excel" -ForegroundColor Cyan
if (-not (Test-Path $Excel)) { Write-Host "Excel file NOT found. Use: ./update.ps1 -Excel 'D:\...\file.xlsb'" -ForegroundColor Red; exit 1 }

$app = New-Object -ComObject Excel.Application
$app.Visible = $false; $app.DisplayAlerts = $false
$wb = $app.Workbooks.Open($Excel)
$ws = $wb.Sheets["Raw data"]
$n = $ws.UsedRange.Rows.Count
Write-Host "Total rows: $n" -ForegroundColor Green

$qty=$ws.Range("J2:J$n").Value2; $rev=$ws.Range("M2:M$n").Value2
$brand=$ws.Range("P2:P$n").Value2; $team=$ws.Range("U2:U$n").Value2
$cat=$ws.Range("V2:V$n").Value2; $year=$ws.Range("X2:X$n").Value2
$mon=$ws.Range("Y2:Y$n").Value2; $ctype=$ws.Range("Z2:Z$n").Value2
$prof=$ws.Range("AC2:AC$n").Value2
$inv=$ws.Range("A2:A$n").Value2; $short=$ws.Range("T2:T$n").Value2; $name=$ws.Range("F2:F$n").Value2
$wb.Close($false); $app.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null

$rows = $qty.GetLength(0)
$mmap=@{'January'=1;'February'=2;'March'=3;'April'=4;'May'=5;'June'=6;'July'=7;'August'=8;'September'=9;'October'=10;'November'=11;'December'=12}

function Idx($map,$ref,$v){ if($map.ContainsKey($v)){return $map[$v]}; $i=$map.Count; $map[$v]=$i; $ref.Value+=$v; return $i }

# ---------- CUBE 1: measure cube (revenue/qty/profit) ----------
Write-Host "Building measure cube..." -ForegroundColor Cyan
$cube=@{}
for($i=1;$i -le $rows;$i++){
  $y=$year[$i,1]; if($null -eq $y){continue}; $y=[int]$y
  $m=$mmap[[string]$mon[$i,1]]; if(-not $m){$m=0}
  $t=[string]$team[$i,1]; if(-not $t){$t='Unknown'}
  $b=[string]$brand[$i,1]; if(-not $b){$b='Unknown'}
  $c=[string]$cat[$i,1]; if(-not $c){$c='Unknown'}
  $ct=[string]$ctype[$i,1]; if(-not $ct){$ct='Unknown'}
  $r=$rev[$i,1]; if($null -eq $r){$r=0}
  $q=$qty[$i,1]; if($null -eq $q){$q=0}
  $p=$prof[$i,1]; if($null -eq $p){$p=0}
  $key="$y`t$m`t$t`t$b`t$c`t$ct"
  if($cube.ContainsKey($key)){$o=$cube[$key];$o[0]+=$r;$o[1]+=$q;$o[2]+=$p}else{$cube[$key]=@([double]$r,[double]$q,[double]$p)}
}
$teams=@{};$tL=@();$brands=@{};$bL=@();$cats=@{};$cL=@();$ctypes=@{};$ctL=@();$years=@{};$yL=@()
$data=@()
foreach($k in $cube.Keys){
  $f=$k -split "`t"; $y=[int]$f[0]; $m=[int]$f[1]
  $ti=Idx $teams ([ref]$tL) $f[2]; $bi=Idx $brands ([ref]$bL) $f[3]
  $ci=Idx $cats ([ref]$cL) $f[4]; $cti=Idx $ctypes ([ref]$ctL) $f[5]
  if(-not $years.ContainsKey($y)){$years[$y]=$true;$yL+=$y}
  $o=$cube[$k]
  $data+= ,@($y,$m,$ti,$bi,$ci,$cti,[math]::Round($o[0]),[math]::Round($o[1],2),[math]::Round($o[2]))
}
$obj=[ordered]@{teams=$tL;brands=$bL;cats=$cL;ctypes=$ctL;years=($yL|Sort-Object);rows=$data}
$json1 = $obj | ConvertTo-Json -Depth 5 -Compress
"window.CUBE = $json1;" | Out-File "$root\data.js" -Encoding utf8
Write-Host ("  -> data.js ({0} buckets)" -f $data.Count) -ForegroundColor Green

# ---------- CUBE 2: customer cube (revenue + invoices) ----------
Write-Host "Building customer cube..." -ForegroundColor Cyan
$agg=@{};$invset=@{};$custMeta=@{}
for($i=1;$i -le $rows;$i++){
  $y=$year[$i,1]; if($null -eq $y){continue}; $y=[int]$y
  $c=[string]$short[$i,1]; if(-not $c -or $c -eq '0'){$c=[string]$name[$i,1]}
  if(-not $c){continue}
  $r=$rev[$i,1]; if($null -eq $r){$r=0}; $r=[double]$r
  $t=[string]$team[$i,1]; if(-not $t){$t='Unknown'}
  $ct=[string]$ctype[$i,1]; if(-not $ct){$ct='Unknown'}
  $key="$y`t$c"
  if($agg.ContainsKey($key)){$agg[$key]+=$r}else{$agg[$key]=$r}
  $ik="$key`t$([string]$inv[$i,1])"; if(-not $invset.ContainsKey($ik)){$invset[$ik]=$true}
  if(-not $custMeta.ContainsKey($c)){$custMeta[$c]=@{}}
  $cm=$custMeta[$c]; $tk="T:$t"; if($cm.ContainsKey($tk)){$cm[$tk]+=$r}else{$cm[$tk]=$r}
  $ck="C:$ct"; if($cm.ContainsKey($ck)){$cm[$ck]+=$r}else{$cm[$ck]=$r}
}
$invcount=@{}
foreach($ik in $invset.Keys){$p=$ik.Substring(0,$ik.LastIndexOf("`t"));if($invcount.ContainsKey($p)){$invcount[$p]++}else{$invcount[$p]=1}}
$custTeam=@{};$custCt=@{}
foreach($c in $custMeta.Keys){
  $cm=$custMeta[$c]
  $custTeam[$c]=($cm.GetEnumerator()|Where-Object{$_.Key -like 'T:*'}|Sort-Object Value -Descending|Select-Object -First 1).Key.Substring(2)
  $custCt[$c]=($cm.GetEnumerator()|Where-Object{$_.Key -like 'C:*'}|Sort-Object Value -Descending|Select-Object -First 1).Key.Substring(2)
}
$teams2=@{};$tL2=@();$ctypes2=@{};$ctL2=@();$custs=@{};$cuL=@()
$data2=@()
foreach($k in $agg.Keys){
  $f=$k -split "`t"; $y=[int]$f[0]; $c=$f[1]
  $ci=Idx $custs ([ref]$cuL) $c; $ti=Idx $teams2 ([ref]$tL2) $custTeam[$c]; $cti=Idx $ctypes2 ([ref]$ctL2) $custCt[$c]
  $ic=if($invcount.ContainsKey($k)){$invcount[$k]}else{0}
  $data2+= ,@($y,$ci,$ti,$cti,[math]::Round($agg[$k]),$ic)
}
$obj2=[ordered]@{custs=$cuL;teams=$tL2;ctypes=$ctL2;rows=$data2}
$json2 = $obj2 | ConvertTo-Json -Depth 5 -Compress
"window.CUST = $json2;" | Out-File "$root\cust.js" -Encoding utf8
Write-Host ("  -> cust.js ({0} buckets, {1} customers)" -f $data2.Count, $cuL.Count) -ForegroundColor Green

# ---------- Commit & push ----------
if ($NoPush) { Write-Host "Data rebuilt. Skipped push (-NoPush)." -ForegroundColor Yellow; exit 0 }
Write-Host "Commit & push to GitHub..." -ForegroundColor Cyan
git -C $root add data.js cust.js
$stamp = Get-Date -Format "yyyy-MM-dd HH:mm"
git -C $root commit -m "Update dashboard data ($stamp)" 2>&1 | Out-Null
git -C $root push origin main 2>&1 | Out-Null
Write-Host "DONE! GitHub Pages will rebuild in ~1 minute." -ForegroundColor Green
Write-Host "Link: https://lethikimnhung-commits.github.io/Sales-Dashboard/" -ForegroundColor Green
