# ──────────────────────────────────────────────────────────────
# VIP 핫딜 대시보드 — 원본 xlsx → data/*.csv 변환기
#
#  회사 DRM(Softcamp)이 디스크에 쓰는 파일을 자동 암호화하기 때문에
#  Excel "다른 이름으로 저장(CSV)"은 깨집니다. 이 스크립트는
#  Excel COM으로 메모리에서 값을 읽어, .NET File.WriteAllText로
#  (DRM 안 걸리는 경로) 깨끗한 UTF-8 CSV를 만듭니다.
#
#  사용법:  PowerShell에서  ./convert.ps1
#  (Downloads 에 핫딜.xlsx, Table.xlsx 가 있다고 가정. 경로는 아래 param 수정)
# ──────────────────────────────────────────────────────────────
param(
  [string]$Hotdeal = "$env:USERPROFILE\Downloads\핫딜.xlsx",
  [string]$Table   = "$env:USERPROFILE\Downloads\Table.xlsx"
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$dataDir = Join-Path $root "data"
if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir | Out-Null }
$enc = [System.Text.UTF8Encoding]::new($false)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false

function Esc($s) { $s = [string]$s; if ($s -match '[",\n]') { '"' + ($s -replace '"', '""') + '"' } else { $s } }

# ── 1) Table.xlsx → data/table_trend.csv (영역별 일별 UV/PV) ──
Write-Host "▶ Table.xlsx 변환 중..."
$wb = $excel.Workbooks.Open($Table)
$ws = $wb.Worksheets.Item(1)
$nCols = $ws.UsedRange.Columns.Count
$data = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(9, $nCols)).Value2
$wb.Close($false)

$rawDates = @(); for ($c = 3; $c -le $nCols; $c++) { $rawDates += [string]$data[1, $c] }
$md = $rawDates | ForEach-Object { $p = $_ -split '/'; [pscustomobject]@{ m = [int]$p[0]; d = [int]$p[1] } }
# 마지막 날짜를 '가장 최근'으로 보고 롤오버 수만큼 역산해 연도 부여
$roll = 0; $prevM = $md[0].m
foreach ($x in $md) { if ($x.m -lt $prevM) { $roll++ }; $prevM = $x.m }
$endYear = (Get-Date).Year
$startYear = $endYear - $roll
$years = @(); $y = $startYear; $prevM = $md[0].m
foreach ($x in $md) { if ($x.m -lt $prevM) { $y++ }; $years += $y; $prevM = $x.m }

$rows = @{ 2 = 'UV'; 3 = 'UV'; 4 = 'UV'; 5 = 'UV'; 6 = 'PV'; 7 = 'PV'; 8 = 'PV'; 9 = 'PV' }
$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine("date,metric,area,value")
for ($r = 2; $r -le 9; $r++) {
  $metric = $rows[$r]; $area = [string]$data[$r, 2]
  for ($i = 0; $i -lt $md.Count; $i++) {
    $dt = "{0:0000}-{1:00}-{2:00}" -f $years[$i], $md[$i].m, $md[$i].d
    $val = $data[$r, ($i + 3)]; if ($null -eq $val) { $val = '' }
    [void]$sb.AppendLine("$dt,$metric,$area,$val")
  }
}
[System.IO.File]::WriteAllText((Join-Path $dataDir "table_trend.csv"), $sb.ToString(), $enc)
Write-Host ("  ✓ table_trend.csv  ({0} ~ {1})" -f ("{0:0000}-{1:00}-{2:00}" -f $years[0], $md[0].m, $md[0].d), ("{0:0000}-{1:00}-{2:00}" -f $years[-1], $md[-1].m, $md[-1].d))

# ── 2) 핫딜.xlsx → data/hotdeal.csv (슬롯·상품 단위 매출) ──
Write-Host "▶ 핫딜.xlsx 변환 중..."
$wb = $excel.Workbooks.Open($Hotdeal)
$ws = $wb.Worksheets.Item(1)
$nrow = $ws.Cells.Item($ws.Rows.Count, 3).End(-4162).Row   # xlUp on col3(영역상세명)
$data = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($nrow, 22)).Value2
$wb.Close($false)

# Total 행(일자 보유) 수집 후 연도 역산 (일자 포맷 M/D → 후반 YY/MM/DD 혼재)
$tot = New-Object System.Collections.ArrayList
for ($r = 2; $r -le $nrow; $r++) {
  $v = [string]$data[$r, 1]; if ($v.Trim() -eq '') { continue }
  if ($v -match '^(\d+)/(\d+)/(\d+)') { $null = $tot.Add([pscustomobject]@{ r = $r; m = [int]$Matches[2]; d = [int]$Matches[3]; y = 2000 + [int]$Matches[1] }) }
  elseif ($v -match '^(\d+)/(\d+)') { $null = $tot.Add([pscustomobject]@{ r = $r; m = [int]$Matches[1]; d = [int]$Matches[2]; y = $null }) }
}
$k = 0; while ($k -lt $tot.Count -and $null -eq $tot[$k].y) { $k++ }
for ($i = $k - 1; $i -ge 0; $i--) { $ny = $tot[$i + 1].y; if ($tot[$i].m -gt $tot[$i + 1].m) { $ny-- }; $tot[$i].y = $ny }
for ($i = $k + 1; $i -lt $tot.Count; $i++) { if ($null -eq $tot[$i].y) { $ny = $tot[$i - 1].y; if ($tot[$i].m -lt $tot[$i - 1].m) { $ny++ }; $tot[$i].y = $ny } }
$rowDate = @{}
foreach ($t in $tot) { $rowDate[$t.r] = ("{0:0000}-{1:00}-{2:00}" -f $t.y, $t.m, $t.d) }

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine("date,slot,row_type,prodcode,prodname,md,bpu,brand,category,UV,PV,cust,ord,qty,rev,h_UV,h_PV,h_cust,h_ord,h_qty,h_rev")
$curDate = $null
for ($r = 2; $r -le $nrow; $r++) {
  if ($rowDate.ContainsKey($r)) { $curDate = $rowDate[$r] }
  $detail = [string]$data[$r, 3]; if ($detail.Trim() -eq '') { continue }
  if ($detail -eq 'Total') { $slot = 'Total'; $rtype = 'TOTAL' }
  elseif ($detail -match '오전') { $slot = '오전'; $rtype = 'SLOT' }
  elseif ($detail -match '오후') { $slot = '오후'; $rtype = 'SLOT' }
  else { $slot = $detail; $rtype = 'SLOT' }
  $vals = @($curDate, $slot, $rtype, (Esc $data[$r, 5]), (Esc $data[$r, 6]), (Esc $data[$r, 7]), (Esc $data[$r, 8]), (Esc $data[$r, 9]), (Esc $data[$r, 10]))
  for ($c = 11; $c -le 22; $c++) { $x = $data[$r, $c]; if ($null -eq $x) { $x = '' }; $vals += [string]$x }
  [void]$sb.AppendLine(($vals -join ','))
}
[System.IO.File]::WriteAllText((Join-Path $dataDir "hotdeal.csv"), $sb.ToString(), $enc)
Write-Host ("  ✓ hotdeal.csv  ({0} ~ {1}, {2}일)" -f $rowDate[$tot[0].r], $rowDate[$tot[-1].r], $tot.Count)

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "`n완료. 이제 git add/commit/push 하면 Streamlit Cloud에 반영됩니다."
