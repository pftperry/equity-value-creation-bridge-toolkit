$ErrorActionPreference = "Stop"
$d = $PSScriptRoot; $p = Join-Path $d "Equity_Value_Creation_Bridge_Toolkit.xlsx"; $ev = Join-Path $d "evidence"
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $true
try {
  $wb = $xl.Workbooks.Open($p); $xl.CalculateFull()
  $total = 0
  foreach ($ws in $wb.Worksheets) {
    $f = 0; $e = 0
    try { $f = $ws.UsedRange.SpecialCells(-4123).Count } catch {}
    try { $e = $ws.UsedRange.SpecialCells(-4123, 16).Count } catch {}
    "{0}: formula cells={1} error cells={2}" -f $ws.Name, $f, $e; $total += $e
  }
  "TOTAL ERROR CELLS=$total"
  $b = $wb.Worksheets.Item("Equity Value Creation Bridge"); $s = $wb.Worksheets.Item("Summary Recommendation"); $l = $wb.Worksheets.Item("Lever Impact Ranking"); $w = $wb.Worksheets.Item("Anonymized Worked Example")
  foreach ($r in 4..19) { "BR row {0}: {1} = {2} | {3}" -f $r, $b.Cells.Item($r,2).Text, $b.Cells.Item($r,4).Text, $b.Cells.Item($r,5).Text }
  foreach ($r in 14..20) { "LR row {0}: {1} = {2} | {3} | {4}" -f $r, $l.Cells.Item($r,1).Text, $l.Cells.Item($r,2).Text, $l.Cells.Item($r,3).Text, $l.Cells.Item($r,4).Text }
  foreach ($r in 16..28) { "SR row {0}: {1} = {2}" -f $r, $s.Cells.Item($r,1).Text, $s.Cells.Item($r,2).Text }
  foreach ($r in 26..30) { "WE row {0}: {1} | {2} | {3} | {4}" -f $r, $w.Cells.Item($r,1).Text, $w.Cells.Item($r,3).Text, $w.Cells.Item($r,4).Text, $w.Cells.Item($r,5).Text }
  "B12 formula: " + $b.Range("D12").Formula; "B27 formula: " + $s.Range("B27").Formula
  $wb.Save()
  function Shot($ws, $rng, $file, $formulas) {
    $ws.Activate(); $xl.ActiveWindow.DisplayFormulas = $formulas; $xl.ActiveWindow.ScrollRow = 1; $xl.ActiveWindow.ScrollColumn = 1
    $r = $ws.Range($rng); $r.CopyPicture(1, 2); Start-Sleep -Milliseconds 400
    $co = $ws.ChartObjects().Add(0, 0, $r.Width, $r.Height); $co.Activate() | Out-Null; $co.Chart.Paste(); Start-Sleep -Milliseconds 400
    $co.Chart.Export((Join-Path $ev $file)) | Out-Null; $co.Delete(); $xl.ActiveWindow.DisplayFormulas = $false; "exported $file"
  }
  Shot $b "A3:F19" "bridge_values_and_formulatext.png" $false
  Shot $s "A3:C29" "summary_recommendation.png" $false
  Shot $w "A3:F38" "worked_example_output.png" $false
  $co = $w.ChartObjects().Item(1); $co.Chart.Export((Join-Path $ev "worked_example_bridge_chart.png")) | Out-Null; "exported chart"
  $wb.Close($false)
} finally { $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }

