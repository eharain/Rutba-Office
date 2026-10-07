# Make the binary fixtures with Office itself: Word 97-2003 documents, and the
# showcase workbook and deck saved by Excel and PowerPoint in their 97-2003
# formats — and the workbook in every older Excel format this Excel still
# writes. The readers in packages/office-formats are judged against these,
# files Office wrote, not files written from what we think the formats say.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\make-binary-fixtures.ps1 [-Out tests\fixtures\binary] [-Only word,excel,charts,conditions,powerpoint,tables]
#
# Office is driven invisibly; each application is quit and the processes this
# run started are reaped at the end of its section, even when a step fails.
#
# Word's Save As hangs for ever in an automated session on this machine when
# the document was built inside a PowerShell function or script block; the
# same steps written at the top level save (see make-binary-word.ps1, and
# make-rich-fixtures.ps1, which met the hang first). So Word makes four
# small documents, each by make-binary-word.ps1 in a PowerShell of its own:
# the characters, the paragraphs and lists, and the structure (a table, a
# link, a footnote, a picture, a header and footer, a page break) — which
# also lets each reader test aim at one thing — and the floating drawings
# (a picture, a text box, two shapes) in a fourth.
#
# This Excel no longer writes the 2.1, 3.0 and 4.0 formats; those, and the
# Word 2.0, 6.0/95, DOS and Write formats no Office here writes, are built
# from their published layouts in the tests that read them.
param([string]$Out = 'tests\fixtures\binary', [string]$Only = 'word,excel,charts,conditions,powerpoint,tables')
$run = $Only.Split(',')

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$outDir = if ([IO.Path]::IsPathRooted($Out)) { $Out } else { Join-Path $root $Out }
$rich = Join-Path $root 'tests\fixtures\rich'
New-Item -ItemType Directory -Force $outDir | Out-Null
function RGB($r, $g, $b) { return $r + ($g * 256) + ($b * 65536) }
function Release($o) { if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} } }
function Remove-Existing($p) { if (Test-Path $p) { Remove-Item $p -Force } }
function Attempt($what, [scriptblock]$block) { try { & $block } catch { "warn: $what - $($_.Exception.Message)" } }
$officeBefore = @(Get-Process -Name WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
function Reap {
  Start-Sleep -Seconds 1
  foreach ($p in (Get-Process -Name WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue)) {
    if ($officeBefore -notcontains $p.Id) { try { Stop-Process -Id $p.Id -Force -ErrorAction Stop; "reaped $($p.ProcessName) $($p.Id)" } catch {} }
  }
}

# ═══════════════════════════════════════════════════════════════════════════
# Word: three documents, each made in a Word of its own
# ═══════════════════════════════════════════════════════════════════════════
# Each document in a PowerShell of its own — make-binary-word.ps1 says why —
# given three minutes, and the Word it started reaped either way.
if ($run -contains 'word') {
  foreach ($kind in @('text', 'paragraphs', 'structure', 'floats')) {
    $target = Join-Path $outDir "$kind.doc"
    $log = Join-Path $env:TEMP "rutba-binary-word-$kind.log"
    $child = Start-Process powershell -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $PSScriptRoot 'make-binary-word.ps1'), '-Out', $target, '-Kind', $kind, '-Picture', (Join-Path $rich 'picture.png') -PassThru -WindowStyle Hidden -RedirectStandardOutput $log
    if ($child.WaitForExit(180000)) { Get-Content $log } else { Stop-Process -Id $child.Id -Force; "warn: $kind.doc - Word did not finish saving in three minutes" }
    Remove-Item $log -ErrorAction SilentlyContinue
    Reap
  }
}

# ═══════════════════════════════════════════════════════════════════════════
# Excel: the showcase workbook in each binary format Excel still writes
# ═══════════════════════════════════════════════════════════════════════════
$xl = $null
if ($run -contains 'excel') { try {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  foreach ($f in @(@(56, 'showcase.xls', '97-2003'), @(39, 'showcase-95.xls', '5.0/95'), @(35, 'showcase-4w.xls', '4.0 workbook'), @(33, 'showcase-4.xls', '4.0'), @(29, 'showcase-3.xls', '3.0'), @(16, 'showcase-2.xls', '2.1'))) {
    $wb = $xl.Workbooks.Open((Join-Path $rich 'showcase.xlsx'), 0, $true)
    # A picture on the Summary sheet, so the binary formats carry one.
    [void]$wb.Worksheets.Item('Summary').Shapes.AddPicture((Join-Path $rich 'picture.png'), 0, -1, 320, 20, 120, 80)
    $p = Join-Path $outDir $f[1]; Remove-Existing $p
    try { $wb.SaveAs($p, $f[0]); "$($f[1]) written as Excel $($f[2]) ($((Get-Item $p).Length) bytes)" } catch { "warn: Excel $($f[2]) - $($_.Exception.Message)" }
    $wb.Close($false)
  }
} finally {
  if ($xl) { try { $xl.Quit() } catch {}; Release $xl }
  Reap
} }

# ═══════════════════════════════════════════════════════════════════════════
# Excel: the charts the showcase has not got
# ═══════════════════════════════════════════════════════════════════════════
$xl = $null
if ($run -contains 'charts') { try {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  # A workbook of its own, saved as 97-2003 and as 5.0/95: a doughnut, a
  # scatter with its axes' titles, columns with a line on a second axis, a
  # radar, columns stacked to 100%, and a 3-D chart on a chart sheet.
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = 'Data'
  function Put($range, $rows) {
    $block = New-Object 'object[,]' $rows.Count, $rows[0].Count
    for ($r = 0; $r -lt $rows.Count; $r++) { for ($k = 0; $k -lt $rows[0].Count; $k++) { $block[$r, $k] = $rows[$r][$k] } }
    $ws.Range($range).Value2 = $block
  }
  Put 'A1:D7' @(@('Month', 'North', 'South', 'East'), @('Jan', 120, 80, 60), @('Feb', 135, 90, 75), @('Mar', 150, 70, 95), @('Apr', 160, 110, 85), @('May', 175, 95, 120), @('Jun', 190, 130, 105))
  Put 'F1:G7' @(@('X', 'Y'), @(1, 2.5), @(2, 3.9), @(3, 5.1), @(4, 8.2), @(5, 9.8), @(6, 13.1))
  function Add-Chart($left, $top, $type, $range, $title) {
    # The data first: Excel will not make a chart with none a doughnut.
    $chart = $ws.ChartObjects().Add($left, $top, 300, 200).Chart
    $chart.SetSourceData($ws.Range($range))
    $chart.ChartType = [int]$type
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = $title
    return $chart
  }
  $c = Add-Chart 10 120 -4120 'A1:B7' 'North, by month'
  $c.HasLegend = $true; $c.Legend.Position = -4152
  $c = Add-Chart 320 120 74 'F1:G7' 'Y against X'
  $c.HasLegend = $false
  $c.Axes(1).HasTitle = $true; $c.Axes(1).AxisTitle.Text = 'Week'
  $c.Axes(2).HasTitle = $true; $c.Axes(2).AxisTitle.Text = 'Orders'
  $c = Add-Chart 10 330 51 'A1:C7' 'North in columns, South as a line'
  $s = $c.SeriesCollection(2); $s.ChartType = 4; $s.AxisGroup = 2
  $c.HasLegend = $true; $c.Legend.Position = -4107
  $c = Add-Chart 320 330 -4151 'A1:D7' 'Three regions round'
  $c = Add-Chart 10 540 53 'A1:D7' 'Each month, as shares'
  $sheet = $wb.Charts.Add()
  $sheet.SetSourceData($ws.Range('A1:D7'))
  $sheet.ChartType = 54
  $sheet.HasTitle = $true
  $sheet.ChartTitle.Text = 'In three dimensions'
  $sheet.Name = 'Chart3D'
  $sheet.Move([Type]::Missing, $ws)
  foreach ($f in @(@(56, 'charts.xls', '97-2003'), @(39, 'charts-95.xls', '5.0/95'))) {
    $p = Join-Path $outDir $f[1]; Remove-Existing $p
    try { $wb.SaveAs($p, $f[0]); "$($f[1]) written as Excel $($f[2]) ($((Get-Item $p).Length) bytes)" } catch { "warn: charts as Excel $($f[2]) - $($_.Exception.Message)" }
  }
  $wb.Close($false)
} finally {
  if ($xl) { try { $xl.Quit() } catch {}; Release $xl }
  Reap
} }

# ═══════════════════════════════════════════════════════════════════════════
# Excel: conditional formats of the Excel 97 kind, as 97-2003
# ═══════════════════════════════════════════════════════════════════════════
# A value over a number in red on pink and bold, a value between two numbers
# in green italic, and every other row filled by a formula.
$xl = $null
if ($run -contains 'conditions') { try {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = 'Scores'
  $block = New-Object 'object[,]' 9, 2
  $block[0, 0] = 'Name'; $block[0, 1] = 'Score'
  $people = @('Asha', 'Bilal', 'Chen', 'Dara', 'Ezra', 'Farah', 'Gita', 'Hugo'); $scores = @(91, 47, 78, 120, 63, 105, 55, 84)
  for ($r = 0; $r -lt 8; $r++) { $block[($r + 1), 0] = $people[$r]; $block[($r + 1), 1] = $scores[$r] }
  $ws.Range('A1:B9').Value2 = $block
  $over = $ws.Range('B2:B9').FormatConditions.Add(1, 5, '100')
  $over.Font.Bold = $true; $over.Font.Color = RGB 192 0 0; $over.Interior.Color = RGB 255 199 206
  $mid = $ws.Range('B2:B9').FormatConditions.Add(1, 1, '60', '90')
  $mid.Font.Italic = $true; $mid.Font.Color = RGB 0 97 0
  $rows = $ws.Range('A2:A9').FormatConditions.Add(2, 0, '=MOD(ROW(),2)=0')
  $rows.Interior.Color = RGB 221 235 247
  $p = Join-Path $outDir 'conditions.xls'; Remove-Existing $p
  $wb.SaveAs($p, 56)
  "conditions.xls written ($((Get-Item $p).Length) bytes)"
  $wb.Close($false)
} finally {
  if ($xl) { try { $xl.Quit() } catch {}; Release $xl }
  Reap
} }

# ═══════════════════════════════════════════════════════════════════════════
# PowerPoint: the showcase deck as 97-2003
# ═══════════════════════════════════════════════════════════════════════════
$pp = $null
if ($run -contains 'powerpoint') { try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pres = $pp.Presentations.Open((Join-Path $rich 'showcase.pptx'), -1, 0, 0)
  $p = Join-Path $outDir 'showcase.ppt'; Remove-Existing $p
  $pres.SaveAs($p, 1)
  "showcase.ppt written ($((Get-Item $p).Length) bytes)"
  $pres.Close()
} finally {
  if ($pp) { try { $pp.Quit() } catch {}; Release $pp }
  Reap
} }

# ═══════════════════════════════════════════════════════════════════════════
# PowerPoint: a table with merged cells, as 97-2003
# ═══════════════════════════════════════════════════════════════════════════
# Cells merged across and down, one set in the middle, one with no fill —
# what the showcase's table has not got.
$pp = $null
if ($run -contains 'tables') { try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pres = $pp.Presentations.Add(0)
  $slide = $pres.Slides.Add(1, 12)
  $table = $slide.Shapes.AddTable(4, 3, 60, 80, 600, 240).Table
  $words = @(@('Across two', '', 'Top right'), @('North', '12', 'Down three'), @('South', '9', ''), @('East', '15', ''))
  for ($r = 1; $r -le 4; $r++) { for ($k = 1; $k -le 3; $k++) { if ($words[$r - 1][$k - 1]) { $table.Cell($r, $k).Shape.TextFrame.TextRange.Text = $words[$r - 1][$k - 1] } } }
  $table.Cell(1, 1).Merge($table.Cell(1, 2))
  $table.Cell(2, 3).Merge($table.Cell(4, 3))
  $table.Cell(2, 3).Shape.TextFrame.VerticalAnchor = 3
  $table.Cell(3, 1).Shape.Fill.Visible = 0
  $p = Join-Path $outDir 'tables.ppt'; Remove-Existing $p
  $pres.SaveAs($p, 1)
  "tables.ppt written ($((Get-Item $p).Length) bytes)"
  $pres.Close()
} finally {
  if ($pp) { try { $pp.Quit() } catch {}; Release $pp }
  Reap
} }
