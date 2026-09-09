# Make the rich fixtures with Office itself: a workbook, a document and a deck
# that use what people actually put in files — charts, shapes, pictures,
# cross-sheet formulas, names, number formats, conditional formats, notes,
# footnotes, tracked changes, fields, equations, animations, speaker notes —
# saved as OOXML and as OpenDocument by the same application, so both trees
# come from one source of truth and the suite is judged against what Office
# wrote, not against what we think Office writes.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\make-rich-fixtures.ps1 [-Out tests\fixtures\rich]
#
# Office is driven invisibly. Nothing here touches the desktop, the keyboard
# or the network; the applications are quit and released at the end even when
# a step fails. Chart data in the document and the deck is Office's default
# — editing it would open an Excel window.
param([string]$Out = 'tests\fixtures\rich')

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$outDir = if ([IO.Path]::IsPathRooted($Out)) { $Out } else { Join-Path $root $Out }
New-Item -ItemType Directory -Force $outDir | Out-Null
$stamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
function RGB($r, $g, $b) { return $r + ($g * 256) + ($b * 65536) }
function Release($o) { if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} } }
function Remove-Existing($p) { if (Test-Path $p) { Remove-Item $p -Force } }
# A decoration that Office refuses is reported and skipped; the file is still
# written with everything else, and the warning names what to fix.
function Attempt($what, [scriptblock]$block) { try { & $block } catch { "warn: $what - $($_.Exception.Message)" } }
# Office processes that were already running belong to somebody; the ones
# this run starts are reaped at the end of each section, because Quit() has
# been seen to leave an Excel behind with references still alive.
$officeBefore = @(Get-Process -Name WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
function Reap {
  Start-Sleep -Seconds 1
  foreach ($p in (Get-Process -Name WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue)) {
    if ($officeBefore -notcontains $p.Id) { try { Stop-Process -Id $p.Id -Force -ErrorAction Stop; "reaped $($p.ProcessName) $($p.Id)" } catch {} }
  }
}

# ── a small colourful picture, made here so nothing is fetched ──────────────
Add-Type -AssemblyName System.Drawing
$png = Join-Path $outDir 'picture.png'
$bmp = New-Object System.Drawing.Bitmap 240, 160
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = 'AntiAlias'
$brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush (New-Object System.Drawing.Rectangle 0, 0, 240, 160), ([System.Drawing.Color]::FromArgb(255, 30, 90, 180)), ([System.Drawing.Color]::FromArgb(255, 250, 190, 40)), 30
$g.FillRectangle($brush, 0, 0, 240, 160)
$g.FillEllipse([System.Drawing.Brushes]::White, 60, 30, 100, 100)
$g.FillEllipse([System.Drawing.Brushes]::Crimson, 85, 55, 50, 50)
$font = New-Object System.Drawing.Font 'Segoe UI', 16, ([System.Drawing.FontStyle]::Bold)
$g.DrawString('Rutba', $font, [System.Drawing.Brushes]::White, 8, 120)
$g.Dispose(); $bmp.Save($png, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
"picture.png written"

$months = @('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec')
$regions = @('North', 'South', 'East', 'West')

# ═══════════════════════════════════════════════════════════════════════════
# Excel: the workbook
# ═══════════════════════════════════════════════════════════════════════════
$xl = $null
try {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  Attempt 'add-ins off' { foreach ($a in $xl.COMAddIns) { if ($a.Connect) { $a.Connect = $false } } }
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 5) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $sales = $wb.Worksheets.Item(1); $sales.Name = 'Sales'
  $summary = $wb.Worksheets.Item(2); $summary.Name = 'Summary'
  $charts = $wb.Worksheets.Item(3); $charts.Name = 'Charts'
  $types = $wb.Worksheets.Item(4); $types.Name = 'Data types'
  $table = $wb.Worksheets.Item(5); $table.Name = 'Table'

  # ── Sales: twelve months by four regions, totals every way ──────────────
  $sales.Range('A1').Value2 = 'Month'
  for ($c = 0; $c -lt 4; $c++) { $sales.Cells.Item(1, 2 + $c).Value2 = $regions[$c] }
  $sales.Range('F1').Value2 = 'Total'; $sales.Range('G1').Value2 = 'Trend'; $sales.Range('H1').Value2 = 'Growth'
  for ($r = 0; $r -lt 12; $r++) {
    $row = 2 + $r
    $sales.Cells.Item($row, 1).Value2 = $months[$r]
    # [double] throughout: PowerShell 5.1 binds a COM property setter to the
    # type it first saw — a string, above — and refuses an Int32 after that.
    for ($c = 0; $c -lt 4; $c++) { $sales.Cells.Item($row, 2 + $c).Value2 = [double](1200 + ($r * 137) + ($c * 310) + ((($r * 7) + ($c * 13)) % 97) * 11) }
    $sales.Cells.Item($row, 6).Formula = "=SUM(B${row}:E${row})"
    if ($r -gt 0) { $sales.Cells.Item($row, 8).Formula = "=IFERROR((F${row}-F$($row - 1))/F$($row - 1),0)" }
  }
  $sales.Range('A14').Value2 = 'Total'; $sales.Range('A15').Value2 = 'Average'; $sales.Range('A16').Value2 = 'Best'; $sales.Range('A17').Value2 = 'Worst'
  foreach ($col in @('B', 'C', 'D', 'E', 'F')) {
    $sales.Range("${col}14").Formula = "=SUM(${col}2:${col}13)"
    $sales.Range("${col}15").Formula = "=AVERAGE(${col}2:${col}13)"
    $sales.Range("${col}16").Formula = "=MAX(${col}2:${col}13)"
    $sales.Range("${col}17").Formula = "=MIN(${col}2:${col}13)"
  }
  $sales.Range('B2:F17').NumberFormat = '#,##0'
  $sales.Range('B15:F15').NumberFormat = '#,##0.0'
  $sales.Range('H3:H13').NumberFormat = '0.0%'
  $head = $sales.Range('A1:H1'); $head.Font.Bold = $true; $head.Font.Color = RGB 255 255 255; $head.Interior.Color = RGB 31 78 121
  $sales.Range('A14:H14').Font.Bold = $true; $sales.Range('A14:H14').Interior.Color = RGB 221 235 247
  $sales.Range('A15:H17').Interior.Color = RGB 242 242 242
  $sales.Range('A1:H17').Borders.LineStyle = 1; $sales.Range('A1:H17').Borders.Weight = 2
  $sales.Range('A1:H17').Borders.Color = RGB 160 160 160
  $sales.Range('F2:F13').Font.Bold = $true
  # conditional formats: a colour scale over the grid, data bars on the totals, icons on growth
  Attempt 'colour scale' { [void]$sales.Range('B2:E13').FormatConditions.AddColorScale(3) }
  Attempt 'data bars' { [void]$sales.Range('F2:F13').FormatConditions.AddDatabar() }
  Attempt 'icon set' { [void]$sales.Range('H3:H13').FormatConditions.AddIconSetCondition() }
  # sparklines, an autofilter, frozen headings, widths, a note, a hyperlink, a validation list
  Attempt 'sparklines' { [void]$sales.Range('G2:G13').SparklineGroups.Add(1, 'B2:E13') }
  Attempt 'autofilter' { [void]$sales.Range('A1:H13').AutoFilter() }
  $sales.Columns.Item(1).ColumnWidth = 12; $sales.Range('B:H').ColumnWidth = 11
  Attempt 'note' { [void]$sales.Range('A1').AddComment("Months of the year $stamp - a plain note, not a threaded comment.") }
  $sales.Range('A19').Value2 = 'Rutba Office'
  Attempt 'hyperlink' { [void]$sales.Hyperlinks.Add($sales.Range('A19'), 'https://office.rutba.io', '', 'The free suite', 'Rutba Office') }
  $sales.Range('A21').Value2 = 'Approved?'
  Attempt 'validation' { $sales.Range('B21').Validation.Add(3, 1, 1, 'Yes,No,Maybe') }
  $sales.Range('B21').Value2 = 'Yes'
  Attempt 'freeze panes' { $sales.Activate(); $win = $wb.Windows.Item(1); $win.SplitRow = 1; $win.SplitColumn = 1; $win.FreezePanes = $true }
  # names the other sheets use
  [void]$wb.Names.Add('SalesTotals', '=Sales!$F$2:$F$13')
  [void]$wb.Names.Add('Regions', '=Sales!$B$1:$E$1')
  [void]$wb.Names.Add('GrandTotal', '=Sales!$F$14')

  # ── Summary: every kind of reference into the other sheets ──────────────
  $summary.Range('A1').Value2 = 'Summary of the year'
  $summary.Range('A1:D1').Merge(); $summary.Range('A1').Font.Size = 18; $summary.Range('A1').Font.Bold = $true
  $summary.Range('A1').Font.Color = RGB 255 255 255; $summary.Range('A1:D1').Interior.Color = RGB 192 0 0
  $summary.Range('A1:D1').HorizontalAlignment = -4108
  $rows = @(
    @('Grand total', '=SUM(Sales!F2:F13)', '#,##0'),
    @('Grand total, by name', '=GrandTotal', '#,##0'),
    @('Best month', '=INDEX(Sales!A2:A13,MATCH(MAX(Sales!F2:F13),Sales!F2:F13,0))', 'General'),
    @('March, looked up', '=VLOOKUP("Mar",Sales!A2:F13,6,FALSE)', '#,##0'),
    @('Months over 8,000', '=COUNTIF(Sales!F2:F13,">8000")', '0'),
    @('J-months, summed', '=SUMIF(Sales!A2:A13,"J*",Sales!F2:F13)', '#,##0'),
    @('North share', '=Sales!B14/GrandTotal', '0.0%'),
    @('Average month, by name', '=AVERAGE(SalesTotals)', '#,##0.0'),
    @('Verdict', '=IF(GrandTotal>90000,"On target","Below target")', 'General'),
    @('As text', '=TEXT(GrandTotal,"#,##0.00")', 'General'),
    @('Data types, added', "='Data types'!B2+'Data types'!B3", '0.000'),
    @('Table total', '=SUM(Table!D2:D7)', '#,##0.00'),
    @('Rounded, nested', '=ROUND(AVERAGE(Sales!B2:E13)*1.1,2)', '0.00'),
    @('Regions counted', '=COUNTA(Regions)', '0')
  )
  for ($i = 0; $i -lt $rows.Count; $i++) {
    $r = 3 + $i
    $summary.Cells.Item($r, 1).Value2 = $rows[$i][0]
    $summary.Cells.Item($r, 2).Formula = $rows[$i][1]
    $summary.Cells.Item($r, 2).NumberFormat = $rows[$i][2]
  }
  $summary.Range('A3:A16').Font.Bold = $true; $summary.Columns.Item(1).ColumnWidth = 26; $summary.Columns.Item(2).ColumnWidth = 16
  $summary.Range('A3:B16').Borders.LineStyle = 1
  $summary.Range('A3:B16').Interior.Color = RGB 255 242 204
  $summary.Range('A18').Value2 = 'Region totals'; $summary.Range('A18').Font.Bold = $true
  for ($c = 0; $c -lt 4; $c++) { $summary.Cells.Item(19, 1 + $c).Formula = "=Sales!$(@('B','C','D','E')[$c])1"; $summary.Cells.Item(20, 1 + $c).Formula = "=Sales!$(@('B','C','D','E')[$c])14"; $summary.Cells.Item(20, 1 + $c).NumberFormat = '#,##0' }
  $pie = $summary.Shapes.AddChart2(-1, 5, 300, 30, 340, 240)
  $pie.Chart.SetSourceData($summary.Range('A19:D20'))
  $pie.Chart.HasTitle = $true; $pie.Chart.ChartTitle.Text = 'Share by region'
  Attempt 'pie labels' { $pie.Chart.SeriesCollection(1).HasDataLabels = $true }
  Attempt 'pie percentages' { $pie.Chart.SeriesCollection(1).DataLabels().ShowPercentage = $true }

  # ── Charts: four kinds, and a page of shapes and a picture ──────────────
  $col = $charts.Shapes.AddChart2(-1, 51, 10, 10, 420, 240)
  $col.Chart.SetSourceData($sales.Range('A1:E13')); $col.Chart.HasTitle = $true; $col.Chart.ChartTitle.Text = 'Sales by region, by month'
  $line = $charts.Shapes.AddChart2(-1, 4, 440, 10, 420, 240)
  $line.Chart.SetSourceData($sales.Range('A1:E13')); $line.Chart.HasTitle = $true; $line.Chart.ChartTitle.Text = 'The year, as lines'
  $bar = $charts.Shapes.AddChart2(-1, 57, 10, 260, 420, 240)
  $bar.Chart.SetSourceData($sales.Range('A1:A13,F1:F13')); $bar.Chart.HasTitle = $true; $bar.Chart.ChartTitle.Text = 'Monthly totals'
  Attempt 'bar labels' { $bar.Chart.SeriesCollection(1).HasDataLabels = $true }
  $area = $charts.Shapes.AddChart2(-1, 76, 440, 260, 420, 240)
  $area.Chart.SetSourceData($sales.Range('A1:E13')); $area.Chart.HasTitle = $true; $area.Chart.ChartTitle.Text = 'Stacked, as areas'
  $shapeKinds = @(@(1, 'Rectangle'), @(5, 'Rounded'), @(9, 'Oval'), @(7, 'Triangle'), @(92, 'Star'), @(33, 'Arrow'), @(52, 'Chevron'), @(10, 'Hexagon'), @(21, 'Heart'), @(17, 'Smile'), @(61, 'Process'), @(105, 'Callout'))
  $palette = @((RGB 192 0 0), (RGB 237 125 49), (RGB 255 192 0), (RGB 112 173 71), (RGB 68 114 196), (RGB 112 48 160), (RGB 0 176 240), (RGB 255 0 102), (RGB 0 128 128), (RGB 128 96 0), (RGB 91 155 213), (RGB 165 165 165))
  for ($i = 0; $i -lt $shapeKinds.Count; $i++) {
    $s = $charts.Shapes.AddShape($shapeKinds[$i][0], 10 + (($i % 6) * 145), 520 + ([Math]::Floor($i / 6) * 110), 120, 80)
    $s.Name = 'Shape ' + $shapeKinds[$i][1]
    $s.Fill.ForeColor.RGB = $palette[$i]
    if ($i % 3 -eq 1) { $s.Fill.TwoColorGradient(1, 1); $s.Fill.BackColor.RGB = RGB 255 255 255 }
    $s.Line.ForeColor.RGB = RGB 40 40 40; $s.Line.Weight = 1.5
    if ($i % 4 -eq 3) { $s.Shadow.Visible = -1 }
    $s.TextFrame2.TextRange.Text = $shapeKinds[$i][1]
    $s.TextFrame2.TextRange.Font.Size = 12; $s.TextFrame2.TextRange.Font.Bold = -1
    $s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB 255 255 255
  }
  [void]$charts.Shapes.AddPicture($png, 0, -1, 10, 750, 240, 160)
  $tb = $charts.Shapes.AddTextbox(1, 270, 750, 300, 160)
  $tb.TextFrame2.TextRange.Text = "A text box beside the picture.`rSecond line, in colour.`rThird line, bold."
  $tb.TextFrame2.TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = RGB 192 0 0
  $tb.TextFrame2.TextRange.Paragraphs(3).Font.Bold = -1
  $tb.Fill.ForeColor.RGB = RGB 255 242 204; $tb.Line.ForeColor.RGB = RGB 191 143 0
  $g1 = $charts.Shapes.AddShape(9, 600, 760, 60, 60); $g1.Name = 'GroupA'; $g1.Fill.ForeColor.RGB = RGB 68 114 196
  $g2 = $charts.Shapes.AddShape(9, 640, 800, 60, 60); $g2.Name = 'GroupB'; $g2.Fill.ForeColor.RGB = RGB 237 125 49
  Attempt 'group' { $grp = $charts.Shapes.Range([object[]]@('GroupA', 'GroupB')).Group(); $grp.Name = 'Two circles' }
  $rot = $charts.Shapes.AddShape(1, 740, 760, 100, 50); $rot.Rotation = 30; $rot.Fill.ForeColor.RGB = RGB 112 173 71; $rot.TextFrame2.TextRange.Text = 'Rotated'
  Attempt 'connector' { [void]$charts.Shapes.AddConnector(2, 130, 560, 155, 560) }

  # ── Data types: one of everything a cell can hold ───────────────────────
  $typeRows = @(
    @('Integer', 42, 'General'), @('Decimal', 3.14159, '0.000'), @('Currency', 1234.5, '£#,##0.00'), @('Percent', 0.256, '0.0%'),
    @('Date', 46275, 'dd/mm/yyyy'), @('Time', 0.75, 'hh:mm'), @('Fraction', 0.375, '# ?/?'), @('Scientific', 123456789, '0.00E+00'),
    @('Boolean', $true, 'General'), @('Negative in red', -987.65, '#,##0.00;[Red]-#,##0.00'), @('Thousands', 9876543, '#,##0'), @('Accounting', 42.5, '_-£* #,##0.00_-;-£* #,##0.00_-;_-£* "-"??_-;_-@_-')
  )
  $types.Range('A1').Value2 = 'Kind'; $types.Range('B1').Value2 = 'Value'; $types.Range('C1').Value2 = 'Hidden'; $types.Range('D1').Value2 = 'Note'
  $types.Range('A1:D1').Font.Bold = $true; $types.Range('A1:D1').Interior.Color = RGB 226 239 218
  for ($i = 0; $i -lt $typeRows.Count; $i++) {
    $r = 2 + $i
    $types.Cells.Item($r, 1).Value2 = $typeRows[$i][0]
    $v = $typeRows[$i][1]
    if ($v -is [bool]) { try { $types.Cells.Item($r, 2).Value2 = [bool]$v } catch { $types.Cells.Item($r, 2).Formula = '=TRUE()' } } else { $types.Cells.Item($r, 2).Value2 = [double]$v }
    $types.Cells.Item($r, 2).NumberFormat = $typeRows[$i][2]
    $types.Cells.Item($r, 3).Value2 = 'hidden ' + $i
  }
  $types.Range('A14').Value2 = 'Division by zero'; $types.Range('B14').Formula = '=1/0'
  $types.Range('A15').Value2 = 'Not available'; $types.Range('B15').Formula = '=NA()'
  $types.Range('A16').Value2 = 'From another sheet'; $types.Range('B16').Formula = '=Sales!F14'; $types.Range('B16').NumberFormat = '#,##0'
  $types.Range('A17').Value2 = 'Long text, wrapped'; $types.Range('B17').Value2 = 'A sentence long enough to need wrapping inside its cell, which is what this column is for.'; $types.Range('B17').WrapText = $true; $types.Rows.Item(17).RowHeight = 60
  $types.Range('A18').Value2 = 'Rich text'; $types.Range('B18').Value2 = 'Bold start, red middle, plain end'
  Attempt 'rich text' { $types.Range('B18').Characters(1, 10).Font.Bold = $true; $types.Range('B18').Characters(12, 10).Font.Color = RGB 255 0 0; $types.Range('B18').Characters(12, 10).Font.Italic = $true }
  $types.Range('A19').Value2 = 'Rotated'; $types.Range('B19').Value2 = 'Up at 45 degrees'; $types.Range('B19').Orientation = 45; $types.Rows.Item(19).RowHeight = 70
  $types.Range('A20').Value2 = 'Merged across'; $types.Range('B20:D20').Merge(); $types.Range('B20').Value2 = 'One cell across three columns'; $types.Range('B20').HorizontalAlignment = -4108; $types.Range('B20').Interior.Color = RGB 255 230 153
  Attempt 'superscript' { $types.Range('A21').Value2 = 'Superscript'; $types.Range('B21').Value2 = 'E = mc2'; $types.Range('B21').Characters(7, 1).Font.Superscript = $true }
  $types.Range('A22').Value2 = 'Unicode'; $types.Range('B22').Value2 = 'Café — naïve • 日本語 • العربية • Ωmega • ✓'
  $types.Range('A23').Value2 = 'Big font'; $types.Range('B23').Value2 = 'Twenty-four point'; $types.Range('B23').Font.Size = 24; $types.Range('B23').Font.Name = 'Georgia'
  $types.Range('A24').Value2 = 'Fill and border'; $types.Range('B24').Value2 = 'Boxed'; $types.Range('B24').Interior.Color = RGB 189 215 238; $types.Range('B24').Borders.LineStyle = 1; $types.Range('B24').Borders.Weight = 4
  $types.Range('A25').Value2 = 'Indented, right'; $types.Range('B25').Value2 = 'right'; $types.Range('B25').HorizontalAlignment = -4152; $types.Range('B25').IndentLevel = 2
  $types.Columns.Item(3).Hidden = $true
  $types.Range('A27').Value2 = 'Grouped rows below'
  $types.Range('A28').Value2 = 'one'; $types.Range('A29').Value2 = 'two'; $types.Range('A30').Value2 = 'three'
  Attempt 'row group' { [void]$types.Rows.Item('28:30').Group() }
  $types.Columns.Item(1).ColumnWidth = 22; $types.Columns.Item(2).ColumnWidth = 34; $types.Columns.Item(4).ColumnWidth = 18
  $types.Range('D2').Value2 = 'plain'; $types.Range('D3').Value2 = 'three decimals'; $types.Range('D4').Value2 = 'pounds'
  $types.Range('A2').Font.Color = RGB 0 112 192; $types.Range('A2').Font.Underline = 2

  # ── Table: a real table with a style and a total row ────────────────────
  $table.Range('A1').Value2 = 'Product'; $table.Range('B1').Value2 = 'Quantity'; $table.Range('C1').Value2 = 'Price'; $table.Range('D1').Value2 = 'Amount'
  $products = @(@('Desk', 4, 249.99), @('Chair', 12, 89.5), @('Lamp', 7, 34.25), @('Shelf', 3, 129), @('Monitor', 6, 219.99), @('Cable', 40, 4.75))
  for ($i = 0; $i -lt $products.Count; $i++) {
    $r = 2 + $i
    $table.Cells.Item($r, 1).Value2 = $products[$i][0]; $table.Cells.Item($r, 2).Value2 = [double]$products[$i][1]; $table.Cells.Item($r, 3).Value2 = [double]$products[$i][2]
    $table.Cells.Item($r, 4).Formula = "=B${r}*C${r}"
  }
  $lo = $table.ListObjects.Add(1, $table.Range('A1:D7'), $null, 1)
  $lo.Name = 'Orders'; $lo.TableStyle = 'TableStyleMedium9'; $lo.ShowTotals = $true
  Attempt 'table total' { $lo.ListColumns.Item(4).TotalsCalculation = 1 }
  $table.Range('C2:D8').NumberFormat = '£#,##0.00'
  $table.Range('A10').Value2 = 'Structured reference'; $table.Range('B10').Formula = '=SUM(Orders[Amount])'; $table.Range('B10').NumberFormat = '£#,##0.00'
  $table.Columns.Item(1).ColumnWidth = 22

  $sales.Activate()
  $xlsx = Join-Path $outDir 'showcase.xlsx'; Remove-Existing $xlsx
  $wb.SaveAs($xlsx, 51)
  "showcase.xlsx written ($((Get-Item $xlsx).Length) bytes)"
  $ods = Join-Path $outDir 'showcase.ods'; Remove-Existing $ods
  $wb.SaveAs($ods, 60)
  "showcase.ods written ($((Get-Item $ods).Length) bytes)"
  $wb.Close($false)
} finally {
  if ($xl) { try { $xl.Quit() } catch {} ; Release $xl }
  Reap
}

# ═══════════════════════════════════════════════════════════════════════════
# Word: the document
# ═══════════════════════════════════════════════════════════════════════════
$wd = $null
try {
  $wd = New-Object -ComObject Word.Application
  $wd.Visible = $false
  $wd.DisplayAlerts = 0
  # Add-ins hook the save: Acrobat's PDFMaker held SaveAs2 for ever in a
  # hidden Word here. Disconnected for this session only.
  Attempt 'add-ins off' { foreach ($a in $wd.COMAddIns) { if ($a.Connect) { $a.Connect = $false } } }
  $doc = $wd.Documents.Add()
  $sel = $wd.Selection
  function Para($text, $style) { $sel.Style = $style; $sel.TypeText($text); $sel.TypeParagraph() }

  $sel.Style = 'Title'; $sel.TypeText('The Rutba Office showcase'); $sel.TypeParagraph()
  $sel.Style = 'Subtitle'; $sel.TypeText("A document with one of everything, written by Word on $stamp"); $sel.TypeParagraph()
  $sel.Style = 'Normal'
  $sel.Font.Bold = $true; $sel.TypeText('Contents'); $sel.Font.Bold = $false; $sel.TypeParagraph()
  # The table of contents is built here at the end, once the headings exist
  # (see 'table of contents' below); this paragraph marks where.
  Attempt 'contents mark' { [void]$doc.Bookmarks.Add('TOCHere', $sel.Range) }
  $sel.TypeParagraph()

  Para 'Text and its formatting' 'Heading 1'
  $sel.Style = 'Normal'
  $sel.TypeText('This paragraph mixes '); $sel.Font.Bold = $true; $sel.TypeText('bold'); $sel.Font.Bold = $false
  $sel.TypeText(', '); $sel.Font.Italic = $true; $sel.TypeText('italic'); $sel.Font.Italic = $false
  $sel.TypeText(', '); $sel.Font.Underline = 1; $sel.TypeText('underlined'); $sel.Font.Underline = 0
  $sel.TypeText(', '); $sel.Font.StrikeThrough = $true; $sel.TypeText('struck'); $sel.Font.StrikeThrough = $false
  $sel.TypeText(', '); $sel.Font.Color = RGB 192 0 0; $sel.TypeText('red'); $sel.Font.Color = -16777216
  $sel.TypeText(', '); $sel.Range.HighlightColorIndex = 7; $sel.TypeText('highlighted'); $sel.Range.HighlightColorIndex = 0
  $sel.TypeText(', '); $sel.Font.Superscript = $true; $sel.TypeText('super'); $sel.Font.Superscript = $false
  $sel.TypeText(' and '); $sel.Font.Subscript = $true; $sel.TypeText('sub'); $sel.Font.Subscript = $false
  $sel.TypeText(', '); $sel.Font.SmallCaps = $true; $sel.TypeText('small caps'); $sel.Font.SmallCaps = $false
  $sel.TypeText(', '); $sel.Font.Name = 'Georgia'; $sel.Font.Size = 16; $sel.TypeText('Georgia at 16'); $sel.Font.Name = 'Calibri'; $sel.Font.Size = 11
  $sel.TypeText(', and a '); $sel.Font.Spacing = 3; $sel.TypeText('spaced'); $sel.Font.Spacing = 0
  $sel.TypeText(' run. Symbols: © ® ™ € £ ¥ → ← ↔ ✓ ★ ½ ¾ ∑ ∞ ≠ ≤. Non-Latin: 日本語のテキスト, 简体中文, العربية من اليمين إلى اليسار, Ελληνικά, кириллица.')
  $sel.TypeParagraph()
  $sel.ParagraphFormat.Alignment = 3; $sel.ParagraphFormat.LeftIndent = 36; $sel.ParagraphFormat.RightIndent = 36
  $sel.ParagraphFormat.SpaceBefore = 6; $sel.ParagraphFormat.SpaceAfter = 6; $sel.ParagraphFormat.LineSpacingRule = 1
  $sel.ParagraphFormat.Shading.BackgroundPatternColor = RGB 226 239 218
  $sel.ParagraphFormat.Borders.Item(-1).LineStyle = 1; $sel.ParagraphFormat.Borders.Item(-3).LineStyle = 1
  $sel.TypeText('A justified, indented paragraph with shading and a border above and below, one and a half spaced. It runs on long enough to wrap onto more than one line so that the justification and the indents can be seen for what they are.')
  $sel.TypeParagraph()
  $sel.ParagraphFormat.Reset(); $sel.ParagraphFormat.Shading.BackgroundPatternColor = -16777216; $sel.ParagraphFormat.Borders.Item(-1).LineStyle = 0; $sel.ParagraphFormat.Borders.Item(-3).LineStyle = 0
  $sel.Style = 'Normal'

  Para 'Lists' 'Heading 1'
  Para 'Bulleted, three levels' 'Heading 2'
  $sel.Style = 'Normal'; $sel.Range.ListFormat.ApplyBulletDefault()
  $sel.TypeText('First point'); $sel.TypeParagraph()
  $sel.Range.ListFormat.ListIndent(); $sel.TypeText('A sub-point'); $sel.TypeParagraph()
  $sel.Range.ListFormat.ListIndent(); $sel.TypeText('A sub-sub-point'); $sel.TypeParagraph()
  $sel.Range.ListFormat.ListOutdent(); $sel.Range.ListFormat.ListOutdent(); $sel.TypeText('Second point'); $sel.TypeParagraph()
  $sel.Range.ListFormat.RemoveNumbers()
  Para 'Numbered' 'Heading 2'
  $sel.Style = 'Normal'; $sel.Range.ListFormat.ApplyNumberDefault()
  foreach ($t in @('Open the file', 'Change something', 'Save it back')) { $sel.TypeText($t); $sel.TypeParagraph() }
  $sel.Range.ListFormat.RemoveNumbers()
  Para 'Outline numbered' 'Heading 2'
  $sel.Style = 'Normal'; $sel.Range.ListFormat.ApplyOutlineNumberDefault()
  $sel.TypeText('Chapter'); $sel.TypeParagraph(); $sel.Range.ListFormat.ListIndent(); $sel.TypeText('Section'); $sel.TypeParagraph()
  $sel.Range.ListFormat.ListIndent(); $sel.TypeText('Clause'); $sel.TypeParagraph(); $sel.Range.ListFormat.RemoveNumbers()

  Para 'A table' 'Heading 1'
  $sel.Style = 'Normal'
  $t = $doc.Tables.Add($sel.Range, 5, 4)
  $t.Style = 'Grid Table 4 - Accent 1'
  $t.Cell(1, 1).Range.Text = 'Region'; $t.Cell(1, 2).Range.Text = 'Q1'; $t.Cell(1, 3).Range.Text = 'Q2'; $t.Cell(1, 4).Range.Text = 'Total'
  $data = @(@('North', 1200, 1350), @('South', 980, 1120), @('East', 1430, 1510))
  for ($i = 0; $i -lt 3; $i++) { $t.Cell(2 + $i, 1).Range.Text = $data[$i][0]; $t.Cell(2 + $i, 2).Range.Text = "$($data[$i][1])"; $t.Cell(2 + $i, 3).Range.Text = "$($data[$i][2])"; $t.Cell(2 + $i, 4).Range.Text = "$($data[$i][1] + $data[$i][2])" }
  $t.Cell(5, 1).Range.Text = 'All regions'; $t.Cell(5, 1).Merge($t.Cell(5, 3)); $t.Cell(5, 2).Range.Text = '7590'
  $t.Cell(5, 2).Shading.BackgroundPatternColor = RGB 255 230 153; $t.Cell(5, 2).Range.Font.Bold = $true
  $t.Rows.Item(1).HeadingFormat = $true
  $t.Range.ParagraphFormat.Alignment = 1
  $sel.EndKey(6) | Out-Null; $sel.TypeParagraph()

  Para 'A picture, a chart and shapes' 'Heading 1'
  $sel.Style = 'Normal'
  [void]$sel.InlineShapes.AddPicture($png)
  $sel.TypeParagraph()
  $sel.Style = 'Caption'; $sel.TypeText('Figure 1: a picture made for this document'); $sel.TypeParagraph(); $sel.Style = 'Normal'
  # No chart in the document: InlineShapes.AddChart2 keeps its data in an
  # Excel Word starts for it, and SaveAs2 then never returns — with the data
  # window open or closed, twice over. The workbook and the deck carry the
  # charts; the corpus has documents with charts other people wrote.
  $sel.EndKey(6) | Out-Null; $sel.TypeParagraph()
  $sel.TypeText('The shapes below float beside this text: a gradient rectangle with words in it, an oval, an arrow, a star and a callout, and a text box.')
  $sel.TypeParagraph()
  "stage: floating shapes"
  Attempt 'floating shapes' {
    $anchor = $sel.Range
    $r1 = $doc.Shapes.AddShape(1, 40, 0, 150, 60, $anchor); $r1.Fill.ForeColor.RGB = RGB 68 114 196; $r1.Fill.TwoColorGradient(1, 1); $r1.Fill.BackColor.RGB = RGB 255 255 255
    $r1.TextFrame.TextRange.Text = 'Gradient'; $r1.TextFrame.TextRange.Font.Bold = $true; $r1.WrapFormat.Type = 4
    $r2 = $doc.Shapes.AddShape(9, 210, 0, 80, 60, $anchor); $r2.Fill.ForeColor.RGB = RGB 237 125 49; $r2.WrapFormat.Type = 4
    $r3 = $doc.Shapes.AddShape(33, 310, 10, 90, 40, $anchor); $r3.Fill.ForeColor.RGB = RGB 112 173 71; $r3.WrapFormat.Type = 4
    $r4 = $doc.Shapes.AddShape(92, 420, 0, 60, 60, $anchor); $r4.Fill.ForeColor.RGB = RGB 255 192 0; $r4.WrapFormat.Type = 4
    $r5 = $doc.Shapes.AddShape(105, 40, 80, 160, 60, $anchor); $r5.Fill.ForeColor.RGB = RGB 255 242 204; $r5.TextFrame.TextRange.Text = 'A callout'; $r5.WrapFormat.Type = 4
    $tb = $doc.Shapes.AddTextbox(1, 220, 80, 200, 60, $anchor); $tb.TextFrame.TextRange.Text = 'A floating text box with its own border.'; $tb.Line.ForeColor.RGB = RGB 192 0 0; $tb.WrapFormat.Type = 4
  }
  for ($i = 0; $i -lt 8; $i++) { $sel.TypeParagraph() }

  "stage: notes, comments, fields, changes"
  Para 'Notes, comments and changes' 'Heading 1'
  $sel.Style = 'Normal'
  $sel.TypeText('This sentence carries a footnote')
  Attempt 'footnote' { [void]$doc.Footnotes.Add($sel.Range, [Type]::Missing, 'The footnote text, at the foot of the page.') }
  $sel.TypeText(' and this one an endnote')
  Attempt 'endnote' { [void]$doc.Endnotes.Add($sel.Range, [Type]::Missing, 'The endnote text, at the end of the document.') }
  $sel.TypeText('. ')
  $start = $sel.Range.Start
  $sel.TypeText('This phrase has a comment on it')
  $commented = $doc.Range($start, $sel.Range.End)
  Attempt 'comment' { [void]$doc.Comments.Add($commented, 'A comment from the author, with the date Word stamps on it.') }
  $sel.TypeText('. ')
  $bmStart = $sel.Range.Start
  $sel.TypeText('This phrase is bookmarked')
  Attempt 'bookmark' { [void]$doc.Bookmarks.Add('Bookmarked', $doc.Range($bmStart, $sel.Range.End)) }
  $sel.TypeText(', this word is a ')
  $hlStart = $sel.Range.Start
  $sel.TypeText('hyperlink')
  Attempt 'hyperlink' { [void]$doc.Hyperlinks.Add($doc.Range($hlStart, $sel.Range.End), 'https://office.rutba.io', '', 'Rutba Office', 'hyperlink') }
  $sel.EndKey(6) | Out-Null
  $sel.TypeText(', and the fields: today is ')
  Attempt 'date field' { [void]$doc.Fields.Add($sel.Range, -1, 'DATE \@ "d MMMM yyyy"', $true) }
  $sel.EndKey(6) | Out-Null
  $sel.TypeText(', this is page ')
  Attempt 'page field' { [void]$doc.Fields.Add($sel.Range, -1, 'PAGE', $true) }
  $sel.EndKey(6) | Out-Null
  $sel.TypeText(' of ')
  Attempt 'numpages field' { [void]$doc.Fields.Add($sel.Range, -1, 'NUMPAGES', $true) }
  $sel.EndKey(6) | Out-Null
  $sel.TypeText('.'); $sel.TypeParagraph()
  Attempt 'tracked changes' {
    $doc.TrackRevisions = $true
    $sel.TypeText('This sentence was inserted with Track Changes on, ')
    $doc.TrackRevisions = $false
    $sel.TypeText('and the next words were deleted the same way: ')
    $delStart = $sel.Range.Start
    $sel.TypeText('these words are gone')
    $sel.TypeText('.'); $sel.TypeParagraph()
    $doc.TrackRevisions = $true
    $doc.Range($delStart, $delStart + 20).Delete() | Out-Null
    $doc.TrackRevisions = $false
  }
  $sel.EndKey(6) | Out-Null

  "stage: equation"
  Para 'An equation' 'Heading 1'
  $sel.Style = 'Normal'
  Attempt 'equation' {
    $eqRange = $sel.Range
    $eqRange.Text = 'x=(-b±√(b^2-4ac))/2a'
    $om = $doc.OMaths.Add($eqRange)
    $om.OMaths.Item(1).BuildUp()
  }
  $sel.EndKey(6) | Out-Null; $sel.TypeParagraph()

  # a landscape section in two columns
  "stage: landscape section"
  Attempt 'landscape section' {
    [void]$doc.Sections.Add()
    $sec = $doc.Sections.Item($doc.Sections.Count)
    $sec.PageSetup.Orientation = 1
    $sec.PageSetup.TextColumns.SetCount(2)
  }
  $sel.EndKey(6) | Out-Null
  Para 'A landscape section in two columns' 'Heading 1'
  $sel.Style = 'Normal'
  for ($i = 0; $i -lt 6; $i++) { $sel.TypeText("Column text, paragraph $($i + 1). Words enough to fill a column and flow into the next, so that the section's orientation and its two columns both show. "); $sel.TypeParagraph() }

  # header, footer with page numbers, a watermark
  Attempt 'header' { $doc.Sections.Item(1).Headers.Item(1).Range.Text = 'Rutba Office showcase — header' }
  Attempt 'footer page numbers' { $doc.Sections.Item(1).Footers.Item(1).PageNumbers.Add(2) | Out-Null }
  "stage: header, footer, watermark"
  Attempt 'watermark' {
    $wm = $doc.Sections.Item(1).Headers.Item(1).Shapes.AddTextEffect(0, 'DRAFT', 'Calibri', 1, $false, $false, 0, 0)
    $wm.Name = 'PowerPlusWaterMarkObject1'; $wm.TextEffect.NormalizedHeight = $false; $wm.Line.Visible = $false
    $wm.Fill.Visible = $true; $wm.Fill.Solid(); $wm.Fill.ForeColor.RGB = RGB 192 192 192; $wm.Fill.Transparency = 0.5
    $wm.Rotation = 315; $wm.LockAspectRatio = $true; $wm.Height = 120; $wm.Width = 360
    $wm.RelativeHorizontalPosition = 1; $wm.RelativeVerticalPosition = 1; $wm.Left = -999995; $wm.Top = -999995
  }

  Attempt 'title property' { $doc.BuiltInDocumentProperties.Item('Title').Value = 'The Rutba Office showcase' }
  Attempt 'author property' { $doc.BuiltInDocumentProperties.Item('Author').Value = 'Rutba Office' }
  "stage: table of contents"
  Attempt 'table of contents' { [void]$doc.TablesOfContents.Add($doc.Bookmarks.Item('TOCHere').Range, $true, 1, 3) }

  "stage: taking the document as Flat OPC"
  # Word's SaveAs never returns from an automated session on this machine —
  # hidden or minimised, add-ins on or off, background save off, SaveAs or
  # SaveAs2, a new document or an opened one, any format. WordOpenXML is the
  # same package as one XML string and involves no save at all, so the
  # document is taken that way and packed by tools/flat-opc-to-docx.mjs.
  # The .odt, .rtf and .doc that SaveAs2 would have written are not made;
  # the corpus carries those from other hands.
  $flat = Join-Path $env:TEMP 'rich-showcase-flat.xml'
  [IO.File]::WriteAllText($flat, $doc.WordOpenXML, [Text.Encoding]::UTF8)
  "flat opc: $((Get-Item $flat).Length) bytes"
  $docx = Join-Path $outDir 'showcase.docx'; Remove-Existing $docx
  & node (Join-Path $root 'tools\flat-opc-to-docx.mjs') $flat $docx
  if (Test-Path $docx) { "showcase.docx written ($((Get-Item $docx).Length) bytes)" } else { "warn: showcase.docx was not packed" }
  Remove-Item $flat -ErrorAction SilentlyContinue
  $doc.Close(0)
} finally {
  if ($wd) { try { $wd.Quit() } catch {} ; Release $wd }
  Reap
}

# ═══════════════════════════════════════════════════════════════════════════
# PowerPoint: the deck
# ═══════════════════════════════════════════════════════════════════════════
$pp = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  Attempt 'add-ins off' { foreach ($a in $pp.COMAddIns) { if ($a.Connect) { $a.Connect = $false } } }
  # With a window: a chart cannot be added to a presentation that has none
  # (AddChart and AddChart2 fail with E_FAIL), so PowerPoint shows itself
  # for the length of this section, minimised at once. The one window this
  # script puts on the desktop.
  $pres = $pp.Presentations.Add(-1)
  Attempt 'minimise' { $pp.WindowState = 2 }
  $pres.PageSetup.SlideWidth = 960; $pres.PageSetup.SlideHeight = 540
  $layouts = $pres.SlideMaster.CustomLayouts
  function Layout($name) { foreach ($l in $layouts) { if ($l.Name -eq $name) { return $l } }; return $layouts.Item(1) }
  $titleLayout = Layout 'Title Slide'; $contentLayout = Layout 'Title and Content'; $blankLayout = Layout 'Blank'; $titleOnly = Layout 'Title Only'

  # 1 — title, on a gradient
  $s1 = $pres.Slides.AddSlide(1, $titleLayout)
  Attempt 'gradient background' { $s1.FollowMasterBackground = 0; $s1.Background.Fill.TwoColorGradient(4, 1) }
  Attempt 'gradient colours' { $s1.Background.Fill.ForeColor.RGB = RGB 31 78 121; $s1.Background.Fill.BackColor.RGB = RGB 192 0 0 }
  $s1.Shapes.Item(1).TextFrame.TextRange.Text = 'The Rutba Office showcase'
  $s1.Shapes.Item(1).TextFrame.TextRange.Font.Color.RGB = RGB 255 255 255
  $s1.Shapes.Item(2).TextFrame.TextRange.Text = "A deck with one of everything, written by PowerPoint on $stamp"
  $s1.Shapes.Item(2).TextFrame.TextRange.Font.Color.RGB = RGB 255 230 153
  Attempt 'notes 1' { $s1.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = 'Speaker notes for the title slide: welcome everyone, then move on.' }

  # 2 — bullets at three levels, and a picture
  $s2 = $pres.Slides.AddSlide(2, $contentLayout)
  $s2.Shapes.Item(1).TextFrame.TextRange.Text = 'Bullets at three levels'
  $body = $s2.Shapes.Item(2).TextFrame.TextRange
  $body.Text = "First point`rSecond point`rA sub-point`rA sub-sub-point`rThird point, in colour"
  $body.Paragraphs(3).IndentLevel = 2; $body.Paragraphs(4).IndentLevel = 3
  $body.Paragraphs(5).Font.Color.RGB = RGB 192 0 0; $body.Paragraphs(5).Font.Bold = -1
  $s2.Shapes.Item(2).Width = 520
  [void]$s2.Shapes.AddPicture($png, 0, -1, 620, 160, 300, 200)
  Attempt 'notes 2' { $s2.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = 'Notes on slide two: the picture is generated, not fetched.' }

  # 3 — the shapes gallery
  $s3 = $pres.Slides.AddSlide(3, $titleOnly)
  $s3.Shapes.Item(1).TextFrame.TextRange.Text = 'Shapes: fills, gradients, outlines, shadows, text, a group, a rotation'
  $kinds = @(@(1, 'Rectangle'), @(5, 'Rounded'), @(9, 'Oval'), @(7, 'Triangle'), @(92, 'Star'), @(33, 'Arrow'), @(52, 'Chevron'), @(10, 'Hexagon'), @(21, 'Heart'), @(17, 'Smile'), @(61, 'Process'), @(105, 'Callout'), @(179, 'Cloud'), @(35, 'Up arrow'), @(4, 'Diamond'), @(13, 'Parallelogram'))
  $palette = @((RGB 192 0 0), (RGB 237 125 49), (RGB 255 192 0), (RGB 112 173 71), (RGB 68 114 196), (RGB 112 48 160), (RGB 0 176 240), (RGB 255 0 102), (RGB 0 128 128), (RGB 128 96 0), (RGB 91 155 213), (RGB 165 165 165), (RGB 255 153 204), (RGB 0 102 204), (RGB 153 204 0), (RGB 204 102 0))
  for ($i = 0; $i -lt $kinds.Count; $i++) {
    $sh = $s3.Shapes.AddShape($kinds[$i][0], 30 + (($i % 8) * 115), 130 + ([Math]::Floor($i / 8) * 150), 100, 110)
    $sh.Name = 'Shape ' + $kinds[$i][1]
    $sh.Fill.ForeColor.RGB = $palette[$i]
    if ($i % 3 -eq 1) { $sh.Fill.TwoColorGradient(1, 1); $sh.Fill.BackColor.RGB = RGB 255 255 255 }
    if ($i % 5 -eq 4) { $sh.Fill.Transparency = 0.4 }
    $sh.Line.ForeColor.RGB = RGB 30 30 30; $sh.Line.Weight = 1.5
    if ($i % 4 -eq 3) { $sh.Shadow.Visible = -1 }
    if ($i -eq 15) { $sh.Rotation = 25 }
    $sh.TextFrame.TextRange.Text = $kinds[$i][1]
    $sh.TextFrame.TextRange.Font.Size = 12; $sh.TextFrame.TextRange.Font.Bold = -1; $sh.TextFrame.TextRange.Font.Color.RGB = RGB 255 255 255
  }
  Attempt 'shape hyperlink' { $s3.Shapes.Item('Shape Heart').ActionSettings.Item(1).Hyperlink.Address = 'https://office.rutba.io' }
  Attempt 'group' { $grp = $s3.Shapes.Range([object[]]@('Shape Cloud', 'Shape Up arrow')).Group(); $grp.Name = 'Grouped pair' }
  Attempt 'connector' { [void]$s3.Shapes.AddConnector(2, 130, 185, 145, 185) }
  Attempt 'transition 3' { $s3.SlideShowTransition.EntryEffect = 3844 }
  Attempt 'notes 3' { $s3.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = 'Sixteen preset shapes, two of them grouped, one rotated, one a hyperlink.' }

  # 4 — a table
  $s4 = $pres.Slides.AddSlide(4, $titleOnly)
  $s4.Shapes.Item(1).TextFrame.TextRange.Text = 'A table with a style'
  $tbl = $s4.Shapes.AddTable(5, 4, 80, 140, 800, 300)
  $cells = @(@('Region', 'Q1', 'Q2', 'Total'), @('North', '1,200', '1,350', '2,550'), @('South', '980', '1,120', '2,100'), @('East', '1,430', '1,510', '2,940'), @('All', '3,610', '3,980', '7,590'))
  for ($r = 0; $r -lt 5; $r++) { for ($c = 0; $c -lt 4; $c++) { $tbl.Table.Cell($r + 1, $c + 1).Shape.TextFrame.TextRange.Text = $cells[$r][$c] } }
  $tbl.Table.Cell(5, 4).Shape.TextFrame.TextRange.Font.Bold = -1
  $tbl.Table.Cell(5, 4).Shape.Fill.ForeColor.RGB = RGB 255 230 153

  # 5 — charts
  $s5 = $pres.Slides.AddSlide(5, $titleOnly)
  $s5.Shapes.Item(1).TextFrame.TextRange.Text = 'Charts: a column chart and a pie'
  Attempt 'column chart' { $c1 = $s5.Shapes.AddChart2(-1, 51, 40, 130, 440, 360); $c1.Chart.HasTitle = $true; $c1.Chart.ChartTitle.Text = 'Columns'; $c1.Chart.ChartData.Workbook.Close() }
  Attempt 'pie chart' { $c2 = $s5.Shapes.AddChart2(-1, 5, 500, 130, 420, 360); $c2.Chart.HasTitle = $true; $c2.Chart.ChartTitle.Text = 'A pie'; $c2.Chart.ChartData.Workbook.Close() }
  Attempt 'notes 5' { $s5.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = 'The chart data is what PowerPoint puts in by default.' }

  # 6 — WordArt, a formatted text box, a callout
  $s6 = $pres.Slides.AddSlide(6, $blankLayout)
  $s6.FollowMasterBackground = 0; $s6.Background.Fill.Solid(); $s6.Background.Fill.ForeColor.RGB = RGB 255 242 204
  Attempt 'wordart' {
    $wa = $s6.Shapes.AddTextEffect(0, 'WordArt', 'Calibri', 60, -1, 0, 60, 40)
    $wa.Fill.ForeColor.RGB = RGB 192 0 0; $wa.Line.ForeColor.RGB = RGB 31 78 121
  }
  $tb = $s6.Shapes.AddTextbox(1, 60, 180, 520, 300)
  $tr = $tb.TextFrame.TextRange
  $tr.Text = "A text box with formatted paragraphs.`rThis one is bold and blue.`rThis one is italic, larger, and right-aligned.`rThis one is bulleted.`rAnd this one is bulleted too."
  $tr.Paragraphs(2).Font.Bold = -1; $tr.Paragraphs(2).Font.Color.RGB = RGB 0 112 192
  $tr.Paragraphs(3).Font.Italic = -1; $tr.Paragraphs(3).Font.Size = 24; $tr.Paragraphs(3).ParagraphFormat.Alignment = 3
  Attempt 'bullets' { $tr.Paragraphs(4).ParagraphFormat.Bullet.Visible = -1; $tr.Paragraphs(5).ParagraphFormat.Bullet.Visible = -1 }
  $tb.Fill.ForeColor.RGB = RGB 255 255 255; $tb.Line.ForeColor.RGB = RGB 191 143 0; $tb.Line.Weight = 2
  $co = $s6.Shapes.AddShape(105, 620, 200, 280, 120); $co.Fill.ForeColor.RGB = RGB 226 239 218
  $co.TextFrame.TextRange.Text = 'A callout pointing at the text box'; $co.TextFrame.TextRange.Font.Color.RGB = RGB 0 0 0
  Attempt 'transition 6' { $s6.SlideShowTransition.EntryEffect = 3844 }

  # 7 — animations, hidden from the show
  $s7 = $pres.Slides.AddSlide(7, $titleOnly)
  $s7.Shapes.Item(1).TextFrame.TextRange.Text = 'Animated, and hidden from the show'
  $a1 = $s7.Shapes.AddShape(1, 100, 180, 220, 120); $a1.Fill.ForeColor.RGB = RGB 68 114 196; $a1.TextFrame.TextRange.Text = 'Flies in'
  $a2 = $s7.Shapes.AddShape(9, 400, 180, 220, 120); $a2.Fill.ForeColor.RGB = RGB 237 125 49; $a2.TextFrame.TextRange.Text = 'Fades in'
  $a3 = $s7.Shapes.AddShape(92, 700, 160, 160, 160); $a3.Fill.ForeColor.RGB = RGB 255 192 0; $a3.TextFrame.TextRange.Text = 'Zooms'
  Attempt 'animations' {
    [void]$s7.TimeLine.MainSequence.AddEffect($a1, 2, 0, 1)
    [void]$s7.TimeLine.MainSequence.AddEffect($a2, 10, 0, 1)
    [void]$s7.TimeLine.MainSequence.AddEffect($a3, 23, 0, 2)
    $s7.SlideShowTransition.Hidden = -1
  }

  # 8 — closing
  $s8 = $pres.Slides.AddSlide(8, $contentLayout)
  $s8.Shapes.Item(1).TextFrame.TextRange.Text = 'Thank you'
  $s8.Shapes.Item(2).TextFrame.TextRange.Text = "Eight slides, sixteen shapes, two charts, a table, a picture, WordArt, animations, notes, a hidden slide, sections, a footer and slide numbers.`rNon-Latin: 日本語 • العربية • Ελληνικά"
  Attempt 'notes 8' { $s8.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = 'Closing notes: thank the audience and take questions.' }

  Attempt 'footers and sections' {
    $hf = $pres.SlideMaster.HeadersFooters
    $hf.SlideNumber.Visible = -1; $hf.Footer.Visible = -1; $hf.Footer.Text = 'Rutba Office showcase'
    foreach ($sl in $pres.Slides) { try { $sl.HeadersFooters.SlideNumber.Visible = -1; $sl.HeadersFooters.Footer.Visible = -1; $sl.HeadersFooters.Footer.Text = 'Rutba Office showcase' } catch {} }
    [void]$pres.SectionProperties.AddBeforeSlide(1, 'Opening')
    [void]$pres.SectionProperties.AddBeforeSlide(3, 'Content')
    [void]$pres.SectionProperties.AddBeforeSlide(8, 'Closing')
  }

  $pptx = Join-Path $outDir 'showcase.pptx'; Remove-Existing $pptx; $pres.SaveAs($pptx, 24); "showcase.pptx written ($((Get-Item $pptx).Length) bytes)"
  $odp = Join-Path $outDir 'showcase.odp'; Remove-Existing $odp; $pres.SaveAs($odp, 35); "showcase.odp written ($((Get-Item $odp).Length) bytes)"
  $pres.Close()
} finally {
  if ($pp) { try { $pp.Quit() } catch {} ; Release $pp }
  Reap
}
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
"done: " + ((Get-ChildItem $outDir | ForEach-Object { $_.Name + ' ' + $_.Length }) -join ', ')
