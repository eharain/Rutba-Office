# What Word, Excel or PowerPoint finds in each of some files, one line of
# JSON a file: the format it says the file is in and the counts of what it
# holds, or the error opening it gave. Run by check-strict-office.ps1, in a
# PowerShell of its own with a time limit; every file is opened read only,
# invisibly, and closed unsaved.
param([string]$App, [string[]]$Files)

$ErrorActionPreference = 'Stop'
# Given as one argument, the names separated by commas.
$Files = @($Files | ForEach-Object { $_ -split ',' } | Where-Object { $_ })
function Release($o) { if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} } }

if ($App -eq 'word') {
  $w = New-Object -ComObject Word.Application
  $w.Visible = $false
  $w.DisplayAlerts = 0
  foreach ($a in $w.COMAddIns) { try { if ($a.Connect) { $a.Connect = $false } } catch {} }
  foreach ($f in $Files) {
    try {
      # ConfirmConversions off, read only, not in the recent list.
      $d = $w.Documents.Open($f, $false, $true, $false)
      [ordered]@{ file = (Split-Path -Leaf $f); format = $d.SaveFormat; paragraphs = $d.Paragraphs.Count; tables = $d.Tables.Count; pictures = $d.InlineShapes.Count;
        shapes = $d.Shapes.Count; headerShapes = $d.Sections.Item(1).Headers.Item(1).Shapes.Count; comments = $d.Comments.Count; footnotes = $d.Footnotes.Count; endnotes = $d.Endnotes.Count; sections = $d.Sections.Count } | ConvertTo-Json -Compress
      $d.Close(0)
    } catch { [ordered]@{ file = (Split-Path -Leaf $f); error = $_.Exception.Message } | ConvertTo-Json -Compress }
  }
  $w.Quit()
  Release $w
}

if ($App -eq 'excel') {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  foreach ($f in $Files) {
    try {
      $wb = $xl.Workbooks.Open($f, 0, $true)
      $charts = 0; $shapes = 0; $comments = 0; $tables = 0
      foreach ($ws in $wb.Worksheets) { $charts += $ws.ChartObjects().Count; $shapes += $ws.Shapes.Count; $comments += $ws.Comments.Count; $tables += $ws.ListObjects.Count }
      $types = $wb.Worksheets.Item('Data types')
      [ordered]@{ file = (Split-Path -Leaf $f); name = $wb.Name; format = $wb.FileFormat; sheets = $wb.Worksheets.Count; charts = $charts; shapes = $shapes; comments = $comments; tables = $tables;
        total = $wb.Worksheets.Item('Sales').Range('F14').Value2; date = $types.Range('B6').Value2; dateText = $types.Range('B6').Text; time = $types.Range('B7').Value2 } | ConvertTo-Json -Compress
      $wb.Close($false)
    } catch { [ordered]@{ file = (Split-Path -Leaf $f); error = $_.Exception.Message } | ConvertTo-Json -Compress }
  }
  $xl.Quit()
  Release $xl
}

if ($App -eq 'powerpoint') {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  foreach ($f in $Files) {
    try {
      # Read only, untitled off, no window.
      $p = $pp.Presentations.Open($f, -1, 0, 0)
      $shapes = 0; $charts = 0; $tables = 0; $pictures = 0
      foreach ($s in $p.Slides) { foreach ($sh in $s.Shapes) { $shapes++; if ($sh.HasChart) { $charts++ }; if ($sh.HasTable) { $tables++ }; if ($sh.Type -eq 13) { $pictures++ } } }
      $notes = ''
      # The notes page's body placeholder (ppPlaceholderBody, 2), wherever it is in the list.
      try { foreach ($ph in $p.Slides.Item(1).NotesPage.Shapes.Placeholders) { if ($ph.PlaceholderFormat.Type -eq 2) { $notes = $ph.TextFrame.TextRange.Text } } } catch {}
      [ordered]@{ file = (Split-Path -Leaf $f); slides = $p.Slides.Count; shapes = $shapes; charts = $charts; tables = $tables; pictures = $pictures; notes = $notes } | ConvertTo-Json -Compress
      $p.Close()
    } catch { [ordered]@{ file = (Split-Path -Leaf $f); error = $_.Exception.Message } | ConvertTo-Json -Compress }
  }
  $pp.Quit()
  Release $pp
}
