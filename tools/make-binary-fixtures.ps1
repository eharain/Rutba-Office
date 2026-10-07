# Make the binary fixtures with Office itself: Word 97-2003 documents, and the
# showcase workbook and deck saved by Excel and PowerPoint in their 97-2003
# formats — and the workbook in every older Excel format this Excel still
# writes. The readers in packages/office-formats are judged against these,
# files Office wrote, not files written from what we think the formats say.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\make-binary-fixtures.ps1 [-Out tests\fixtures\binary] [-Only word,excel,powerpoint]
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
param([string]$Out = 'tests\fixtures\binary', [string]$Only = 'word,excel,powerpoint')
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
    $p = Join-Path $outDir $f[1]; Remove-Existing $p
    try { $wb.SaveAs($p, $f[0]); "$($f[1]) written as Excel $($f[2]) ($((Get-Item $p).Length) bytes)" } catch { "warn: Excel $($f[2]) - $($_.Exception.Message)" }
    $wb.Close($false)
  }
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
