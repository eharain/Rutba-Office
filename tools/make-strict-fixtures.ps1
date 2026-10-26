# Make the Strict fixtures with Office itself: the showcase document,
# workbook and deck (tools/make-rich-fixtures.ps1, written by Word, Excel and
# PowerPoint) saved again by the same applications as ISO/IEC 29500 Strict
# Open XML, so the Strict reader is judged against what Office writes and
# against the same content as the Transitional fixtures.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\make-strict-fixtures.ps1 [-Out tests\fixtures\strict]
#
# Office is driven invisibly; each application is quit and the processes this
# run started are reaped at the end, even when a step fails. Word's Save As is
# written at the top level, not in a function or script block, where it has
# been seen to wait for ever in an automated session (make-binary-fixtures.ps1).
param([string]$Out = 'tests\fixtures\strict')

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$outDir = if ([IO.Path]::IsPathRooted($Out)) { $Out } else { Join-Path $root $Out }
$rich = Join-Path $root 'tests\fixtures\rich'
New-Item -ItemType Directory -Force $outDir | Out-Null
function Release($o) { if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} } }
function Remove-Existing($p) { if (Test-Path $p) { Remove-Item $p -Force } }
$officeBefore = @(Get-Process -Name WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
function Reap {
  Start-Sleep -Seconds 1
  foreach ($p in (Get-Process -Name WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue)) {
    if ($officeBefore -notcontains $p.Id) { try { Stop-Process -Id $p.Id -Force -ErrorAction Stop; "reaped $($p.ProcessName) $($p.Id)" } catch {} }
  }
}

# Word: wdFormatStrictOpenXMLDocument is 24, saved by make-strict-word.ps1 in
# a PowerShell of its own, given three minutes, the Word it started reaped.
$target = Join-Path $outDir 'showcase-strict.docx'
$log = Join-Path $env:TEMP 'rutba-strict-word.log'
$child = Start-Process powershell -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $PSScriptRoot 'make-strict-word.ps1'), '-In', (Join-Path $rich 'showcase.docx'), '-Out', $target, '-Format', 24 -PassThru -WindowStyle Hidden -RedirectStandardOutput $log
if ($child.WaitForExit(180000)) { Get-Content $log } else { Stop-Process -Id $child.Id -Force; "warn: showcase-strict.docx - Word did not finish saving in three minutes" }
Remove-Item $log -ErrorAction SilentlyContinue
Reap

# Excel: xlOpenXMLStrictWorkbook is 61.
$xl = $null
try {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $p = Join-Path $outDir 'showcase-strict.xlsx'
  Remove-Existing $p
  $wb = $xl.Workbooks.Open((Join-Path $rich 'showcase.xlsx'), 0, $true)
  $wb.SaveAs($p, 61)
  $wb.Close($false)
  "showcase-strict.xlsx written by Excel ($((Get-Item $p).Length) bytes)"
} catch {
  "warn: Excel - $($_.Exception.Message)"
} finally {
  if ($xl) { try { $xl.Quit() } catch {}; Release $xl }
  Reap
}

# PowerPoint: ppSaveAsStrictOpenXMLPresentation is 38.
$pp = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $p = Join-Path $outDir 'showcase-strict.pptx'
  Remove-Existing $p
  $pres = $pp.Presentations.Open((Join-Path $rich 'showcase.pptx'), -1, 0, 0)
  $pres.SaveAs($p, 38)
  $pres.Close()
  "showcase-strict.pptx written by PowerPoint ($((Get-Item $p).Length) bytes)"
} catch {
  "warn: PowerPoint - $($_.Exception.Message)"
} finally {
  if ($pp) { try { $pp.Quit() } catch {}; Release $pp }
  Reap
}
