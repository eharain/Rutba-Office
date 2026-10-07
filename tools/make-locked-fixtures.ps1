# Make the password-protected 97-2003 fixtures with Office itself: a Word
# document, workbooks and a deck saved with a password in their binary
# formats, so the readers' decryption is judged against files Office wrote.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\make-locked-fixtures.ps1 [-Out tests\fixtures\binary] [-Only word,excel,powerpoint]
#
# Every file's password is "Rutba-1". What each one is:
#
#   locked.doc       three paragraphs saved by Word 97-2003 with a password (RC4 CryptoAPI)
#   locked.xls       conditions.xls saved by Excel 97-2003 with a password (RC4 CryptoAPI, 128-bit)
#   locked-95.xls    the same saved as Excel 5.0/95, which hides it with XOR obfuscation
#   locked.ppt       a two-slide deck with a picture, saved by PowerPoint 97-2003 with a password
#
# Office is driven invisibly; the processes this run started are reaped at
# the end of each section, even when a step fails. Word's document is made
# by make-locked-word.ps1 in a PowerShell of its own: Word's Save As hangs
# here when it is driven from inside a script's blocks (see
# make-binary-fixtures.ps1).
param([string]$Out = 'tests\fixtures\binary', [string]$Only = 'word,excel,powerpoint')
$run = $Only.Split(',')
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$outDir = if ([IO.Path]::IsPathRooted($Out)) { $Out } else { Join-Path $root $Out }
$rich = Join-Path $root 'tests\fixtures\rich'
$password = 'Rutba-1'
$officeBefore = @(Get-Process -Name WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
function Release($o) { if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} } }
function Remove-Existing($p) { if (Test-Path $p) { Remove-Item $p -Force } }
function Reap {
  Start-Sleep -Seconds 1
  foreach ($p in (Get-Process -Name WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue)) {
    if ($officeBefore -notcontains $p.Id) { try { Stop-Process -Id $p.Id -Force -ErrorAction Stop; "reaped $($p.ProcessName) $($p.Id)" } catch {} }
  }
}

# ── Word ────────────────────────────────────────────────────────────────────
if ($run -contains 'word') {
  $target = Join-Path $outDir 'locked.doc'
  $log = Join-Path $env:TEMP 'rutba-locked-word.log'
  $child = Start-Process powershell -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $PSScriptRoot 'make-locked-word.ps1'), '-Out', $target -PassThru -WindowStyle Hidden -RedirectStandardOutput $log
  if ($child.WaitForExit(180000)) { Get-Content $log } else { Stop-Process -Id $child.Id -Force; 'warn: locked.doc - Word did not finish saving in three minutes' }
  Remove-Item $log -ErrorAction SilentlyContinue
  Reap
}

# ── Excel ───────────────────────────────────────────────────────────────────
$xl = $null
if ($run -contains 'excel') { try {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $source = Join-Path $outDir 'conditions.xls'
  foreach ($f in @(@(56, 'locked.xls'), @(39, 'locked-95.xls'))) {
    $wb = $xl.Workbooks.Open($source, 0, $true)
    $p = Join-Path $outDir $f[1]; Remove-Existing $p
    try {
      $wb.SaveAs($p, $f[0], $password)
      "$($f[1]) written ($((Get-Item $p).Length) bytes)"
    } catch { "warn: $($f[1]) - $($_.Exception.Message)" }
    $wb.Close($false)
  }
} catch { "warn: Excel - $($_.Exception.Message)" } finally {
  if ($xl) { try { $xl.Quit() } catch {}; Release $xl }
  Reap
} }

# ── PowerPoint ──────────────────────────────────────────────────────────────
$pp = $null
if ($run -contains 'powerpoint') { try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pres = $pp.Presentations.Add(0)
  $s1 = $pres.Slides.Add(1, 1)
  $s1.Shapes.Item(1).TextFrame.TextRange.Text = 'A locked deck'
  $s1.Shapes.Item(2).TextFrame.TextRange.Text = 'Opened with its password'
  $s2 = $pres.Slides.Add(2, 2)
  $s2.Shapes.Item(1).TextFrame.TextRange.Text = 'Second slide'
  $s2.Shapes.Item(2).TextFrame.TextRange.Text = "First point`rSecond point"
  [void]$s2.Shapes.AddPicture((Join-Path $rich 'picture.png'), 0, -1, 500, 300, 160, 100)
  $target = Join-Path $outDir 'locked.ppt'
  Remove-Existing $target
  $pres.Password = $password
  $pres.SaveAs($target, 1)
  $pres.Close()
  "locked.ppt written ($((Get-Item $target).Length) bytes)"
} catch { "warn: PowerPoint - $($_.Exception.Message)" } finally {
  if ($pp) { try { $pp.Quit() } catch {}; Release $pp }
  Reap
} }
