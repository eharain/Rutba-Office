# Word, Excel and PowerPoint reopen what Rutba Office saves: each rich
# fixture saved by Rutba Office as it is, saved after an edit, and saved as
# Strict Open XML, and each of Office's own Strict fixtures edited here and
# saved (tools/strict-saves.mjs). Each is opened in the application that made
# the original, and what it finds is set beside what it finds in the
# original; in an edited save it must also find the edit. The gate runs it as
# its office:reopen pass.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\check-strict-office.ps1 [-Dir out\strict-saves]
#
# Run through the machine's window lock. Each application runs in a
# PowerShell of its own, given three minutes. Only the copies of Word, Excel
# and PowerPoint started for automation (-Embedding) that were not running
# before are ended at the end: one a person opens meanwhile is left alone.
param([string]$Dir = 'out\strict-saves')

# Continue: Node's warnings on stderr would otherwise end the run here.
$ErrorActionPreference = 'Continue'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$out = if ([IO.Path]::IsPathRooted($Dir)) { $Dir } else { Join-Path $root $Dir }
& node (Join-Path $root 'tools\strict-saves.mjs') $out 2>$null
if ($LASTEXITCODE -ne 0) { "FAIL the saves could not be made (tools/strict-saves.mjs ended with $LASTEXITCODE)"; exit 1 }
function Automation { @(Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE' OR Name='EXCEL.EXE' OR Name='POWERPNT.EXE'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -match '(-|/)Embedding' }) }
$officeBefore = @(Automation | ForEach-Object { $_.ProcessId })
$failed = 0
$EDIT = 'Edited in Rutba Office'

foreach ($pair in @(@('word', 'docx'), @('excel', 'xlsx'), @('powerpoint', 'pptx'))) {
  $app, $ext = $pair
  $files = @(
    (Join-Path $root "tests\fixtures\rich\showcase.$ext"),
    (Join-Path $out "saved.$ext"),
    (Join-Path $out "edited.$ext"),
    (Join-Path $root "tests\fixtures\strict\showcase-strict.$ext"),
    (Join-Path $out "saved-as-strict.$ext"),
    (Join-Path $out "strict-edited.$ext")
  )
  $log = Join-Path $env:TEMP "rutba-strict-$app.log"
  $child = Start-Process powershell -ArgumentList (@('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $PSScriptRoot 'check-strict-app.ps1'), '-App', $app, '-Files') + ($files -join ',')) -PassThru -WindowStyle Hidden -RedirectStandardOutput $log
  if (-not $child.WaitForExit(180000)) { Stop-Process -Id $child.Id -Force; "FAIL $app did not finish in three minutes"; $failed++ }
  $lines = @(Get-Content $log -ErrorAction SilentlyContinue | Where-Object { $_ -match '^\{' } | ForEach-Object { $_ | ConvertFrom-Json })
  $said = @(Get-Content $log -ErrorAction SilentlyContinue | Where-Object { $_ -notmatch '^\{' -and $_.Trim() })
  Remove-Item $log -ErrorAction SilentlyContinue
  foreach ($p in (Automation)) {
    if ($officeBefore -notcontains $p.ProcessId) { try { Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop } catch {} }
  }
  "== $app"
  $lines | ForEach-Object { $_ | ConvertTo-Json -Compress }
  # An application that answered for fewer files than it was given (not installed, or stopped) is a failure, not a pass of none.
  foreach ($f in $files) {
    if (-not ($lines | Where-Object { $_.file -eq (Split-Path -Leaf $f) })) { "FAIL $(Split-Path -Leaf $f): $app gave no answer for it$(if ($said) { " ($($said[-1]))" })"; $failed++ }
  }
  $reference = $lines | Where-Object { $_.file -eq "showcase.$ext" } | Select-Object -First 1
  if (-not $reference -or $reference.error) { continue }
  foreach ($line in $lines | Where-Object { $_.file -ne "showcase.$ext" }) {
    if ($line.error) { "FAIL $($line.file): $($line.error)"; $failed++; continue }
    $differs = @()
    foreach ($prop in $reference.PSObject.Properties) {
      if ($prop.Name -in @('file', 'name', 'format', 'paragraphs', 'notes', 'start', 'edited')) { continue }
      if ("$($line.($prop.Name))" -ne "$($prop.Value)") { $differs += "$($prop.Name) $($line.($prop.Name)) for $($prop.Value)" }
    }
    # An edited save shows the edit: the first paragraph, cell A31 of Sales, the first slide's notes.
    if ($line.file -match 'edited' -and "$($line.start)$($line.edited)$($line.notes)" -notmatch $EDIT) { $differs += "the edit is not there" }
    if ($differs.Count) { "FAIL $($line.file): $($differs -join '; ')"; $failed++ } else { "ok   $($line.file): opens with what showcase.$ext has" }
  }
}
if ($failed) { "$failed check(s) failed"; exit 1 } else { 'every save reopens in the application that made its original' }
