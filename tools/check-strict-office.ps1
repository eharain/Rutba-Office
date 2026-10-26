# Word, Excel and PowerPoint reopen Rutba Office's Strict saves: each rich
# fixture saved by Rutba Office as Strict Open XML, and each of Office's own
# Strict fixtures edited here and saved (tools/strict-saves.mjs), opened in
# the application that made the original, with what it finds set beside what
# it finds in the Transitional original and in its own Strict file.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\check-strict-office.ps1 [-Dir out\strict-saves]
#
# Run through the machine's window lock. Each application runs in a
# PowerShell of its own, given three minutes, and any Office process this run
# started is ended at the end.
param([string]$Dir = 'out\strict-saves')

# Continue: Node's warnings on stderr would otherwise end the run here.
$ErrorActionPreference = 'Continue'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$out = if ([IO.Path]::IsPathRooted($Dir)) { $Dir } else { Join-Path $root $Dir }
& node (Join-Path $root 'tools\strict-saves.mjs') $out 2>$null
$officeBefore = @(Get-Process -Name WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$failed = 0

foreach ($pair in @(@('word', 'docx'), @('excel', 'xlsx'), @('powerpoint', 'pptx'))) {
  $app, $ext = $pair
  $files = @(
    (Join-Path $root "tests\fixtures\rich\showcase.$ext"),
    (Join-Path $root "tests\fixtures\strict\showcase-strict.$ext"),
    (Join-Path $out "saved-as-strict.$ext"),
    (Join-Path $out "strict-edited.$ext")
  )
  $log = Join-Path $env:TEMP "rutba-strict-$app.log"
  $child = Start-Process powershell -ArgumentList (@('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $PSScriptRoot 'check-strict-app.ps1'), '-App', $app, '-Files') + ($files -join ',')) -PassThru -WindowStyle Hidden -RedirectStandardOutput $log
  if (-not $child.WaitForExit(180000)) { Stop-Process -Id $child.Id -Force; "warn: $app did not finish in three minutes"; $failed++ }
  $lines = @(Get-Content $log -ErrorAction SilentlyContinue | Where-Object { $_ -match '^\{' } | ForEach-Object { $_ | ConvertFrom-Json })
  Remove-Item $log -ErrorAction SilentlyContinue
  foreach ($p in (Get-Process -Name WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue)) {
    if ($officeBefore -notcontains $p.Id) { try { Stop-Process -Id $p.Id -Force -ErrorAction Stop } catch {} }
  }
  "== $app"
  $lines | ForEach-Object { $_ | ConvertTo-Json -Compress }
  $reference = $lines | Where-Object { $_.file -eq "showcase.$ext" } | Select-Object -First 1
  foreach ($line in $lines | Where-Object { $_.file -ne "showcase.$ext" }) {
    if ($line.error) { "FAIL $($line.file): $($line.error)"; $failed++; continue }
    $differs = @()
    foreach ($prop in $reference.PSObject.Properties) {
      if ($prop.Name -in @('file', 'name', 'format', 'paragraphs', 'notes')) { continue }
      if ("$($line.($prop.Name))" -ne "$($prop.Value)") { $differs += "$($prop.Name) $($line.($prop.Name)) for $($prop.Value)" }
    }
    if ($differs.Count) { "FAIL $($line.file): $($differs -join '; ')"; $failed++ } else { "ok   $($line.file): opens with what showcase.$ext has" }
  }
}
if ($failed) { "$failed check(s) failed"; exit 1 } else { 'every Strict save reopens in the application that made its original' }
