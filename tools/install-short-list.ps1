# The installer with a short list of apps (docs/APP-PACKS.md 5.3), on a
# machine with no copy of Rutba Office: /PACKS=word,sheets writes only those
# two packs, registers only their kinds of file, records the choice, and the
# installed copy opens a Documents window. For a CI runner or a clean
# machine: it installs per user, over nothing.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\install-short-list.ps1 `
#     -Installer apps\desktop\release\Rutba-Office-2.20.0-win-x64.exe
#
# Exit codes: 0 proven; 1 not installed; 2 a copy is already installed or
# running, nothing was done; 3 a check failed, each failure printed.
param([string]$Installer, [string]$Packs = 'word,sheets')
$keys = @('HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
$before = @(Get-ItemProperty $keys -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -match 'Rutba' })
if ($before.Count -gt 0 -or (Get-Process -Name 'Rutba Office' -ErrorAction SilentlyContinue)) { 'A copy of Rutba Office is already here; this check is for a machine without one.'; exit 2 }

$p = Start-Process -FilePath $Installer -ArgumentList @('/S', "/PACKS=$Packs") -Wait -PassThru
"installer exit {0}" -f $p.ExitCode
$exe = Join-Path $env:LOCALAPPDATA 'Programs\Rutba Office\Rutba Office.exe'
if (-not (Test-Path $exe)) { "no installed exe at $exe"; exit 1 }
$failed = @()
$wanted = @($Packs -split ',' | Where-Object { $_ })

# Only the packs chosen are on disk.
$packDir = Join-Path (Split-Path $exe) 'resources\packs'
$found = @(Get-ChildItem $packDir -Filter '*.pack.json' -ErrorAction SilentlyContinue | ForEach-Object { $_.Name -replace '\.pack\.json$', '' } | Sort-Object)
"packs on disk: {0}" -f ($found -join ',')
foreach ($id in $wanted) { if ($found -notcontains $id) { $failed += "pack $id was chosen and is not on disk" } }
foreach ($id in $found) { if ($wanted -notcontains $id) { $failed += "pack $id was not chosen and is on disk" } }

# The choice recorded.
$chosen = (Get-ItemProperty 'HKCU:\Software\Rutba Office' -ErrorAction SilentlyContinue).Packs
"Packs: $chosen"
foreach ($id in $wanted) { if (",$chosen," -notmatch ",$id,") { $failed += "Packs does not name $id" } }

# The chosen apps' kinds of file registered, and no other app's.
function ProgId($ext) { try { (Get-ItemProperty "HKCU:\Software\Classes\$ext" -ErrorAction Stop).'(default)' } catch { $null } }
foreach ($ext in @('.docx', '.xlsx')) { $id = ProgId $ext; "$ext -> $id"; if (-not $id) { $failed += "$ext is not registered" } }
foreach ($ext in @('.pptx', '.eml', '.pdf', '.jpg', '.ics', '.vcf')) { $id = ProgId $ext; "$ext -> $id"; if ($id -match 'Rutba') { $failed += "$ext is registered to Rutba Office with its app not installed" } }

# The installed copy opens a Documents window.
$dataDir = Join-Path $env:TEMP ('rutba-shortlist-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
$env:RUTBA_WINDOW_DISPLAY = 'offscreen'
$proc = Start-Process -FilePath $exe -ArgumentList @('--app=word', "--user-data-dir=$dataDir") -PassThru
$titles = @()
for ($i = 0; $i -lt 30; $i++) {
  Start-Sleep -Seconds 1
  $titles = @(Get-Process -Name 'Rutba Office' -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle } | ForEach-Object { $_.MainWindowTitle })
  if ($titles -match 'Rutba Documents' -or $titles -contains 'Error') { break }
}
"windows after --app=word: {0}" -f ($titles -join ' | ')
if (-not ($titles -match 'Rutba Documents')) {
  $failed += 'the installed copy did not open a Documents window'
  $log = Join-Path $dataDir 'errors.log'
  if (Test-Path $log) { 'errors.log:'; Get-Content $log | Select-Object -First 12 }
}
function Tree($id) { $out = @($id); foreach ($c in (Get-CimInstance Win32_Process -Filter "ParentProcessId=$id")) { $out += Tree $c.ProcessId }; $out }
foreach ($id in (Tree $proc.Id)) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }

if ($failed.Count) { $failed | ForEach-Object { "FAILED: $_" }; exit 3 }
'proven: only the chosen apps installed and registered, and a Documents window opened.'
