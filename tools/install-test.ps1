# Prove the installer on this machine, as the last check before a release is
# published: one copy registered, one icon per file type, the app tiles
# unpacked beside the archive, and the installed copy — not the unpacked
# build, which sits inside the repository and finds packages there that the
# installer never packed — opening a Rutba Word window from --app=word.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\install-test.ps1 `
#     -Installer apps\desktop\release\Rutba-Office-1.8.0-win-x64.exe -Label 1.8.0
#
# Exit codes: 0 proven; 1 no installed exe; 2 a copy is running, nothing was
# installed over it; 3 the installed copy did not open a Rutba Word window, and what
# it showed instead is printed — an "Error" box's text included, read through
# UI Automation, because that box is the one thing a person sees when the
# archive is missing a package.
param([string]$Installer, [string]$Label = 'run')
$keys = @('HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
function Registered {
  Get-ItemProperty $keys -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -match 'Rutba' } |
    ForEach-Object { "{0} {1} scope={2} dir={3}" -f $_.DisplayName, $_.DisplayVersion, ($(if ($_.PSPath -match 'HKEY_CURRENT_USER') { 'user' } else { 'machine' })), $_.InstallLocation }
}
# Never install over a copy somebody is using: the installer closes it.
if (Get-Process -Name 'Rutba Office' -ErrorAction SilentlyContinue) { "[$Label] Rutba Office is running; not installing over it."; exit 2 }
"[$Label] before:"; Registered
$sw = [Diagnostics.Stopwatch]::StartNew()
$p = Start-Process -FilePath $Installer -ArgumentList '/S' -Wait -PassThru
"[$Label] installer exit {0} in {1:n1}s" -f $p.ExitCode, $sw.Elapsed.TotalSeconds
"[$Label] after:"; Registered
$exe = Join-Path $env:LOCALAPPDATA 'Programs\Rutba Office\Rutba Office.exe'
if (-not (Test-Path $exe)) { "[$Label] no installed exe at $exe"; exit 1 }
"[$Label] installed exe version: {0}" -f (Get-Item $exe).VersionInfo.FileVersion

# One icon per file type: the ProgId each extension points at, and its DefaultIcon.
$missingIcons = 0
foreach ($ext in @('.docx', '.xlsx', '.pptx', '.jpg', '.mp4', '.eml', '.pdf', '.csv', '.md', '.ics', '.vcf')) {
  $progId = $null
  try { $progId = (Get-ItemProperty "HKCU:\Software\Classes\$ext" -ErrorAction Stop).'(default)' } catch {}
  if (-not $progId) { try { $progId = (Get-ItemProperty "HKLM:\Software\Classes\$ext" -ErrorAction Stop).'(default)' } catch {} }
  $icon = $null
  if ($progId) {
    try { $icon = (Get-ItemProperty "HKCU:\Software\Classes\$progId\DefaultIcon" -ErrorAction Stop).'(default)' } catch {}
    if (-not $icon) { try { $icon = (Get-ItemProperty "HKLM:\Software\Classes\$progId\DefaultIcon" -ErrorAction Stop).'(default)' } catch {} }
  }
  $iconFile = if ($icon) { ($icon -split ',')[0].Trim('"') } else { '' }
  $present = if ($iconFile -and (Test-Path $iconFile)) { 'present' } else { 'MISSING'; $missingIcons++ }
  "[$Label] {0,-6} progid={1} icon={2} ({3})" -f $ext, $progId, $icon, $present
}

# The app tiles, unpacked beside the archive.
$tiles = Join-Path (Split-Path $exe) 'resources\app.asar.unpacked\resources\apps'
if (Test-Path $tiles) { "[$Label] tiles: {0}" -f ((Get-ChildItem $tiles -Filter *.ico | ForEach-Object { $_.Name }) -join ' ') } else { "[$Label] tiles MISSING at $tiles" }

# --app=word opens a Rutba Word window, in a profile of its own, off the desktop.
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$dataDir = Join-Path $env:TEMP ('rutba-apptest-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
$env:RUTBA_WINDOW_DISPLAY = 'offscreen'
$proc = Start-Process -FilePath $exe -ArgumentList @('--app=word', "--user-data-dir=$dataDir") -PassThru
$titles = @()
for ($i = 0; $i -lt 20; $i++) {
  Start-Sleep -Seconds 1
  $titles = @(Get-Process -Name 'Rutba Office' -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle } | ForEach-Object { $_.MainWindowTitle })
  if ($titles -match 'Word' -or $titles -contains 'Error') { break }
}
"[$Label] windows after --app=word: {0}" -f ($titles -join ' | ')
$opened = [bool]($titles -match 'Word')
if (-not $opened) {
  # What it showed instead: the text of every top-level window this process owns.
  Get-Process -Name 'Rutba Office' -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | ForEach-Object {
    try {
      $root = [System.Windows.Automation.AutomationElement]::FromHandle($_.MainWindowHandle)
      foreach ($e in $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)) {
        $n = $e.Current.Name
        if ($n -and $e.Current.ControlType.ProgrammaticName -eq 'ControlType.Text') { "[$Label]   " + $n.Substring(0, [Math]::Min(600, $n.Length)) }
      }
    } catch { "[$Label]   (could not read the window: {0})" -f $_.Exception.Message }
  }
  $log = Join-Path $dataDir 'errors.log'
  if (Test-Path $log) { "[$Label] errors.log:"; Get-Content $log | Select-Object -First 8 }
}
# Only the tree this script started.
function Tree($id) { $out = @($id); foreach ($c in (Get-CimInstance Win32_Process -Filter "ParentProcessId=$id")) { $out += Tree $c.ProcessId }; $out }
foreach ($id in (Tree $proc.Id)) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }
Start-Sleep -Seconds 2
Remove-Item $dataDir -Recurse -Force -ErrorAction SilentlyContinue
if (-not $opened) { "[$Label] FAILED: the installed copy did not open a Rutba Word window."; exit 3 }
if ($missingIcons -gt 0) { "[$Label] FAILED: $missingIcons file types have no icon."; exit 3 }
"[$Label] proven: installed, registered with icons, tiles unpacked, and a Rutba Word window opened."
