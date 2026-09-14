# Prove that an installed copy updates itself, on this machine, against the
# release that was just published: install an OLDER release, launch it off the
# desktop, wait for its own updater to find and download the current release,
# close it the way a person does, and read the version the installer left
# behind. Nothing is simulated - the copy under test asks GitHub, downloads
# the real installer and runs it, exactly as an installed copy anywhere does.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\update-test.ps1 `
#     -Older apps\desktop\release\Rutba-Office-1.11.0-win-x64.exe -Expect 1.12.0
#
# Exit codes: 0 proven; 1 no installed exe; 2 a copy is running, nothing was
# installed over it; 3 the update was not downloaded in time; 4 it was
# downloaded but the version after the quit is not the one expected.
param([string]$Older, [string]$Expect, [string]$Label = 'update')
$keys = @('HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
function Registered {
  Get-ItemProperty $keys -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -match 'Rutba' } |
    ForEach-Object { "{0} {1} scope={2}" -f $_.DisplayName, $_.DisplayVersion, ($(if ($_.PSPath -match 'HKEY_CURRENT_USER') { 'user' } else { 'machine' })) }
}
function Tree($id) { $out = @($id); foreach ($c in (Get-CimInstance Win32_Process -Filter "ParentProcessId=$id")) { $out += Tree $c.ProcessId }; $out }

# Never install over a copy somebody is using: the installer closes it.
if (Get-Process -Name 'Rutba Office' -ErrorAction SilentlyContinue) { "[$Label] Rutba Office is running; not installing over it."; exit 2 }

"[$Label] before:"; Registered
$sw = [Diagnostics.Stopwatch]::StartNew()
$p = Start-Process -FilePath $Older -ArgumentList '/S' -Wait -PassThru
"[$Label] older installer exit {0} in {1:n1}s" -f $p.ExitCode, $sw.Elapsed.TotalSeconds
$exe = Join-Path $env:LOCALAPPDATA 'Programs\Rutba Office\Rutba Office.exe'
if (-not (Test-Path $exe)) { "[$Label] no installed exe at $exe"; exit 1 }
$before = (Get-Item $exe).VersionInfo.FileVersion
"[$Label] installed exe version before: $before"

# The updater's own cache, emptied so what appears in it was fetched now.
# The name is electron-builder's `updaterCacheDirName` in the installed
# copy's resources\app-update.yml.
$cache = Join-Path $env:LOCALAPPDATA '@rutbaoffice-desktop-updater'
Remove-Item $cache -Recurse -Force -ErrorAction SilentlyContinue
$pending = Join-Path $cache 'pending'

# The launcher, in a profile of its own (so automatic updates are at their
# default: on), off the desktop. The service checks 25 s after launch and
# downloads in the background; every window shows the update prompt when it has.
$dataDir = Join-Path $env:TEMP ('rutba-updatetest-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
$env:RUTBA_WINDOW_DISPLAY = 'offscreen'
$proc = Start-Process -FilePath $exe -ArgumentList @('--app=home', "--user-data-dir=$dataDir") -PassThru
$sw.Restart()
$info = $null
for ($i = 0; $i -lt 900; $i++) {
  Start-Sleep -Seconds 1
  $file = Join-Path $pending 'update-info.json'
  if (Test-Path $file) { Start-Sleep -Seconds 2; $info = (Get-Content $file -Raw) -replace '\s+', ' '; break }
  if ($proc.HasExited) { "[$Label] the launcher exited on its own after {0:n0}s" -f $sw.Elapsed.TotalSeconds; break }
}
if (-not $info) {
  "[$Label] FAILED: no update was downloaded within {0:n0}s (cache: {1})" -f $sw.Elapsed.TotalSeconds, $cache
  if (Test-Path $cache) { Get-ChildItem $cache -Recurse | ForEach-Object { "[$Label]   " + $_.FullName + ' ' + $_.Length } }
  foreach ($id in (Tree $proc.Id)) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }
  Remove-Item $dataDir -Recurse -Force -ErrorAction SilentlyContinue
  exit 3
}
"[$Label] downloaded after {0:n0}s: {1}" -f $sw.Elapsed.TotalSeconds, $info
$got = Get-ChildItem $pending -Filter *.exe | ForEach-Object { "{0} ({1:n1} MB)" -f $_.Name, ($_.Length / 1MB) }
"[$Label] pending: $got"

# The window's own word for it - the prompt, and its button - read through UI Automation.
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$chip = ''
for ($i = 0; $i -lt 20 -and -not $chip; $i++) {
  Start-Sleep -Seconds 1
  Get-Process -Name 'Rutba Office' -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | ForEach-Object {
    try {
      $root = [System.Windows.Automation.AutomationElement]::FromHandle($_.MainWindowHandle)
      foreach ($e in $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)) {
        $n = $e.Current.Name
        if ($n -and $n -match 'ready to install|Restart and update|Update ready') { $chip = $n }
      }
    } catch {}
  }
}
"[$Label] window says: {0}" -f $(if ($chip) { $chip } else { '(no update prompt read)' })

# Close the window the way a person does; the service installs on quit.
$sw.Restart()
$null = $proc.CloseMainWindow()
for ($i = 0; $i -lt 60; $i++) { Start-Sleep -Seconds 1; if ($proc.HasExited) { break } }
if (-not $proc.HasExited) { "[$Label] the launcher did not quit when its window closed; stopping it"; foreach ($id in (Tree $proc.Id)) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue } }
"[$Label] quit after {0:n0}s; waiting for the installer" -f $sw.Elapsed.TotalSeconds
for ($i = 0; $i -lt 240; $i++) {
  Start-Sleep -Seconds 1
  $running = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -like 'Rutba-Office*' -or $_.ProcessName -eq 'Rutba Office' })
  if (-not $running.Count) { break }
}
Start-Sleep -Seconds 3
$after = (Get-Item $exe).VersionInfo.FileVersion
"[$Label] installed exe version after: $after ({0:n0}s after the quit)" -f $sw.Elapsed.TotalSeconds
"[$Label] after:"; Registered
Remove-Item $dataDir -Recurse -Force -ErrorAction SilentlyContinue
if ($after -ne $Expect) { "[$Label] FAILED: expected $Expect after the update, found $after."; exit 4 }
"[$Label] proven: $before updated itself to $after - found, downloaded and installed by the installed copy's own updater."
