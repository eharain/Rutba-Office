// Kinds of file that do something when opened, rather than show something:
// programs, scripts, installers, shortcuts, and the disk images that carry
// them past the system's own warning — the extensions mail is used to
// deliver malware under, on all three systems. Opening one of these goes
// through a warning in the main process, whatever asked for it.

export const RUNS_WHEN_OPENED = new Set([
  'exe', 'com', 'scr', 'pif', 'bat', 'cmd', 'msi', 'msp', 'cpl', 'dll', 'hta', 'jar',
  'js', 'jse', 'vbs', 'vbe', 'wsf', 'wsh', 'ps1', 'psm1', 'reg', 'lnk', 'url', 'scf', 'inf', 'msc', 'chm',
  'application', 'appref-ms', 'iso', 'img', 'vhd', 'vhdx',
  'app', 'command', 'pkg', 'dmg', 'sh', 'desktop', 'appimage', 'deb', 'rpm',
]);

/**
 * The extension a file is opened under, lower case. Windows drops trailing
 * dots and spaces and an alternate stream's name when it makes or opens a
 * file, so "run.exe " and "run.exe::$DATA" are programs; whatever follows the
 * first character that cannot be in an extension is not part of it.
 */
export function openedExtension(name) {
  const base = String(name ?? '').split(/[\\/]/).pop().replace(/[. ]+$/, '');
  if (!base.includes('.')) return '';
  const last = base.split('.').pop();
  return (/^[A-Za-z0-9_~+-]*/.exec(last)[0] || '').toLowerCase();
}

/** Whether opening `name` would run it. */
export const runsWhenOpened = (name) => RUNS_WHEN_OPENED.has(openedExtension(name));
