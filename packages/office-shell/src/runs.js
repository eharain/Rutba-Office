// Kinds of file that do something when opened, rather than show something:
// programs, scripts, installers, shortcuts, and the disk images that carry
// them past the system's own warning — the extensions mail is used to
// deliver malware under, on all three systems. Opening one of these goes
// through a warning in the main process, whatever asked for it.

export const RUNS_WHEN_OPENED = new Set([
  'exe', 'com', 'scr', 'pif', 'bat', 'cmd', 'msi', 'msp', 'cpl', 'dll', 'hta', 'jar',
  'js', 'jse', 'vbs', 'vbe', 'wsf', 'wsh', 'ps1', 'psm1', 'reg', 'lnk', 'url', 'scf', 'inf', 'msc', 'chm',
  'application', 'appref-ms', 'iso', 'img', 'vhd', 'vhdx',
  'msix', 'msixbundle', 'appx', 'appxbundle', 'appinstaller', 'msu', 'gadget', 'sct', 'wsc', 'ws', 'ps1xml', 'ps2', 'psc1', 'psd1', 'msh', 'mshxml',
  'shb', 'shs', 'scpt', 'xbap', 'ins', 'isp', 'mst', 'hlp', 'diagcab', 'settingcontent-ms', 'library-ms', 'searchconnector-ms', 'search-ms', 'xll', 'wll', 'cab',
  'py', 'pyw', 'pl', 'rb', 'jnlp', 'vb', 'vba', 'mcf', 'cnt', 'crt', 'der', 'website', 'theme', 'themepack',
  'app', 'command', 'pkg', 'dmg', 'sh', 'desktop', 'appimage', 'deb', 'rpm',
  'ade', 'adp', 'mda', 'mdb', 'mde', 'mdt', 'mdw', 'wsb', 'cer', 'bas', 'htc', 'ksh', 'msh1', 'msh2', 'ps2xml', 'psc2', 'pyc', 'pyz', 'vbp', 'vsmacros',
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

/** The mark Windows puts on a file from the internet; SmartScreen and other programs' protected modes read it. */
export const FROM_INTERNET = '[ZoneTransfer]\r\nZoneId=3\r\n';

/**
 * Marks a file that came in by mail as from the internet, as a browser marks a
 * download. `write` writes the mark, which on Windows is a stream of the file.
 * Answers whether it is marked: a disk that cannot hold streams cannot be, and
 * the file is still the person's.
 */
export async function markFromInternet(filePath, { write, platform }) {
  if (platform !== 'win32' || !/^[a-z]:[\\/][^:]*$/i.test(String(filePath))) return false;
  try {
    await write(`${filePath}:Zone.Identifier`, FROM_INTERNET);
    return true;
  } catch {
    return false;
  }
}

/** Whether opening `name` would run it. */
export const runsWhenOpened = (name) => RUNS_WHEN_OPENED.has(openedExtension(name));
