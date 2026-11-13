// Where the Office profile sits, for the Locker's warning (D-2026-11-13-20,
// the owner: "Warn (Recommended)"). An opened locked file is copied into the
// profile (working.js), and the profile's own permissions keep that copy the
// person's. A drive formatted exFAT or FAT keeps no permissions at all, and a
// network share keeps the server's, so there Windows cannot keep the copy to
// this account: the Locker says so in a plain line. Nothing is refused.

import { execFile } from 'node:child_process';
import fsp from 'node:fs/promises';
import path from 'node:path';

/** File systems that keep no permissions: anyone with the drive can read what is on it. */
const NO_PERMISSIONS = new Set(['EXFAT', 'FAT', 'FAT32', 'FAT12', 'FAT16']);
/** Win32_LogicalDisk's DriveType for a network drive. */
const NETWORK_DRIVE = 4;

/**
 * Windows PowerShell by its full path: a bare name is looked for in the
 * current folder first, so a program of that name in the folder Office was
 * started from would run instead (review of 13 Nov).
 */
const POWERSHELL = path.win32.join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');

/** What the asking printed: "<DriveType> <FileSystem>", the file system empty for a drive with none. */
export function parseDrive(out) {
  const [type = '', fileSystem = ''] = String(out || '').trim().split(/\s+/);
  if (!/^\d+$/.test(type)) return null;
  return { driveType: Number(type), fileSystem };
}

/** What Windows says of a drive letter: { fileSystem, driveType }, or null when it cannot be asked. */
function askWindows(letter) {
  const id = `${letter.toUpperCase()}:`;
  const script = `$d = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='${id}'"; if ($d) { "{0} {1}" -f $d.DriveType, $d.FileSystem }`;
  return new Promise((resolve) => {
    execFile(POWERSHELL, ['-NoProfile', '-NonInteractive', '-Command', script], { encoding: 'utf8', timeout: 15_000, windowsHide: true }, (err, out) => {
      resolve(err ? null : parseDrive(out));
    });
  });
}

/**
 * 'share' when the profile is on a network share, 'portable' when its drive
 * keeps no permissions (exFAT or FAT), or null: a drive that keeps them, a
 * system other than Windows (where the copy's folder is 0700), or a drive
 * that cannot be asked about (nothing is claimed then). Not seen: a volume
 * mounted in a folder of another drive, with no letter of its own, is judged
 * by that drive's letter; and the computer's own drive reached by a share
 * name (\\localhost\C$) is taken for a share.
 */
export async function profileDrive(dir, { platform = process.platform, ask = askWindows, realpath = fsp.realpath } = {}) {
  if (platform !== 'win32' || !dir) return null;
  // A junction or a subst-style link: the drive it really is on.
  const real = await realpath(dir).catch(() => dir);
  const p = path.win32.normalize(String(real));
  if (/^\\\\\?\\UNC\\/i.test(p) || (/^\\\\/.test(p) && !/^\\\\[?.]\\/.test(p))) return 'share';
  const letter = /^(?:\\\\\?\\)?([A-Za-z]):/.exec(p)?.[1];
  if (!letter) return null;
  const said = await ask(letter);
  if (!said) return null;
  if (said.driveType === NETWORK_DRIVE) return 'share';
  if (NO_PERMISSIONS.has(String(said.fileSystem).toUpperCase())) return 'portable';
  return null;
}
