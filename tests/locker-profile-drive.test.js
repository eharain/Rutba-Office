// The Locker warns when the Office profile is where Windows cannot keep the
// opened copies to this account (D-2026-11-13-20, the owner: "Warn
// (Recommended)"): an exFAT or FAT drive, or a network share.
import test from 'node:test';
import assert from 'node:assert/strict';
import { profileDrive, parseDrive } from '../apps/desktop/main/extras/locker/profile-drive.js';

const at = (said, real = null) => ({ platform: 'win32', ask: async () => said, realpath: async (d) => real ?? d });

test('a profile on an exFAT or FAT drive is on a portable drive, whatever the case of the name', async () => {
  for (const fileSystem of ['exFAT', 'EXFAT', 'FAT32', 'FAT']) {
    assert.equal(await profileDrive('E:\\Rutba Office', at({ driveType: 2, fileSystem })), 'portable', fileSystem);
  }
});

test('a profile on a mapped network drive or a UNC path is on a share', async () => {
  assert.equal(await profileDrive('Z:\\Office', at({ driveType: 4, fileSystem: 'NTFS' })), 'share');
  let asked = false;
  const unc = { platform: 'win32', ask: async () => { asked = true; return null; }, realpath: async (d) => d };
  assert.equal(await profileDrive('\\\\server\\home\\Office', unc), 'share');
  assert.equal(await profileDrive('\\\\?\\UNC\\server\\home\\Office', unc), 'share');
  assert.equal(asked, false, 'a UNC path needs no question');
});

test('a profile on NTFS or ReFS, on another system, or on a drive that cannot be asked about, has no warning', async () => {
  assert.equal(await profileDrive('C:\\Users\\a\\AppData\\Roaming\\Rutba Office', at({ driveType: 3, fileSystem: 'NTFS' })), null);
  assert.equal(await profileDrive('D:\\Office', at({ driveType: 3, fileSystem: 'ReFS' })), null);
  assert.equal(await profileDrive('C:\\Office', at(null)), null, 'nothing claimed when Windows does not say');
  assert.equal(await profileDrive('/home/a/.config/Rutba Office', { platform: 'linux', ask: async () => { throw new Error('not asked'); } }), null);
  assert.equal(await profileDrive('', at({ driveType: 2, fileSystem: 'exFAT' })), null);
  assert.equal(await profileDrive('\\\\?\\C:\\Office', at({ driveType: 3, fileSystem: 'NTFS' })), null, 'a long path is its drive letter');
});

test('a folder that is a link to another drive is judged by the drive it really is on', async () => {
  const seen = [];
  const opts = { platform: 'win32', realpath: async () => 'F:\\Moved\\Office', ask: async (letter) => { seen.push(letter); return { driveType: 2, fileSystem: 'exFAT' }; } };
  assert.equal(await profileDrive('C:\\Users\\a\\AppData\\Roaming\\Rutba Office', opts), 'portable');
  assert.deepEqual(seen, ['F']);
});

test('on this computer the profile folder is asked about for real and answers without a throw', { skip: process.platform !== 'win32' && 'asks Windows' }, async () => {
  const said = await profileDrive(process.env.APPDATA || process.cwd());
  assert.ok(said === null || said === 'portable' || said === 'share');
});

test('what PowerShell prints is read as the drive type and its file system, and anything else as nothing', () => {
  assert.deepEqual(parseDrive('3 NTFS\r\n'), { driveType: 3, fileSystem: 'NTFS' });
  assert.deepEqual(parseDrive('2 exFAT'), { driveType: 2, fileSystem: 'exFAT' });
  assert.deepEqual(parseDrive('4 '), { driveType: 4, fileSystem: '' }, 'a network drive Windows gives no file system for');
  assert.equal(parseDrive(''), null);
  assert.equal(parseDrive(null), null);
  assert.equal(parseDrive('Get-CimInstance : Access denied'), null);
});
