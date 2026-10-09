// rutba://file and rutba://thumb on a network share.
//
// They served any path asked for, so a share a document or a page named was
// looked at without a click, and looking at a share hands it the person's
// Windows sign-in. A local file is still served as asked; a share only once
// the person reached it in this run, by opening a file there, picking it in
// a dialog or browsing the folder.

import test from 'node:test';
import assert from 'node:assert/strict';
import { isNetworkPath, grantPlace, mayServe, forgetPlaces } from '../packages/office-shell/src/electron/grants.js';

const WIN = 'win32';

test('the network forms are known on Windows, and nowhere else', () => {
  for (const p of ['\\\\server\\share\\a.png', '//server/share/a.png', '\\\\?\\UNC\\server\\share\\a.png', '\\\\.\\pipe\\x']) assert.equal(isNetworkPath(p, WIN), true, p);
  for (const p of ['C:\\Users\\me\\a.png', 'D:/photos/a.png', 'relative\\a.png', '', null]) assert.equal(isNetworkPath(p, WIN), false, String(p));
  assert.equal(isNetworkPath('//home/me/a.png', 'linux'), false, 'two slashes are a local path on Linux');
});

test('a local file is served as asked; a share only once reached', () => {
  forgetPlaces();
  assert.equal(mayServe('C:\\Users\\me\\Pictures\\a.png', WIN), true);
  assert.equal(mayServe('\\\\attacker\\share\\x.png', WIN), false, 'a share named by a document is not looked at');
  assert.equal(mayServe('\\\\.\\pipe\\x', WIN), false);
});

test('a folder browsed or picked opens it and what is under it, and nothing beside it', () => {
  forgetPlaces();
  grantPlace('\\\\nas\\photos\\2026', WIN);
  assert.equal(mayServe('\\\\nas\\photos\\2026\\a.jpg', WIN), true);
  assert.equal(mayServe('\\\\NAS\\Photos\\2026\\Trip\\b.jpg', WIN), true, 'Windows paths are the same in any case');
  assert.equal(mayServe('//nas/photos/2026/c.jpg', WIN), true, 'forward slashes are the same place');
  assert.equal(mayServe('\\\\nas\\photos\\2025\\a.jpg', WIN), false, 'a folder beside it is not');
  assert.equal(mayServe('\\\\nas\\photos', WIN), false, 'nor the one above');
  assert.equal(mayServe('\\\\nas\\photos\\2026-old\\a.jpg', WIN), false, 'a name that only starts the same is not inside it');
});

test('a file opened opens only itself', () => {
  forgetPlaces();
  grantPlace('\\\\nas\\video\\talk.mp4', WIN);
  assert.equal(mayServe('\\\\nas\\video\\talk.mp4', WIN), true);
  assert.equal(mayServe('\\\\nas\\video\\other.mp4', WIN), false);
  grantPlace('C:\\local\\a.png', WIN);
  assert.equal(mayServe('\\\\nas\\video\\other.mp4', WIN), false, 'a local grant opens no share');
  forgetPlaces();
});
