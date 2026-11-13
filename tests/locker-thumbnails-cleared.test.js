// The thumbnail folder emptied once, on the first start of 2.26.0, and never
// again (owner, D-2026-11-13-29): packages/office-shell/src/electron/thumbs.js.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { tempDir } from './temp-dirs.js';
import { clearThumbnailsOnce } from '../packages/office-shell/src/electron/thumbs.js';
import { removeTreeSync } from '../packages/office-shell/src/remove-tree.js';

test('the thumbnail folder is emptied on the first start, noted, and left alone at every start after', async () => {
  const userData = tempDir('rl-thumbs-');
  const dir = path.join(userData, 'thumbnails');
  fs.mkdirSync(dir);
  for (let i = 0; i < 5; i++) fs.writeFileSync(path.join(dir, `${i}.jpg`), Buffer.alloc(100, i));
  const first = clearThumbnailsOnce(dir);
  assert.equal(first.cleared, true);
  await first.done;
  assert.equal(fs.existsSync(dir), false, 'every thumbnail made before is gone');
  assert.deepEqual(fs.readdirSync(userData).filter((n) => n.startsWith('thumbnails.old-')), [], 'and the folder it was moved to');
  assert.equal(fs.existsSync(`${dir}.cleared-2.26`), true, 'noted');
  // What Pictures makes from then on stays.
  fs.mkdirSync(dir);
  fs.writeFileSync(path.join(dir, 'new.jpg'), Buffer.alloc(10));
  const second = clearThumbnailsOnce(dir);
  assert.equal(second.cleared, false);
  await second.done;
  assert.deepEqual(fs.readdirSync(dir), ['new.jpg']);
});

test('a profile with no thumbnails yet is noted as done, and a folder a start left half removed goes', async () => {
  const userData = tempDir('rl-thumbs-');
  const dir = path.join(userData, 'thumbnails');
  const left = path.join(userData, 'thumbnails.old-1-2');
  fs.mkdirSync(left);
  fs.writeFileSync(path.join(left, 'a.jpg'), Buffer.alloc(10));
  const out = clearThumbnailsOnce(dir);
  assert.equal(out.cleared, true);
  await out.done;
  assert.equal(fs.existsSync(left), false);
  assert.equal(fs.existsSync(`${dir}.cleared-2.26`), true);
});

test('a folder that can be neither moved aside nor emptied is not noted, and is cleared at the next start', async () => {
  const userData = tempDir('rl-thumbs-');
  const dir = path.join(userData, 'thumbnails');
  fs.mkdirSync(dir);
  fs.writeFileSync(path.join(dir, 'a.jpg'), Buffer.alloc(10));
  const held = () => { throw Object.assign(new Error('busy'), { code: 'EBUSY' }); };
  const failed = clearThumbnailsOnce(dir, { ops: { rename: held, rm: held } });
  assert.equal(failed.cleared, false);
  assert.equal(fs.existsSync(`${dir}.cleared-2.26`), false, 'not noted');
  // Moved aside it could not be, but emptied in place.
  const inPlace = clearThumbnailsOnce(dir, { ops: { rename: held, rm: removeTreeSync } });
  assert.equal(inPlace.cleared, true);
  await inPlace.done;
  assert.equal(fs.existsSync(dir), false);
  assert.equal(fs.existsSync(`${dir}.cleared-2.26`), true);
});
