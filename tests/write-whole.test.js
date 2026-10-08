// A file the windows write is written whole or not at all: the new bytes go
// beside it and are renamed over it, so a write that fails part way leaves
// the old file, and no stray file is left beside it.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { writeWhole } from '../packages/office-shell/src/write-whole.js';

test('a file is replaced whole, and nothing is left beside it', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-whole-'));
  const file = path.join(dir, 'clip.webm');
  fs.writeFileSync(file, 'old');
  await writeWhole(file, Buffer.from('new bytes'));
  assert.equal(fs.readFileSync(file, 'utf8'), 'new bytes');
  await writeWhole(path.join(dir, 'notes.txt'), 'words', 'utf8');
  assert.equal(fs.readFileSync(path.join(dir, 'notes.txt'), 'utf8'), 'words');
  assert.deepEqual(fs.readdirSync(dir).sort(), ['clip.webm', 'notes.txt']);
});

test('a write that fails leaves the old file as it was', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-whole-'));
  const file = path.join(dir, 'kept.txt');
  fs.writeFileSync(file, 'the old words');
  // Data that cannot be written: the write throws before anything is renamed.
  await assert.rejects(writeWhole(file, { not: 'bytes' }));
  assert.equal(fs.readFileSync(file, 'utf8'), 'the old words');
  assert.deepEqual(fs.readdirSync(dir), ['kept.txt']);
});
