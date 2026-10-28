// The taskbar's copy of an app's tile (taskbar-tile.js): named by its content,
// so an icon changed by an update reaches the taskbar under a new name.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { taskbarTile } from '../packages/office-shell/src/electron/taskbar-tile.js';

test('a tile is copied under a name of its content, the same tile to the same name, a changed one to a new name', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-tiles-'));
  try {
    const tile = path.join(dir, 'mail.ico');
    const into = path.join(dir, 'tiles');
    fs.writeFileSync(tile, 'first icon');
    const first = taskbarTile('mail', tile, into);
    assert.match(path.basename(first), /^mail-[0-9a-f]{12}\.ico$/);
    assert.equal(fs.readFileSync(first, 'utf8'), 'first icon');
    assert.equal(taskbarTile('mail', tile, into), first, 'the same tile, the same name');

    fs.writeFileSync(tile, 'second icon');
    const second = taskbarTile('mail', tile, into);
    assert.notEqual(second, first, 'an updated tile is a new name');
    assert.equal(fs.readFileSync(second, 'utf8'), 'second icon');
    assert.ok(fs.existsSync(first), 'the earlier copy stays for a pin made with it');
    assert.deepEqual(fs.readdirSync(into).filter((f) => f.endsWith('.part')), []);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a tile that cannot be read is handed back as it is', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-tiles-'));
  try {
    const gone = path.join(dir, 'gone.ico');
    assert.equal(taskbarTile('mail', gone, path.join(dir, 'tiles')), gone);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
