// The app's own temp folders, left by a crash or held at quit, go on a later start.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { sweepTemps } from '../packages/office-shell/src/electron/temp-sweep.js';

test('a day-old folder of the app\'s goes; a new one, and anyone else\'s, stay', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'sweep-'));
  try {
    const old = path.join(dir, 'rutba-office-abc');
    const obj = path.join(dir, 'rutba-object-def');
    const fresh = path.join(dir, 'rutba-office-new');
    const other = path.join(dir, 'someone-else');
    for (const d of [old, obj, fresh, other]) { fs.mkdirSync(d); fs.writeFileSync(path.join(d, 'scratch.bin'), 'x'); }
    const twoDays = (Date.now() - 2 * 24 * 3600 * 1000) / 1000;
    for (const d of [old, obj, other]) fs.utimesSync(d, twoDays, twoDays);
    assert.equal(sweepTemps(dir), 2);
    assert.ok(!fs.existsSync(old) && !fs.existsSync(obj));
    assert.ok(fs.existsSync(fresh) && fs.existsSync(other));
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
