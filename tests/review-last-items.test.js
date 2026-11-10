// The last small faults the reviews of 10 Nov 2026 found, each reproduced on the code as it stood.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { Base } from '../packages/database/src/base.js';
import { rebuildSource } from '../apps/desktop/main/mail.js';
import { newWalk, spent } from '../packages/mailbox/src/pst/ndb.js';
import { createLocker } from '../apps/desktop/main/extras/locker/locker.js';
import { LARGE } from '../apps/desktop/main/extras/locker/folder.js';

const CHEAP = { N: 1024, r: 8, p: 1 };
const tmp = (pre = 'rli-') => fs.mkdtempSync(path.join(os.tmpdir(), pre));

test('a long whole number is saved whole, and comes back as it was typed', () => {
  const b = Base.open(':memory:');
  b.createTable({ name: 'T', fields: [{ name: 'Key', type: 'integer' }, { name: 'Note', type: 'text' }] });
  for (const long of ['9007199254740993', '-9007199254740993', '9223372036854775807', '-9223372036854775808']) {
    const rowid = b.insert('T', { Key: long, Note: 'a' });
    const row = b.db.prepare('SELECT CAST("Key" AS TEXT) AS k FROM T WHERE rowid = ?').get(rowid);
    assert.equal(row.k, long);
    b.update('T', rowid, { Key: long, Note: 'b' });
    assert.equal(b.db.prepare('SELECT CAST("Key" AS TEXT) AS k FROM T WHERE rowid = ?').get(rowid).k, long, 'given back unchanged');
  }
  assert.throws(() => b.insert('T', { Key: '9223372036854775808' }), /exactly/);
  b.close();
});

test('a rebuilt message says which attachment it could not include', () => {
  const missing = [];
  const text = rebuildSource(
    { text: 'hello', attachments: [{ filename: 'kept.txt', stored: 'a' }, { filename: 'gone.txt', stored: 'b' }, { filename: 'never.txt' }] },
    (a) => {
      if (a.stored === 'a') return Buffer.from('here');
      throw new Error('not there');
    },
    (a) => missing.push(a.filename),
  );
  assert.match(text, /kept\.txt/);
  assert.deepEqual(missing, ['gone.txt', 'never.txt']);
});

test('a mail file has one allowance across all its items, not one each', () => {
  const file = { left: 100 };
  const raw = new Uint8Array(60);
  spent(newWalk(file), raw);
  assert.throws(() => spent(newWalk(file), raw), /damaged/, 'a second item, its own limit nowhere near');
});

test('both versions of a large file can be saved and removed on their own', async () => {
  const dir = tmp();
  const folder = tmp('rli-folder-');
  const locker = createLocker({ dir, cost: CHEAP });
  await locker.setup('password one');
  const { id } = await locker.addFolder(folder, { afterLock: 'delete', hidden: false, keepVersions: 3 });
  const name = 'big.bin';
  const write = (size) => {
    const bytes = crypto.randomBytes(size);
    fs.writeFileSync(path.join(folder, name), bytes);
    return bytes;
  };
  const first = write(LARGE + 2048);
  await locker.lockWaiting(id);
  const second = write(LARGE + 4096);
  await locker.lockWaiting(id);
  const d = path.join(folder, `${name}.rlock.d`);
  // What a second computer leaves: its list carries the same number as this one's.
  const lists = fs.readdirSync(d).filter((f) => f.startsWith('list-')).sort();
  fs.renameSync(path.join(d, lists[0]), path.join(d, lists[1].replace(/-[0-9a-f]+\.rlist$/, '-aaaaaa.rlist')));

  const [row] = (await locker.scan(id)).filter((e) => e.state === 'conflict');
  assert.equal(row.twoLists.length, 2);
  assert.deepEqual(row.twoLists.map((l) => l.size).sort((a, b) => a - b), [first.length, second.length]);
  const copy = async (list) => {
    const out = path.join(tmp('rli-out-'), name);
    await locker.exportCopy(id, name, out, { allowConflict: true, lockedRel: row.lockedRel, list });
    return fs.readFileSync(out);
  };
  const saved = await Promise.all(row.twoLists.map((l) => copy(l.file)));
  assert.ok(saved.some((b) => b.equals(first)), 'the first version');
  assert.ok(saved.some((b) => b.equals(second)), 'and the second');

  // The older of the two, by its row: the choice is not always the one the file opens from.
  const going = row.twoLists.find((l) => l.size === first.length);
  await locker.deleteFile(id, name, 'conflict', { lockedRel: row.lockedRel, list: going.file });
  assert.deepEqual((await locker.scan(id)).map((e) => e.state), ['locked']);
  const out = path.join(tmp('rli-out-'), name);
  await locker.exportCopy(id, name, out);
  assert.ok(fs.readFileSync(out).equals(second), 'the other one stayed, whole');
});
