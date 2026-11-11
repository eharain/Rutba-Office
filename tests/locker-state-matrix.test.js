// Every thing a person can do to a file, against every state a file can be in.
// A refusal must be the Locker's own error with its code, and must change
// nothing on the disk; what is allowed is checked by what comes back out.
// Then the same operations from a second locker sharing the folder, because a
// folder in the cloud is opened by two computers and neither may break the other.
import { test } from 'node:test';
import { tempDir } from './temp-dirs.js';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { createLocker } from '../apps/desktop/main/extras/locker/locker.js';
import { LARGE } from '../apps/desktop/main/extras/locker/folder.js';

const CHEAP = { N: 1024, r: 8, p: 1 };
const tmp = (pre = 'rl-m-') => tempDir(pre);
const OLD = crypto.randomBytes(3100);
const PLAIN = crypto.randomBytes(3000);
const SHARE = 'a password for sending';

async function locker(choices = {}) {
  const folder = tmp('rl-m-folder-');
  const l = createLocker({ dir: tmp(), cost: CHEAP });
  await l.setup('password one');
  const { id } = await l.addFolder(folder, { afterLock: 'delete', hidden: false, keepVersions: 3, ...choices });
  return { locker: l, folder, id, at: (name) => path.join(folder, name) };
}

/** Two versions locked, so there is an earlier one to go back to. */
async function twoVersions(c, name = 'a.txt', [first, second] = [OLD, PLAIN]) {
  fs.writeFileSync(c.at(name), first);
  await c.locker.lockWaiting(c.id);
  fs.writeFileSync(c.at(name), second);
  await c.locker.lockWaiting(c.id);
}

const STATES = {
  locked: { make: (c) => twoVersions(c), rel: 'a.txt', row: 'locked' },
  'locked, name hidden': { choices: { hidden: true }, make: (c) => twoVersions(c), rel: 'a.txt', row: 'locked' },
  'locked, in a subfolder': {
    make: async (c) => { fs.mkdirSync(c.at('in/deep'), { recursive: true }); await twoVersions(c, 'in/deep/a.txt'); },
    rel: 'in/deep/a.txt',
    row: 'locked',
  },
  waiting: { make: (c) => fs.writeFileSync(c.at('a.txt'), PLAIN), rel: 'a.txt', row: 'waiting' },
  excluded: { make: (c) => fs.writeFileSync(c.at('~$a.docx'), PLAIN), rel: '~$a.docx', row: 'excluded' },
  damaged: {
    make: async (c) => {
      await twoVersions(c);
      const bytes = fs.readFileSync(c.at('a.txt.rlock'));
      fs.writeFileSync(c.at('a.txt.rlock'), bytes.subarray(0, 140));
    },
    rel: 'a.txt.rlock',
    row: 'damaged',
  },
  'locked by a newer version': {
    make: async (c) => {
      await twoVersions(c);
      const bytes = fs.readFileSync(c.at('a.txt.rlock'));
      bytes.write('RLOCKV2\n', 0, 'latin1');
      fs.writeFileSync(c.at('a.txt.rlock'), bytes);
    },
    rel: 'a.txt.rlock',
    row: 'damaged',
  },
  'two versions from the cloud': {
    make: async (c) => {
      await twoVersions(c);
      fs.copyFileSync(c.at('a.txt.rlock'), c.at('a (conflicted copy).txt.rlock'));
    },
    rel: 'a.txt',
    row: 'conflict',
  },
  'plain beside its locked copy': {
    make: async (c) => {
      await twoVersions(c);
      fs.writeFileSync(c.at('a.txt'), Buffer.from('a different plain file'));
    },
    rel: 'a.txt',
    row: 'locked',
  },
};

const outDest = () => path.join(tmp('rl-m-out-'), 'out.bin');

/** Each operation: run it, then say whether what it did was right. */
const OPS = {
  export: {
    run: async (c, rel) => { const dest = outDest(); await c.locker.exportCopy(c.id, rel, dest); return dest; },
    good: (c, dest) => assert.ok(fs.readFileSync(dest).equals(PLAIN), 'the copy holds the file'),
    unchanged: true,
  },
  unlock: {
    run: (c, rel) => c.locker.unlockForGood(c.id, rel),
    good: async (c, _r, rel) => {
      assert.ok(fs.readFileSync(path.join(c.folder, ...rel.split('/'))).equals(PLAIN), 'the plain file is back');
      assert.equal((await c.locker.scan(c.id)).find((e) => e.rel === rel).state, 'waiting');
    },
  },
  rename: {
    run: (c, rel) => c.locker.renameFile(c.id, rel, 'b.txt'),
    good: async (c, _r, rel) => {
      const dir = path.posix.dirname(rel);
      const newRel = dir === '.' ? 'b.txt' : `${dir}/b.txt`;
      const list = await c.locker.scan(c.id);
      assert.equal(list.some((e) => e.rel === rel), false, 'the old name is gone');
      const row = list.find((e) => e.rel === newRel);
      assert.ok(row, 'the new name is listed');
      if (row.state === 'locked') {
        const dest = outDest();
        await c.locker.exportCopy(c.id, newRel, dest);
        assert.ok(fs.readFileSync(dest).equals(PLAIN), 'the same contents under the new name');
      } else assert.ok(fs.readFileSync(path.join(c.folder, ...newRel.split('/'))).equals(PLAIN));
    },
  },
  delete: {
    run: (c, rel) => c.locker.deleteFile(c.id, rel),
    good: async (c, _r, rel) => assert.equal((await c.locker.scan(c.id)).some((e) => e.rel === rel), false),
  },
  'lock it': {
    run: (c, rel) => c.locker.lockOne(c.id, rel),
    good: async (c, _r, rel) => {
      assert.equal((await c.locker.scan(c.id)).find((e) => e.rel === rel).state, 'locked');
      const dest = outDest();
      await c.locker.exportCopy(c.id, rel, dest);
      assert.ok(fs.readFileSync(dest).equals(PLAIN));
    },
  },
  open: {
    run: (c, rel) => c.locker.openPlain(c.id, rel),
    good: async (c, out) => {
      assert.ok(fs.readFileSync(out.path).equals(PLAIN), 'the temporary copy holds the file');
      await c.locker.lock();
      assert.equal(fs.existsSync(out.path), false, 'locking clears the temporary copy');
    },
  },
  share: {
    run: async (c, rel) => { const dest = outDest(); await c.locker.exportShare(c.id, rel, dest, SHARE); return dest; },
    good: async (c, dest) => {
      const back = outDest();
      await c.locker.openShared(dest, back, SHARE);
      assert.ok(fs.readFileSync(back).equals(PLAIN), 'the shared file opens to the file');
    },
    unchanged: true,
  },
  versions: {
    run: (c, rel) => c.locker.versions(c.id, rel),
    good: (c, list) => assert.deepEqual([list[0].current, list[1]?.current], [true, false], 'now, and the earlier one'),
    unchanged: true,
  },
  'go back a version': {
    run: (c, rel) => c.locker.restore(c.id, rel, 1),
    good: async (c, _r, rel) => {
      const dest = outDest();
      await c.locker.exportCopy(c.id, rel, dest);
      assert.ok(fs.readFileSync(dest).equals(OLD), 'the earlier contents are the newest');
    },
  },
};

const LOCKED = { export: 'ok', unlock: 'ok', rename: 'ok', delete: 'ok', 'lock it': 'state', open: 'ok', share: 'ok', versions: 'ok', 'go back a version': 'ok' };
const NOT_READABLE = { export: 'state', unlock: 'state', rename: 'state', delete: 'gone', 'lock it': 'state', open: 'state', share: 'state', versions: 'state', 'go back a version': 'state' };
const EXPECTED = {
  locked: LOCKED,
  'locked, name hidden': LOCKED,
  'locked, in a subfolder': LOCKED,
  waiting: { export: 'state', unlock: 'state', rename: 'ok', delete: 'ok', 'lock it': 'ok', open: 'state', share: 'state', versions: 'state', 'go back a version': 'state' },
  excluded: { export: 'state', unlock: 'state', rename: 'ok', delete: 'ok', 'lock it': 'state', open: 'state', share: 'state', versions: 'state', 'go back a version': 'state' },
  damaged: NOT_READABLE,
  'locked by a newer version': NOT_READABLE,
  'two versions from the cloud': { export: 'conflict', unlock: 'conflict', rename: 'conflict', delete: 'conflict', 'lock it': 'state', open: 'conflict', share: 'conflict', versions: 'conflict', 'go back a version': 'conflict' },
  // Which of the two a name means is the locked one; the plain file beside it is never overwritten.
  'plain beside its locked copy': { export: 'ok', unlock: 'exists', rename: 'skip', delete: 'skip', 'lock it': 'skip', open: 'ok', share: 'ok', versions: 'ok', 'go back a version': 'ok' },
};

/** Every file under the folder with its hash, the Locker's own private ones included. */
function fingerprint(root) {
  const out = [];
  const go = (dir) => {
    for (const e of fs.readdirSync(dir, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name))) {
      const p = path.join(dir, e.name);
      if (e.isDirectory()) go(p);
      else out.push(`${path.relative(root, p)} ${crypto.createHash('sha256').update(fs.readFileSync(p)).digest('hex')}`);
    }
  };
  go(root);
  return out.join('\n');
}

for (const [stateName, state] of Object.entries(STATES)) {
  for (const [opName, op] of Object.entries(OPS)) {
    const want = EXPECTED[stateName][opName];
    if (want === 'skip') continue;
    test(`${stateName}: ${opName} ${want === 'ok' ? 'works' : want === 'gone' ? 'removes what cannot be read' : `is refused (${want})`}`, async () => {
      const c = await locker(state.choices);
      await state.make(c);
      assert.equal((await c.locker.scan(c.id)).find((e) => e.rel === state.rel)?.state, state.row, 'the file is in the state this test is about');
      const before = fingerprint(c.folder);
      if (want === 'ok') {
        const result = await op.run(c, state.rel);
        await op.good(c, result, state.rel);
        if (op.unchanged) assert.equal(fingerprint(c.folder), before, 'the folder is as it was');
      } else if (want === 'gone') {
        await op.run(c, state.rel);
        assert.equal((await c.locker.scan(c.id)).some((e) => e.rel === state.rel), false);
      } else {
        await assert.rejects(() => op.run(c, state.rel), (err) => {
          assert.equal(err.name, 'LockerError', `the Locker's own refusal, not ${err.name}: ${err.message}`);
          assert.equal(err.code, want, err.message);
          return true;
        });
        assert.equal(fingerprint(c.folder), before, 'a refusal changes nothing');
      }
    });
  }
}

// The way out of two versions: each can be read, and either can be removed.

test('two versions from the cloud: each can be saved on its own, and they differ', async () => {
  const c = await locker();
  await twoVersions(c);
  fs.copyFileSync(c.at('a.txt.rlock'), c.at('a (conflicted copy).txt.rlock'));
  const rows = (await c.locker.scan(c.id)).filter((e) => e.state === 'conflict');
  assert.deepEqual(rows.map((e) => e.lockedRel).sort(), ['a (conflicted copy).txt.rlock', 'a.txt.rlock']);
  for (const row of rows) {
    const dest = outDest();
    await c.locker.exportCopy(c.id, row.rel, dest, { allowConflict: true, lockedRel: row.lockedRel });
    assert.ok(fs.readFileSync(dest).equals(PLAIN));
  }
});

test('two versions that differ: asking for one by its place on the disk gives that one', async () => {
  // A cloud conflict copy made by another computer holds other bytes under the same file id.
  const c = await locker();
  await twoVersions(c);
  const first = fs.readFileSync(c.at('a.txt.rlock'));
  fs.writeFileSync(c.at('a.txt'), OLD);
  await c.locker.lockWaiting(c.id);
  fs.renameSync(c.at('a.txt.rlock'), c.at('a (conflicted copy).txt.rlock'));
  fs.writeFileSync(c.at('a.txt.rlock'), first);
  const rows = (await c.locker.scan(c.id)).filter((e) => e.state === 'conflict');
  const got = {};
  for (const row of rows) {
    const dest = outDest();
    await c.locker.exportCopy(c.id, row.rel, dest, { allowConflict: true, lockedRel: row.lockedRel });
    got[row.lockedRel] = fs.readFileSync(dest);
  }
  assert.ok(got['a.txt.rlock'].equals(PLAIN), 'the first is its own contents');
  assert.ok(got['a (conflicted copy).txt.rlock'].equals(OLD), 'the second is its own contents');
});

test('two versions: removing one leaves the other as the file, with its earlier versions', async () => {
  const c = await locker();
  await twoVersions(c);
  fs.copyFileSync(c.at('a.txt.rlock'), c.at('a (conflicted copy).txt.rlock'));
  await c.locker.deleteFile(c.id, 'a.txt', 'conflict', { lockedRel: 'a (conflicted copy).txt.rlock' });
  assert.equal(fs.existsSync(c.at('a (conflicted copy).txt.rlock')), false);
  const rows = await c.locker.scan(c.id);
  assert.deepEqual(rows.map((e) => [e.rel, e.state]), [['a.txt', 'locked']]);
  assert.ok((await c.locker.versions(c.id, 'a.txt')).length >= 2, 'the earlier versions stay');
});

test('stopping a folder with two versions of a file changes nothing, and says why', async () => {
  const c = await locker();
  await twoVersions(c);
  fs.copyFileSync(c.at('a.txt.rlock'), c.at('a (conflicted copy).txt.rlock'));
  const before = fingerprint(c.folder);
  const out = await c.locker.removeFolder(c.id);
  assert.equal(out.removed, false);
  assert.equal(out.conflicts, 1);
  assert.equal(fingerprint(c.folder), before);
});

test('a plain file beside a locked one of its name: each is named by its row, and neither touches the other', async () => {
  const beside = Buffer.from('a different plain file');
  const setUp = async () => {
    const c = await locker();
    await twoVersions(c);
    fs.writeFileSync(c.at('a.txt'), beside);
    return c;
  };
  const read = async (c, rel) => { const dest = outDest(); await c.locker.exportCopy(c.id, rel, dest); return fs.readFileSync(dest); };
  let c = await setUp();
  await c.locker.renameFile(c.id, 'a.txt', 'b.txt', 'waiting');
  assert.ok(fs.readFileSync(c.at('b.txt')).equals(beside));
  assert.ok((await read(c, 'a.txt')).equals(PLAIN), 'the locked one is untouched');
  c = await setUp();
  await c.locker.renameFile(c.id, 'a.txt', 'b.txt', 'locked');
  assert.ok((await read(c, 'b.txt')).equals(PLAIN));
  assert.ok(fs.readFileSync(c.at('a.txt')).equals(beside), 'the plain one is untouched');
  c = await setUp();
  await c.locker.deleteFile(c.id, 'a.txt', 'waiting');
  assert.equal(fs.existsSync(c.at('a.txt')), false);
  assert.ok((await read(c, 'a.txt')).equals(PLAIN), 'the locked one stays');
  c = await setUp();
  await c.locker.deleteFile(c.id, 'a.txt', 'locked');
  assert.equal(fs.existsSync(c.at('a.txt.rlock')), false);
  assert.ok(fs.readFileSync(c.at('a.txt')).equals(beside), 'the plain one stays');
});

test('a locked file damaged inside: nothing half-written is left, and the locked copy stays', async () => {
  const c = await locker();
  await twoVersions(c);
  const bytes = fs.readFileSync(c.at('a.txt.rlock'));
  bytes[bytes.length - 40] ^= 0xff;
  fs.writeFileSync(c.at('a.txt.rlock'), bytes);
  const dest = outDest();
  await assert.rejects(() => c.locker.exportCopy(c.id, 'a.txt', dest), (e) => e.name === 'LockerError');
  assert.equal(fs.existsSync(dest), false, 'no partial copy');
  await assert.rejects(() => c.locker.unlockForGood(c.id, 'a.txt'), (e) => e.name === 'LockerError');
  assert.equal(fs.existsSync(c.at('a.txt')), false, 'no partial plain file');
  assert.ok(fs.readFileSync(c.at('a.txt.rlock')).equals(bytes), 'the locked copy is as it was');
  assert.equal(fs.readdirSync(c.folder).some((n) => n.endsWith('.part')), false, 'no working file left in the folder');
});

// The same folder, from two lockers: nothing one does may stop the other.

async function twoLockers() {
  const c = await locker();
  await twoVersions(c);
  const second = createLocker({ dir: tmp(), cost: CHEAP });
  await second.setup('password two');
  const { id } = await second.addFolder(c.folder, {}, { password: 'password one' });
  return { first: c, second: { locker: second, folder: c.folder, id, at: c.at } };
}

/** What the other locker finds after each: a.txt still locked with these bytes, under another name, plain again, or gone. */
const SEEN_BY_OTHER = {
  export: { rel: 'a.txt', locked: PLAIN },
  open: { rel: 'a.txt', locked: PLAIN },
  share: { rel: 'a.txt', locked: PLAIN },
  versions: { rel: 'a.txt', locked: PLAIN },
  'go back a version': { rel: 'a.txt', locked: OLD },
  rename: { rel: 'b.txt', locked: PLAIN, gone: 'a.txt' },
  unlock: { rel: 'a.txt', plain: PLAIN },
  delete: { gone: 'a.txt' },
};

for (const [opName, op] of Object.entries(OPS)) {
  if (opName === 'lock it') continue;
  for (const [who, other] of [['first', 'second'], ['second', 'first']]) {
    test(`two lockers on one folder: ${opName} from the ${who} leaves the ${other} able to read it`, async () => {
      const both = await twoLockers();
      const result = await op.run(both[who], 'a.txt');
      await op.good(both[who], result, 'a.txt');
      if (opName === 'open') await both[who].locker.unlock({ password: who === 'first' ? 'password one' : 'password two' });
      await both[other].locker.healFolderKeys();
      const them = both[other];
      const list = await them.locker.scan(them.id);
      assert.equal(list.some((e) => e.state === 'damaged'), false, JSON.stringify(list.map((e) => [e.rel, e.state, e.message])));
      const seen = SEEN_BY_OTHER[opName];
      const row = (rel) => list.find((e) => e.rel === rel);
      if (seen.gone) assert.equal(row(seen.gone), undefined, `${seen.gone} is gone for the ${other} too`);
      if (seen.plain) {
        assert.equal(row(seen.rel)?.state, 'waiting');
        assert.ok(fs.readFileSync(them.at(seen.rel)).equals(seen.plain));
      }
      if (seen.locked) {
        assert.equal(row(seen.rel)?.state, 'locked', JSON.stringify(list.map((e) => [e.rel, e.state])));
        const dest = outDest();
        await them.locker.exportCopy(them.id, seen.rel, dest);
        assert.ok(fs.readFileSync(dest).equals(seen.locked), `the ${other} reads what the ${who} left`);
      }
    });
  }
}

test('two lockers on one folder: a file locked by one opens for the other, and each change of password leaves the other in', async () => {
  const { first, second } = await twoLockers();
  fs.writeFileSync(first.at('new.txt'), PLAIN);
  await first.locker.lockWaiting(first.id);
  const read = async (side) => {
    const dest = outDest();
    await side.locker.exportCopy(side.id, 'new.txt', dest);
    return fs.readFileSync(dest);
  };
  assert.ok((await read(second)).equals(PLAIN));
  await first.locker.changePassword('password one', 'a changed one');
  await second.locker.lock();
  await second.locker.unlock({ password: 'password two' });
  assert.ok((await read(second)).equals(PLAIN), 'the second locker is not locked out');
  await second.locker.changePassword('password two', 'also changed');
  await first.locker.lock();
  await first.locker.unlock({ password: 'a changed one' });
  assert.ok((await read(first)).equals(PLAIN), 'nor is the first');
});

test('a large file: each operation works on one kept in pieces', async () => {
  const c = await locker();
  const big = crypto.randomBytes(LARGE + 4096);
  fs.writeFileSync(c.at('big.bin'), big);
  await c.locker.lockWaiting(c.id);
  assert.ok(fs.existsSync(c.at('big.bin.rlock.d')));
  const dest = outDest();
  await c.locker.exportCopy(c.id, 'big.bin', dest);
  assert.ok(fs.readFileSync(dest).equals(big), 'export');
  await c.locker.renameFile(c.id, 'big.bin', 'large.bin');
  assert.ok(fs.existsSync(c.at('large.bin.rlock.d')) && !fs.existsSync(c.at('big.bin.rlock.d')), 'rename on the disk');
  const dest2 = outDest();
  await c.locker.exportCopy(c.id, 'large.bin', dest2);
  assert.ok(fs.readFileSync(dest2).equals(big), 'the same contents after rename');
  await c.locker.unlockForGood(c.id, 'large.bin');
  assert.ok(fs.readFileSync(c.at('large.bin')).equals(big), 'unlock');
  assert.equal(fs.existsSync(c.at('large.bin.rlock.d')), false, 'the pieces went');
  await c.locker.lockWaiting(c.id);
  await c.locker.deleteFile(c.id, 'large.bin');
  assert.deepEqual(await c.locker.scan(c.id), [], 'delete');
});
