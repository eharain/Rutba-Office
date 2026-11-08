// The Locker's one promise: a person's file is never lost. Each test here is a
// way of losing one, and each says what has to still be true afterwards.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { newHeader, parseHeader, verifyLocked, openLocked, readHeader, HEADER_SIZE, META_FRAME, DATA_FRAME, CHUNK, MODE_FOLDER } from '../apps/desktop/main/extras/locker/format.js';
import { sealFile, writePlain, sweepWorking, isWorkingFile } from '../apps/desktop/main/extras/locker/place.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const KEY = crypto.randomBytes(32);
const OTHER = crypto.randomBytes(32);

function tmp() {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'rl-test-'));
}

function makeFile(dir, name, bytes) {
  const f = path.join(dir, name);
  fs.writeFileSync(f, bytes);
  return f;
}

const random = (n) => crypto.randomBytes(n);
const sameBytes = (a, b) => Buffer.compare(a, b) === 0;

test('a file comes back from its locked copy exactly as it was: empty, short, one frame, many frames', async () => {
  const dir = tmp();
  const stage = path.join(dir, 'stage');
  for (const bytes of [Buffer.alloc(0), Buffer.from('hello'), random(CHUNK), random(CHUNK * 3 + 17)]) {
    const src = makeFile(dir, 'plain.bin', bytes);
    const target = path.join(dir, 'plain.rlock');
    const out = await sealFile({ source: src, target, rootKey: KEY, headerBuf: newHeader({ folderId: Buffer.alloc(16, 1) }), meta: { name: 'plain.bin' }, stageDir: stage });
    assert.equal(out.size, bytes.length);
    const dest = path.join(dir, 'back.bin');
    const back = await writePlain({ locked: target, dest, rootKey: KEY, overwrite: true });
    assert.equal(back.sha256, crypto.createHash('sha256').update(bytes).digest('hex'));
    assert.ok(sameBytes(fs.readFileSync(dest), bytes));
    fs.rmSync(target);
  }
});

test('the name and the size inside the metadata are unreadable, and the frames are not the plain bytes', async () => {
  const dir = tmp();
  const secret = 'Lease, signed in private.pdf';
  const bytes = Buffer.from('the same words over and over '.repeat(5000));
  const src = makeFile(dir, 'a.bin', bytes);
  const target = path.join(dir, 'a.rlock');
  await sealFile({ source: src, target, rootKey: KEY, headerBuf: newHeader(), meta: { name: secret }, stageDir: dir });
  const raw = fs.readFileSync(target);
  assert.equal(raw.indexOf(secret), -1, 'the name does not appear');
  assert.equal(raw.indexOf('the same words'), -1, 'the bytes do not appear');
  const { meta } = await verifyLocked(target, KEY);
  assert.equal(meta.name, secret);
});

test('a wrong key never changes a file, and never reads it', async () => {
  const dir = tmp();
  const src = makeFile(dir, 'a.bin', random(CHUNK + 5));
  const target = path.join(dir, 'a.rlock');
  await sealFile({ source: src, target, rootKey: KEY, headerBuf: newHeader(), meta: { name: 'a.bin' }, stageDir: dir });
  const before = fs.readFileSync(target);
  await assert.rejects(() => openLocked(target, OTHER), (e) => e.code === 'bad-key');
  await assert.rejects(() => writePlain({ locked: target, dest: path.join(dir, 'x.bin'), rootKey: OTHER }), (e) => e.code === 'bad-key');
  assert.ok(sameBytes(fs.readFileSync(target), before), 'the locked file is exactly as it was');
  assert.ok(!fs.existsSync(path.join(dir, 'x.bin')), 'no plain file was started');
});

test('every byte of a locked file is protected: a change anywhere is found, and the first frame is the key check', async () => {
  const dir = tmp();
  const src = makeFile(dir, 'a.bin', random(CHUNK * 2 + 100));
  const target = path.join(dir, 'a.rlock');
  await sealFile({ source: src, target, rootKey: KEY, headerBuf: newHeader(), meta: { name: 'a.bin' }, stageDir: dir });
  const good = fs.readFileSync(target);
  const offsets = [20, HEADER_SIZE + 5, HEADER_SIZE + META_FRAME + 10, HEADER_SIZE + META_FRAME + DATA_FRAME + 10, good.length - 3];
  for (const at of offsets) {
    const bad = Buffer.from(good);
    bad[at] ^= 1;
    fs.writeFileSync(target, bad);
    await assert.rejects(async () => {
      const o = await openLocked(target, KEY);
      try {
        for await (const c of o.frames()) void c;
      } finally {
        await o.close();
      }
    }, (e) => ['bad-key', 'tampered', 'format'].includes(e.code), `byte ${at}`);
  }
});

test('a locked file cut short is found, at every length', async () => {
  const dir = tmp();
  const src = makeFile(dir, 'a.bin', random(CHUNK * 2 + 100));
  const target = path.join(dir, 'a.rlock');
  await sealFile({ source: src, target, rootKey: KEY, headerBuf: newHeader(), meta: { name: 'a.bin' }, stageDir: dir });
  const good = fs.readFileSync(target);
  for (const len of [0, 10, HEADER_SIZE, HEADER_SIZE + 10, HEADER_SIZE + META_FRAME, HEADER_SIZE + META_FRAME + 5, HEADER_SIZE + META_FRAME + DATA_FRAME, good.length - 1]) {
    fs.writeFileSync(target, good.subarray(0, len));
    await assert.rejects(async () => {
      const o = await openLocked(target, KEY);
      try {
        for await (const c of o.frames()) void c;
      } finally {
        await o.close();
      }
    }, (e) => ['truncated', 'format', 'tampered'].includes(e.code), `length ${len}`);
  }
});

test('frames moved around, or swapped between files, or a frame cut from the end, are found', async () => {
  const dir = tmp();
  const srcA = makeFile(dir, 'a.bin', random(CHUNK * 3 + 100));
  const srcB = makeFile(dir, 'b.bin', random(CHUNK * 3 + 100));
  const a = path.join(dir, 'a.rlock');
  const b = path.join(dir, 'b.rlock');
  await sealFile({ source: srcA, target: a, rootKey: KEY, headerBuf: newHeader(), meta: { name: 'a' }, stageDir: dir });
  await sealFile({ source: srcB, target: b, rootKey: KEY, headerBuf: newHeader(), meta: { name: 'b' }, stageDir: dir });
  const goodA = fs.readFileSync(a);
  const goodB = fs.readFileSync(b);
  const f1 = HEADER_SIZE + META_FRAME;
  const f2 = f1 + DATA_FRAME;
  const read = async () => {
    const o = await openLocked(a, KEY);
    try {
      for await (const c of o.frames()) void c;
    } finally {
      await o.close();
    }
  };
  // Two frames swapped in place.
  const swapped = Buffer.from(goodA);
  goodA.copy(swapped, f1, f2, f2 + DATA_FRAME);
  goodA.copy(swapped, f2, f1, f1 + DATA_FRAME);
  fs.writeFileSync(a, swapped);
  await assert.rejects(read, (e) => e.code === 'tampered');
  // A frame from another file.
  const mixed = Buffer.from(goodA);
  goodB.copy(mixed, f1, f1, f1 + DATA_FRAME);
  fs.writeFileSync(a, mixed);
  await assert.rejects(read, (e) => e.code === 'tampered');
  // The last frame cut off: nothing else after it, so it must fail on the frame that should have been marked last.
  fs.writeFileSync(a, goodA.subarray(0, f2 + DATA_FRAME));
  await assert.rejects(read, (e) => ['tampered', 'truncated'].includes(e.code));
  // Bytes added after the end.
  fs.writeFileSync(a, Buffer.concat([goodA, Buffer.from([1])]));
  await assert.rejects(read, (e) => e.code === 'tampered');
});

test('a save keeps the file\'s unchanged frames byte for byte, and only those', async () => {
  const dir = tmp();
  const bytes = random(CHUNK * 5);
  const src = makeFile(dir, 'a.bin', bytes);
  const target = path.join(dir, 'a.rlock');
  const headerBuf = newHeader();
  await sealFile({ source: src, target, rootKey: KEY, headerBuf, meta: { name: 'a' }, stageDir: dir });
  const v1 = fs.readFileSync(target);
  const copy = Buffer.from(bytes);
  copy.writeUInt8(copy[CHUNK * 2 + 9] ^ 1, CHUNK * 2 + 9);
  const src2 = makeFile(dir, 'a2.bin', copy);
  const out = await sealFile({ source: src2, target, rootKey: KEY, headerBuf: fs.readFileSync(target).subarray(0, HEADER_SIZE), meta: { name: 'a' }, stageDir: dir, previous: target });
  assert.equal(out.reused, 4, 'four of five frames did not change');
  const v2 = fs.readFileSync(target);
  const f = (i) => [HEADER_SIZE + META_FRAME + i * DATA_FRAME, HEADER_SIZE + META_FRAME + (i + 1) * DATA_FRAME];
  assert.ok(sameBytes(v1.subarray(...f(0)), v2.subarray(...f(0))));
  assert.ok(!sameBytes(v1.subarray(...f(2)), v2.subarray(...f(2))), 'the changed frame was sealed anew');
  const back = path.join(dir, 'back.bin');
  await writePlain({ locked: target, dest: back, rootKey: KEY });
  assert.ok(sameBytes(fs.readFileSync(back), copy));
});

test('a frame from an earlier save put back into a later one is found', async () => {
  const dir = tmp();
  const bytes = random(CHUNK * 3);
  const src = makeFile(dir, 'a.bin', bytes);
  const target = path.join(dir, 'a.rlock');
  const headerBuf = newHeader();
  await sealFile({ source: src, target, rootKey: KEY, headerBuf, meta: { name: 'a' }, stageDir: dir });
  const v1 = fs.readFileSync(target);
  const copy = Buffer.from(bytes);
  copy.writeUInt8(copy[CHUNK + 5] ^ 1, CHUNK + 5);
  const src2 = makeFile(dir, 'a2.bin', copy);
  await sealFile({ source: src2, target, rootKey: KEY, headerBuf: v1.subarray(0, HEADER_SIZE), meta: { name: 'a' }, stageDir: dir, previous: target });
  const v2 = fs.readFileSync(target);
  // The whole-file hash in the metadata must catch a frame of v1 inside v2.
  const f1 = HEADER_SIZE + META_FRAME + DATA_FRAME;
  const mixed = Buffer.from(v2);
  v1.copy(mixed, f1, f1, f1 + DATA_FRAME);
  fs.writeFileSync(target, mixed);
  await assert.rejects(async () => {
    const o = await openLocked(target, KEY);
    try {
      for await (const c of o.frames()) void c;
    } finally {
      await o.close();
    }
  }, (e) => e.code === 'tampered');
});

test('a failure at every step of a save leaves the old locked file whole, and no plain file touched', async () => {
  const dir = tmp();
  const original = random(CHUNK * 2 + 9);
  const src = makeFile(dir, 'a.bin', original);
  const target = path.join(dir, 'a.rlock');
  const headerBuf = newHeader();
  await sealFile({ source: src, target, rootKey: KEY, headerBuf, meta: { name: 'a' }, stageDir: path.join(dir, 's') });
  const old = fs.readFileSync(target);
  const edited = makeFile(dir, 'b.bin', random(CHUNK * 2 + 9));
  for (const step of ['staged', 'verified', 'copied', 'swapped']) {
    const opts = {
      source: edited,
      target,
      rootKey: KEY,
      headerBuf: old.subarray(0, HEADER_SIZE),
      meta: { name: 'a' },
      stageDir: path.join(dir, 's'),
      previous: target,
      hooks: { at: (s) => { if (s === step) throw new Error(`stopped at ${s}`); }, rename: step === 'copied' ? async () => { const e = new Error('other drive'); e.code = 'EXDEV'; throw e; } : undefined, retries: 0 },
    };
    if (step === 'copied') fs.rmSync(target, { force: true }), fs.writeFileSync(target, old);
    if (step === 'swapped') {
      // A failure after the swap: the new file is the file now, and whole.
      await assert.rejects(() => sealFile(opts), /stopped at swapped/);
      const read = await verifyLocked(target, KEY);
      assert.equal(read.size, CHUNK * 2 + 9);
      fs.writeFileSync(target, old);
      continue;
    }
    await assert.rejects(() => sealFile(opts), (e) => /stopped/.test(e.message) || e.code === 'EXDEV', step);
    assert.ok(sameBytes(fs.readFileSync(target), old), `the old file is untouched after a failure at ${step}`);
    assert.ok(sameBytes(fs.readFileSync(src), original), 'the plain file is untouched');
    assert.equal(fs.readdirSync(dir).filter(isWorkingFile).length, 0, 'no working file left in the folder');
  }
});

test('a rename the system refuses is retried, then given up with the old file whole', async () => {
  const dir = tmp();
  const src = makeFile(dir, 'a.bin', random(1000));
  const target = path.join(dir, 'a.rlock');
  const headerBuf = newHeader();
  await sealFile({ source: src, target, rootKey: KEY, headerBuf, meta: { name: 'a' }, stageDir: dir });
  const old = fs.readFileSync(target);
  let tries = 0;
  const busy = async () => { tries += 1; const e = new Error('held'); e.code = 'EBUSY'; throw e; };
  await assert.rejects(() => sealFile({ source: src, target, rootKey: KEY, headerBuf, meta: { name: 'a' }, stageDir: dir, previous: target, hooks: { rename: busy, retries: 2 } }), (e) => e.code === 'EBUSY');
  assert.equal(tries, 3);
  assert.ok(sameBytes(fs.readFileSync(target), old));
  // And one that goes through on the second try is a normal save.
  tries = 0;
  const real = fsp.rename;
  const flaky = async (a, b) => { tries += 1; if (tries === 1) { const e = new Error('held'); e.code = 'EPERM'; throw e; } return real(a, b); };
  await sealFile({ source: src, target, rootKey: KEY, headerBuf, meta: { name: 'a' }, stageDir: dir, previous: target, hooks: { rename: flaky } });
  assert.equal(tries, 2);
});

test('across drives: copied beside the target, read back, renamed; a bad copy never replaces the file', async () => {
  const dir = tmp();
  const src = makeFile(dir, 'a.bin', random(CHUNK + 1));
  const target = path.join(dir, 'a.rlock');
  await sealFile({ source: src, target, rootKey: KEY, headerBuf: newHeader(), meta: { name: 'a' }, stageDir: path.join(dir, 'p') });
  const old = fs.readFileSync(target);
  const src2 = makeFile(dir, 'b.bin', random(CHUNK + 1));
  const realRename = fsp.rename;
  let first = true;
  const rename = async (a, b) => {
    if (first) {
      first = false;
      const e = new Error('other drive');
      e.code = 'EXDEV';
      throw e;
    }
    return realRename(a, b);
  };
  // A copy damaged before it is read back: the verify step refuses it.
  await assert.rejects(() => sealFile({ source: src2, target, rootKey: KEY, headerBuf: old.subarray(0, HEADER_SIZE), meta: { name: 'a' }, stageDir: path.join(dir, 'p'), hooks: { rename, at: (s) => {
    if (s === 'copied') {
      const near = fs.readdirSync(dir).find(isWorkingFile);
      const p = path.join(dir, near);
      const b = fs.readFileSync(p);
      b[b.length - 1] ^= 1;
      fs.writeFileSync(p, b);
    }
  }, retries: 0 } }), (e) => e.code === 'tampered' || e.code === 'failed');
  assert.ok(sameBytes(fs.readFileSync(target), old), 'the old file stands');
  assert.equal(fs.readdirSync(dir).filter(isWorkingFile).length, 0);
  // And without damage it goes through.
  first = true;
  await sealFile({ source: src2, target, rootKey: KEY, headerBuf: old.subarray(0, HEADER_SIZE), meta: { name: 'a' }, stageDir: path.join(dir, 'p'), hooks: { rename } });
  const back = path.join(dir, 'back.bin');
  await writePlain({ locked: target, dest: back, rootKey: KEY });
  assert.ok(sameBytes(fs.readFileSync(back), fs.readFileSync(src2)));
});

test('a program killed at each step of a save: the old file or the new one is whole, never neither', () => {
  const dir = tmp();
  const script = path.join(here, 'fixtures', 'locker-crash-child.js');
  const original = random(CHUNK * 4 + 11);
  const src = makeFile(dir, 'a.bin', original);
  const target = path.join(dir, 'a.rlock');
  const edited = makeFile(dir, 'b.bin', random(CHUNK * 4 + 11));
  const stage = path.join(dir, 'stage');
  const keyHex = KEY.toString('hex');
  const run = (step, mode) => spawnSync(process.execPath, [script, src, edited, target, stage, keyHex, step, mode], { encoding: 'utf8', env: { ...process.env, LOCKER_FORMAT: pathToFileURL(path.resolve(here, '../apps/desktop/main/extras/locker/format.js')).href, LOCKER_PLACE: pathToFileURL(path.resolve(here, '../apps/desktop/main/extras/locker/place.js')).href } });
  assert.equal(run('none', 'first').status, 0, 'the first lock finished');
  const firstRead = crypto.createHash('sha256').update(original).digest('hex');
  for (const step of ['staged', 'verified', 'swapped']) {
    const r = run(step, 'save');
    assert.equal(r.status, 9, `killed at ${step}: ${r.stderr}`);
    // Whichever is under the real name reads back whole.
    const read = spawnSync(process.execPath, [script, src, edited, target, stage, keyHex, 'read', 'read'], { encoding: 'utf8', env: { ...process.env, LOCKER_FORMAT: pathToFileURL(path.resolve(here, '../apps/desktop/main/extras/locker/format.js')).href, LOCKER_PLACE: pathToFileURL(path.resolve(here, '../apps/desktop/main/extras/locker/place.js')).href } });
    assert.equal(read.status, 0, `a whole file at ${step}: ${read.stderr}`);
    const got = read.stdout.trim();
    const newHash = crypto.createHash('sha256').update(fs.readFileSync(edited)).digest('hex');
    if (step === 'swapped') assert.equal(got, newHash, 'after the swap the new one is the file');
    else assert.equal(got, firstRead, `before the swap the old one is still the file (${step})`);
    assert.ok(fs.existsSync(src), 'the plain file was not touched');
  }
  // What a kill left behind is only working files, and the sweep removes them.
  return sweepWorking(stage).then(() => assert.equal(fs.readdirSync(stage).filter((n) => !isWorkingFile(n)).length, 0));
});

test('writing a plain file: refuses to overwrite unless asked, and a failure leaves nothing half-written', async () => {
  const dir = tmp();
  const bytes = random(CHUNK + 10);
  const src = makeFile(dir, 'a.bin', bytes);
  const target = path.join(dir, 'a.rlock');
  await sealFile({ source: src, target, rootKey: KEY, headerBuf: newHeader(), meta: { name: 'a.bin' }, stageDir: dir });
  const dest = path.join(dir, 'out.bin');
  fs.writeFileSync(dest, 'precious');
  await assert.rejects(() => writePlain({ locked: target, dest, rootKey: KEY }), (e) => e.code === 'exists');
  assert.equal(fs.readFileSync(dest, 'utf8'), 'precious');
  await assert.rejects(() => writePlain({ locked: target, dest, rootKey: KEY, overwrite: true, hooks: { at: (s) => { if (s === 'written') throw new Error('stop'); } } }), /stop/);
  assert.equal(fs.readFileSync(dest, 'utf8'), 'precious', 'the file already there is untouched by a failed write');
  assert.equal(fs.readdirSync(dir).filter(isWorkingFile).length, 0);
  await writePlain({ locked: target, dest, rootKey: KEY, overwrite: true });
  assert.ok(sameBytes(fs.readFileSync(dest), bytes));
});

test('the locked file\'s header says what it is without any key', async () => {
  const dir = tmp();
  const src = makeFile(dir, 'a.bin', random(10));
  const target = path.join(dir, 'a.rlock');
  const headerBuf = newHeader({ folderId: Buffer.alloc(16, 9) });
  await sealFile({ source: src, target, rootKey: KEY, headerBuf, meta: { name: 'a' }, stageDir: dir });
  const h = await readHeader(target);
  assert.equal(h.mode, MODE_FOLDER);
  assert.ok(h.folderId.equals(Buffer.alloc(16, 9)));
  assert.ok(parseHeader(headerBuf).fileId.equals(h.fileId));
});
