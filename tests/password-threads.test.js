/**
 * Passwords hashed off the document service's thread (apps/desktop/main/
 * password-host.js): the threads give exactly what the engine gives
 * itself, for known vectors and for every kind of password Office hashes;
 * an answer worked out ahead is the one the engine uses; and the service's
 * thread goes on answering while 100,000 rounds are spun for Protect
 * Workbook, Restrict Editing, Encrypt with Password and opening an
 * encrypted file.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { createPasswordHost, passwordHost } from '../apps/desktop/main/password-host.js';
import { createDocumentService } from '../apps/desktop/main/documents.js';
import {
  hashPassword, checkPassword, checkWordPassword, passwordAttrs, wordProtectionAttrs, wordLegacyHex,
  rememberPasswordHash, offerPasswordSalt, collectPasswordHashes, newPasswordSpec,
} from '@rutba/ooxml/protection';
import { spinPassword, encryptPackage, decryptPackage, EncryptedFileError, isEncryptedPackage } from '@rutba/office-formats/crypt';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml';
import { WRONG_PASSWORD } from '../packages/sheet-view/src/view.js';

// The known vector of sheet-workbook-protection.test.js, worked out with .NET's SHA512.
const SALT = 'AAECAwQFBgcICQoLDA0ODw==';
const RUTBA_HASH = 'VuujxQaYHeJ14+FUG9kIYQMcmdZUeAqNzfZavjHt5PbCYp6Sc2+KgimXS6nhmw/lRB0EHrODhQMGCqig3NYTYQ==';

/** The algorithm as ECMA-376 writes it, step by step, with a hash object per round. */
function longhand(password, salt, spinCount, alg = 'sha512') {
  let h = createHash(alg).update(Buffer.from(salt, 'base64')).update(Buffer.from(password, 'utf16le')).digest();
  for (let i = 0; i < spinCount; i++) {
    const it = Buffer.alloc(4);
    it.writeUInt32LE(i, 0);
    h = createHash(alg).update(h).update(it).digest();
  }
  return h.toString('base64');
}

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-pwthreads-'));
const write = (name, bytes) => {
  const p = path.join(dir, name);
  fs.writeFileSync(p, bytes);
  return p;
};
const host = createPasswordHost({ threads: 2 });
test.after(async () => {
  await host.stop();
  await passwordHost().stop();
  fs.rmSync(dir, { recursive: true, force: true });
});

test('a thread hashes a password exactly as the engine does, for the known vector and every kind Office writes', async () => {
  assert.equal(await host.hash({ password: 'Rutba', salt: SALT, spinCount: 100000, algorithm: 'SHA-512' }), RUTBA_HASH);
  const cases = [
    ['Rutba', SALT, 100000, 'SHA-512'],
    ['', SALT, 100000, 'SHA-512'],
    ['pässwörd ✓ پاس', 'q83vASNFZ4mrze8BI0VniQ==', 5000, 'SHA-512'],
    [wordLegacyHex('Contract'), 'ABEiM0RVZneImaq7zN3u/w==', 100000, 'SHA-512'],
    ['legacy', SALT, 50000, 'SHA-1'],
    ['two five six', SALT, 1000, 'SHA-256'],
    ['no rounds', SALT, 0, 'SHA-512'],
  ];
  const threaded = await Promise.all(cases.map(([password, salt, spinCount, algorithm]) => host.hash({ password, salt, spinCount, algorithm })));
  cases.forEach(([password, salt, spinCount, algorithm], i) => {
    const alg = { 'SHA-512': 'sha512', 'SHA-1': 'sha1', 'SHA-256': 'sha256' }[algorithm];
    assert.equal(threaded[i], hashPassword(password, { salt, spinCount, algorithm }), `${algorithm} × ${spinCount}`);
    assert.equal(threaded[i], longhand(password, salt, spinCount, alg), `${algorithm} × ${spinCount}, as the standard writes it`);
  });
  assert.ok(host.stats.threaded >= cases.length + 1, 'worked out on the threads');
  assert.equal(host.stats.inPlace, 0);
});

test('the key an encrypted file is spun from, and the file itself, come back from a thread as the engine makes them', async () => {
  const salt = Buffer.from('00112233445566778899aabbccddeeff', 'hex');
  for (const alg of ['sha512', 'sha1']) {
    const spun = await host.spin({ password: 'Tr0ub4dor&3', salt, spinCount: 100000, alg });
    assert.ok(Buffer.isBuffer(spun));
    assert.equal(spun.toString('hex'), spinPassword('Tr0ub4dor&3', salt, 100000, alg).toString('hex'), alg);
  }
  const plain = buildDocx({ paragraphs: ['For the board only.'] });
  const locked = encryptPackage(plain, 'Open1', { spinCount: 20000 });
  const out = await host.decrypt(locked, 'Open1');
  const here = decryptPackage(locked, 'Open1');
  assert.ok(Buffer.isBuffer(out.bytes));
  assert.equal(Buffer.compare(out.bytes, here.bytes), 0);
  assert.equal(out.kind, 'agile');
  assert.equal(out.integrity, true);
  await assert.rejects(host.decrypt(locked, 'open1'), (err) => err instanceof EncryptedFileError && err.code === 'password');
  const changed = Buffer.from(locked);
  // A byte of the encrypted package changed: the HMAC refuses it, from the thread as here.
  const at = changed.length - 700;
  changed[at] ^= 0xff;
  let codeHere = null;
  try { decryptPackage(changed, 'Open1'); } catch (err) { codeHere = err.code; }
  await assert.rejects(host.decrypt(changed, 'Open1'), (err) => err instanceof EncryptedFileError && err.code === codeHere);
});

test('a host that cannot use a thread works the same answers out in place', async () => {
  const inPlace = createPasswordHost({ worker: false });
  assert.equal(await inPlace.hash({ password: 'Rutba', salt: SALT, spinCount: 100000, algorithm: 'SHA-512' }), RUTBA_HASH);
  assert.equal(inPlace.stats.inPlace, 1);
  assert.equal(inPlace.stats.threaded, 0);
});

test('a hash worked out ahead is used once, a check lists what it would hash without hashing, and an offered salt is written', () => {
  // Remembered: the next hash of that spec answers with it, once.
  const spec = { password: 'ahead', salt: SALT, spinCount: 100000, algorithm: 'SHA-512' };
  rememberPasswordHash(spec, 'from-the-thread');
  assert.equal(hashPassword('ahead', spec), 'from-the-thread');
  assert.equal(hashPassword('ahead', spec), longhand('ahead', SALT, 100000));

  // Collected: the check runs, nothing is spun, each hash it wanted is listed.
  const a = { algorithmName: 'SHA-512', hashValue: RUTBA_HASH, saltValue: SALT, spinCount: '100000' };
  const t0 = performance.now();
  const listed = collectPasswordHashes(() => checkPassword('Rutba', a));
  assert.ok(performance.now() - t0 < 60, 'nothing spun');
  assert.deepEqual(listed, [{ password: 'Rutba', salt: SALT, spinCount: 100000, algorithm: 'SHA-512' }]);
  assert.deepEqual(collectPasswordHashes(() => checkPassword('x', { sheet: '1' })), [], 'no password: nothing to hash');
  const word = wordProtectionAttrs('Contract', { salt: SALT, spinCount: 1000 });
  const read = Object.fromEntries(Object.entries(word).map(([k, v]) => [k.slice(2), v]));
  assert.deepEqual(collectPasswordHashes(() => checkWordPassword('Contract', read)).map((s) => s.password), [wordLegacyHex('Contract'), 'Contract'], 'Word tries its key, then the password itself');
  // ...and with what was listed remembered, the real check is right without spinning.
  for (const s of collectPasswordHashes(() => checkPassword('Rutba', a))) rememberPasswordHash(s, hashPassword(s.password, s));
  assert.equal(checkPassword('Rutba', a), true);

  // Offered: a new password is written with the salt its hash was made for.
  const fresh = newPasswordSpec('Budget2026');
  offerPasswordSalt(fresh);
  const attrs = passwordAttrs('Budget2026', { prefix: 'workbook' });
  assert.equal(attrs.workbookSaltValue, fresh.salt);
  assert.equal(attrs.workbookHashValue, longhand('Budget2026', fresh.salt, 100000));
  assert.notEqual(passwordAttrs('Budget2026').saltValue, fresh.salt, 'each offered salt is used once');
  const wordFresh = newPasswordSpec('Contract', { word: true });
  assert.equal(wordFresh.password, wordLegacyHex('Contract'));
  offerPasswordSalt(wordFresh);
  assert.equal(wordProtectionAttrs('Contract')['w:salt'], wordFresh.salt);
});

/** Resolves once `promise` settles, with how many 10 ms ticks of this thread ran meanwhile. */
async function ticksWhile(promise) {
  let ticks = 0;
  const timer = setInterval(() => { ticks++; }, 10);
  try {
    await promise.catch(() => {});
  } finally {
    clearInterval(timer);
  }
  return ticks;
}

test('Protect and Unprotect Workbook and Sheet hash on the threads: the service answers meanwhile, and the file is what Excel writes', async () => {
  const doc = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }), recoveryDir: null });
  const file = write('book.xlsx', buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['a', 1]] }, { name: 'Two', rows: [] }] }));
  const s = doc.open({ path: file, kind: 'sheet' });
  const other = doc.open({ path: write('other.xlsx', buildXlsx({ sheets: [{ name: 'S', rows: [] }] })), kind: 'sheet' });
  const before = passwordHost().stats.threaded;

  // No password, no hashing: answered at once, as always.
  const plain = doc.apply({ id: s.id, ops: [{ op: 'protect' }] });
  assert.equal(typeof plain.then, 'undefined');
  doc.apply({ id: s.id, ops: [{ op: 'unprotect' }] });

  const protecting = doc.apply({ id: s.id, ops: [{ op: 'protectWorkbook', password: 'Budget2026' }] });
  assert.equal(typeof protecting.then, 'function', 'a password: answered with a promise');
  // The thread is free: another workbook's edit is answered before the hash is done.
  const typed = doc.apply({ id: other.id, ops: [{ op: 'select', row: 0, col: 0 }, { op: 'setCell', row: 0, col: 0, value: 'typed' }] });
  assert.equal(typeof typed.then, 'undefined');
  assert.ok((await ticksWhile(protecting)) >= 3, 'the thread ticked while the password was hashed');
  const locked = await protecting;
  assert.equal(locked.model.workbookProtection.structure, true);

  // A wrong password: Excel's words, after the hash, with the thread free.
  const wrong = doc.apply({ id: s.id, ops: [{ op: 'unprotectWorkbook', password: 'budget2026' }] });
  assert.ok((await ticksWhile(wrong)) >= 3);
  await assert.rejects(wrong, (err) => err.message === WRONG_PASSWORD);
  // An edit sent while the check waits is answered after it, in order.
  const order = [];
  const right = doc.apply({ id: s.id, ops: [{ op: 'unprotectWorkbook', password: 'Budget2026' }] }).then(() => order.push('unprotect'));
  const after = Promise.resolve(doc.apply({ id: s.id, ops: [{ op: 'addSheet' }] })).then(() => order.push('addSheet'));
  await Promise.all([right, after]);
  assert.deepEqual(order, ['unprotect', 'addSheet'], 'the sheet added after the workbook was unprotected, not refused before');

  // Protect Sheet with a password and the saved hash, byte for byte the documented algorithm.
  await doc.apply({ id: s.id, ops: [{ op: 'protect', password: 'Sheet1' }] });
  await doc.apply({ id: s.id, ops: [{ op: 'protectWorkbook', password: 'Budget2026' }] });
  doc.save({ id: s.id });
  const pkg = OoxmlPackage.read(fs.readFileSync(file));
  const wb = /<workbookProtection\b[^>]*>/.exec(pkg.text('xl/workbook.xml'))[0];
  const wsalt = /workbookSaltValue="([^"]+)"/.exec(wb)[1];
  assert.equal(/workbookHashValue="([^"]+)"/.exec(wb)[1], longhand('Budget2026', wsalt, 100000));
  // On the active sheet: the one Insert Sheet added last.
  const sp = pkg.partNames().filter((p) => /^xl\/worksheets\/sheet\d+\.xml$/.test(p)).map((p) => /<sheetProtection\b[^>]*>/.exec(pkg.text(p))?.[0]).find(Boolean);
  assert.equal(/hashValue="([^"]+)"/.exec(sp)[1], longhand('Sheet1', /saltValue="([^"]+)"/.exec(sp)[1], 100000));
  await assert.rejects(Promise.resolve().then(() => doc.apply({ id: s.id, ops: [{ op: 'unprotect', password: 'sheet1' }] })), (err) => err.message === WRONG_PASSWORD);
  await doc.apply({ id: s.id, ops: [{ op: 'unprotect', password: 'Sheet1' }] });
  assert.ok(passwordHost().stats.threaded - before >= 6, 'each password hashed on a thread');
  doc.close({ id: s.id });
  doc.close({ id: other.id });
});

test('Restrict Editing on and off hash on the threads, and a wrong password keeps Word\'s sentence', async () => {
  const doc = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }), recoveryDir: null });
  const file = write('memo.docx', buildDocx({ paragraphs: ['One', 'Two'] }));
  const d = doc.open({ path: file, kind: 'doc' });
  const on = doc.apply({ id: d.id, ops: [{ op: 'setProtection', edit: 'readOnly', password: 'Contract' }] });
  assert.equal(typeof on.then, 'function');
  assert.ok((await ticksWhile(on)) >= 3);
  await on;
  const wrong = doc.apply({ id: d.id, ops: [{ op: 'stopProtection', password: 'contract' }] });
  assert.ok((await ticksWhile(wrong)) >= 3);
  await assert.rejects(wrong, /That password is not right, so the protection stays on\./);
  doc.save({ id: d.id });
  const settings = OoxmlPackage.read(fs.readFileSync(file)).text('word/settings.xml');
  const salt = /w:salt="([^"]+)"/.exec(settings)[1];
  assert.equal(/w:hash="([^"]+)"/.exec(settings)[1], longhand(wordLegacyHex('Contract'), salt, 100000));
  await doc.apply({ id: d.id, ops: [{ op: 'stopProtection', password: 'Contract' }] });
  assert.equal(doc.model({ id: d.id }).protection?.enforced ?? false, false);
  doc.close({ id: d.id });
});

test('Encrypt with Password spins the key on a thread before the save, and an encrypted file opens with its password tried on a thread', async () => {
  const doc = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }), recoveryDir: null });
  const file = write('board.docx', buildDocx({ paragraphs: ['For the board only.'] }));
  const d = doc.open({ path: file, kind: 'doc' });
  const before = passwordHost().stats.threaded;
  doc.setPassword({ id: d.id, password: 'Tr0ub4dor&3' });
  const saving = doc.saveAsync({ id: d.id });
  assert.ok((await ticksWhile(saving)) >= 3, 'the thread ticked while the key was spun');
  await saving;
  const saved = fs.readFileSync(file);
  assert.equal(isEncryptedPackage(saved), true);
  assert.match(OoxmlPackage.read(decryptPackage(saved, 'Tr0ub4dor&3').bytes).text('word/document.xml'), /For the board only\./);
  doc.close({ id: d.id });

  const wrong = doc.openAsync({ path: file, kind: 'doc', password: 'tr0ub4dor&3' });
  assert.ok((await ticksWhile(wrong)) >= 3);
  assert.deepEqual([(await wrong).locked, (await wrong).wrong], [true, true]);
  const opened = await doc.openAsync({ path: file, kind: 'doc', password: 'Tr0ub4dor&3' });
  assert.equal(opened.encrypted, true);
  assert.ok(passwordHost().stats.threaded - before >= 3, 'the key and both tries on a thread');
  // The synchronous open, which the checks and tests use, still answers the same.
  assert.equal(doc.open({ path: file, kind: 'doc', password: 'nope' }).wrong, true);
  doc.close({ id: opened.id });
});
