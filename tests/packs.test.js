// App packs (docs/APP-PACKS.md): made and signed by tools/make-packs.mjs
// from a small built tree made here, then found and checked by the core's
// loader (packages/office-shell/src/electron/packs.js). The key is made for
// the test and thrown away with its folder; nothing secret is in the
// repository.
import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { findPacks, readArchive, canonical, keyIdOf, PACK_FORMAT } from '../packages/office-shell/src/electron/packs.js';
import { makePacks, readPrivateKey, signManifest } from '../tools/make-packs.mjs';
import { makeKeyPair, writeKeyPair, insideRepository } from '../tools/make-pack-key.mjs';

const require = createRequire(import.meta.url);
const asar = require('@electron/asar');
const VERSION = '9.9.9';
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-packs-'));
test.after(() => fs.rmSync(tmp, { recursive: true, force: true }));

const keys = makeKeyPair();
const privateKey = crypto.createPrivateKey(keys.privatePem);
const publicKey = keys.publicPem;

/** A small built tree: Contacts, Mail (needs Contacts, a tile kept on disk, a file past a megabyte) and Text recognition. */
function tree() {
  const root = path.join(tmp, 'tree');
  const put = (p, body) => {
    const f = path.join(root, ...p.split('/'));
    fs.mkdirSync(path.dirname(f), { recursive: true });
    fs.writeFileSync(f, body);
  };
  put('contacts/pack.json', JSON.stringify({ name: 'Contacts', main: 'main/index.js', ipc: { contacts: ['list'] } }));
  put('contacts/main/index.js', 'export function register() {}\n');
  put('mail/pack.json', JSON.stringify({
    name: 'Mail',
    needs: ['contacts'],
    uses: ['ocr'],
    main: 'main/index.js',
    windows: 'renderer/mail.js',
    unpack: ['tiles'],
    ipc: { mail: ['list', 'send'] },
    events: ['mail:new'],
    tiles: [{ app: 'mail', icon: 'tiles/mail.ico' }],
    fileTypes: [{ ext: 'eml', name: 'Email message', mimeType: 'message/rfc822', role: 'Editor' }],
  }));
  put('mail/main/index.js', 'export function register() {}\n');
  put('mail/main/store.js', 'export const store = 1;\n');
  put('mail/renderer/mail.js', 'export default function Mail() {}\n');
  put('mail/tiles/mail.ico', Buffer.from([0, 0, 1, 0, 1, 0, 16, 16]));
  put('mail/data/big.bin', crypto.randomBytes(1536 * 1024));
  put('ocr/pack.json', JSON.stringify({ name: 'Text recognition', main: 'main/ocr.js', tools: ['read_text_in_picture'] }));
  put('ocr/main/ocr.js', 'export function register() {}\n');
  put('ocr/vendor/tessdata/eng.bin', crypto.randomBytes(4096));
  return root;
}

const built = path.join(tmp, 'built');
const made = await makePacks({ tree: tree(), out: built, version: VERSION, privateKey });

/** A copy of the installed packs to spoil, so each check starts from good ones. */
let copies = 0;
function installed() {
  const dir = path.join(tmp, `installed-${++copies}`);
  fs.cpSync(built, dir, { recursive: true });
  return dir;
}
const refusal = (result, id) => result.refused.find((r) => r.id === id);

test('make-packs: an archive and a signed manifest per pack, every file with its SHA-512', () => {
  assert.deepEqual(made.map((m) => m.id), ['contacts', 'mail', 'ocr']);
  const signed = JSON.parse(fs.readFileSync(path.join(built, 'mail.pack.json'), 'utf8'));
  assert.equal(signed.format, PACK_FORMAT);
  assert.equal(signed.keyId, keys.keyId);
  assert.equal(signed.manifest.version, VERSION);
  assert.deepEqual(Object.keys(signed.manifest.files).sort(), ['data/big.bin', 'main/index.js', 'main/store.js', 'renderer/mail.js', 'tiles/mail.ico']);
  assert.equal(signed.manifest.files['tiles/mail.ico'].unpacked, true);
  const big = fs.readFileSync(path.join(tmp, 'tree', 'mail', 'data', 'big.bin'));
  assert.equal(signed.manifest.files['data/big.bin'].sha512, crypto.createHash('sha512').update(big).digest('base64'));
  assert.ok(crypto.verify(null, Buffer.from(canonical(signed.manifest)), crypto.createPublicKey(publicKey), Buffer.from(signed.signature, 'base64')));
  // pack.json itself is not in the archive; the tile is beside it, not in it.
  const table = readArchive(path.join(built, 'mail.asar'));
  assert.equal(table.has('pack.json'), false);
  assert.equal(table.get('tiles/mail.ico').unpacked, true);
  assert.ok(fs.existsSync(path.join(built, 'mail.asar.unpacked', 'tiles', 'mail.ico')));
});

test('the loader reads an archive as Electron\'s own tool wrote it', () => {
  const archive = path.join(built, 'mail.asar');
  const table = readArchive(archive);
  const listed = asar.listPackage(archive, {}).map((p) => p.replace(/\\/g, '/').replace(/^\//, ''));
  const files = listed.filter((p) => !('files' in asar.statFile(archive, p)));
  assert.deepEqual([...table.keys()].sort(), files.sort());
  const entry = table.get('main/store.js');
  const fd = fs.openSync(archive, 'r');
  const bytes = Buffer.alloc(entry.size);
  fs.readSync(fd, bytes, 0, entry.size, entry.offset);
  fs.closeSync(fd);
  assert.deepEqual(bytes, asar.extractFile(archive, 'main/store.js'));
});

test('good packs load, what a pack needs first, with paths into their archives', () => {
  const dir = installed();
  const { packs, refused } = findPacks({ dir, publicKey, version: VERSION });
  assert.deepEqual(refused, []);
  assert.deepEqual(packs.map((p) => p.id), ['contacts', 'mail', 'ocr']);
  const mail = packs.find((p) => p.id === 'mail');
  assert.equal(mail.name, 'Mail');
  assert.equal(mail.main, path.join(dir, 'mail.asar', 'main', 'index.js'));
  assert.equal(mail.windows, path.join(dir, 'mail.asar', 'renderer', 'mail.js'));
  assert.equal(mail.tiles[0].icon, path.join(dir, 'mail.asar.unpacked', 'tiles', 'mail.ico'));
  assert.deepEqual(mail.ipc, { mail: ['list', 'send'] });
  assert.deepEqual(mail.events, ['mail:new']);
  assert.deepEqual(mail.fileTypes.map((t) => t.ext), ['eml']);
  assert.deepEqual(mail.needs, ['contacts']);
  assert.deepEqual(mail.uses, ['ocr']);
  assert.deepEqual(packs.find((p) => p.id === 'ocr').tools, ['read_text_in_picture']);
  // A folder given relative to where the core runs reads the same.
  const relative = path.relative(process.cwd(), dir);
  if (!path.isAbsolute(relative)) assert.equal(findPacks({ dir: relative, publicKey, version: VERSION }).packs.length, 3);
});

test('no packs folder: nothing to load, nothing refused', () => {
  assert.deepEqual(findPacks({ dir: path.join(tmp, 'none'), publicKey, version: VERSION }), { packs: [], refused: [] });
});

test('a pack of another version is refused, saying both versions', () => {
  const { packs, refused } = findPacks({ dir: installed(), publicKey, version: '9.9.10' });
  assert.deepEqual(packs, []);
  const mail = refusal({ refused }, 'mail');
  assert.equal(mail.code, 'version');
  assert.match(mail.reason, /The Mail pack is version 9\.9\.9 and this Rutba Office is 9\.9\.10/);
});

test('a file changed inside the archive is refused, and so is the pack that needs it', () => {
  const dir = installed();
  const archive = path.join(dir, 'contacts.asar');
  const entry = readArchive(archive).get('main/index.js');
  const fd = fs.openSync(archive, 'r+');
  const b = Buffer.alloc(1);
  fs.readSync(fd, b, 0, 1, entry.offset);
  b[0] ^= 0x20;
  fs.writeSync(fd, b, 0, 1, entry.offset);
  fs.closeSync(fd);
  const result = findPacks({ dir, publicKey, version: VERSION });
  assert.equal(refusal(result, 'contacts').code, 'changed');
  assert.match(refusal(result, 'contacts').reason, /The Contacts pack's file main\/index\.js has been changed since it was signed/);
  assert.equal(refusal(result, 'mail').code, 'needs');
  assert.equal(refusal(result, 'mail').reason, 'The Mail pack was not loaded. It needs contacts, which could not be loaded either.');
  assert.deepEqual(result.packs.map((p) => p.id), ['ocr']);
});

test('a file changed beside the archive (unpacked) is refused', () => {
  const dir = installed();
  fs.writeFileSync(path.join(dir, 'mail.asar.unpacked', 'tiles', 'mail.ico'), Buffer.from([1, 2, 3, 4, 5, 6, 7, 8]));
  assert.equal(refusal(findPacks({ dir, publicKey, version: VERSION }), 'mail').code, 'changed');
});

test('a pack signed with another key is refused', () => {
  const other = makeKeyPair();
  const result = findPacks({ dir: installed(), publicKey: other.publicPem, version: VERSION });
  assert.deepEqual(result.packs, []);
  assert.equal(refusal(result, 'mail').code, 'signature');
  assert.match(refusal(result, 'mail').reason, /The mail pack is not signed by Rutba/);
  // Either of two keys will do: the core can carry the next key before the last is retired.
  assert.equal(findPacks({ dir: installed(), publicKey: [other.publicPem, publicKey], version: VERSION }).packs.length, 3);
});

test('a manifest edited after signing is refused', () => {
  const dir = installed();
  const file = path.join(dir, 'ocr.pack.json');
  const signed = JSON.parse(fs.readFileSync(file, 'utf8'));
  signed.manifest.version = '9.9.10';
  fs.writeFileSync(file, JSON.stringify(signed));
  assert.equal(refusal(findPacks({ dir, publicKey, version: '9.9.10' }), 'ocr').code, 'signature');
});

test('a pack whose needs are missing is refused, naming them', () => {
  const dir = installed();
  for (const f of ['contacts.asar', 'contacts.pack.json']) fs.rmSync(path.join(dir, f));
  const result = findPacks({ dir, publicKey, version: VERSION });
  assert.equal(refusal(result, 'mail').code, 'needs');
  assert.equal(refusal(result, 'mail').reason, 'The Mail pack was not loaded. It needs contacts, which is not installed: add it with Add or remove apps.');
  assert.deepEqual(result.packs.map((p) => p.id), ['ocr']);
});

test('a file the manifest does not list is refused, in the archive or beside it', () => {
  const dir = installed();
  fs.writeFileSync(path.join(dir, 'mail.asar.unpacked', 'tiles', 'extra.dll'), 'x');
  const beside = refusal(findPacks({ dir, publicKey, version: VERSION }), 'mail');
  assert.equal(beside.code, 'unknown');
  assert.match(beside.reason, /does not list: tiles\/extra\.dll/);

  // The archive made again from its files and one more, the signed manifest kept.
  const dir2 = installed();
  const src = path.join(tmp, 'extra-src');
  fs.cpSync(path.join(tmp, 'tree', 'ocr'), src, { recursive: true });
  fs.rmSync(path.join(src, 'pack.json'));
  fs.writeFileSync(path.join(src, 'main', 'payload.js'), 'globalThis.x = 1;\n');
  const files = ['main/ocr.js', 'main/payload.js', 'vendor/tessdata/eng.bin'].map((p) => path.join(src, ...p.split('/')));
  return asar.createPackageFromFiles(src, path.join(dir2, 'ocr.asar'), files).then(() => {
    const inside = refusal(findPacks({ dir: dir2, publicKey, version: VERSION }), 'ocr');
    assert.equal(inside.code, 'unknown');
    assert.match(inside.reason, /does not list: main\/payload\.js/);
  });
});

test('a missing archive or file is refused', () => {
  const dir = installed();
  fs.rmSync(path.join(dir, 'ocr.asar'));
  assert.equal(refusal(findPacks({ dir, publicKey, version: VERSION }), 'ocr').code, 'missing');
  const dir2 = installed();
  fs.rmSync(path.join(dir2, 'mail.asar.unpacked', 'tiles', 'mail.ico'));
  assert.equal(refusal(findPacks({ dir: dir2, publicKey, version: VERSION }), 'mail').code, 'missing');
});

test('a signed manifest that names a path out of its folder is refused', () => {
  const dir = installed();
  const file = path.join(dir, 'ocr.pack.json');
  const { manifest } = JSON.parse(fs.readFileSync(file, 'utf8'));
  manifest.files['../escape.js'] = { size: 1, sha512: 'x' };
  fs.writeFileSync(file, JSON.stringify(signManifest(manifest, privateKey)));
  assert.equal(refusal(findPacks({ dir, publicKey, version: VERSION }), 'ocr').code, 'format');
  // And a manifest that names itself differently from its file.
  const dir2 = installed();
  const file2 = path.join(dir2, 'ocr.pack.json');
  const signed2 = JSON.parse(fs.readFileSync(file2, 'utf8'));
  fs.writeFileSync(file2, JSON.stringify(signManifest({ ...signed2.manifest, id: 'mail' }, privateKey)));
  assert.equal(refusal(findPacks({ dir: dir2, publicKey, version: VERSION }), 'ocr').code, 'id');
});

test('the private key is never read from or written into the repository', () => {
  const repoFile = path.join(path.dirname(fileURLToPath(import.meta.url)), 'fixtures', 'pack-key.pem');
  assert.equal(insideRepository(repoFile), true);
  assert.throws(() => readPrivateKey({ file: repoFile, env: {} }), /never read from inside the repository/);
  assert.throws(() => writeKeyPair(path.dirname(repoFile)), /never kept inside the repository/);
  assert.throws(() => readPrivateKey({ env: {} }), /No pack key/);
  // From the environment, as the release machine gives it.
  assert.equal(keyIdOf(crypto.createPublicKey(readPrivateKey({ env: { RUTBA_PACK_KEY: keys.privatePem } }))), keys.keyId);
  // And a key pair written outside it, the private half not readable by others where the system allows.
  const made = writeKeyPair(path.join(tmp, 'keys'));
  assert.equal(made.keyId, keyIdOf(fs.readFileSync(made.publicFile, 'utf8')));
  assert.equal(keyIdOf(crypto.createPublicKey(readPrivateKey({ file: made.privateFile, env: {} }))), made.keyId);
  assert.throws(() => writeKeyPair(path.join(tmp, 'keys')), /already exists/);
});

test('the signed bytes do not depend on the order keys were written in', () => {
  assert.equal(canonical({ b: 1, a: { d: [1, { f: 2, e: 3 }], c: 'x' } }), canonical({ a: { c: 'x', d: [1, { e: 3, f: 2 }] }, b: 1 }));
  assert.equal(canonical({ a: undefined, b: null }), '{"b":null}');
});
