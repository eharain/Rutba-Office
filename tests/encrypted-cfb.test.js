/**
 * The compound-file writer, read back by the compound-file reader.
 *
 * A password-protected file is a compound file, so this suite now writes
 * one: storages, streams in the mini stream and in regular sectors, the
 * directory's red-black trees, and a FAT long enough to need DIFAT sectors.
 * Everything written is read back with the reader that already opens .msg
 * and 97-2003 files, and the \x06DataSpaces streams are compared with the
 * ones inside a file another program encrypted.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { randomBytes } from 'node:crypto';
import { CompoundFile } from '@rutba/office-formats/cfb';
import { writeCompoundFile, compareNames } from '@rutba/office-formats/cfb-write';
import { dataSpacesStreams } from '@rutba/office-formats/crypt';

const FIXTURES = path.join(path.dirname(fileURLToPath(import.meta.url)), 'fixtures');
const NOSTREAM = 0xffffffff;

/** The colour byte of each entry, which the reader does not keep. */
function colours(bytes, cfb) {
  const out = [];
  const perSector = 512 / 128;
  let s = cfb.dirStart;
  let i = 0;
  while (s < 0xfffffffa) {
    for (let k = 0; k < perSector; k++, i++) out[i] = bytes[(s + 1) * 512 + k * 128 + 67];
    s = cfb.fat[s];
  }
  return out;
}

/** Every storage's children form a binary search tree that is a valid red-black tree. */
function assertRedBlack(bytes, cfb) {
  const colour = colours(bytes, cfb);
  assert.equal(colour[0], 1, 'the root entry is black');
  for (const storage of cfb.entries.filter((e) => e.type === 1 || e.type === 5)) {
    if (storage.child === NOSTREAM) continue;
    assert.equal(colour[storage.child], 1, `${storage.name}'s tree has a black root`);
    const blackHeight = (id, parentRed) => {
      if (id === NOSTREAM) return 1;
      const e = cfb.entries[id];
      const red = colour[id] === 0;
      assert.ok(!(red && parentRed), `no red node under a red node (${e.name})`);
      if (e.left !== NOSTREAM) assert.ok(compareNames(cfb.entries[e.left].name, e.name) < 0, 'left sorts before');
      if (e.right !== NOSTREAM) assert.ok(compareNames(cfb.entries[e.right].name, e.name) > 0, 'right sorts after');
      const l = blackHeight(e.left, red);
      const r = blackHeight(e.right, red);
      assert.equal(l, r, `every path below ${e.name} has as many black nodes`);
      return l + (red ? 0 : 1);
    };
    blackHeight(storage.child, false);
  }
}

test('what the writer writes, the reader reads back', () => {
  const small = randomBytes(1000);
  const edge = randomBytes(4096); // exactly the cutoff: a regular stream
  const under = randomBytes(4095); // one under: the mini stream
  const large = randomBytes(70000);
  const many = Array.from({ length: 37 }, (_, i) => ({ path: ['Many', `Entry ${String(i).padStart(2, '0')}${'x'.repeat(i % 5)}`], data: Buffer.from(`value ${i}`) }));
  const bytes = writeCompoundFile([
    { path: ['Small'], data: small },
    { path: ['Edge'], data: edge },
    { path: ['Under'], data: under },
    { path: ['Deep', 'Deeper', 'Large'], data: large },
    { path: ['Deep', 'Empty'], data: new Uint8Array(0) },
    ...many,
  ]);
  assert.equal(bytes.length % 512, 0);
  const cfb = new CompoundFile(bytes);
  assert.equal(cfb.major, 3);
  assert.equal(cfb.sectorSize, 512);
  assert.equal(cfb.miniCutoff, 4096);
  const read = (p) => Buffer.from(cfb.read(cfb.find(p)));
  assert.deepEqual(read(['Small']), small);
  assert.deepEqual(read(['Edge']), edge);
  assert.deepEqual(read(['Under']), under);
  assert.deepEqual(read(['Deep', 'Deeper', 'Large']), large);
  assert.equal(read(['Deep', 'Empty']).length, 0);
  for (const m of many) assert.equal(read(m.path).toString(), m.data.toString());
  assert.equal(cfb.childrenOf(cfb.find(['Many'])).length, 37);
  assert.equal(cfb.find(['Deep']).type, 1, 'a storage');
  assertRedBlack(bytes, cfb);
});

test('a file past 109 FAT sectors chains its FAT through DIFAT sectors', () => {
  // 109 FAT sectors map 109 × 128 × 512 bytes, about 7 MB; a document with
  // photographs in it is bigger than that.
  const big = randomBytes(8 * 1024 * 1024);
  const bytes = writeCompoundFile([{ path: ['EncryptedPackage'], data: big }, { path: ['EncryptionInfo'], data: randomBytes(900) }]);
  const cfb = new CompoundFile(bytes);
  assert.ok(cfb.difatCount >= 1, 'DIFAT sectors are used');
  assert.ok(cfb.fatCount > 109);
  assert.ok(Buffer.from(cfb.read(cfb.find(['EncryptedPackage']))).equals(big));
  assert.equal(cfb.read(cfb.find(['EncryptionInfo'])).length, 900);
});

test('the \\x06DataSpaces streams match the ones another program wrote', () => {
  // The fixture's Version, DataSpaceMap and DataSpaceInfo streams are the
  // known answer for [MS-OFFCRYPTO] 2.1's structures; its transform stream
  // is misnamed ("Primary" without the \x06) but the same 200 bytes long.
  const fixture = new CompoundFile(fs.readFileSync(path.join(FIXTURES, 'encrypted.docx')));
  const ours = new Map(dataSpacesStreams().map((s) => [s.path.join('/'), Buffer.from(s.data)]));
  for (const p of [['\u0006DataSpaces', 'Version'], ['\u0006DataSpaces', 'DataSpaceMap'], ['\u0006DataSpaces', 'DataSpaceInfo', 'StrongEncryptionDataSpace']]) {
    assert.deepEqual(ours.get(p.join('/')), Buffer.from(fixture.read(fixture.find(p))), p.join('/'));
  }
  const primary = ours.get(['\u0006DataSpaces', 'TransformInfo', 'StrongEncryptionTransform', '\u0006Primary'].join('/'));
  assert.equal(primary.length, fixture.find(['\u0006DataSpaces', 'TransformInfo', 'StrongEncryptionTransform', 'Primary']).size);
  // TransformInfoHeader: its length (88 bytes before the name), type 1, the
  // encryption transform's id and name, versions 1.0; then no encryption
  // name, block size 0, cipher mode 0, reserved 4.
  assert.equal(primary.readUInt32LE(0), 88);
  assert.equal(primary.readUInt32LE(4), 1);
  assert.equal(primary.subarray(12, 88).toString('utf16le'), '{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}');
  assert.equal(primary.subarray(92, 170).toString('utf16le'), 'Microsoft.Container.EncryptionTransform');
  assert.deepEqual([...primary.subarray(184)], [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 4, 0, 0, 0]);
});

test('entry names are ordered as the spec orders them and refused when they cannot be written', () => {
  assert.ok(compareNames('B', 'AA') < 0, 'shorter first');
  assert.equal(compareNames('abc', 'ABC'), 0, 'case does not matter');
  assert.ok(compareNames('ABD', 'abc') > 0);
  assert.throws(() => writeCompoundFile([{ path: ['a/b'], data: Buffer.from('x') }]), /cannot name/);
  assert.throws(() => writeCompoundFile([{ path: ['x'.repeat(32)], data: Buffer.from('x') }]), /cannot name/);
  assert.throws(() => writeCompoundFile([{ path: ['Same'], data: Buffer.from('x') }, { path: ['SAME'], data: Buffer.from('y') }]), /named twice/);
});
