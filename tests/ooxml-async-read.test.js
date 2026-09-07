/**
 * The async read door, proven equivalent to the sync one.
 *
 * WHY: the OOXML engine runs on the server today for exactly one reason —
 * `zlib.inflateRawSync` (and Buffer, which is a shim, not an architecture).
 * `OoxmlPackage.readAsync` removes the zlib dependence from the read path by
 * pre-inflating every deflated entry through the web-standard
 * `DecompressionStream('deflate-raw')`, then handing back the same package
 * object everything else already consumes synchronously.
 *
 * "Same" is the entire claim, so these tests are equivalence tests: same part
 * names, byte-identical part contents, byte-identical deterministic write()
 * output, and a Workbook that answers identically over either package. The
 * stream inflater is forced with the internal `{ inflate: 'stream' }` option —
 * no monkey-patching of globals — so the browser path is what actually runs
 * under Node here, on a Node that has zlib and would otherwise never take it.
 *
 * The CRC-32 fallback gets the same treatment: `crc32Js` is what a browser
 * build will run, so it is pinned against `zlib.crc32` and against published
 * check values, not merely trusted.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import zlib from 'node:zlib';
import { randomBytes } from 'node:crypto';
import { buildXlsx, OoxmlPackage, Workbook } from '@rutba/ooxml';
import { crc32, crc32Js, inflateRawStream } from '@rutba/ooxml/zip';

const workbookBytes = () => buildXlsx({
  sheets: [
    { name: 'Data', rows: [['a', 'b&<>"é'], [1, 2]] },
    { name: 'Calc', rows: [[10, 20, '=A1+B1']] },
  ],
  definedNames: [{ name: 'Rng', ref: 'Data!$A$1:$B$2' }],
});

/** Every part, byte for byte, plus names and order. */
function assertPackagesEqual(actual, expected) {
  assert.deepStrictEqual(actual.partNames(), expected.partNames());
  for (const name of expected.partNames()) {
    const a = actual.read(name);
    const e = expected.read(name);
    assert.ok(Buffer.isBuffer(a), name + ': async-read part data is a Buffer');
    assert.ok(a.equals(e), name + ': part bytes identical');
  }
}

test('readAsync equals sync read: part names, part bytes, text', async () => {
  const bytes = workbookBytes();
  const syncPkg = OoxmlPackage.read(bytes);
  const asyncPkg = await OoxmlPackage.readAsync(bytes);
  assertPackagesEqual(asyncPkg, syncPkg);
  assert.equal(asyncPkg.text('xl/workbook.xml'), syncPkg.text('xl/workbook.xml'));
  assert.equal(asyncPkg.kind(), 'sheet');
});

test('forced stream inflater: identical bytes, pre-inflated, nothing marked modified', async () => {
  const bytes = workbookBytes();
  const syncPkg = OoxmlPackage.read(bytes);
  const asyncPkg = await OoxmlPackage.readAsync(bytes, { inflate: 'stream' });

  // Every deflated entry was inflated UP FRONT — downstream reads never need
  // zlib (or any async step) again. That is the property that lets Workbook,
  // Document and every part() consumer stay synchronous and untouched.
  for (const entry of asyncPkg.entries) {
    if (entry.method === 8) assert.notEqual(entry._raw, null, entry.name + ': pre-inflated');
  }

  // Pre-inflating is a read, not an edit: fidelity accounting is untouched and
  // untouched entries will round-trip from their original compressed bytes.
  assert.deepStrictEqual(asyncPkg.modifiedParts(), []);

  assertPackagesEqual(asyncPkg, syncPkg);
  assert.equal(asyncPkg.text('xl/workbook.xml'), syncPkg.text('xl/workbook.xml'));
});

test('a Workbook over an async-read package answers like the sync one', async () => {
  const bytes = workbookBytes();
  const syncWb = new Workbook(OoxmlPackage.read(bytes));
  const asyncWb = new Workbook(await OoxmlPackage.readAsync(bytes, { inflate: 'stream' }));

  assert.deepStrictEqual(asyncWb.sheetNames(), syncWb.sheetNames());
  for (const [sheet, ref] of [['Data', 'A1'], ['Data', 'B1'], ['Data', 'A2'], ['Data', 'B2'], ['Calc', 'C1']]) {
    assert.deepStrictEqual(asyncWb.getCell(sheet, ref), syncWb.getCell(sheet, ref), sheet + '!' + ref);
  }
  assert.equal(asyncWb.getCell('Data', 'B1'), 'b&<>"é');
  assert.equal(asyncWb.getCell('Calc', 'C1'), '=A1+B1');
});

test('deterministic write(): async-read package writes byte-identical output', async () => {
  const bytes = workbookBytes();
  const syncOut = OoxmlPackage.read(bytes).write();
  const asyncOut = (await OoxmlPackage.readAsync(bytes, { inflate: 'stream' })).write();
  assert.ok(asyncOut.equals(syncOut), 'write() after readAsync equals write() after read');
});

test('a stored (method 0) entry works through the async read', async () => {
  const pkg = OoxmlPackage.fromParts([
    {
      name: '[Content_Types].xml',
      data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        + '<Default Extension="txt" ContentType="text/plain"/></Types>',
    },
    { name: 'stored.txt', data: 'kept verbatim, never deflated' },
  ]);
  pkg.part('stored.txt').method = 0; // serialize() will store it uncompressed
  const bytes = pkg.write();

  const reread = await OoxmlPackage.readAsync(bytes, { inflate: 'stream' });
  const entry = reread.part('stored.txt');
  assert.equal(entry.method, 0);
  assert.equal(entry._raw, null, 'stored entries are not pre-inflated — nothing to inflate');
  assert.equal(reread.text('stored.txt'), 'kept verbatim, never deflated');

  const syncReread = OoxmlPackage.read(bytes);
  assertPackagesEqual(reread, syncReread);
  assert.ok(reread.write().equals(syncReread.write()));
});

test('readAsync accepts Uint8Array and ArrayBuffer input', async () => {
  const bytes = workbookBytes();
  const expected = OoxmlPackage.read(bytes);

  const fromU8 = await OoxmlPackage.readAsync(new Uint8Array(bytes), { inflate: 'stream' });
  assertPackagesEqual(fromU8, expected);

  const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  const fromAb = await OoxmlPackage.readAsync(ab, { inflate: 'stream' });
  assertPackagesEqual(fromAb, expected);
});

test('readAsync rejects garbage with the same errors read throws', async () => {
  await assert.rejects(OoxmlPackage.readAsync(Buffer.alloc(0)), { name: 'OoxmlError', message: /empty file/ });
  await assert.rejects(
    OoxmlPackage.readAsync(Buffer.from('this is not a zip archive, not even slightly')),
    { name: 'OoxmlError', message: /missing zip signature/ },
  );
});

test('inflateRawStream matches inflateRawSync byte for byte', async () => {
  const inputs = [
    Buffer.alloc(0),
    Buffer.from('hello world'),
    Buffer.from('abc123'.repeat(50_000)), // ~300KB inflated: forces multi-chunk stream output
    randomBytes(64 * 1024), // incompressible
  ];
  for (const input of inputs) {
    const deflated = zlib.deflateRawSync(input, { level: 9 });
    const viaStream = await inflateRawStream(deflated);
    const viaZlib = zlib.inflateRawSync(deflated);
    assert.deepStrictEqual(new Uint8Array(viaStream), new Uint8Array(viaZlib), 'input of ' + input.length + ' bytes');
  }
});

test('inflateRawStream rejects corrupt deflate data, as inflateRawSync throws on it', async () => {
  const garbage = Buffer.from([0xff, 0xff, 0xff, 0xff, 0xff]);
  assert.throws(() => zlib.inflateRawSync(garbage));
  await assert.rejects(inflateRawStream(garbage));
});

test('crc32Js agrees with zlib and with published check values', () => {
  // The CRC-32 check value from the spec (ITU-T V.42 / the zlib manual):
  // crc32 of the ASCII bytes "123456789" is 0xCBF43926, and of nothing is 0.
  assert.equal(crc32Js(Buffer.alloc(0)), 0);
  assert.equal(crc32Js(Buffer.from('123456789')), 0xcbf43926);

  const inputs = [
    Buffer.from('a'),
    Buffer.from('The quick brown fox jumps over the lazy dog'),
    Buffer.from([0x00, 0xff, 0x00, 0xff]),
    randomBytes(1),
    randomBytes(255),
    randomBytes(4096),
    randomBytes(70_000),
  ];
  for (const input of inputs) {
    const js = crc32Js(input);
    assert.equal(js, crc32(input), 'exported crc32, input of ' + input.length + ' bytes');
    if (typeof zlib.crc32 === 'function') {
      assert.equal(js, zlib.crc32(input) >>> 0, 'zlib.crc32, input of ' + input.length + ' bytes');
    }
  }
});
