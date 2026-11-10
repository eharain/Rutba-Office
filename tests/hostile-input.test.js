// Files and messages made to do harm. Each of these was found by a peer review
// on 10 Nov 2026: a path that handed a Windows sign-in to another computer, one
// email or invitation that stopped the app, a little file that asked for more
// memory than the computer has. A test here is about time and memory as much
// as about the answer.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { isNetworkPath, mayServe, forgetPlaces, grantPlace } from '../packages/office-shell/src/electron/grants.js';
import { dispositionFor } from '../packages/office-shell/src/electron/blobs.js';
import { parseParameters } from '../packages/mailbox/src/mime.js';
import { occurrences } from '../packages/calendar/src/recur.js';
import { readBth } from '../packages/mailbox/src/pst/ltp.js';
import { inflateRaw, inflateLimit, readZip, writeZip, ZipEntry } from '../packages/ooxml/src/zip.js';
import { inflate } from '@rutba/pdf/reader';
import { CompoundFile } from '../packages/office-formats/src/cfb.js';
import { readOdf } from '../packages/office-formats/src/odf.js';
import { bodyDocument } from '../apps/desktop/renderer/apps/mail/parts.js';

/** Done within this long, or it is the fault this test is about. */
function quickly(ms, fn) {
  const at = Date.now();
  const out = fn();
  assert.ok(Date.now() - at < ms, `took ${Date.now() - at} ms`);
  return out;
}

test('a network path is known in every spelling Windows accepts, and none is served unasked', () => {
  forgetPlaces();
  for (const p of ['\\\\evil\\share\\a.png', '//evil/share/a.png', '/\\evil\\share\\a.png', '\\/evil/share/a.png', '\\\\?\\UNC\\evil\\share\\a.png', '\\\\.\\pipe\\x', '\\??\\UNC\\evil\\share\\a.png']) {
    assert.equal(isNetworkPath(p, 'win32'), true, p);
    assert.equal(mayServe(p, 'win32'), false, p);
  }
  for (const p of ['C:\\Users\\me\\a.png', 'C:/Users/me/a.png', 'a\\b.png', '/a/b.png']) assert.equal(isNetworkPath(p, 'win32'), false, p);
  grantPlace('/\\server\\share', 'win32');
  assert.equal(mayServe('\\\\server\\share\\in\\a.png', 'win32'), true, 'a place the person reached stays reachable by its other spelling');
});

test('the mail frame draws only the pictures the app holds for that message', () => {
  const page = bodyDocument({ html: '<p>hello</p>' }, {});
  const remote = bodyDocument({ html: '<p>hello</p>' }, { remote: true });
  for (const doc of [page, remote]) {
    const csp = /Content-Security-Policy" content="([^"]+)"/.exec(doc)[1];
    const img = csp.split(';').find((part) => part.trim().startsWith('img-src')).trim().split(/\s+/);
    assert.ok(img.includes('rutba://blob'), csp);
    assert.equal(img.includes('rutba:'), false, 'not every rutba: address');
  }
});

test('an attachment with a name in any script can be served', async () => {
  for (const name of ['رپورٹ.pdf', 'تقرير.pdf', 'रिपोर्ट.pdf', 'plain.pdf', 'with "quotes".pdf', 'a\r\nSet-Cookie: x=1.pdf', 'lone \ud800 half.pdf']) {
    const header = dispositionFor(name);
    assert.doesNotThrow(() => new Response('x', { headers: { 'content-disposition': header } }), name);
    assert.equal(/[\r\n]/.test(header), false, 'no header can be started from a name');
    const encoded = /filename\*=UTF-8''(.+)$/.exec(header)[1];
    if (!/[\r\n"\ud800]/.test(name)) assert.equal(decodeURIComponent(encoded), name);
  }
});

test('MIME parameters: an index of four billion costs nothing, and a key may be any word', () => {
  const got = quickly(500, () => parseParameters('attachment; name*4294967294=x; name*0=a'));
  assert.equal(got.params?.name ?? got.name, 'ax');
  assert.doesNotThrow(() => parseParameters('attachment; __proto__*0=x; constructor*1=y'));
  const normal = parseParameters("attachment; filename*0*=utf-8''%E2%82%AC; filename*1*=uro");
  assert.equal(normal.params?.filename ?? normal.filename, '€uro');
  const shuffled = parseParameters('attachment; name*2=c; name*0=a; name*1=b');
  assert.equal(shuffled.params?.name ?? shuffled.name, 'abc');
});

const at = (iso) => ({ at: Date.parse(iso), date: iso.slice(0, 10), time: iso.slice(11, 19), utc: true });
const series = (rrule, extra = {}) => ({ uid: 'u', start: at('2020-01-06T09:00:00Z'), end: at('2020-01-06T10:00:00Z'), rrule, ...extra });

test('recurrence: an interval too large for a date stops at once', () => {
  for (const freq of ['DAILY', 'WEEKLY', 'MONTHLY', 'YEARLY']) {
    quickly(1500, () => occurrences(series({ freq, interval: 2e15 }), { from: Date.parse('2020-01-01'), to: Date.parse('2030-01-01') }));
    quickly(1500, () => occurrences(series({ freq, interval: 2e15 }), {}));
  }
});

test('recurrence: a rule that never lands on a day ends', () => {
  quickly(3000, () => occurrences(series({ freq: 'DAILY', byMonth: [13] }), {}));
});

test('recurrence: a series that began long ago still shows this week', () => {
  const list = occurrences({ uid: 'u', start: at('2012-01-02T09:00:00Z'), end: at('2012-01-02T10:00:00Z'), rrule: { freq: 'DAILY' } }, { from: Date.parse('2026-11-01T00:00:00Z'), to: Date.parse('2026-11-08T00:00:00Z'), limit: 500 });
  assert.equal(list.length, 7);
});

test('recurrence: one occurrence moved past the end of the range does not end the series', () => {
  const moved = { uid: 'u', recurrenceId: at('2020-01-07T09:00:00Z'), start: at('2020-06-01T09:00:00Z'), end: at('2020-06-01T10:00:00Z') };
  const list = occurrences(series({ freq: 'DAILY' }), { from: Date.parse('2020-01-06T00:00:00Z'), to: Date.parse('2020-01-10T00:00:00Z'), siblings: [moved] });
  assert.deepEqual(list.map((o) => new Date(o.start).toISOString().slice(0, 10)), ['2020-01-06', '2020-01-08', '2020-01-09']);
});

test('recurrence: a hundred thousand excluded dates do not slow a long series', () => {
  const exdates = Array.from({ length: 100000 }, (_, i) => ({ at: Date.parse('2500-01-01') + i * 1000 }));
  quickly(4000, () => occurrences(series({ freq: 'DAILY' }, { exdates }), { from: Date.parse('2400-01-01'), to: Date.parse('2410-01-01') }));
});

test('a .pst index that names itself ends', () => {
  const bytes = Buffer.alloc(8 + 6 * 4000);
  // Header: signature, key and entry size, one more level than there is, root.
  const header = new Uint8Array(8);
  const dv = new DataView(header.buffer);
  header[0] = 0xb5;
  header[1] = 2;
  header[2] = 6;
  header[3] = 200;
  dv.setUint32(4, 0x20, true);
  for (let i = 0; i < 4000; i += 1) bytes.writeUInt32LE(0x20, 2 + i * 6);
  const heap = { get: (hid) => (hid === 0x99 ? header : new Uint8Array(bytes.subarray(0, 6 * 4000))) };
  quickly(1500, () => readBth(heap, 0x99));
});

test('a zip entry that unpacks past the limit is refused in a sentence', () => {
  const bomb = zlib.deflateRawSync(Buffer.alloc(8 * 1024 * 1024));
  assert.throws(() => inflateRaw(bomb, 'word/document.xml', 1024 * 1024), (e) => /word\/document\.xml unpacks to more/.test(e.message) && e.name === 'Error');
  assert.equal(inflateRaw(bomb, 'x', 16 * 1024 * 1024).length, 8 * 1024 * 1024, 'within it, read as ever');
  assert.ok(inflateLimit(10) >= 1024 * 1024 * 1024);
  assert.equal(inflateLimit(1024 * 1024 * 1024), 4 * 1024 * 1024 * 1024, 'a large video stored in a file is not refused');
});

test('a compressed PDF stream that unpacks past the limit is refused, by name and not as damage', () => {
  const bomb = zlib.deflateSync(Buffer.alloc(8 * 1024 * 1024));
  assert.throws(() => inflate(bomb, 1024 * 1024), (e) => /more than this program can hold/.test(e.message) && e.name !== 'RangeError');
  assert.equal(inflate(bomb, 16 * 1024 * 1024).length, 8 * 1024 * 1024);
});

test('nothing unpacks a file without a limit', () => {
  for (const file of ['packages/ooxml/src/zip.js', 'packages/mailbox/src/zip-file.js', 'packages/pdf/src/reader.js']) {
    const source = fs.readFileSync(path.join(import.meta.dirname, '..', file), 'utf8');
    for (const m of source.matchAll(/zlib\.inflate(?:Raw)?Sync\(([^\n]*)/g)) assert.match(m[1], /maxOutputLength/, `${file}: ${m[0]}`);
  }
});

test('a compound file whose table chain names itself is read in bounded memory', () => {
  const size = 4 * 1024 * 1024;
  const b = Buffer.alloc(size);
  Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(b);
  b.writeUInt16LE(0x3e, 24);
  b.writeUInt16LE(3, 26);
  b.writeUInt16LE(0xfffe, 28);
  b.writeUInt16LE(9, 30);
  b.writeUInt16LE(6, 32);
  b.writeUInt32LE(1, 44);
  b.writeUInt32LE(0xfffffffe, 48);
  b.writeUInt32LE(0x1000, 56);
  b.writeUInt32LE(0xfffffffe, 60);
  b.writeUInt32LE(0, 68); // the chain starts at sector 0
  b.writeUInt32LE(1, 72);
  for (let i = 0; i < 109; i += 1) b.writeUInt32LE(0, 76 + i * 4);
  // Sector 0 lists sectors and names itself as the next.
  for (let i = 0; i < 127; i += 1) b.writeUInt32LE(i, 512 + i * 4);
  b.writeUInt32LE(0, 512 + 127 * 4);
  const before = process.memoryUsage().arrayBuffers;
  quickly(5000, () => { try { new CompoundFile(b); } catch (err) { assert.ok(err instanceof Error); } });
  assert.ok(process.memoryUsage().arrayBuffers - before < 64 * 1024 * 1024, 'a table no larger than the file');
});

function ods(content) {
  const entry = (name, text) => {
    const e = new ZipEntry({ name, method: 0, crc: 0, compressedSize: 0, uncompressedSize: 0, compressed: Buffer.alloc(0), dosTime: 0, dosDate: 0x2821 });
    e.data = Buffer.from(text, 'utf8');
    return e;
  };
  return writeZip([
    entry('mimetype', 'application/vnd.oasis.opendocument.spreadsheet'),
    entry('content.xml', content),
  ]);
}

test('a spreadsheet of repeated cells is refused in a sentence before it fills memory', () => {
  const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"';
  const wide = '<table:table-cell office:value-type="float" office:value="1" table:number-columns-repeated="65536"/>'.repeat(50);
  const row = `<table:table-row table:number-rows-repeated="65536">${wide}</table:table-row>`;
  const bytes = ods(`<office:document-content ${NS}><office:body><office:spreadsheet><table:table table:name="S">${row.repeat(20)}</table:table></office:spreadsheet></office:body></office:document-content>`);
  assert.throws(() => quickly(30000, () => readOdf(bytes)), /more cells in it than this program can open/);
});

test('an ordinary spreadsheet with the usual padding still opens', () => {
  const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"';
  const row = '<table:table-row><table:table-cell office:value-type="float" office:value="7"><text:p>7</text:p></table:table-cell><table:table-cell table:number-columns-repeated="1023"/></table:table-row>';
  const pad = '<table:table-row table:number-rows-repeated="1048575"><table:table-cell table:number-columns-repeated="1024"/></table:table-row>';
  const out = quickly(5000, () => readOdf(ods(`<office:document-content ${NS}><office:body><office:spreadsheet><table:table table:name="S">${row}${pad}</table:table></office:spreadsheet></office:body></office:document-content>`)));
  assert.ok(JSON.stringify(out).includes('"7"'));
  assert.ok(readZip);
});

test('OFFSET asked for a whole sheet is cut to the cells in use', async () => {
  const { Spreadsheet } = await import('@rutba/formula');
  const s = new Spreadsheet();
  s.addSheet('S');
  s.setCell('S', 0, 0, 5);
  s.setCell('S', 1, 1, 7);
  s.setCell('S', 3, 3, '=SUM(OFFSET(A1,0,0,1048576,16384))');
  quickly(3000, () => s.recalculate?.());
  assert.equal(quickly(3000, () => s.getValue('S', 3, 3)), 12);
  s.setCell('S', 4, 3, '=ROWS(OFFSET(A1,0,0,5,2))');
  assert.equal(s.getValue('S', 4, 3), 5, 'a smaller one keeps its shape');
});
