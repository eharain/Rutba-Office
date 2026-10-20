// dBASE and FoxPro tables (the 2.8 plan's old files): the fields as a
// header row, a row a record, dates as dates, numbers with their decimals,
// logicals, memos from a .dbt or .fpt beside the table, deleted records left
// out, and the code page the table names.
import test from 'node:test';
import assert from 'node:assert/strict';
import { readDbf, isDbf } from '../packages/office-formats/src/dbf.js';

/** A .dbf of `fields` [name, type, length, decimals] and `records` (arrays of raw strings, '*' first to delete). */
function makeDbf({ version = 0x03, fields, records, codePage = 0x03 }) {
  const headerSize = 32 + fields.length * 32 + 1;
  const recordSize = 1 + fields.reduce((n, f) => n + f[2], 0);
  const out = Buffer.alloc(headerSize + recordSize * records.length + 1, 0);
  out[0] = version;
  out[1] = 124; out[2] = 10; out[3] = 21;
  out.writeUInt32LE(records.length, 4);
  out.writeUInt16LE(headerSize, 8);
  out.writeUInt16LE(recordSize, 10);
  out[29] = codePage;
  fields.forEach(([name, type, length, decimals = 0], i) => {
    const at = 32 + i * 32;
    out.write(name, at, 'latin1');
    out.write(type, at + 11, 'latin1');
    out[at + 16] = length;
    out[at + 17] = decimals;
  });
  out[headerSize - 1] = 0x0d;
  records.forEach((rec, r) => {
    const at = headerSize + r * recordSize;
    const deleted = rec[0] === '*';
    out[at] = deleted ? 0x2a : 0x20;
    let off = 1;
    (deleted ? rec.slice(1) : rec).forEach((value, c) => {
      const len = fields[c][2];
      if (Buffer.isBuffer(value)) value.copy(out, at + off);
      else out.write(String(value).padEnd(len, ' ').slice(0, len), at + off, 'latin1');
      off += len;
    });
  });
  out[out.length - 1] = 0x1a;
  return new Uint8Array(out);
}

test('a dBASE III table reads as a header row and a row a record, with its numbers, dates and logicals', () => {
  const dbf = makeDbf({
    fields: [['NAME', 'C', 12], ['PRICE', 'N', 8, 2], ['SOLD', 'D', 8], ['STOCK', 'L', 1]],
    records: [['Tea', '    3.50', '20240115', 'T'], ['*', 'Gone', '    1.00', '20240101', 'F'], ['Caf\xe9', '   12.00', '        ', 'N']],
  });
  assert.ok(isDbf(dbf));
  const { sheets, notes } = readDbf(dbf, { name: 'PRICES.DBF' });
  const [s] = sheets;
  assert.equal(s.name, 'PRICES');
  assert.deepEqual(s.rows[0], ['NAME', 'PRICE', 'SOLD', 'STOCK']);
  assert.deepEqual(s.rows[1], ['Tea', 3.5, 45306, true]);
  assert.deepEqual(s.rows[2], ['Café', 12, null, false]);
  assert.equal(s.rows.length, 3);
  assert.equal(s.styles.B2.numFmt, '0.00');
  assert.equal(s.styles.C2.numFmt, 'yyyy-mm-dd');
  assert.equal(s.styles.A1.bold, true);
  assert.deepEqual(notes, ['1 deleted record was left out.']);
});

test('a memo field reads from its .dbt, and says so when the memo file is missing', () => {
  const dbf = makeDbf({ version: 0x83, fields: [['ID', 'N', 4], ['NOTES', 'M', 10]], records: [['   1', '         1'], ['   2', '         2']] });
  const dbt = Buffer.alloc(512 * 3, 0);
  dbt.write('First note\x1a\x1a', 512, 'latin1');
  dbt.write('Second, longer note\x1a\x1a', 1024, 'latin1');
  const { sheets } = readDbf(dbf, { memo: new Uint8Array(dbt), memoKind: 'dbt' });
  assert.deepEqual(sheets[0].rows.slice(1), [[1, 'First note'], [2, 'Second, longer note']]);
  const alone = readDbf(dbf);
  assert.match(alone.notes.join(' '), /memo file was not beside the table/);
});

test('a Visual FoxPro table reads its integers, doubles, currency and date-times, and its .fpt memos', () => {
  const int = Buffer.alloc(4); int.writeInt32LE(-42);
  const dbl = Buffer.alloc(8); dbl.writeDoubleLE(2.25);
  const cur = Buffer.alloc(8); cur.writeBigInt64LE(123450000n);
  const dt = Buffer.alloc(8); dt.writeInt32LE(2451545); dt.writeInt32LE(12 * 3600000, 4);
  const memo = Buffer.alloc(4); memo.writeInt32LE(8);
  const dbf = makeDbf({ version: 0x30, fields: [['N', 'I', 4], ['D', 'B', 8], ['Y', 'Y', 8], ['T', 'T', 8], ['M', 'M', 4]], records: [[int, dbl, cur, dt, memo]] });
  // The header takes the first 512 bytes; with 64-byte blocks the first memo is block 8.
  const fpt = Buffer.alloc(1024, 0);
  fpt.writeUInt16BE(64, 6);
  fpt.writeUInt32BE(1, 512);
  fpt.writeUInt32BE(5, 516);
  fpt.write('Hello', 520, 'latin1');
  const { sheets } = readDbf(dbf, { memo: new Uint8Array(fpt), memoKind: 'fpt' });
  const [n, d, y, t, m] = sheets[0].rows[1];
  assert.equal(n, -42);
  assert.equal(d, 2.25);
  assert.equal(y, 12345);
  // 2000-01-01 at noon.
  assert.equal(t, 36526.5);
  assert.equal(m, 'Hello');
});

test('the code page the table names decodes its words', () => {
  const dbf = makeDbf({ codePage: 0x65, fields: [['W', 'C', 6]], records: [[Buffer.from([0x8f, 0xe0, 0xa8, 0xa2, 0xa5, 0xe2])]] });
  assert.equal(readDbf(dbf).sheets[0].rows[1][0], 'Привет');
  assert.equal(isDbf(new Uint8Array(Buffer.from('not a table at all, just words in a file'))), false);
});
