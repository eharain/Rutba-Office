// Access databases (the 2.8 plan's old files), read only: a sheet per table,
// the field names first, dates as dates, currency as numbers, binary fields
// left out and said, and a password asked for. The reader is mdb-reader; a
// stand-in for it here gives the tables as it gives them.
import test from 'node:test';
import assert from 'node:assert/strict';
import { readAccess, isAccess } from '../packages/office-formats/src/access.js';

class FakeReader {
  constructor(buffer, options) {
    if (buffer[0] === 0x50 && options?.password !== 'open sesame') throw new Error('Wrong password');
  }
  getTableNames() { return ['Customers', 'Orders']; }
  getTable(name) {
    const tables = {
      Customers: {
        rowCount: 2,
        columns: [{ name: 'ID', type: 'long' }, { name: 'Name', type: 'text', size: 100 }, { name: 'Photo', type: 'ole' }, { name: 'Joined', type: 'datetime' }],
        data: [{ ID: 1, Name: 'Amira', Photo: Buffer.from([1]), Joined: new Date(Date.UTC(2024, 0, 15)) }, { ID: 2, Name: 'Ben', Photo: null, Joined: null }],
      },
      Orders: {
        rowCount: 1,
        columns: [{ name: 'OrderID', type: 'long' }, { name: 'Total', type: 'currency' }, { name: 'Shipped', type: 'boolean' }, { name: 'At', type: 'datetime' }],
        data: [{ OrderID: 7, Total: '12.5000', Shipped: true, At: new Date(Date.UTC(2024, 0, 15, 12, 0, 0)) }],
      },
    };
    const t = tables[name];
    return { rowCount: t.rowCount, getColumns: () => t.columns, getData: ({ columns }) => t.data.map((r) => Object.fromEntries(columns.map((c) => [c, r[c]]))) };
  }
}

test('a sheet a table, the fields first, dates and currency as numbers, binary fields left out and said', () => {
  const { sheets, notes } = readAccess(new Uint8Array(64), { Reader: FakeReader });
  assert.deepEqual(sheets.map((s) => s.name), ['Customers', 'Orders']);
  assert.deepEqual(sheets[0].rows[0], ['ID', 'Name', 'Joined']);
  assert.deepEqual(sheets[0].rows[1], [1, 'Amira', 45306]);
  assert.deepEqual(sheets[0].rows[2], [2, 'Ben']);
  assert.equal(sheets[0].styles.C2.numFmt, 'yyyy-mm-dd');
  assert.deepEqual(sheets[1].rows[1], [7, 12.5, true, 45306.5]);
  assert.equal(sheets[1].styles.B2.numFmt, '#,##0.00');
  assert.equal(sheets[1].styles.D2.numFmt, 'yyyy-mm-dd hh:mm:ss');
  assert.deepEqual(notes, [{ code: 'binary-left-out', count: 1 }, { code: 'access-read-only' }]);
});

test('a database with a password says so, and opens with the right one', () => {
  const locked = new Uint8Array(64);
  locked[0] = 0x50;
  assert.throws(() => readAccess(locked, { Reader: FakeReader }), (err) => err.password === true);
  assert.equal(readAccess(locked, { Reader: FakeReader, password: 'open sesame' }).sheets.length, 2);
});

test('an Access file is known by its first page', () => {
  const b = new Uint8Array(64);
  b.set(Buffer.from('Standard Jet DB', 'latin1'), 4);
  assert.ok(isAccess(b));
  b.set(Buffer.from('Standard ACE DB', 'latin1'), 4);
  assert.ok(isAccess(b));
  assert.equal(isAccess(new Uint8Array(64)), false);
});
