// Other sources (D-2026-10-22-13, step 5): a table read from a CSV or
// spreadsheet file at a path or a web address, from a table or query of
// another Data Files database (a protected one as the person who says who
// they are there), or from a web address answering JSON, kept as a copy in
// the file, read only, read again on open, on its schedule and on Refresh.
// The web addresses are stand-in servers in the test.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import { Base, tableFromRows, tableFromRecords, flatten, valueAt } from '../packages/database/src/index.js';
import { createBaseService } from '../apps/desktop/main/base.js';
import { buildXlsx } from '../packages/ooxml/src/build.js';
import { writeOds } from '../packages/office-formats/src/odf-write.js';

const tmp = (name) => fs.mkdtempSync(path.join(os.tmpdir(), name));
const keptSecrets = () => {
  const kept = new Map();
  return { kept, secrets: { available: () => true, get: (k) => kept.get(k) ?? null, set: (k, v) => kept.set(k, v), delete: (k) => kept.delete(k) } };
};
const tableOf = (s, name) => s.tables.find((t) => t.name === name);

/** A stand-in web server: `handle(req, url)` answers { status, type, body, headers }. */
async function server(handle) {
  const srv = http.createServer((req, res) => {
    const url = new URL(req.url, 'http://127.0.0.1');
    const out = handle(req, url) || { status: 404, body: 'not here' };
    res.writeHead(out.status || 200, { 'content-type': out.type || 'application/json', ...(out.headers || {}) });
    res.end(typeof out.body === 'string' || Buffer.isBuffer(out.body) ? out.body : JSON.stringify(out.body));
  });
  await new Promise((r) => srv.listen(0, '127.0.0.1', r));
  return { base: `http://127.0.0.1:${srv.address().port}`, close: () => new Promise((r) => srv.close(r)) };
}

test('rows and JSON records made into a table: the header named, types worked out, nested fields flattened one level', () => {
  const t = tableFromRows([['ignore me'], ['Name', 'Price', 'Added', 'In stock'], ['Bread', '2.50', '2026-10-01', 'yes'], ['', '', '', ''], ['Milk', '1,200.10', '2026-10-02', 'no']], { header: 2 });
  assert.deepEqual(t.fields.map((f) => [f.name, f.type]), [['Name', 'text'], ['Price', 'currency'], ['Added', 'date'], ['In stock', 'yesno']]);
  assert.deepEqual(t.rows, [{ Name: 'Bread', Price: 2.5, Added: '2026-10-01', 'In stock': 1 }, { Name: 'Milk', Price: 1200.1, Added: '2026-10-02', 'In stock': 0 }]);
  assert.deepEqual(flatten({ id: 1, address: { city: 'Leeds', geo: { lat: 1 } }, tags: ['a'] }), { id: 1, 'address.city': 'Leeds', 'address.geo': '{"lat":1}', tags: '["a"]' });
  assert.deepEqual(valueAt({ data: { items: [1, 2] } }, 'data.items'), [1, 2]);
  const j = tableFromRecords([{ id: 1, name: 'A', address: { city: 'Leeds' }, ok: true }, { id: 2, name: 'B', extra: 3.5 }]);
  assert.deepEqual(j.fields.map((f) => [f.name, f.type]), [['id', 'integer'], ['name', 'text'], ['address.city', 'text'], ['ok', 'yesno'], ['extra', 'number']]);
  assert.deepEqual(j.rows[1], { id: 2, name: 'B', 'address.city': null, ok: null, extra: 3.5 });
});

test('a CSV file at a path: linked as a copy, read only, read again on Refresh and on its schedule, a changed set of columns said', async () => {
  const dir = tmp('src-csv-');
  const file = path.join(dir, 'shop.rbase');
  const csv = path.join(dir, 'prices.csv');
  fs.writeFileSync(csv, 'Item,Price\r\nFlour,1.80\r\nSugar,0.95\r\n');
  const svc = createBaseService({ secrets: keptSecrets().secrets });
  try {
    Base.open(file).close();
    let s = await svc.open({ file });
    const id = s.id;
    const tried = await svc.trySource({ kind: 'file', settings: { where: csv } });
    assert.equal(tried.count, 2);
    assert.deepEqual(tried.fields.map((f) => f.name), ['Item', 'Price']);
    const out = await svc.linkSource({ id, kind: 'file', name: 'Supplier prices', settings: { where: csv } });
    s = out.summary;
    const t = tableOf(s, 'Supplier prices');
    assert.equal(t.source, 'file');
    assert.equal(t.every, 15);
    assert.equal(t.readOnly, true);
    assert.equal(t.count, 2);
    assert.equal(s.links.length, 0);
    assert.equal(s.sources[0].kind, 'file');
    assert.ok(svc.fields({ id, table: 'Supplier prices' }).every((f) => f.readOnly));
    await assert.rejects(svc.insert({ id, table: 'Supplier prices', values: { Item: 'Salt' } }), /read from its source/);
    assert.deepEqual(svc.rows({ id, table: 'Supplier prices' }).rows.map((r) => [r.Item, r.Price]), [['Flour', 1.8], ['Sugar', 0.95]]);

    // The file changed: on screen it is not read before it is due, Refresh reads it.
    fs.writeFileSync(csv, 'Item,Price\r\nFlour,1.90\r\nSugar,0.95\r\nSalt,0.40\r\n');
    s = await svc.refreshLinked({ id, table: 'Supplier prices' });
    assert.equal(tableOf(s, 'Supplier prices').count, 2);
    s = await svc.refreshLinked({ id, table: 'Supplier prices', now: true });
    assert.equal(tableOf(s, 'Supplier prices').count, 3);
    assert.equal(svc.rows({ id, table: 'Supplier prices' }).rows[0].Price, 1.9);

    // A column added and one taken away: the table's design follows, and says so.
    fs.writeFileSync(csv, 'Item,Pack size\r\nFlour,1 kg\r\n');
    s = await svc.refreshLinked({ id, table: 'Supplier prices', now: true });
    assert.deepEqual(tableOf(s, 'Supplier prices').note, { added: ['Pack size'], removed: ['Price'], changed: [] });
    assert.deepEqual(svc.fields({ id, table: 'Supplier prices' }).map((f) => f.name), ['Item', 'Pack size']);
    assert.deepEqual(svc.rows({ id, table: 'Supplier prices' }).rows.map((r) => r['Pack size']), ['1 kg']);

    // On its schedule: nothing due yet; every 5 minutes, and its last read long ago, it is.
    assert.equal(await svc.refreshDue({ id }), null);
    s = svc.setSourceEvery({ id, table: 'Supplier prices', every: 5 });
    assert.equal(tableOf(s, 'Supplier prices').every, 5);
    fs.writeFileSync(csv, 'Item,Pack size\r\nFlour,1 kg\r\nRice,2 kg\r\n');
    const other = Base.open(file);
    other.db.prepare('UPDATE rutba_linked SET synced = ?').run(new Date(Date.now() - 6 * 60_000).toISOString());
    other.close();
    s = await svc.refreshDue({ id });
    assert.equal(tableOf(s, 'Supplier prices').count, 2);
    // By hand only: never due.
    svc.setSourceEvery({ id, table: 'Supplier prices', every: 0 });
    assert.equal(await svc.refreshDue({ id }), null);

    // Gone from where it was: the copy stays, and the bar says why.
    fs.rmSync(csv);
    s = await svc.refreshLinked({ id, table: 'Supplier prices', now: true });
    assert.equal(tableOf(s, 'Supplier prices').count, 2);
    assert.match(tableOf(s, 'Supplier prices').linkError, /not there/);
    // Opened again: read on open, the copy there meanwhile.
    svc.close({ id });
    s = await svc.open({ file });
    assert.equal(tableOf(s, 'Supplier prices').count, 2);
    // Unlinked: its table and its link go.
    s = svc.unlinkTable({ id: s.id, table: 'Supplier prices' });
    assert.equal(s.tables.length, 0);
    assert.equal(s.sources.length, 0);
  } finally {
    svc.closeAll();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('spreadsheets as sources: .xlsx with a date column, .ods, .xls, and a CSV at a web address', async () => {
  const dir = tmp('src-sheets-');
  const file = path.join(dir, 'shop.rbase');
  const xlsx = path.join(dir, 'stock.xlsx');
  fs.writeFileSync(xlsx, buildXlsx({ sheets: [
    { name: 'Notes', rows: [['nothing here']] },
    { name: 'Stock', rows: [['Item', 'Count', 'Checked'], ['Flour', 12, 46296], ['Sugar', 4, 46297]], styles: { 'C2:C3': { numFmt: 'dd/mm/yyyy' } } },
  ] }));
  const ods = path.join(dir, 'stock.ods');
  fs.writeFileSync(ods, writeOds({ sheets: [{ name: 'Counts', rows: [[{ value: 'Item' }, { value: 'Count' }, { value: 'Open' }], [{ value: 'Tea' }, { value: 3 }, { value: true }], [{ value: 'Jam' }, { value: 7.5 }, { value: false }]] }] }));
  const web = await server((req, url) => (url.pathname === '/list.csv' ? { type: 'text/csv', body: 'Code,Name\n1,One\n2,Two\n' } : null));
  const svc = createBaseService({ secrets: keptSecrets().secrets });
  try {
    Base.open(file).close();
    const { id } = await svc.open({ file });
    const first = await svc.trySource({ kind: 'file', settings: { where: xlsx } });
    assert.deepEqual(first.sheets, ['Notes', 'Stock']);
    const tried = await svc.trySource({ kind: 'file', settings: { where: xlsx, sheet: 'Stock' } });
    assert.deepEqual(tried.fields.map((f) => [f.name, f.type]), [['Item', 'text'], ['Count', 'integer'], ['Checked', 'date']]);
    await svc.linkSource({ id, kind: 'file', name: 'Stock', settings: { where: xlsx, sheet: 'Stock' } });
    assert.deepEqual(svc.rows({ id, table: 'Stock' }).rows.map((r) => [r.Item, r.Count, r.Checked]), [['Flour', 12, '2026-10-01'], ['Sugar', 4, '2026-10-02']]);

    const fromOds = await svc.trySource({ kind: 'file', settings: { where: ods } });
    assert.deepEqual(fromOds.fields.map((f) => [f.name, f.type]), [['Item', 'text'], ['Count', 'number'], ['Open', 'yesno']]);
    assert.deepEqual(fromOds.rows.map((r) => [r.Item, r.Count, r.Open]), [['Tea', 3, 1], ['Jam', 7.5, 0]]);

    const xls = await svc.trySource({ kind: 'file', settings: { where: path.resolve('tests/fixtures/binary/showcase.xls') } });
    assert.ok(xls.fields.length > 0 && xls.sheets.length > 0);

    const s = (await svc.linkSource({ id, kind: 'file', name: 'Codes', settings: { where: `${web.base}/list.csv` } })).summary;
    assert.equal(tableOf(s, 'Codes').count, 2);
    await assert.rejects(svc.trySource({ kind: 'file', settings: { where: `${web.base}/gone.csv` } }), /answered 404/);
  } finally {
    svc.closeAll();
    await web.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a sheet of weeks with a totals row: the dates stay dates beside the word Total, and sums read as the sheet shows them', async () => {
  const dir = tmp('src-takings-');
  const file = path.join(dir, 'shop.rbase');
  const xlsx = path.join(dir, 'takings.xlsx');
  fs.writeFileSync(xlsx, buildXlsx({ sheets: [
    { name: 'Takings', rows: [['Week starting', 'Total'], [46027, 1494.34 + 589.55 + 948.51 + 555.99], [46034, 0.1 + 0.2], ['Total', 3588.3899999999994 + 0.30000000000000004]], styles: { 'A2:A3': { numFmt: 'dd/mm/yyyy' } } },
  ] }));
  const svc = createBaseService({ secrets: keptSecrets().secrets });
  try {
    Base.open(file).close();
    await svc.open({ file });
    const tried = await svc.trySource({ kind: 'file', settings: { where: xlsx } });
    assert.deepEqual(tried.fields.map((f) => [f.name, f.type]), [['Week starting', 'text'], ['Total', 'number']]);
    assert.deepEqual(tried.rows.map((r) => [r['Week starting'], r.Total]), [['2026-01-05', 3588.39], ['2026-01-12', 0.3], ['Total', 3588.69]]);
  } finally {
    svc.closeAll();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('another Data Files database as a source: a table and a query read only; a protected one asks who you are and its roles apply', async () => {
  const dir = tmp('src-rbase-');
  const file = path.join(dir, 'shop.rbase');
  const supplier = path.join(dir, 'supplier.rbase');
  const guarded = path.join(dir, 'guarded.rbase');
  const make = (at) => {
    const b = Base.open(at);
    b.createTable({ name: 'Products', fields: [{ name: 'Name', type: 'text' }, { name: 'Price', type: 'currency' }, { name: 'Cost', type: 'currency' }] });
    b.insert('Products', { Name: 'Bread', Price: 2.5, Cost: 1.1 });
    b.insert('Products', { Name: 'Milk', Price: 1.2, Cost: 0.6 });
    b.saveQuery('Cheap', { tables: ['Products'], columns: [{ table: 'Products', field: 'Name' }, { table: 'Products', field: 'Price', criteria: ['<2'] }] });
    return b;
  };
  make(supplier).close();
  const g = make(guarded);
  g.protect({ name: 'Owner', password: 'owner-pass' });
  g.saveRole('Shop', { tables: { default: 'read', each: { Products: { access: 'read', fields: { Cost: 'hidden' } } } } });
  g.addPerson({ name: 'Ann', role: 'Shop', password: 'ann-pass' });
  g.close();
  const { kept, secrets } = keptSecrets();
  const svc = createBaseService({ secrets });
  try {
    Base.open(file).close();
    const { id } = await svc.open({ file });
    const items = svc.sourceItems({ where: supplier });
    assert.deepEqual(items.items, [{ kind: 'table', name: 'Products' }, { kind: 'query', name: 'Cheap' }]);
    await svc.linkSource({ id, kind: 'rbase', name: 'Their products', settings: { where: supplier }, item: { kind: 'table', name: 'Products' } });
    let s = (await svc.linkSource({ id, kind: 'rbase', name: 'Their cheap', settings: { where: supplier }, item: { kind: 'query', name: 'Cheap' } })).summary;
    assert.deepEqual(svc.rows({ id, table: 'Their products' }).rows.map((r) => [r.Name, r.Price, r.Cost]), [['Bread', 2.5, 1.1], ['Milk', 1.2, 0.6]]);
    assert.deepEqual(svc.rows({ id, table: 'Their cheap' }).rows.map((r) => r.Name), ['Milk']);
    assert.equal(tableOf(s, 'Their cheap').readOnly, true);
    // The other file is never changed by being read.
    const before = fs.statSync(supplier).mtimeMs;
    const w = Base.open(supplier);
    w.insert('Products', { Name: 'Eggs', Price: 1.5, Cost: 0.8 });
    w.close();
    s = await svc.refreshLinked({ id, table: 'Their cheap', now: true });
    assert.deepEqual(svc.rows({ id, table: 'Their cheap' }).rows.map((r) => r.Name), ['Milk', 'Eggs']);
    assert.ok(fs.statSync(supplier).mtimeMs >= before);

    // Protected: who are you there?
    await assert.rejects(svc.trySource({ kind: 'rbase', settings: { where: guarded } }), /protected: say who you are/);
    await assert.rejects(svc.trySource({ kind: 'rbase', settings: { where: guarded, person: 'Ann' }, password: 'wrong' }), /do not match/);
    s = (await svc.linkSource({ id, kind: 'rbase', name: 'Guarded products', settings: { where: guarded, person: 'Ann' }, password: 'ann-pass', item: { kind: 'table', name: 'Products' } })).summary;
    // Ann's role hides the cost: it is not read.
    assert.deepEqual(svc.fields({ id, table: 'Guarded products' }).map((f) => f.name), ['ID', 'Name', 'Price']);
    assert.equal(svc.rows({ id, table: 'Guarded products' }).rows.length, 2);
    // Her password is in the protected store, not the file.
    assert.ok([...kept.values()].some((v) => String(v).includes('ann-pass')));
    svc.close({ id });
    assert.ok(!fs.readFileSync(file).includes(Buffer.from('ann-pass')));
  } finally {
    svc.closeAll();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a web address answering JSON: the records found at a path, paging by next link and by page number, a refused key, the key kept out of the file', async () => {
  const dir = tmp('src-json-');
  const file = path.join(dir, 'shop.rbase');
  const people = Array.from({ length: 7 }, (_, i) => ({ id: i + 1, name: `Person ${i + 1}`, address: { city: i % 2 ? 'Leeds' : 'York' }, active: i % 3 !== 0 }));
  const seen = [];
  const web = await server((req, url) => {
    seen.push(url.pathname + url.search);
    if (req.headers['x-api-key'] !== 'secret-key-123') return { status: 401, body: { error: 'no key' } };
    if (url.pathname === '/next') {
      const at = Number(url.searchParams.get('from') || 0);
      const page = people.slice(at, at + 3);
      return { body: { data: { items: page }, links: { next: at + 3 < people.length ? `/next?from=${at + 3}` : null } } };
    }
    if (url.pathname === '/pages') {
      const n = Number(url.searchParams.get('page'));
      return { body: { results: people.slice((n - 1) * 3, n * 3) } };
    }
    return null;
  });
  const { kept, secrets } = keptSecrets();
  const svc = createBaseService({ secrets });
  try {
    Base.open(file).close();
    const { id } = await svc.open({ file });
    const settings = { where: `${web.base}/next`, headers: [{ name: 'X-Api-Key', secret: true }], records: 'data.items', paging: { kind: 'next', next: 'links.next' } };
    await assert.rejects(svc.trySource({ kind: 'json', settings, headers: { 'X-Api-Key': 'wrong' } }), /refused the key or token \(401\)/);
    const tried = await svc.trySource({ kind: 'json', settings, headers: { 'X-Api-Key': 'secret-key-123' } });
    assert.equal(tried.count, 7);
    assert.equal(tried.pages, 3);
    assert.deepEqual(tried.fields.map((f) => [f.name, f.type]), [['id', 'integer'], ['name', 'text'], ['address.city', 'text'], ['active', 'yesno']]);
    let s = (await svc.linkSource({ id, kind: 'json', name: 'People', settings, headers: { 'X-Api-Key': 'secret-key-123' } })).summary;
    assert.equal(tableOf(s, 'People').count, 7);
    assert.deepEqual(svc.rows({ id, table: 'People' }).rows.slice(0, 2).map((r) => [r.id, r.name, r['address.city'], r.active]), [[1, 'Person 1', 'York', 0], [2, 'Person 2', 'Leeds', 1]]);
    assert.ok([...kept.keys()].some((k) => k.endsWith(':X-Api-Key')));

    // By page number, from 1, until a page comes back short.
    s = (await svc.linkSource({ id, kind: 'json', name: 'People by page', settings: { where: `${web.base}/pages`, headers: [{ name: 'X-Api-Key', secret: true }], records: 'results', paging: { kind: 'page', param: 'page', start: 1 } }, headers: { 'X-Api-Key': 'secret-key-123' } })).summary;
    assert.equal(tableOf(s, 'People by page').count, 7);
    assert.ok(seen.includes('/pages?page=3') && !seen.includes('/pages?page=4'));

    // The key taken away from the store: refused, the copy kept and the bar saying why.
    for (const k of [...kept.keys()]) kept.delete(k);
    s = await svc.refreshLinked({ id, table: 'People', now: true });
    assert.equal(tableOf(s, 'People').count, 7);
    assert.match(tableOf(s, 'People').linkError, /refused the key/);
    svc.close({ id });
    assert.ok(!fs.readFileSync(file).includes(Buffer.from('secret-key-123')));
    for (const s2 of ['-wal']) if (fs.existsSync(file + s2)) assert.ok(!fs.readFileSync(file + s2).includes(Buffer.from('secret-key-123')));
  } finally {
    svc.closeAll();
    await web.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
