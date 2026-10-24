// Data Files' rough edges the how-to scripts found (2.18, Part A2 of the
// plan): UK dates and spreadsheets imported, a file's own IDs and its money
// kept, a type change and a missing related record refused in plain words,
// a criterion naming another field, an average's heading and rounding, a
// query kept from the SQL view, a record named by its number and date, and
// a report of a query with its totals.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Base, importRows, importSheets, inferType, designToSql, criterionSql, expressionSql } from '../packages/database/src/index.js';
import { buildXlsx } from '../packages/ooxml/src/build.js';
import { writeOds } from '../packages/office-formats/src/odf-write.js';
import { readWorkbook } from '../apps/desktop/main/base-workbook.js';

const fresh = () => Base.open(':memory:');

test('UK dates import as dates: day first, unless a day above 12 says otherwise, and the import says how it read them', () => {
  assert.deepEqual(inferType(['12/04/2025', '01/05/2025']), { type: 'date', order: 'dmy' });
  assert.deepEqual(inferType(['04/12/2025', '05/13/2025']), { type: 'date', order: 'mdy' });
  assert.deepEqual(inferType(['13/04/2025', '04/13/2025']), { type: 'text', size: 255 });
  assert.deepEqual(inferType(['12/04/2025 14:30', '01/05/2025 09:00']), { type: 'datetime', order: 'dmy' });
  // No such day: words.
  assert.equal(inferType(['31/02/2025']).type, 'text');
  const b = fresh();
  const made = importRows(b, 'Members', [['Name', 'Joined', 'Seen'], ['Amira', '12/04/2025', '21/10/2026 09:15'], ['Ben', '3/5/2025', '']]);
  assert.deepEqual(b.fields('Members').map((f) => [f.name, f.type]), [['ID', 'autonumber'], ['Name', 'text'], ['Joined', 'date'], ['Seen', 'datetime']]);
  assert.deepEqual(made.dates, [{ field: 'Joined', order: 'dmy' }, { field: 'Seen', order: 'dmy' }]);
  assert.deepEqual(b.rows('Members').map((r) => [r.Joined, r.Seen]), [['2025-04-12', '2026-10-21T09:15:00'], ['2025-05-03', null]]);
  // An ISO column says nothing about its order.
  assert.deepEqual(importRows(b, 'Iso', [['On'], ['2025-04-12']]).dates, []);
});

test('a column of amounts written with two places, or with a money sign, imports as Currency', () => {
  assert.deepEqual(inferType(['1.10', '3.00', '12.50']), { type: 'currency' });
  assert.deepEqual(inferType(['£1.10', '£3']), { type: 'currency' });
  assert.deepEqual(inferType(['1.5', '2.25']), { type: 'number' });
  assert.deepEqual(inferType(['3', '4']), { type: 'integer' });
  const b = fresh();
  importRows(b, 'Sale Lines', [['Item', 'Unit Price', 'Quantity'], ['Milk', '1.10', '2'], ['Bread', '£1,250.00', '1']]);
  const f = b.fields('Sale Lines');
  assert.equal(f.find((x) => x.name === 'Unit Price').type, 'currency');
  assert.equal(f.find((x) => x.name === 'Quantity').type, 'integer');
  assert.deepEqual(b.rows('Sale Lines').map((r) => r['Unit Price']), [1.1, 1250]);
});

test("a CSV's own ID numbers are the table's key, so the file's records relate by them", () => {
  const b = fresh();
  importRows(b, 'Products', [['ID', 'Name'], ['3', 'Milk'], ['5', 'Bread']]);
  importRows(b, 'Sale Lines', [['ID', 'Product', 'Quantity'], ['1', '5', '2'], ['2', '3', '1']]);
  assert.deepEqual(b.fields('Products').map((x) => [x.name, x.type]), [['ID', 'autonumber'], ['Name', 'text']]);
  b.addRelationship({ from: { table: 'Sale Lines', field: 'Product' }, to: { table: 'Products', field: 'ID' } });
  assert.deepEqual(b.lookupPage({ table: 'Products', field: 'ID', values: [5] }).items, [{ value: 5, text: 'Bread' }]);
  // The next one is numbered after the file's highest.
  assert.equal(b.insert('Products', { Name: 'Tea' }), 6);
  // Two the same: not a key; the table numbers its own.
  assert.equal(importRows(b, 'Twice', [['ID', 'A'], [1, 'x'], [1, 'y']]).key, null);
});

test('workbooks import as Worksheets reads them: each sheet, or the one chosen; dates and money by their formats', () => {
  const xlsx = buildXlsx({ sheets: [
    { name: 'Sales', styles: { B2: { numFmt: 'dd/mm/yyyy' }, B3: { numFmt: 'dd/mm/yyyy' }, C2: { numFmt: '"£"#,##0.00' }, C3: { numFmt: '"£"#,##0.00' } }, rows: [['ID', 'Sold On', 'Total', 'Paid'], [1, 45759, 1.1, '=TRUE()'], [2, 45760, 3, '=FALSE()']] },
    { name: 'Empty', rows: [] },
    { name: 'Notes', rows: [['Note'], ['Open late on Friday']] },
  ] });
  const sheets = readWorkbook(xlsx, '.xlsx');
  assert.deepEqual(sheets.map((s) => s.name), ['Sales', 'Notes']);
  // The window reads the file's bytes into a Uint8Array of their own: read the same.
  assert.deepEqual(readWorkbook(new Uint8Array(xlsx), 'xlsx').map((s) => s.name), ['Sales', 'Notes']);
  const b = fresh();
  const [sales, notes] = importSheets(b, { sheets });
  assert.deepEqual(sales.fields.map((f) => [f.name, f.type]), [['ID', 'autonumber'], ['Sold On', 'date'], ['Total', 'currency'], ['Paid', 'yesno']]);
  assert.deepEqual(b.rows('Sales').map((r) => [r.ID, r['Sold On'], r.Total, r.Paid]), [[1, '2025-04-12', 1.1, 1], [2, '2025-04-13', 3, 0]]);
  assert.equal(notes.table, 'Notes');
  assert.deepEqual(readWorkbook(xlsx, '.xlsx', { only: 'Notes' }).map((s) => s.name), ['Notes']);
  assert.throws(() => readWorkbook(xlsx, '.xlsx', { only: 'Nope' }), /no sheet called Nope/);

  const ods = writeOds({ sheets: [{ name: 'Stock', rows: [
    [{ value: 'Name' }, { value: 'Price' }, { value: 'Arrived' }],
    [{ value: 'Milk' }, { value: 1.1 }, { value: 45759, format: 'yyyy-mm-dd' }],
    [{ value: 'Bread' }, { value: 2 }, { value: 45760, format: 'yyyy-mm-dd' }],
  ] }] });
  const [stock] = readWorkbook(ods, 'ods');
  importSheets(b, { sheets: [stock] });
  assert.deepEqual(b.fields('Stock').map((f) => [f.name, f.type]), [['ID', 'autonumber'], ['Name', 'text'], ['Price', 'number'], ['Arrived', 'date']]);
  assert.deepEqual(b.rows('Stock').map((r) => [r.Name, r.Price, r.Arrived]), [['Milk', 1.1, '2025-04-12'], ['Bread', 2, '2025-04-13']]);
});

test('a type change that a value does not fit is refused saying which records, and can clear them', () => {
  const b = fresh();
  b.createTable({ name: 'Members', fields: [{ name: 'Name', type: 'text' }, { name: 'Joined', type: 'text' }, { name: 'Code', type: 'text' }] });
  for (const [n, j, c] of [['Amira', '2025-04-12', 'A1'], ['Ben', 'soon', 'A1'], ['Cal', '03/05/2025', null], ['Dee', 'next week', 'B2']]) b.insert('Members', { Name: n, Joined: j, Code: c });
  assert.throws(() => b.changeField('Members', 'Joined', { type: 'date' }), (e) => {
    assert.equal(e.message, '2 records of Members have a Joined that is not a date: ID 2 ("soon") and ID 4 ("next week"). Clear them and change the type, or put them right first.');
    return true;
  });
  // Nothing changed.
  assert.equal(b.fields('Members').find((f) => f.name === 'Joined').type, 'text');
  b.changeField('Members', 'Joined', { type: 'date' }, { clear: true });
  // A UK date is read as one; the ones that did not fit are empty.
  assert.deepEqual(b.rows('Members').map((r) => r.Joined), ['2025-04-12', null, '2025-05-03', null]);
  assert.throws(() => b.changeField('Members', 'Code', { required: true }), /1 record of Members has no Code: ID 3\. Fill it in first/);
  assert.throws(() => b.changeField('Members', 'Code', { unique: true }), /ID 1 and ID 2 of Members have the same Code \("A1"\)/);
  // Text to a number: its words read as a number, never quietly made nought.
  b.addField('Members', { name: 'Paid', type: 'text' });
  b.update('Members', 1, { Paid: '£1,250.50' });
  b.update('Members', 2, { Paid: 'lots' });
  assert.throws(() => b.changeField('Members', 'Paid', { type: 'currency' }), /ID 2 \("lots"\)/);
  b.update('Members', 2, { Paid: '3' });
  b.changeField('Members', 'Paid', { type: 'currency' });
  assert.deepEqual(b.rows('Members').map((r) => r.Paid), [1250.5, 3, null, null]);
});

test('a record pointing at one that is not there, or pointed at, is refused naming them', () => {
  const b = fresh();
  b.createTable({ name: 'Products', fields: [{ name: 'Name', type: 'text' }] });
  b.createTable({ name: 'Sale Lines', fields: [{ name: 'Product', type: 'integer' }, { name: 'Quantity', type: 'integer' }] });
  b.addRelationship({ from: { table: 'Sale Lines', field: 'Product' }, to: { table: 'Products', field: 'ID' } });
  const milk = b.insert('Products', { Name: 'Milk' });
  const line = b.insert('Sale Lines', { Product: milk, Quantity: 2 });
  b.insert('Sale Lines', { Product: milk, Quantity: 1 });
  assert.throws(() => b.insert('Sale Lines', { Product: 99, Quantity: 1 }), /^Error: Products has no record 99, so Product cannot point at it\.$/);
  assert.throws(() => b.update('Sale Lines', line, { Product: '42' }), /Products has no record 42/);
  assert.throws(() => b.remove('Products', [milk]), (e) => {
    assert.match(e.message, /^2 records of Sale Lines point at this one \(ID 1 and ID 2\)\. Delete or change those first/);
    return true;
  });
});

test('a criterion naming another field compares with that field', () => {
  const b = fresh();
  b.createTable({ name: 'Products', fields: [{ name: 'Name', type: 'text' }, { name: 'In Stock', type: 'integer' }, { name: 'Reorder Level', type: 'integer' }] });
  for (const [n, s, r] of [['Milk', 3, 5], ['Bread', 10, 4], ['Tea', 4, 4]]) b.insert('Products', { Name: n, 'In Stock': s, 'Reorder Level': r });
  const run = (criteria) => b.runQuery({ design: { tables: ['Products'], columns: [{ table: 'Products', field: 'Name' }, { table: 'Products', field: 'In Stock', criteria: [criteria] }] } }).rows.map((r) => r.Name);
  assert.deepEqual(run('<= [Reorder Level]'), ['Milk', 'Tea']);
  assert.deepEqual(run('< [Reorder Level] * 2'), ['Milk', 'Tea']);
  assert.deepEqual(run('[Reorder Level]'), ['Tea']);
  assert.deepEqual(run('Between [Reorder Level] And 10'), ['Bread', 'Tea']);
  // Quoted, it is the words.
  assert.deepEqual(run('"[Reorder Level]"'), []);
  assert.throws(() => run('<= [Reorder Levels]'), /No table in the query has a field called Reorder Levels/);
  assert.throws(() => criterionSql('"x"', '<= [Reorder Level]'), /names a field/);
});

test("an average's heading says Average; Avg in an expression says where totals are; Decimals round it", () => {
  const b = fresh();
  b.createTable({ name: 'Sales', fields: [{ name: 'Total', type: 'currency' }] });
  for (const t of [10, 10, 10.01]) b.insert('Sales', { Total: t });
  const out = b.runQuery({ design: { tables: ['Sales'], columns: [{ table: 'Sales', field: 'Total', total: 'avg', decimals: 2 }, { table: 'Sales', field: 'Total', total: 'sum' }] } });
  assert.deepEqual(out.columns, ['Average of Total', 'Sum of Total']);
  assert.equal(out.rows[0]['Average of Total'], 10);
  assert.throws(() => expressionSql('Round(Avg([Total]), 2)', () => '"Total"'), /Avg totals a column, which an expression here cannot\. Tick Totals, set this column's Total to Average, and choose its Decimals beside it to round it\./);
  // A report made in 2.17 of a query naming "Avg of Total" still finds its column.
  b.saveQuery('Average sale', { tables: ['Sales'], columns: [{ table: 'Sales', field: 'Total', total: 'avg' }] });
  b.saveReport({ name: 'Old', title: 'Old', source: { query: 'Average sale' }, columns: [{ field: 'Avg of Total', label: 'Avg of Total' }], group: null, sort: null, totals: [] });
  const data = b.reportData({ name: 'Old' });
  assert.equal(Math.round(data.groups[0].rows[0]['Average of Total'] * 10000), 100033);
  assert.equal(data.layout.columns[0].field, 'Average of Total');
});

test('a query written in SQL is kept as written, run as a saved query is, and only for those who may change the design', () => {
  const b = fresh();
  b.createTable({ name: 'Sales', fields: [{ name: 'Payment', type: 'text' }, { name: 'Total', type: 'currency' }] });
  for (const [p, t] of [['Cash', 2.5], ['Card', 10], ['Cash', 4]]) b.insert('Sales', { Payment: p, Total: t });
  const sql = 'SELECT "Payment", SUM("Total") AS "Takings" FROM "Sales" GROUP BY "Payment" ORDER BY "Takings" DESC';
  b.saveQuery('Takings by payment', { sql: `${sql};` });
  assert.deepEqual(b.queries().find((x) => x.name === 'Takings by payment').design, { sql });
  assert.deepEqual(b.runQuery({ name: 'Takings by payment' }).rows, [{ Payment: 'Card', Takings: 10 }, { Payment: 'Cash', Takings: 6.5 }]);
  const page = b.queryRows({ name: 'Takings by payment', search: 'cash' });
  assert.deepEqual(page.columns, [{ name: 'Payment', type: 'text' }, { name: 'Takings', type: 'number' }]);
  assert.equal(page.count, 1);
  assert.equal(b.sqlOf({ sql }), sql);
  assert.throws(() => b.saveQuery('Bad', { sql: 'DELETE FROM "Sales"' }), /starts with SELECT/);
  assert.throws(() => b.saveQuery('Bad', { sql: 'SELECT * FROM "Nowhere"' }), /no such table/);
  // A report of it: its amounts summed.
  const report = b.makeReport({ query: 'Takings by payment' });
  assert.deepEqual(report.totals, [{ field: 'Takings', fn: 'sum' }]);
  assert.equal(b.reportData({ name: report.name }).totals['sum:Takings'], 16.5);
  // A restricted role does not see it, and cannot run SQL of its own by a design.
  b.protect({ name: 'Owner', password: 'pass-word-1' });
  b.saveRole('Till', { tables: { Sales: { access: 'read' } } });
  b.addPerson({ name: 'Sam', role: 'Till', password: 'pass-word-2' });
  b.signIn('Sam', 'pass-word-2');
  assert.equal(b.queries().some((x) => x.name === 'Takings by payment'), false);
  assert.throws(() => b.runQuery({ design: { sql } }), /role may change the design/);
  assert.throws(() => b.queryRows({ design: { sql } }), /role may change the design/);
});

test('a record named by nothing of its own is named by its number and date', () => {
  const b = fresh();
  b.createTable({ name: 'Sales', fields: [{ name: 'Sold On', type: 'date' }, { name: 'Payment', type: 'text' }, { name: 'Total', type: 'currency' }] });
  b.createTable({ name: 'Sale Lines', fields: [{ name: 'Sale', type: 'integer' }, { name: 'Quantity', type: 'integer' }] });
  b.addRelationship({ from: { table: 'Sale Lines', field: 'Sale' }, to: { table: 'Sales', field: 'ID' } });
  for (const [d, p] of [['2026-10-21', 'Cash'], ['2026-10-21', 'Cash'], ['2026-10-22', 'Card']]) b.insert('Sales', { 'Sold On': d, Payment: p, Total: 1 });
  b.insert('Sale Lines', { Sale: 1, Quantity: 2 });
  b.insert('Sale Lines', { Sale: 1, Quantity: 1 });
  assert.deepEqual(b.shownOf('Sales'), { number: 'ID', date: 'Sold On', noun: 'Sale' });
  assert.deepEqual(b.lookupPage({ table: 'Sales', field: 'ID', values: [1, 3] }).items.map((o) => o.text).sort(), ['Sale 1, 21/10/2026', 'Sale 3, 22/10/2026']);
  // A line is named by its sale still, as before.
  assert.equal(b.shownOf('Sale Lines'), 'Sale');
  assert.equal(b.wordsOf('Sale Lines', 1), 'Sale 1, 21/10/2026');
  // A print template's fields name the sale so too.
  const doc = b.makeDocument({ table: 'Sale Lines' });
  const data = b.documentData({ layout: doc, keys: [1] });
  assert.ok(JSON.stringify(data.documents[0]).includes('Sale 1, 21/10/2026'), JSON.stringify(data.documents[0]).slice(0, 400));
  // Payments that name one sale each still name it.
  const c = fresh();
  c.createTable({ name: 'Sales', fields: [{ name: 'Sold On', type: 'date' }, { name: 'Payment', type: 'text' }] });
  for (const p of ['Cash', 'Card']) c.insert('Sales', { 'Sold On': '2026-10-21', Payment: p });
  assert.equal(c.shownOf('Sales'), 'Payment');
});

test('a design without the new pieces opens and runs as before', () => {
  const b = fresh();
  b.createTable({ name: 'Items', fields: [{ name: 'Name', type: 'text' }, { name: 'Price', type: 'number' }] });
  b.insert('Items', { Name: 'Tea', Price: 3.5 });
  const { sql, params } = designToSql({ tables: ['Items'], columns: [{ table: 'Items', field: 'Price', criteria: ['> 3'] }] });
  assert.match(sql, /"Items"\."Price" > \?/);
  assert.deepEqual(params, [3]);
});
