// Mail merge from a Data Files database: its tables and saved queries offered
// as Word offers an Access file's, the first table taken until one is chosen;
// each field's values as the merge gives them (an amount with its two places,
// a yes/no as Yes or No); the file read and never changed, an older one left
// at its version; and the letter keeping which table it merges, as Word
// writes an Access source, read back with it.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import { Base } from '@rutba/database';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml/build';
import { readMergeSource } from '../apps/desktop/main/mailmerge-source.js';

function members(file) {
  for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
  const b = Base.open(file);
  b.createTable({ name: 'Members', fields: [{ name: 'First Name', type: 'text' }, { name: 'City', type: 'text' }, { name: 'Fee', type: 'currency' }, { name: 'Paid', type: 'yesno' }] });
  b.createTable({ name: 'Visits', fields: [{ name: 'When', type: 'date' }] });
  b.insert('Members', { 'First Name': 'Amira', City: 'Lahore', Fee: 12.5, Paid: true });
  b.insert('Members', { 'First Name': 'Ben', City: 'London', Fee: 20, Paid: false });
  b.saveQuery('Unpaid', { tables: ['Members'], columns: [{ table: 'Members', field: 'First Name' }, { table: 'Members', field: 'Paid', criteria: ['No'], show: false }] });
  b.close();
}

test("a database's tables and queries are the lists to choose from, the first table's records taken until one is chosen", () => {
  const file = path.join(os.tmpdir(), `merge-${process.pid}.rbase`);
  try {
    members(file);
    const src = readMergeSource(file);
    assert.equal(src.kind, 'rbase');
    assert.deepEqual(src.sheets, ['Members', 'Visits', 'Unpaid']);
    assert.equal(src.sheet, 'Members');
    assert.deepEqual(src.fields, ['ID', 'First Name', 'City', 'Fee', 'Paid']);
    assert.deepEqual(src.records, [['1', 'Amira', 'Lahore', '12.50', 'Yes'], ['2', 'Ben', 'London', '20.00', 'No']]);
    const q = readMergeSource(file, { sheet: 'Unpaid' });
    assert.equal(q.sheet, 'Unpaid');
    assert.deepEqual(q.fields, ['First Name']);
    assert.deepEqual(q.records, [['Ben']]);
  } finally {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
  }
});

test('the merge reads the database and changes nothing in it: an older file keeps its version', () => {
  const file = path.join(os.tmpdir(), `merge-old-${process.pid}.rbase`);
  try {
    members(file);
    const raw = new DatabaseSync(file);
    raw.exec("UPDATE rutba_meta SET value = '2' WHERE key = 'version'; DROP TABLE rutba_reports;");
    raw.close();
    readMergeSource(file);
    const after = new DatabaseSync(file);
    assert.equal(after.prepare("SELECT value FROM rutba_meta WHERE key = 'version'").get().value, '2');
    assert.equal(after.prepare("SELECT COUNT(*) AS n FROM sqlite_master WHERE name = 'rutba_reports'").get().n, 0);
    after.close();
    // And a SQLite file that is not one of Data Files' is refused in words.
    const plain = path.join(os.tmpdir(), `merge-plain-${process.pid}.rbase`);
    const p = new DatabaseSync(plain);
    p.exec('CREATE TABLE t (a)');
    p.close();
    assert.throws(() => readMergeSource(plain), /no tables to take recipients from/);
    fs.rmSync(plain, { force: true });
  } finally {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
  }
});

test('a letter keeps which table it merges, as Word writes an Access source, and reads it back', () => {
  const view = openDocx(buildDocx({ paragraphs: ['Dear'] }));
  const engine = view.doc.doc;
  engine.setMailMerge({ type: 'formLetters', source: { kind: 'rbase', path: 'C:\\Club\\Members.rbase', sheet: 'Members' } });
  const settings = engine.pkg.text('word/settings.xml');
  assert.match(settings, /<w:dataType w:val="native"\/><w:connectString w:val="Data Source=C:\\Club\\Members.rbase;Mode=Read;"\/><w:query w:val="SELECT \* FROM `Members` "\/>/);
  const back = openDocx(view.save()).doc.doc.mailMerge();
  assert.equal(back.path, 'C:\\Club\\Members.rbase');
  assert.equal(back.sheet, 'Members');
});
