// Save on a .csv or an .ods: what the file held is still there, and what the
// format has no place for is said before it goes.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createDocumentService } from '../apps/desktop/main/documents.js';
import { readOdf, validationFromOdf } from '../packages/office-formats/src/odf.js';
import { validationToOdf, writeOds } from '../packages/office-formats/src/odf-write.js';
import { buildXlsx } from '@rutba/ooxml/build';

const doc = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-sheet-save-'));
let n = 0;
const file = (name, bytes) => {
  const p = path.join(dir, `${++n}-${name}`);
  fs.writeFileSync(p, bytes);
  return p;
};
const saveAgain = (p) => {
  const opened = doc.open({ path: p });
  doc.save({ id: opened.id, path: p });
  doc.close?.({ id: opened.id });
  return fs.readFileSync(p);
};

test('a .csv saved unchanged is the same file', () => {
  for (const text of [
    'code,amount\r\n00123,"1,234.50"\r\n007,1.50\r\n',
    'a;b;c\n1,5;2,75;x\n',
    'long,short\n12345678901234567890,1\n4111111111111111,2\n',
    'one,two,three\nonly\nleft,right\n',
    `${String.fromCharCode(0xfeff)}name,qty
Åse,+5
`,
    'plain,row',
  ]) {
    assert.equal(saveAgain(file('same.csv', text)).toString('utf8'), text);
  }
});

test('a code is text: its zeros are there to see, and a card number keeps its end', () => {
  const opened = doc.open({ path: file('codes.csv', 'code\n00123\n4111111111111111111\n42\n') });
  const out = path.join(dir, 'codes-out.ods');
  doc.save({ id: opened.id, path: out });
  const rows = readOdf(fs.readFileSync(out)).sheets[0].rows;
  assert.equal(rows[1][0].type, 'string');
  assert.equal(rows[1][0].text, '00123');
  assert.equal(rows[2][0].text, '4111111111111111111');
  assert.equal(rows[3][0].type, 'float', 'a plain number is still a number');
});

test('a .tsv keeps its tabs, and a .csv saved as .tsv takes the new one', () => {
  assert.equal(saveAgain(file('t.tsv', 'a\tb\n1\t2\n')).toString(), 'a\tb\n1\t2\n');
  const opened = doc.open({ path: file('c.csv', 'a;b\n1;2\n') });
  const out = path.join(dir, 'c-out.tsv');
  doc.save({ id: opened.id, path: out });
  assert.equal(fs.readFileSync(out, 'utf8'), 'a\tb\n1\t2\n');
});

test('a plain .csv has nothing to warn about, and a formula or a second sheet does', () => {
  const p = file('plain.csv', 'a,b\n1,2\n');
  const plain = doc.open({ path: p });
  assert.deepEqual(doc.saveLosses({ id: plain.id, path: p }).lost, []);
  const book = doc.open({ path: file('book.xlsx', buildXlsx({ sheets: [{ name: 'One', rows: [[1, '=A1+1']] }, { name: 'Two', rows: [[2]] }] })) });
  const lost = doc.saveLosses({ id: book.id, path: path.join(dir, 'x.csv') }).lost;
  assert.ok(lost.includes('sheets'));
  assert.ok(lost.includes('formulas'));
  assert.deepEqual(doc.saveLosses({ id: book.id, path: path.join(dir, 'x.xlsx') }).lost, [], 'a workbook holds it all');
});

test('saving to .ods keeps validations and named ranges, through a second save', () => {
  const opened = doc.open({
    path: file('full.xlsx', buildXlsx({
      definedNames: [{ name: 'Rate', ref: 'One!$B$2' }, { name: 'Span', ref: "'Second one'!$A$1:$B$2" }],
      sheets: [
        { name: 'One', rows: [['a', 1], ['b', 2], ['c', 3]], validations: [
          { type: 'list', formula1: '"yes,no"', sqref: 'A1:A2', errorTitle: 'Pick one', error: 'Yes or no', prompt: 'Choose' },
          { type: 'whole', operator: 'between', formula1: '1', formula2: '10', sqref: 'B1' },
          { type: 'decimal', operator: 'greaterThan', formula1: '0.5', sqref: 'B2' },
          { type: 'textLength', operator: 'lessThanOrEqual', formula1: '4', sqref: 'C1' },
          { type: 'list', formula1: '$A$1:$A$3', sqref: 'D5' },
        ] },
        { name: 'Second one', rows: [['x']] },
      ],
    })),
  });
  const session = opened.id;
  const out = path.join(dir, 'full.ods');
  doc.save({ id: session, path: out });
  assert.deepEqual(doc.saveLosses({ id: session, path: out }).lost, [], 'nothing here the format cannot hold');

  const check = (bytes) => {
    const odf = readOdf(bytes);
    const one = odf.sheets[0];
    assert.deepEqual(odf.names.map((x) => [x.name, x.ref]).sort(), [['Rate', 'One!$B$2'], ['Span', "'Second one'!$A$1:$B$2"]]);
    const rules = one.validated.map((v) => [v.top, v.left, odf.validations.get(v.name)?.type]);
    assert.deepEqual(rules.filter(([r, c]) => c === 0 && r < 2).map((x) => x[2]), ['list', 'list']);
    assert.ok(rules.some(([r, c, t]) => r === 0 && c === 1 && t === 'whole'));
    assert.ok(rules.some(([r, c, t]) => r === 1 && c === 1 && t === 'decimal'));
    assert.ok(rules.some(([r, c, t]) => r === 0 && c === 2 && t === 'textLength'));
    assert.ok(rules.some(([r, c, t]) => r === 4 && c === 3 && t === 'list'), 'on a cell with nothing in it');
    const first = odf.validations.get(one.validated[0].name);
    assert.equal(first.formula1, '"yes,no"');
    assert.equal(first.error, 'Yes or no');
    assert.equal(first.prompt, 'Choose');
  };
  check(fs.readFileSync(out));
  check(saveAgain(out));
});

test('an .ods opened and saved keeps its comments and its hidden rows and columns, twice over', () => {
  const cell = (text, comment) => ({ value: text, text, ...(comment ? { comment } : {}) });
  const p = file('notes.ods', writeOds({
    sheets: [{
      name: 'One',
      hiddenRows: [1],
      hiddenCols: [1],
      rows: [
        [cell('a', { author: 'Sam', text: 'First line\nSecond line' }), cell('b')],
        [cell('c'), cell('d')],
        [null, { comment: { author: '', text: 'On an empty cell' } }],
        [cell('e'), cell('f')],
      ],
    }],
  }));
  const check = (bytes) => {
    const one = readOdf(bytes).sheets[0];
    assert.deepEqual(one.notes.map((x) => [x.row, x.col, x.author, x.text]), [[0, 0, 'Sam', 'First line\nSecond line'], [2, 1, '', 'On an empty cell']]);
    assert.deepEqual(one.hiddenRows, [1]);
    assert.deepEqual(one.hiddenCols, [1]);
  };
  check(fs.readFileSync(p));
  check(saveAgain(p));
  check(saveAgain(p));
});

test('each rule says the same thing after a trip through the format', () => {
  for (const rule of [
    { type: 'list', formula1: '"a,b c,d"' },
    { type: 'list', formula1: '$D$1:$D$9' },
    { type: 'whole', operator: 'between', formula1: '1', formula2: '10' },
    { type: 'whole', operator: 'notBetween', formula1: '-3', formula2: '3' },
    { type: 'decimal', operator: 'greaterThanOrEqual', formula1: '0.25' },
    { type: 'decimal', operator: 'notEqual', formula1: '5' },
    { type: 'textLength', operator: 'between', formula1: '1', formula2: '5' },
    { type: 'textLength', operator: 'lessThan', formula1: '9' },
  ]) {
    assert.deepEqual(validationFromOdf(validationToOdf(rule)), rule);
  }
});

test('a rule the format cannot say is named before Save, not dropped in silence', () => {
  const opened = doc.open({ path: file('custom.xlsx', buildXlsx({ sheets: [{ name: 'One', rows: [[1]], validations: [{ type: 'custom', formula1: 'A1>0', sqref: 'A1' }] }] })) });
  assert.deepEqual(doc.saveLosses({ id: opened.id, path: path.join(dir, 'c.ods') }).lost, ['validations']);
});

test('an .ods whose rule cannot be read says so too', () => {
  assert.equal(validationFromOdf('of:cell-content-is-date() and cell-content()>1'), null);
});
