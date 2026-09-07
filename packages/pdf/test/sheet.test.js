'use strict';

const test = require('node:test');
const assert = require('node:assert');
const { Sheet } = require('../src/sheet');
const { renderPaper } = require('../src/paper');

const fresh = (opts = {}) => new Sheet({ created: '2026-09-01T09:00:00Z', ...opts });
const rows = (n, prefix = 'Line') =>
  Array.from({ length: n }, (_, i) => ({ d: `${prefix} ${i + 1}`, a: (i + 1).toFixed(2) }));

test('wrapping fits the column and uses every bit of it', () => {
  const sheet = fresh();
  const width = 180;
  const lines = sheet.wrap(
    'Consultancy services rendered during August, including two site visits and a written report',
    width, { size: 10 },
  );
  assert.ok(lines.length > 1, 'it did wrap');
  for (const line of lines) {
    assert.ok(sheet.doc.widthOf(line, { size: 10 }) <= width, `"${line}" overflows`);
  }
  // Greedy means no line could have taken the next word as well.
  for (let i = 0; i < lines.length - 1; i++) {
    const next = lines[i + 1].split(' ')[0];
    assert.ok(sheet.doc.widthOf(`${lines[i]} ${next}`, { size: 10 }) > width,
      `"${lines[i]}" left room for "${next}"`);
  }
});

test('a word too long for its column is broken rather than allowed to overflow', () => {
  // A part number or a URL in a narrow cell. An ugly break beats a value
  // sitting on top of the column beside it.
  const sheet = fresh();
  const lines = sheet.wrap('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789', 40, { size: 10 });
  assert.ok(lines.length > 1);
  for (const line of lines) assert.ok(sheet.doc.widthOf(line, { size: 10 }) <= 40);
  assert.strictEqual(lines.join(''), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789', 'nothing lost');
});

test('explicit newlines are honoured, blank lines included', () => {
  const sheet = fresh();
  assert.deepStrictEqual(sheet.wrap('one\n\ntwo', 300), ['one', '', 'two']);
});

test('content that outgrows the paper takes another sheet', () => {
  const one = fresh();
  one.table({ columns: [{ key: 'd', label: 'D', flex: 1 }], rows: rows(5) });
  assert.strictEqual(one.end().pages.length, 1);

  const many = fresh();
  many.table({ columns: [{ key: 'd', label: 'D', flex: 1 }], rows: rows(300) });
  assert.ok(many.end().pages.length > 3, 'three hundred rows is more than one page');
});

test('nothing is ever drawn below the bottom margin', () => {
  const sheet = fresh();
  sheet.table({ columns: [{ key: 'd', label: 'D', flex: 1 }], rows: rows(120) });
  const doc = sheet.end();
  for (const page of doc.pages) {
    // Every Tm in the stream, back in top-down terms.
    for (const match of page.content.matchAll(/1 0 0 1 [\d.]+ ([\d.]+) Tm/g)) {
      const fromTop = page.height - Number(match[1]);
      assert.ok(fromTop < page.height - 20, `text at ${fromTop.toFixed(1)} is off the paper`);
    }
  }
});

test('a table takes its headings with it onto the next page', () => {
  // Otherwise page two is a grid of numbers with no column names.
  const sheet = fresh();
  sheet.table({
    columns: [{ key: 'd', label: 'Description', flex: 3 }, { key: 'a', label: 'Amount', width: 80, align: 'right' }],
    rows: rows(200),
  });
  const doc = sheet.end();
  assert.ok(doc.pages.length > 2);
  for (const [i, page] of doc.pages.entries()) {
    assert.ok(page.content.includes('(Description)'), `page ${i + 1} has no column headings`);
  }
});

test('the headings stop repeating once the table ends', () => {
  const sheet = fresh();
  sheet.table({ columns: [{ key: 'd', label: 'Description', flex: 1 }], rows: rows(3) });
  assert.strictEqual(sheet.continuation, null);
  for (let i = 0; i < 80; i++) sheet.paragraph('Terms and conditions continue at length.');
  const pages = sheet.end().pages;
  assert.ok(pages.length > 1);
  assert.ok(!pages[pages.length - 1].content.includes('(Description)'),
    'a later page inherited a heading from a table that had finished');
});

test('every page says which page it is, and how many there are', () => {
  const sheet = fresh();
  sheet.table({ columns: [{ key: 'd', label: 'D', flex: 1 }], rows: rows(200) });
  const doc = sheet.end();
  const total = doc.pages.length;
  doc.pages.forEach((page, i) => {
    assert.ok(page.content.includes(`(Page ${i + 1} of ${total})`),
      `page ${i + 1} does not number itself`);
  });
});

test('finishing twice does not draw the footers twice', () => {
  const sheet = fresh();
  sheet.paragraph('One page.');
  sheet.end();
  const doc = sheet.end();
  assert.strictEqual((doc.pages[0].content.match(/\(Page 1 of 1\)/g) || []).length, 1);
});

test('a cell too tall for the page continues onto the next one', () => {
  // Not drawn off the bottom edge, and not chased onto a page that could
  // not hold it either. Both of those lose text a reader was sent.
  const sheet = fresh();
  const enormous = Array.from({ length: 1500 }, (_, i) => 'word' + i).join(' ');
  sheet.table({ columns: [{ key: 'd', label: 'D', width: 120 }], rows: [{ d: enormous }] });
  const doc = sheet.end();
  assert.ok(doc.pages.length > 2, `one cell over three pages, not ${doc.pages.length}`);

  const drawn = doc.pages
    .flatMap((page) => [...page.content.matchAll(/\(([^()]*)\) Tj/g)].map((m) => m[1]))
    .join(' ');
  for (const word of ['word0', 'word750', 'word1499']) {
    assert.ok(new RegExp(`\\b${word}\\b`).test(drawn), `${word} was lost`);
  }

  // And every line of it landed on the paper.
  for (const page of doc.pages) {
    for (const match of page.content.matchAll(/1 0 0 1 [\d.]+ ([\d.]+) Tm/g)) {
      const fromTop = page.height - Number(match[1]);
      assert.ok(fromTop < page.height - 20, `text at ${fromTop.toFixed(1)} is off the paper`);
    }
  }
});

test('side-by-side blocks end level with the tallest, not the last', () => {
  const sheet = fresh();
  const top = sheet.y;
  sheet.columns([
    { lines: ['one line'] },
    { lines: ['a', 'b', 'c', 'd', 'e'] },
  ]);
  assert.ok(sheet.y > top + 4 * 10, 'the cursor cleared the five-line column');
});

test('an empty table says so instead of showing a bare heading', () => {
  const bytes = renderPaper({
    title: 'Statement', reference: 'S-1', created: '2026-09-01T09:00:00Z',
    table: { columns: [{ key: 'd', label: 'Invoice', flex: 1 }], rows: [], empty: 'Nothing outstanding.' },
  });
  assert.ok(bytes.toString('latin1').includes('(Nothing outstanding.)'));
});

test('a document description turns into a document', () => {
  const bytes = renderPaper({
    title: 'Invoice',
    reference: 'INV-0001',
    letterhead: { name: 'Tech-Style (UK) Ltd', lines: ['London'] },
    party: { label: 'Billed to', name: 'Acme (Northern) Ltd', lines: ['Manchester'] },
    meta: [['Invoice date', '2026-08-01'], ['Due date', '']],
    table: {
      columns: [{ key: 'd', label: 'Description', flex: 3 }, { key: 'a', label: 'Amount (GBP)', width: 90, align: 'right' }],
      rows: [{ d: 'Consultancy', a: '100.00' }],
    },
    totals: [['Subtotal', '100.00'], ['Total (GBP)', '120.00']],
    sections: [{ heading: 'How to pay', rows: [['Sort code', '04-00-04']] }],
    footer: 'Registered in England.',
    created: '2026-09-01T09:00:00Z',
  });
  const text = bytes.toString('latin1');
  assert.ok(text.includes('(INVOICE)'), 'the document names itself');
  assert.ok(text.includes('(INV-0001)'));
  assert.ok(text.includes('(Acme \\(Northern\\) Ltd)'), 'the bracket survived');
  assert.ok(text.includes('(Amount \\(GBP\\))'), 'currency is named in the heading');
  assert.ok(text.includes('(Registered in England.)'));
  assert.ok(!text.includes('(Due date)'), 'an empty meta row is omitted, not shown blank');
});
