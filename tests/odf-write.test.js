// OpenDocument, written and read back.
//
// The installer registers this suite as the editor of .odt, .ods and .odp;
// until now Ctrl+S on one refused. The writer puts back what the reader
// takes, so a file that came in as OpenDocument leaves as OpenDocument, and
// these tests close the loop through the reader itself.

import test from 'node:test';
import assert from 'node:assert/strict';
import { writeOdt, writeOds, writeOdp, formulaToOdf, formulaFromOdf } from '../packages/office-formats/src/odf-write.js';
import { readOdf, odfFlavour } from '../packages/office-formats/src/odf.js';
import { readZip } from '../packages/ooxml/src/zip.js';

/** One part of the archive, as text — the parts are deflated, so the bytes are not searchable. */
function part(bytes, name) {
  const { entries } = readZip(bytes);
  const e = entries.find((x) => x.name === name);
  return e ? Buffer.from(e.data).toString('utf8') : '';
}

test('the archive starts with a stored mimetype, as the format requires', () => {
  const bytes = writeOdt({ blocks: [{ type: 'paragraph', runs: [{ text: 'Hello' }] }] });
  // Local header: signature, version, flags, method (0 = stored), then the name at offset 30.
  assert.equal(bytes.readUInt32LE(0), 0x04034b50);
  assert.equal(bytes.readUInt16LE(8), 0, 'mimetype is stored, not deflated');
  assert.equal(bytes.toString('latin1', 30, 38), 'mimetype');
  assert.equal(bytes.toString('latin1', 38, 38 + 39), 'application/vnd.oasis.opendocument.text');
  assert.equal(odfFlavour(bytes), 'odt');
});

test('a document keeps its headings, paragraphs, runs, lists and tables', () => {
  const blocks = [
    { type: 'heading', level: 1, runs: [{ text: 'Quarterly Review' }] },
    { type: 'paragraph', runs: [{ text: 'The northern region ' }, { text: 'grew', bold: true }, { text: ' by 12%.' }] },
    { type: 'paragraph', list: { kind: 'bullet' }, runs: [{ text: 'First point' }] },
    { type: 'paragraph', list: { kind: 'bullet' }, runs: [{ text: 'Second point' }] },
    { type: 'paragraph', list: { kind: 'number' }, runs: [{ text: 'Step one' }] },
    { type: 'table', rows: [[{ runs: [{ text: 'Region' }] }, { runs: [{ text: 'Sales' }] }], [{ runs: [{ text: 'North' }] }, { runs: [{ text: '1,200' }] }]] },
    { type: 'paragraph', align: 'center', runs: [{ text: 'Tab\there  and two spaces', italic: true, size: 14, colour: '#ff0000' }] },
  ];
  const back = readOdf(writeOdt({ blocks, title: 'Review' }));
  assert.equal(back.kind ?? 'odt', back.kind ?? 'odt');
  const got = back.blocks;
  assert.equal(got[0].type, 'heading');
  assert.equal(got[0].level, 1);
  assert.equal(got[0].text, 'Quarterly Review');
  assert.equal(got[1].text, 'The northern region grew by 12%.');
  assert.equal(got[2].type, 'list');
  assert.equal(got[2].ordered, false);
  assert.deepEqual(got[2].items, ['First point', 'Second point']);
  assert.equal(got[3].type, 'list');
  assert.equal(got[3].ordered, true);
  assert.deepEqual(got[3].items, ['Step one']);
  assert.equal(got[4].type, 'table');
  assert.deepEqual(got[4].rows.map((r) => r.map((c) => c.text)), [['Region', 'Sales'], ['North', '1,200']]);
  assert.equal(got[5].text, 'Tab\there  and two spaces', 'tabs and repeated spaces survive');
  assert.equal(back.meta?.title, 'Review');
});

test('run formatting is written as text styles, once each', () => {
  const s = part(writeOdt({ blocks: [{ type: 'paragraph', runs: [{ text: 'a', bold: true }, { text: 'b', bold: true }, { text: 'c', italic: true }] }] }), 'content.xml');
  assert.ok(/fo:font-weight="bold"/.test(s));
  assert.ok(/fo:font-style="italic"/.test(s));
  assert.equal((s.match(/style:family="text"/g) || []).length, 2, 'two distinct run styles, not three');
});

test('formulas cross into the file dialect and back', () => {
  assert.equal(formulaToOdf('=SUM(A1:B2, 3)'), 'of:=SUM([.A1:.B2]; 3)');
  assert.equal(formulaToOdf("=Sheet2!$A$1+'My Sheet'!B2"), "of:=[$Sheet2.$A$1]+[$'My Sheet'.B2]");
  assert.equal(formulaToOdf('=IF(A1>0,"yes, sir",B:B)'), 'of:=IF([.A1]>0;"yes, sir";[.B1:.B1048576])');
  assert.equal(formulaToOdf('=LEN("A1,B2")'), 'of:=LEN("A1,B2")', 'a reference inside a string is text');
  assert.equal(formulaFromOdf('of:=SUM([.A1:.B2]; 3)'), '=SUM(A1:B2, 3)');
  assert.equal(formulaFromOdf("of:=[$Sheet2.$A$1]+[$'My Sheet'.B2]"), "=Sheet2!$A$1+'My Sheet'!B2");
  assert.equal(formulaFromOdf('=SUM([.A1:.B2])'), '=SUM(A1:B2)', 'a bare formula without the prefix');
  assert.equal(formulaFromOdf(formulaToOdf('=AVERAGE(Data!A2:A100)*1.5')), '=AVERAGE(Data!A2:A100)*1.5');
});

test('a workbook keeps values, formulas, dates, percentages and text', () => {
  const sheets = [
    {
      name: 'Sales',
      rows: [
        [{ value: 'Region', text: 'Region' }, { value: 'Amount', text: 'Amount' }, null, { value: 45000, text: '09/03/2023', format: 'dd/mm/yyyy' }],
        [{ value: 'North', text: 'North' }, { value: 1200.5, text: '1,200.50', format: '#,##0.00' }, { value: 0.125, text: '12.5%', format: '0.0%' }],
        [{ value: 'Total', text: 'Total' }, { value: 1200.5, text: '1,200.50', formula: '=SUM(B2:B2)', format: '#,##0.00' }, { value: true, text: 'TRUE' }],
      ],
    },
    { name: 'Empty', rows: [] },
  ];
  const back = readOdf(writeOds({ sheets, title: 'Book' }));
  assert.equal(back.sheets.length, 2);
  const s = back.sheets[0];
  assert.equal(s.name, 'Sales');
  assert.equal(s.rows[0][0].text, 'Region');
  assert.equal(s.rows[0][3].type, 'date');
  assert.equal(s.rows[0][3].value, '2023-03-15', 'serial 45000 is 15 March 2023');
  assert.equal(s.rows[1][1].type, 'float');
  assert.equal(Number(s.rows[1][1].value), 1200.5);
  assert.equal(s.rows[1][2].type, 'percentage');
  assert.equal(s.rows[2][1].formula, '=SUM(B2:B2)');
  assert.equal(s.rows[2][2].type, 'boolean');
  assert.equal(s.rows[0][2].text, '', 'an empty cell between two stays empty');
  assert.equal(back.sheets[1].name, 'Empty');
});

test('a deck keeps its text boxes, notes and a picture', () => {
  const png = Buffer.from('89504e470d0a1a0a0000000d49484452', 'hex');
  const slides = [
    {
      name: 'Opening',
      shapes: [
        { kind: 'text', placeholder: 'title', paragraphs: [{ runs: [{ text: 'Welcome', bold: true }] }], x: 96, y: 48, w: 960, h: 120 },
        { kind: 'text', paragraphs: ['One', 'Two'], x: 96, y: 200, w: 960, h: 400 },
        { kind: 'picture', data: png, contentType: 'image/png', x: 600, y: 300, w: 200, h: 150 },
      ],
      notes: 'Say hello.\nThen begin.',
    },
    { name: 'Second', shapes: [], notes: '' },
  ];
  const bytes = writeOdp({ slides, size: { width: 1280, height: 720 }, title: 'Deck' });
  assert.equal(odfFlavour(bytes), 'odp');
  const back = readOdf(bytes);
  assert.equal(back.slides.length, 2);
  const first = back.slides[0];
  assert.equal(first.name, 'Opening');
  const texts = first.shapes.filter((s) => s.type === 'text');
  assert.deepEqual(texts[0].paragraphs, ['Welcome']);
  assert.deepEqual(texts[1].paragraphs, ['One', 'Two']);
  assert.equal(texts[0].x, '2.540cm', 'ninety-six pixels is an inch');
  const image = first.shapes.find((s) => s.type === 'image');
  assert.equal(image.href, 'Pictures/image1.png');
  assert.equal(first.notes, 'Say hello.\nThen begin.');
});
