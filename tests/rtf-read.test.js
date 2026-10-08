// Rich Text opened as it was made — Word's, WordPad's or this suite's own:
// headings, looks, links, lists at their levels in their formats, tables
// with their widths, merges and shading, pictures and the page, where it
// came in as paragraphs of runs and tables of plain cells.
//
// tests/fixtures/rtf/word-made.rtf is Microsoft Word's own Rich Text: the
// page in tests/html-read.test.js opened in the suite, saved as .docx, and
// saved again from Word as .rtf (its author and operator then set to
// "Rutba Office").
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { readRtfDocument } from '../packages/office-formats/src/rtf-read.js';
import { writeRtfDocument } from '../packages/office-formats/src/doc-export.js';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const WORD = path.join(path.dirname(fileURLToPath(import.meta.url)), 'fixtures', 'rtf', 'word-made.rtf');
const cell = (c) => (c.covered ? '-' : c.blocks.map((b) => b.runs.map((r) => r.text).join('')).join('/') + (c.rowspan > 1 ? `^${c.rowspan}` : '') + (c.colspan > 1 ? `<${c.colspan}` : ''));
const textOf = (b) => (b.runs || []).map((r) => r.text ?? '').join('');

test('Word\'s Rich Text is read whole: the heading, looks, a link, lists at their levels, the table\'s widths, merges and shading, the picture, the page', () => {
  const doc = readRtfDocument(fs.readFileSync(WORD));
  const [heading, para] = doc.blocks;
  assert.equal(heading.heading, 1);
  assert.equal(textOf(heading), 'Figures for the quarter');
  assert.equal(heading.runs.find((r) => r.text === 'for').italic, true);
  assert.equal(para.align, 'center');
  const run = (t) => para.runs.find((r) => r.text === t);
  assert.equal(run('bold').bold, true);
  assert.deepEqual([run('red').color, run('red').size], ['#C00000', 14]);
  assert.equal(run('slanted').italic, true);
  assert.equal(run('link').link, 'https://example.org/');
  assert.ok(para.runs.some((r) => r.br), 'the line break');
  const items = doc.blocks.filter((b) => b.list);
  const level = (b) => doc.lists.get(b.list.style)[b.list.level];
  assert.deepEqual(items.map((b) => [textOf(b), b.list.level, level(b).kind]), [['First bullet', 0, 'bullet'], ['Second bullet', 0, 'bullet'], ['Inside it', 1, 'bullet'], ['Third', 0, 'number'], ['Fourth', 0, 'number']]);
  assert.deepEqual([level(items[3]).format, level(items[3]).start], ['a', 3], 'lettered from c');
  assert.equal(level(items[2]).char, '◦');
  const table = doc.blocks.find((b) => b.type === 'table');
  assert.deepEqual(table.columns.map(Math.round), [200, 100, 100]);
  assert.deepEqual(table.rows.map((r) => r.map(cell).join('|')), ['Region|Q1|Q2', 'North and South^2|both quarters<2|-', '-|140|150']);
  assert.equal(table.rows[1][0].fill, '#DDDDDD');
  const img = doc.blocks.flatMap((b) => b.runs || []).find((r) => r.image)?.image;
  assert.deepEqual([Math.round(img.width), Math.round(img.height)], [48, 24]);
  assert.equal(Buffer.from(doc.images.get(img.href)).subarray(1, 4).toString(), 'PNG');
  const pre = doc.blocks.find((b) => textOf(b).includes('kept'));
  assert.deepEqual(pre.runs.map((r) => (r.br ? '\n' : r.text)), ['  kept', '\n', '    as typed']);
  assert.equal(pre.runs[0].font, 'Consolas');
  assert.ok(doc.blocks.find((b) => textOf(b) === 'Quoted').indentLeft > 0);
  assert.deepEqual([doc.page.width, doc.page.height], [816, 1056]);
});

test('this suite\'s own Rich Text comes back as it went out: lists as lists, a merge down, widths', () => {
  const lists = new Map([['L', [{ kind: 'number', format: '1', prefix: '', suffix: '.', start: 1, display: 1, indent: 48, hanging: 24 }, { kind: 'number', format: 'a', prefix: '', suffix: ')', start: 1, display: 1, indent: 72, hanging: 24 }]]]);
  const p = (text, extra = {}) => ({ type: 'paragraph', heading: null, runs: [{ text }], ...extra });
  const model = {
    lists,
    blocks: [
      p('Title', { heading: 1 }),
      p('One', { list: { id: 'n1', style: 'L', level: 0 } }), p('One a', { list: { id: 'n1', style: 'L', level: 1 } }), p('Two', { list: { id: 'n1', style: 'L', level: 0 } }),
      { type: 'table', columns: [192, 96], rows: [[{ blocks: [p('Down')], colspan: 1, rowspan: 2 }, { blocks: [p('1')], colspan: 1, rowspan: 1 }], [{ covered: true }, { blocks: [p('2')], colspan: 1, rowspan: 1 }]] },
      p('After', { runs: [{ text: 'After ' }, { text: 'link', link: 'https://example.org/' }] }),
    ],
  };
  const back = readRtfDocument(writeRtfDocument(model));
  assert.equal(back.blocks[0].heading, 1);
  const items = back.blocks.filter((b) => b.list);
  assert.deepEqual(items.map((b) => [textOf(b), b.list.level, back.lists.get(b.list.style)[b.list.level].format, back.lists.get(b.list.style)[b.list.level].suffix]), [['One', 0, '1', '.'], ['One a', 1, 'a', ')'], ['Two', 0, '1', '.']]);
  assert.equal(new Set(items.map((b) => b.list.id)).size, 1, 'one list');
  const table = back.blocks.find((b) => b.type === 'table');
  assert.deepEqual(table.columns.map(Math.round), [192, 96]);
  assert.deepEqual(table.rows.map((r) => r.map(cell).join('|')), ['Down^2|1', '-|2']);
  assert.equal(back.blocks.at(-1).runs.find((r) => r.text === 'link').link, 'https://example.org/');
});

test('WordPad\'s lists — a \\pntext label before the words — are lists, the label left to the list', () => {
  const rtf = '{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Calibri;}{\\f1 Symbol;}}\n{\\pntext\\f1\\\'B7\\tab}{\\*\\pn\\pnlvlblt\\pnf1\\pnindent0{\\pntxtb\\\'B7}}\\fi-360\\li720 Apples\\par\n{\\pntext\\f1\\\'B7\\tab}Pears\\par\n\\pard Plain after\\par\n{\\pntext 1.\\tab}{\\*\\pn\\pnlvlbody\\pndec{\\pntxta .}}\\fi-360\\li720 First\\par\n}';
  const doc = readRtfDocument(rtf);
  assert.deepEqual(doc.blocks.map((b) => [textOf(b), b.list ? doc.lists.get(b.list.style)[0].kind : null]), [['Apples', 'bullet'], ['Pears', 'bullet'], ['Plain after', null], ['First', 'number']]);
});

test('opened in the suite, an .rtf is a document: its heading styled, its lists labelled, its table merged and its picture drawn', () => {
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-rtf-')), 'word-made.rtf');
  fs.copyFileSync(WORD, file);
  const s = docs.open({ path: file });
  assert.equal(s.converted?.from, 'rtf');
  const m = docs.model({ id: s.id });
  const index = (t) => m.blocks.findIndex((b) => textOf(b) === t);
  assert.equal(m.blocks[0].style, 'Heading1');
  assert.deepEqual(['First bullet', 'Inside it', 'Third', 'Fourth'].map((t) => m.listLabels?.[index(t)]?.label), ['•', '◦', 'c.', 'd.']);
  assert.ok(m.blocks.some((b) => /^t\d+:r1:c0$/.test(b.container || '') && textOf(b) === 'North and South'));
  assert.ok(m.drawings.some((d) => d.kind === 'picture'));
});
