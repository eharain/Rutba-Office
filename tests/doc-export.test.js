// A document written as HTML, plain text or Rich Text from the whole
// document model — its headings, its lists counted as their levels say,
// its tables with their widths and spans, its links and pictures — where
// each of them was written from the editor's plain lines, a heading, a list
// item and a table cell each a paragraph like any other.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { listLabels, writeHtmlDocument, writePlainDocument, writeRtfDocument } from '../packages/office-formats/src/doc-export.js';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64');
const para = (text, extra = {}) => ({ type: 'paragraph', heading: null, runs: [{ text }], ...extra });
const item = (text, id, level, style = 'L') => para(text, { list: { id, style, level } });

/** A document of every kind of thing: a heading, looks, a link, a tab, lists two deep, a merged table, a picture. */
const model = () => ({
  title: 'Every kind',
  lists: new Map([
    ['L', [
      { kind: 'number', format: '1', prefix: '', suffix: '.', start: 1, display: 1, indent: 48, hanging: 24 },
      { kind: 'number', format: 'a', prefix: '', suffix: ')', start: 1, display: 1, indent: 72, hanging: 24 },
      { kind: 'number', format: 'i', prefix: '(', suffix: ')', start: 1, display: 1, indent: 96, hanging: 24 },
    ]],
    ['B', [{ kind: 'bullet', char: '•', indent: 48, hanging: 24 }]],
    ['D', [{ kind: 'number', format: '1', prefix: '', suffix: '.', start: 3, display: 1, indent: 48, hanging: 24 }, { kind: 'number', format: '1', prefix: '', suffix: '.', start: 1, display: 2, indent: 72, hanging: 24 }]],
  ]),
  images: new Map([['Pictures/dot.png', PNG]]),
  page: { width: 816, height: 1056, top: 96, bottom: 96, left: 96, right: 96 },
  blocks: [
    { type: 'paragraph', heading: 1, runs: [{ text: 'Every kind' }] },
    { type: 'paragraph', heading: null, align: 'center', runs: [{ text: 'Plain, ' }, { text: 'bold', bold: true }, { text: ', ' }, { text: 'red', color: '#C00000' }, { text: ' and a ' }, { text: 'link', link: 'https://example.org/' }, { tab: true }, { text: 'tabbed' }] },
    item('One', 'n1', 0), item('One a', 'n1', 1), item('One a i', 'n1', 2), item('One a ii', 'n1', 2), item('One b', 'n1', 1), item('Two', 'n1', 0), item('Two a', 'n1', 1),
    item('A bullet', 'b1', 0, 'B'),
    para('Between'),
    item('Three', 'n1', 0),
    item('Starts at three', 'd1', 0, 'D'), item('Under it', 'd1', 1, 'D'),
    {
      type: 'table', columns: [192, 96, 96],
      rows: [
        [{ blocks: [para('Region')], colspan: 1, rowspan: 1, fill: '#DDDDDD', border: true }, { blocks: [para('Q1')], colspan: 1, rowspan: 1, border: true }, { blocks: [para('Q2')], colspan: 1, rowspan: 1, border: true }],
        [{ blocks: [para('North and South')], colspan: 1, rowspan: 2, border: true }, { blocks: [para('both')], colspan: 2, rowspan: 1, border: true }, { covered: true }],
        [{ covered: true }, { blocks: [para('140')], colspan: 1, rowspan: 1, border: true }, { blocks: [para('150')], colspan: 1, rowspan: 1, border: true }],
      ],
    },
    { type: 'paragraph', heading: null, runs: [{ image: { href: 'Pictures/dot.png', width: 48, height: 24, name: 'Dot' } }] },
    para('After the page break', { pageBreakBefore: true }),
  ],
});

test('list labels count as their levels say: a level inside starts again under each item, a list carries on after a paragraph, a start and the levels above shown', () => {
  const doc = model();
  const labels = listLabels(doc);
  const of = (text) => labels.get(doc.blocks.find((b) => b.runs?.[0]?.text === text))?.label;
  assert.deepEqual(['One', 'One a', 'One a i', 'One a ii', 'One b', 'Two', 'Two a', 'A bullet', 'Three', 'Starts at three', 'Under it'].map(of),
    ['1.', 'a)', '(i)', '(ii)', 'b)', '2.', 'a)', '•', '3.', '3.', '3.1.']);
  assert.equal(of('Between'), undefined);
});

test('as HTML: the heading, looks, a link, nested lists in their styles and numbers, the table with its widths, spans and shading, the picture in it', () => {
  const html = writeHtmlDocument(model());
  assert.match(html, /<title>Every kind<\/title>/);
  assert.match(html, /<h1>Every kind<\/h1>/);
  assert.match(html, /<p style="text-align:center">Plain, <strong>bold<\/strong>, <span style="color:#C00000">red<\/span> and a <a href="https:\/\/example\.org\/">link<\/a>&emsp;tabbed<\/p>/);
  assert.match(html, /<ol style="list-style-type:decimal"><li>One<ol style="list-style-type:lower-alpha"><li>One a<ol style="list-style-type:lower-roman"><li>One a i<\/li><li>One a ii<\/li><\/ol><\/li><li>One b<\/li><\/ol><\/li><li>Two<ol style="list-style-type:lower-alpha"><li>Two a<\/li><\/ol><\/li><\/ol>/);
  assert.match(html, /<ol style="list-style-type:decimal" start="3"><li>Three<\/li><\/ol>/, 'the list carries on at three');
  assert.match(html, /<col style="width:192px"><col style="width:96px"><col style="width:96px">/);
  assert.match(html, /<td style="background-color:#DDDDDD;border:1px solid #000"><p>Region<\/p><\/td>/);
  assert.match(html, /<td rowspan="2" style="border:1px solid #000"><p>North and South<\/p><\/td><td colspan="2"/);
  assert.match(html, /<img src="data:image\/png;base64,iVBOR[^"]+" width="48" height="24" alt="Dot">/);
  assert.match(html, /<p style="break-before:page">After the page break<\/p>/);
  assert.match(html, /@page \{ size: 8\.50in 11\.00in; margin: 1\.00in/);
});

test('as plain text: list items after their labels at their levels, a table\'s cells between tabs', () => {
  const text = writePlainDocument(model());
  assert.match(text, /^Every kind\n\nPlain, bold, red and a link\ttabbed\n\n1\. One\n {4}a\) One a\n {8}\(i\) One a i\n {8}\(ii\) One a ii\n {4}b\) One b\n2\. Two\n {4}a\) Two a\n• A bullet\n\nBetween\n\n3\. Three\n3\. Starts at three\n {4}3\.1\. Under it\n\nRegion\tQ1\tQ2\nNorth and South\tboth\t\n\t140\t150\n/);
});

test('as Rich Text: outline level, labels hung at their indents, a HYPERLINK field, the grid\'s widths, merges across and down, shading, a picture, a page break', () => {
  const rtf = writeRtfDocument(model());
  assert.match(rtf, /\{\\info\{\\title Every kind\}\}/);
  assert.match(rtf, /\{\\pard\\outlinelevel0\\sa120 \\b\\fs32 Every kind\\b0\\par\}/);
  assert.match(rtf, /\{\\pard\\qc\\sa120 Plain, \{\\b bold\}, \{\\cf1 red\} and a \{\\field\{\\\*\\fldinst\{HYPERLINK "https:\/\/example\.org\/"\}\}\{\\fldrslt\{link\}\}\}\\tab tabbed\\par\}/);
  assert.match(rtf, /\{\\pard\\li720\\fi-360\\sa120 1\.\\tab One\\par\}\n\{\\pard\\li1080\\fi-360\\sa120 a\)\\tab One a\\par\}\n\{\\pard\\li1440\\fi-360\\sa120 \(i\)\\tab One a i\\par\}/);
  assert.match(rtf, /\\trowd\\trgaph108\\clcbpat\d\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrb\\brdrs\\clbrdrr\\brdrs\\cellx2880\\clbrdrt[^\n]*\\cellx4320\\clbrdrt[^\n]*\\cellx5760\n/);
  assert.match(rtf, /\\trowd\\trgaph108\\clvmgf\\clbrdrt[^\n]*\\cellx2880\\clbrdrt[^\n]*\\cellx5760\n\\intbl North and South\\cell\n\\intbl both\\cell\n\\row/, 'down two rows, and across two columns to the grid\'s edge');
  assert.match(rtf, /\\trowd\\trgaph108\\clvmrg\\clbrdrt[^\n]*\\cellx2880[^\n]*\\cellx4320[^\n]*\\cellx5760\n\\intbl \\cell\n\\intbl 140\\cell/);
  assert.match(rtf, /\{\\pict\\pngblip\\picw48\\pich24\\picwgoal720\\pichgoal360 89504e47/);
  assert.match(rtf, /\{\\pard\\pagebb\\sa120 After the page break\\par\}/);
  // A table that says nothing about borders is not ruled by the model.
  const bare = writeRtfDocument({ ...model(), blocks: [{ type: 'table', columns: [96], rows: [[{ blocks: [para('x')], colspan: 1, rowspan: 1 }]] }] });
  assert.doesNotMatch(bare, /\\brdrs/);
});

test('a Word document saved as .rtf, .html and .txt keeps its heading, its list numbers and its table', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-exports-'));
  const file = path.join(dir, 'figures.docx');
  // Two steps made a numbered list in the editor, as a person makes one.
  const view = openDocx(buildDocx({ styles: true, paragraphs: [
    { text: 'Quarterly figures', style: 'Heading1' },
    { text: 'First step' },
    { text: 'Second step' },
    { table: { rows: [['Region', 'Q1'], ['North', '120']] } },
  ] }));
  const at = (text) => view.render({ pages: false }).blocks.findIndex((b) => b.text === text);
  view.setSelection({ block: at('First step'), offset: 0 }, { block: at('Second step'), offset: 2 });
  view.setParagraphFormat({ list: 'number' });
  fs.writeFileSync(file, view.save());
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const s = docs.open({ path: file });
  const read = (ext) => { const out = path.join(dir, `figures.${ext}`); docs.save({ id: s.id, path: out }); return fs.readFileSync(out, 'utf8'); };
  const html = read('html');
  assert.match(html, /<h1>Quarterly figures<\/h1>/);
  assert.match(html, /<ol style="list-style-type:decimal"><li>First step<\/li><li>Second step<\/li><\/ol>/);
  assert.match(html, /<table[\s\S]*<td[^>]*><p>Region<\/p><\/td><td[^>]*><p>Q1<\/p><\/td>[\s\S]*<td[^>]*><p>North<\/p><\/td>/);
  const rtf = read('rtf');
  assert.match(rtf, /\\outlinelevel0[^\n]*Quarterly figures/);
  assert.match(rtf, /1\.\\tab First step\\par\}\n\{\\pard[^ ]* 2\.\\tab Second step/);
  assert.match(rtf, /\\intbl Region\\cell\n\\intbl Q1\\cell\n\\row/);
  const txt = read('txt');
  assert.match(txt, /^Quarterly figures\n\n1\. First step\n2\. Second step\n\nRegion\tQ1\nNorth\t120\n$/);
});
