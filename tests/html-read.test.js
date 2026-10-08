// A web page opened as a document: its headings, looks, links, lists at
// their levels, tables with their spans and its embedded pictures, read
// leniently the way a browser reads a page — where it came in as its words
// alone, one paragraph a block. A document saved as HTML and opened again
// comes back as it went out.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { parseHtml, readHtmlDocument } from '../packages/office-formats/src/html-read.js';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const PAGE = `<!DOCTYPE html><html><head><title>Quarterly &amp; more</title><style>p { color: red }</style><script>document.write('<p>no</p>')</script></head>
<body>
<!-- a comment <p>not this</p> -->
<h1>Figures <em>for</em> the quarter</h1>
<p style="text-align:center">Plain, <b>bold</b>, <span style="color:#c00000;font-size:14pt">red</span>, <i>slanted</i> and a <a href="https://example.org/">link</a><br>a new line &copy; 2026</p>
<ul><li>First bullet<li>Second bullet<ul><li>Inside it</ul></li></ul>
<ol type="a" start="3"><li>Third<li>Fourth</ol>
<table border="1"><colgroup><col width="200"><col width="100"><col width="100"></colgroup>
<tr><th>Region<th>Q1<th>Q2
<tr><td rowspan="2" bgcolor="#dddddd">North and South<td colspan="2">both quarters
<tr><td>140<td>150
</table>
<p><img src="data:image/png;base64,${PNG}" width="48" height="24" alt="Dot"></p>
<pre>  kept
    as typed</pre>
<blockquote>Quoted</blockquote>
</body></html>`;

test('a page is read leniently: unclosed items and cells end where the next begins, scripts, styles and comments are not text', () => {
  const names = (n) => (n.children || []).filter((c) => typeof c !== 'string').map((c) => c.name);
  const root = parseHtml('<ul><li><p>one<li>two</ul><table><tr><td>a<td>b<tr><td>c</table><p>x<div>y</div>');
  assert.deepEqual(names(root), ['ul', 'table', 'p', 'div']);
  assert.deepEqual(names(root.children[0]), ['li', 'li'], 'the second item is not inside the first');
  assert.deepEqual(names(root.children[1]).length, 2);
  const doc = readHtmlDocument(PAGE);
  assert.equal(doc.title, 'Quarterly & more');
  assert.ok(!JSON.stringify(doc.blocks).includes('no</p>') && !JSON.stringify(doc.blocks).includes('not this') && !JSON.stringify(doc.blocks).includes('color: red'));
});

test('a page\'s headings, looks, links, lists, table and picture are read as a document holds them', () => {
  const doc = readHtmlDocument(PAGE);
  const [h1, para] = doc.blocks;
  assert.equal(h1.heading, 1);
  assert.deepEqual(h1.runs.map((r) => [r.text, Boolean(r.italic)]), [['Figures ', false], ['for', true], [' the quarter', false]]);
  assert.equal(para.align, 'center');
  const look = (t) => para.runs.find((r) => r.text === t);
  assert.equal(look('bold').bold, true);
  assert.deepEqual([look('red').color, look('red').size], ['#C00000', 14]);
  assert.equal(look('slanted').italic, true);
  assert.equal(look('link').link, 'https://example.org/');
  assert.ok(para.runs.some((r) => r.br));
  assert.ok(para.runs.some((r) => /© 2026/.test(r.text || '')));
  const items = doc.blocks.filter((b) => b.list).map((b) => [b.runs[0].text, b.list.level, doc.lists.get(b.list.style)[b.list.level].kind]);
  assert.deepEqual(items, [['First bullet', 0, 'bullet'], ['Second bullet', 0, 'bullet'], ['Inside it', 1, 'bullet'], ['Third', 0, 'number'], ['Fourth', 0, 'number']]);
  const lettered = doc.lists.get(doc.blocks.find((b) => b.runs?.[0]?.text === 'Third').list.style)[0];
  assert.deepEqual([lettered.format, lettered.start], ['a', 3]);
  const table = doc.blocks.find((b) => b.type === 'table');
  assert.deepEqual(table.columns, [200, 100, 100]);
  const cell = (c) => (c.covered ? '-' : c.blocks.map((b) => b.runs.map((r) => r.text).join('')).join('/') + (c.rowspan > 1 ? `^${c.rowspan}` : '') + (c.colspan > 1 ? `<${c.colspan}` : ''));
  assert.deepEqual(table.rows.map((r) => r.map(cell).join('|')), ['Region|Q1|Q2', 'North and South^2|both quarters<2|-', '-|140|150']);
  assert.equal(table.rows[0][0].blocks[0].runs[0].bold, true, 'a header cell is bold');
  assert.equal(table.rows[1][0].fill, '#DDDDDD');
  assert.ok(table.rows[0][0].border);
  const img = doc.blocks.flatMap((b) => b.runs || []).find((r) => r.image)?.image;
  assert.deepEqual([img.width, img.height, img.name], [48, 24, 'Dot']);
  assert.deepEqual(Buffer.from(doc.images.get(img.href)), Buffer.from(PNG, 'base64'));
  const pre = doc.blocks.find((b) => b.runs?.some((r) => r.text === '  kept'));
  assert.deepEqual(pre.runs.map((r) => (r.br ? '\n' : r.text)), ['  kept', '\n', '    as typed']);
  assert.equal(pre.runs[0].font, 'Consolas');
  assert.ok(doc.blocks.find((b) => b.runs?.[0]?.text === 'Quoted').indentLeft > 0);
});

test('opened in the suite, a page is a document: its heading styled, its lists labelled, its table merged and its picture drawn', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-html-'));
  const file = path.join(dir, 'page.html');
  fs.writeFileSync(file, PAGE);
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const s = docs.open({ path: file });
  assert.equal(s.converted?.from, 'html');
  const m = docs.model({ id: s.id });
  const text = (b) => (b.runs || []).map((r) => r.text || '').join('');
  const index = (t) => m.blocks.findIndex((b) => text(b) === t);
  assert.equal(m.blocks[0].style, 'Heading1');
  assert.deepEqual(['First bullet', 'Inside it', 'Third', 'Fourth'].map((t) => m.listLabels?.[index(t)]?.label), ['•', '◦', 'c.', 'd.']);
  assert.ok(m.blocks.some((b) => /^t\d+:r1:c0$/.test(b.container || '') && text(b) === 'North and South'));
  assert.ok(m.drawings.some((d) => d.kind === 'picture' && d.name === 'Dot'));

  // Saved as HTML and opened again: the same headings, list labels, table and picture.
  const out = path.join(dir, 'again.html');
  docs.save({ id: s.id, path: out });
  const back = docs.model({ id: docs.open({ path: out }).id });
  const indexBack = (t) => back.blocks.findIndex((b) => text(b) === t);
  assert.equal(back.blocks[0].style, 'Heading1');
  assert.deepEqual(['First bullet', 'Inside it', 'Third', 'Fourth'].map((t) => back.listLabels?.[indexBack(t)]?.label), ['•', '◦', 'c.', 'd.']);
  assert.ok(back.blocks.some((b) => /^t\d+:r1:c0$/.test(b.container || '') && text(b) === 'North and South'));
  assert.ok(back.drawings.some((d) => d.kind === 'picture'));
});
