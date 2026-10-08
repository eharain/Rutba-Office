// OpenDocument text, opened as it was made.
//
// An .odt used to arrive as plain lines: its tables as tab-separated text,
// its lists as bullet characters, its runs without their looks and its
// pictures dropped. It is now read whole and written as the .docx the Word
// engine edits: headings, runs in their looks, links, tabs and breaks,
// lists as numbering at their levels, tables with their widths, spans and
// shading, pictures, page breaks and the page itself. The file here is
// written the way LibreOffice writes one: automatic styles over named
// ones, list styles, header rows and covered cells.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { readZip, writeZip, ZipEntry } from '@rutba/ooxml/zip';
import { writeOdt } from '../packages/office-formats/src/odf-write.js';
import { readOdt } from '../packages/office-formats/src/odt.js';
import { odtToDocx } from '../packages/office-formats/src/odt-docx.js';
import { readDocxDocument } from '../packages/office-formats/src/docx-read.js';
import { writeOdtDocument } from '../packages/office-formats/src/odt-write.js';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml/build';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64');
const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.3"';

const STYLES = `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NS}><office:styles>`
  + '<style:style style:name="Standard" style:family="paragraph"/>'
  + '<style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard"><style:text-properties fo:font-size="14pt"/></style:style>'
  + '<style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="paragraph" style:parent-style-name="Heading" style:default-outline-level="1"><style:text-properties fo:font-size="18pt" fo:font-weight="bold"/></style:style>'
  + '</office:styles><office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2.5cm" fo:margin-right="2.5cm"/></style:page-layout></office:automatic-styles>'
  + '<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"/></office:master-styles></office:document-styles>';

const CONTENT = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS}><office:automatic-styles>`
  + '<style:style style:name="P1" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:text-align="center"/></style:style>'
  + '<style:style style:name="P2" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:break-before="page"/></style:style>'
  + '<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>'
  + '<style:style style:name="T2" style:family="text"><style:text-properties fo:font-style="italic" fo:color="#c9211e" fo:font-size="14pt" style:font-name="Liberation Mono"/></style:style>'
  + '<style:style style:name="Table1.A" style:family="table-column"><style:table-column-properties style:column-width="5cm"/></style:style>'
  + '<style:style style:name="Table1.B" style:family="table-column"><style:table-column-properties style:column-width="3cm"/></style:style>'
  + '<style:style style:name="Table1.A1" style:family="table-cell"><style:table-cell-properties fo:background-color="#dddddd" fo:border="0.5pt solid #000000"/></style:style>'
  + '<text:list-style style:name="L1"><text:list-level-style-bullet text:level="1" text:bullet-char="•"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" fo:text-indent="-0.635cm" fo:margin-left="1.27cm"/></style:list-level-properties></text:list-level-style-bullet></text:list-style>'
  + '<text:list-style style:name="L2"><text:list-level-style-number text:level="1" style:num-suffix="." style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment fo:text-indent="-0.635cm" fo:margin-left="1.27cm"/></style:list-level-properties></text:list-level-style-number>'
  + '<text:list-level-style-number text:level="2" style:num-suffix=")" style:num-format="a"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment fo:text-indent="-0.635cm" fo:margin-left="1.905cm"/></style:list-level-properties></text:list-level-style-number></text:list-style>'
  + '</office:automatic-styles><office:body><office:text>'
  + '<text:h text:style-name="Heading_20_1" text:outline-level="1">The showcase</text:h>'
  + '<text:p text:style-name="Standard">Plain <text:span text:style-name="T1">bold</text:span> and <text:span text:style-name="T2">red italic</text:span> with a <text:a xlink:type="simple" xlink:href="https://example.org/">link</text:a>.<text:tab/>tab<text:line-break/>next line</text:p>'
  + '<text:p text:style-name="P1">Centred</text:p>'
  + '<text:list text:style-name="L1"><text:list-item><text:p>First bullet</text:p></text:list-item><text:list-item><text:p>Second bullet</text:p></text:list-item></text:list>'
  + '<text:list text:style-name="L2"><text:list-item><text:p>Step one</text:p><text:list><text:list-item><text:p>Inside</text:p></text:list-item></text:list></text:list-item><text:list-item><text:p>Step two</text:p></text:list-item></text:list>'
  + '<table:table table:name="Table1"><table:table-column table:style-name="Table1.A"/><table:table-column table:style-name="Table1.B" table:number-columns-repeated="2"/>'
  + '<table:table-header-rows><table:table-row><table:table-cell table:style-name="Table1.A1" office:value-type="string"><text:p>Region</text:p></table:table-cell><table:table-cell table:style-name="Table1.A1"><text:p>Q1</text:p></table:table-cell><table:table-cell table:style-name="Table1.A1"><text:p>Q2</text:p></table:table-cell></table:table-row></table:table-header-rows>'
  + '<table:table-row><table:table-cell table:number-rows-spanned="2"><text:p>North and South</text:p></table:table-cell><table:table-cell><text:p>120</text:p></table:table-cell><table:table-cell><text:p>135</text:p></table:table-cell></table:table-row>'
  + '<table:table-row><table:covered-table-cell/><table:table-cell table:number-columns-spanned="2"><text:p>both quarters</text:p></table:table-cell><table:covered-table-cell/></table:table-row>'
  + '</table:table>'
  + '<text:p text:style-name="Standard"><draw:frame draw:name="Logo" text:anchor-type="as-char" svg:width="2cm" svg:height="1cm"><draw:image xlink:href="Pictures/logo.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame></text:p>'
  + '<text:p text:style-name="P2">After the page break.</text:p>'
  + '</office:text></office:body></office:document-content>';

/** The file: our own archive's mimetype and manifest, with these parts in it and a picture. */
function odt() {
  const { entries } = readZip(Buffer.from(writeOdt({ blocks: [] })));
  for (const e of entries) {
    if (e.name === 'content.xml') e.data = Buffer.from(CONTENT, 'utf8');
    if (e.name === 'styles.xml') e.data = Buffer.from(STYLES, 'utf8');
  }
  const pic = new ZipEntry({ name: 'Pictures/logo.png', method: 0, crc: 0, compressedSize: 0, uncompressedSize: 0, compressed: Buffer.alloc(0), dosTime: 0, dosDate: 0x2821, flags: 0, externalAttrs: 0, comment: Buffer.alloc(0) });
  pic.data = PNG;
  return writeZip([...entries, pic]);
}

const docxParts = (bytes) => new Map(readZip(Buffer.from(bytes)).entries.map((e) => [e.name, e.data]));

test('an .odt is read whole: headings, runs in their looks, links, lists at their levels, tables with spans, pictures, the page', () => {
  const doc = readOdt(odt());
  const [heading, para, centred] = doc.blocks;
  assert.equal(heading.heading, 1);
  assert.deepEqual(para.runs.map((r) => r.text ?? (r.tab ? '\t' : r.br ? '\n' : '')), ['Plain ', 'bold', ' and ', 'red italic', ' with a ', 'link', '.', '\t', 'tab', '\n', 'next line']);
  assert.equal(para.runs[1].bold, true);
  assert.deepEqual([para.runs[3].italic, para.runs[3].color, para.runs[3].size, para.runs[3].font], [true, '#C9211E', 14, 'Liberation Mono']);
  assert.equal(para.runs[5].link, 'https://example.org/');
  assert.equal(centred.align, 'center');
  const items = doc.blocks.filter((b) => b.list);
  assert.deepEqual(items.map((b) => [b.runs[0].text, b.list.style, b.list.level]), [['First bullet', 'L1', 0], ['Second bullet', 'L1', 0], ['Step one', 'L2', 0], ['Inside', 'L2', 1], ['Step two', 'L2', 0]]);
  assert.deepEqual(doc.lists.get('L2').map((l) => [l.kind, l.format, l.suffix]), [['number', '1', '.'], ['number', 'a', ')']]);
  const table = doc.blocks.find((b) => b.type === 'table');
  assert.deepEqual(table.columns.map((w) => Math.round(w)), [189, 113, 113]);
  assert.equal(table.rows[1][0].rowspan, 2);
  assert.equal(table.rows[2][1].colspan, 2);
  assert.equal(table.rows[0][0].fill, '#DDDDDD');
  assert.ok(doc.blocks.some((b) => b.runs?.some((r) => r.image?.href === 'Pictures/logo.png')));
  assert.equal(doc.blocks.at(-1).pageBreakBefore, true);
  assert.equal(Math.round(doc.page.width), 794);
});

test('written as a .docx: heading style, run looks, a link, numbering, the table grid and its merges, the picture, the page', () => {
  const parts = docxParts(odtToDocx(readOdt(odt())));
  const body = parts.get('word/document.xml').toString('utf8');
  assert.match(body, /<w:pStyle w:val="Heading1"\/><\/w:pPr><w:r><w:t xml:space="preserve">The showcase<\/w:t>/);
  assert.match(body, /<w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">bold<\/w:t>/);
  assert.match(body, /<w:rFonts w:ascii="Liberation Mono"[^>]*\/><w:i\/><w:color w:val="C9211E"\/><w:sz w:val="28"\/>/);
  assert.match(body, /<w:hyperlink r:id="rId\d+"><w:r><w:rPr><w:color w:val="0563C1"\/><w:u w:val="single"\/><\/w:rPr><w:t xml:space="preserve">link<\/w:t>/);
  assert.match(body, /<w:tab\/>[\s\S]*<w:br\/>/);
  assert.match(body, /<w:jc w:val="center"\/>/);
  assert.match(body, /<w:numPr><w:ilvl w:val="1"\/><w:numId w:val="2"\/><\/w:numPr><\/w:pPr><w:r><w:t xml:space="preserve">Inside/);
  const numbering = parts.get('word/numbering.xml').toString('utf8');
  assert.match(numbering, /<w:numFmt w:val="bullet"\/><w:lvlText w:val="•"\/>/);
  assert.match(numbering, /<w:numFmt w:val="decimal"\/><w:lvlText w:val="%1\."\/>/);
  assert.match(numbering, /<w:numFmt w:val="lowerLetter"\/><w:lvlText w:val="%2\)"\/>/);
  assert.match(body, /<w:tblGrid><w:gridCol w:w="2835"\/><w:gridCol w:w="1701"\/><w:gridCol w:w="1701"\/><\/w:tblGrid>/);
  assert.match(body, /<w:vMerge w:val="restart"\/>[\s\S]*North and South[\s\S]*<w:vMerge\/>[\s\S]*<w:gridSpan w:val="2"\/>[\s\S]*both quarters/);
  assert.match(body, /<w:shd w:val="clear" w:color="auto" w:fill="DDDDDD"\/>/);
  assert.match(body, /<w:tblBorders>/);
  assert.match(body, /<wp:extent cx="720000" cy="360000"\/><wp:docPr id="1" name="Logo"\/>/, 'the picture at 2 × 1 cm');
  assert.ok(parts.has('word/media/image1.png'));
  assert.match(body, /<w:pageBreakBefore\/>[\s\S]*After the page break/);
  assert.match(body, /<w:pgSz w:w="11906" w:h="16838"\/><w:pgMar w:top="1134" w:right="1417" w:bottom="1134" w:left="1417"/);
});

test('opened in the suite, the .odt is a document with its lists numbered, its table and its picture', () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-odt-')), 'showcase.odt');
  fs.writeFileSync(file, odt());
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const s = docs.open({ path: file });
  assert.equal(s.kind, 'doc');
  const m = docs.model({ id: s.id });
  const text = (b) => (b.runs || []).map((r) => r.text || '').join('');
  const index = (t) => m.blocks.findIndex((b) => text(b) === t);
  const label = (t) => m.listLabels?.[index(t)]?.label;
  assert.deepEqual(['First bullet', 'Step one', 'Inside', 'Step two'].map(label), ['•', '1.', 'a)', '2.'], 'bullets and numbers at their levels, as the list styles said');
  assert.ok(m.blocks.some((b) => /^t\d+:r\d+:c\d+$/.test(b.container || '')), 'a table');
  const logo = m.drawings.find((d) => d.name === 'Logo');
  assert.deepEqual([logo?.kind, Math.round(logo?.widthPx), Math.round(logo?.heightPx)], ['picture', 76, 38], 'the picture at its size');
});

test('saved to .odt a heading is not indented by its level', () => {
  const content = readZip(Buffer.from(writeOdt({ blocks: [{ type: 'heading', level: 2, runs: [{ text: 'Two' }] }] }))).entries.find((e) => e.name === 'content.xml').data.toString('utf8');
  assert.doesNotMatch(content, /fo:margin-left/);
});

/** What a reader of either file sees: each block's kind, words, looks, list, spans and pictures. */
const outline = (doc) => {
  const list = (b) => (b.list ? (() => { const lv = doc.lists.get(b.list.style)?.[b.list.level] || {}; return [b.list.level, lv.kind, lv.format ?? lv.char, lv.suffix ?? ''].join(' '); })() : null);
  const runs = (b) => b.runs.map((r) => (r.image ? `[${Math.round(r.image.width)}x${Math.round(r.image.height)}]` : r.tab ? '\\t' : r.br ? '\\n' : [r.text, r.bold && 'b', r.italic && 'i', r.color, r.size, r.font, r.link].filter(Boolean).join('|')));
  const block = (b) => (b.type === 'table'
    ? { table: b.columns.map((w) => Math.round(w || 0)), rows: b.rows.map((row) => row.map((c) => (c.covered ? '-' : `${c.blocks.map((x) => x.runs?.map((r) => r.text).join('')).join('/')}${c.rowspan > 1 ? '^' + c.rowspan : ''}${c.colspan > 1 ? '<' + c.colspan : ''}${c.fill ? ' ' + c.fill : ''}`)).join('|')) }
    : { h: b.heading || 0, align: b.align || null, list: list(b), brk: Boolean(b.pageBreakBefore), runs: runs(b) });
  return doc.blocks.map(block);
};

test('saved to .odt from the suite, an .odt keeps what it was made of: headings, run looks, lists and their labels, the table and its spans, the picture, the page', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-odt-save-'));
  const file = path.join(dir, 'showcase.odt');
  fs.writeFileSync(file, odt());
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const s = docs.open({ path: file });
  const out = path.join(dir, 'saved.odt');
  assert.equal(docs.save({ id: s.id, path: out }).format, 'odt');
  const before = readOdt(fs.readFileSync(file));
  const after = readOdt(fs.readFileSync(out));
  assert.deepEqual(outline(after), outline(before));
  const href = after.blocks.flatMap((b) => b.runs || []).find((r) => r.image)?.image.href;
  assert.deepEqual(Buffer.from(after.images.get(href)), PNG, 'the picture itself, in the archive the frame names');
  assert.deepEqual(['width', 'height', 'top', 'left'].map((k) => Math.round(after.page[k])), ['width', 'height', 'top', 'left'].map((k) => Math.round(before.page[k])));
  // Opened again, its lists count as they did.
  const again = docs.open({ path: out });
  const m = docs.model({ id: again.id });
  const text = (b) => (b.runs || []).map((r) => r.text || '').join('');
  const label = (t) => m.listLabels?.[m.blocks.findIndex((b) => text(b) === t)]?.label;
  assert.deepEqual(['First bullet', 'Step one', 'Inside', 'Step two'].map(label), ['•', '1.', 'a)', '2.']);
});

test('a Word document saved as .odt keeps its heading, its run looks and a merge down a column', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [
    { text: 'Quarterly figures', style: 'Heading1' },
    { runs: [{ text: 'Totals in ' }, { text: 'bold', bold: true }, { text: ' and ' }, { text: 'red', colour: 'C00000' }] },
    { table: { rows: [['Region', 'Q1'], ['North', '1'], ['', '2']] } },
    { text: 'After the table' },
  ] }));
  const blocks = view.render({ pages: false }).blocks;
  view.setSelection({ block: blocks.findIndex((b) => b.text === 'North'), offset: 0 }, { block: blocks.findIndex((b) => /:r2:c0$/.test(b.container || '')), offset: 0 });
  view.tableOp('mergeCells');
  const doc = readOdt(writeOdtDocument(readDocxDocument(view.save())));
  const [heading, para, table] = doc.blocks;
  assert.equal(heading.heading, 1);
  assert.deepEqual(heading.runs.map((r) => r.text), ['Quarterly figures']);
  assert.deepEqual(para.runs.map((r) => [r.text, Boolean(r.bold), r.color || null]), [['Totals in ', false, null], ['bold', true, null], [' and ', false, null], ['red', false, '#C00000']]);
  assert.equal(table.type, 'table');
  assert.deepEqual(table.rows.map((row) => row.map((c) => (c.covered ? '-' : c.blocks.map((b) => b.runs.map((r) => r.text).join('')).filter(Boolean).join('/') + (c.rowspan > 1 ? '^' + c.rowspan : ''))).join('|')), ['Region|Q1', 'North^2|1', '-|2']);
  assert.equal(doc.blocks.at(-1).runs[0].text, 'After the table');
});

test('a list that carries on after a paragraph is written to carry on its numbering, and read back as the same list', () => {
  const lists = new Map([['WWNum1', [{ kind: 'number', format: '1', prefix: '', suffix: '.', start: 1, display: 1, indent: 48, hanging: 24 }]]]);
  const item = (text) => ({ type: 'paragraph', heading: null, list: { id: 'n1', style: 'WWNum1', level: 0 }, runs: [{ text }] });
  const bytes = writeOdtDocument({ blocks: [item('One'), item('Two'), { type: 'paragraph', heading: null, runs: [{ text: 'Between' }] }, item('Three')], lists });
  const content = readZip(Buffer.from(bytes)).entries.find((e) => e.name === 'content.xml').data.toString('utf8');
  assert.match(content, /<text:list text:style-name="L_WWNum1" text:continue-numbering="true">/);
  const doc = readOdt(bytes);
  const ids = doc.blocks.filter((b) => b.list).map((b) => b.list.id);
  assert.equal(new Set(ids).size, 1, 'one list, its third item numbered on');
});
