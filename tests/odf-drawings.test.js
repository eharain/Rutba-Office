// OpenDocument files, drawn as they were made — the converters' half.
//
// An .odp used to arrive as its words alone, poured into title and body
// placeholders: its pictures, its shapes, its tables and where each stood
// were dropped. Now its slides are its own size and hold each drawing
// where it stood: text boxes, pictures, shapes in their fill and outline
// with their words, lines and tables, and the notes. An .ods keeps its
// column widths, the rows and columns it hides and the panes it freezes.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { readZip, writeZip, ZipEntry } from '@rutba/ooxml/zip';
import { writeOdp, writeOds } from '../packages/office-formats/src/odf-write.js';
import { readOdf, lengthPx } from '../packages/office-formats/src/odf.js';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64');

/** An archive the ODF writer made, with parts swapped for `edits` (name → edit(xml)) and `added` parts. */
function rework(bytes, edits = {}, added = {}) {
  const { entries } = readZip(Buffer.from(bytes));
  for (const e of entries) if (edits[e.name]) e.data = Buffer.from(edits[e.name](e.data.toString('utf8')), 'utf8');
  const extra = Object.entries(added).map(([name, xml]) => {
    const z = new ZipEntry({ name, method: 8, crc: 0, compressedSize: 0, uncompressedSize: 0, compressed: Buffer.alloc(0), dosTime: 0, dosDate: 0x2821, flags: 0, externalAttrs: 0, comment: Buffer.alloc(0) });
    z.data = Buffer.from(xml, 'utf8');
    return z;
  });
  return writeZip([...entries, ...extra]);
}

const open = (bytes, ext) => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-odf-')), 'file.' + ext);
  fs.writeFileSync(file, bytes);
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  return { docs, id: docs.open({ path: file }).id };
};

test('ODF lengths read in every unit', () => {
  assert.equal(Math.round(lengthPx('2.54cm')), 96);
  assert.equal(Math.round(lengthPx('25.4mm')), 96);
  assert.equal(lengthPx('1in'), 96);
  assert.equal(lengthPx('72pt'), 96);
  assert.equal(lengthPx('nonsense'), null);
});

test('an .odp keeps its pictures, shapes, lines, tables and positions, on slides of its own size', () => {
  const base = writeOdp({
    title: 'Plan', size: { width: 1056, height: 794 },
    slides: [{ shapes: [{ kind: 'text', placeholder: 'title', paragraphs: ['Quarterly plan'], x: 48, y: 30, w: 900, h: 90 }, { kind: 'picture', data: PNG, contentType: 'image/png', x: 600, y: 300, w: 200, h: 150 }], notes: 'Start with the numbers.' }],
  });
  const drawings = '<draw:custom-shape draw:name="Callout" draw:style-name="gr1" svg:x="2cm" svg:y="5cm" svg:width="6cm" svg:height="3cm"><text:p>Ship it</text:p><draw:enhanced-geometry draw:type="round-rectangle"/></draw:custom-shape>'
    + '<draw:ellipse draw:style-name="gr2" svg:x="10cm" svg:y="5cm" svg:width="3cm" svg:height="3cm"/>'
    + '<draw:line svg:x1="1cm" svg:y1="15cm" svg:x2="20cm" svg:y2="15cm"/>'
    + '<draw:frame draw:name="Figures" svg:x="2cm" svg:y="10cm" svg:width="10cm" svg:height="3cm"><table:table><table:table-column table:number-columns-repeated="2"/><table:table-row><table:table-cell><text:p>Region</text:p></table:table-cell><table:table-cell><text:p>Sales</text:p></table:table-cell></table:table-row><table:table-row><table:table-cell><text:p>North</text:p></table:table-cell><table:table-cell><text:p>120</text:p></table:table-cell></table:table-row></table:table></draw:frame>';
  const styles = '<style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#ff0000" draw:stroke="none"/></style:style><style:style style:name="gr2" style:family="graphic" style:parent-style-name="named"><style:graphic-properties draw:fill="none"/></style:style>';
  const bytes = rework(base, {
    'content.xml': (xml) => xml.replace('</draw:page>', () => drawings + '</draw:page>').replace(/<office:automatic-styles>/, () => '<office:automatic-styles>' + styles),
    'styles.xml': (xml) => xml.replace('<office:styles/>', () => '<office:styles><style:style style:name="named" style:family="graphic"><style:graphic-properties svg:stroke-color="#00aa00"/></style:style></office:styles>'),
  });
  const odf = readOdf(bytes);
  assert.deepEqual(odf.slides[0].shapes.map((s) => s.type), ['text', 'image', 'shape', 'shape', 'line', 'table']);
  assert.equal(odf.slides[0].shapes[2].fill, '#ff0000');
  assert.equal(odf.slides[0].shapes[3].line, '#00aa00', 'a parent style fills in what the shape\'s own does not say');

  const { docs, id } = open(bytes, 'odp');
  const m = docs.model({ id, slide: 0 });
  assert.equal(Math.round(m.size.width), 1056);
  assert.equal(Math.round(m.size.height), 794, 'the slides are the presentation\'s own size');
  const shapes = m.slide.shapes;
  const words = (s) => (s.text?.paragraphs || []).map((p) => p.plain).join('|');
  const title = shapes.find((s) => words(s) === 'Quarterly plan');
  assert.deepEqual([Math.round(title.geometry.x), Math.round(title.geometry.y)], [48, 30], 'where it stood');
  assert.equal(shapes.filter((s) => s.kind === 'picture').length, 1);
  const callout = shapes.find((s) => words(s) === 'Ship it');
  assert.equal(callout.fill?.color?.toLowerCase(), '#ff0000');
  assert.equal(callout.line?.type, 'none');
  assert.ok(shapes.some((s) => s.kind === 'table'), 'the table');
  assert.ok(shapes.some((s) => /Connector/.test(s.name || '')), 'the line');
  assert.equal(m.slide.notes, 'Start with the numbers.');
});

test('an .ods keeps its column widths, hidden rows and columns, and frozen panes', () => {
  const base = writeOds({ title: 'T', sheets: [{ name: 'Data', rows: [] }] });
  const table = '<table:table table:name="Data"><table:table-column table:style-name="co1"/><table:table-column table:visibility="collapse"/><table:table-column/>'
    + ['A,B,C', '1,2,3', '4,5,6'].map((line, r) => `<table:table-row${r === 1 ? ' table:visibility="collapse"' : ''}>` + line.split(',').map((v) => `<table:table-cell office:value-type="string"><text:p>${v}</text:p></table:table-cell>`).join('') + '</table:table-row>').join('')
    + '</table:table>';
  const settings = '<?xml version="1.0" encoding="UTF-8"?><office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"><office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views"><config:config-item-map-entry><config:config-item-map-named config:name="Tables"><config:config-item-map-entry config:name="Data">'
    + '<config:config-item config:name="HorizontalSplitMode" config:type="short">2</config:config-item><config:config-item config:name="VerticalSplitMode" config:type="short">2</config:config-item><config:config-item config:name="HorizontalSplitPosition" config:type="int">1</config:config-item><config:config-item config:name="VerticalSplitPosition" config:type="int">1</config:config-item>'
    + '</config:config-item-map-entry></config:config-item-map-named></config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set></office:settings></office:document-settings>';
  const bytes = rework(base, {
    'content.xml': (xml) => xml.replace(/<table:table\b[\s\S]*<\/table:table>/, () => table)
      .replace(/<office:automatic-styles\/>|<office:automatic-styles>/, () => '<office:automatic-styles><style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="3cm"/></style:style>')
      .replace(/(<office:automatic-styles>[\s\S]*?)(?=<office:body)/, (m) => (m.includes('</office:automatic-styles>') ? m : m + '</office:automatic-styles>')),
  }, { 'settings.xml': settings });
  const sheet = readOdf(bytes).sheets[0];
  assert.deepEqual(sheet.hiddenCols, [1]);
  assert.deepEqual(sheet.hiddenRows, [1]);
  assert.equal(Math.round(sheet.widths[0]), 113);
  assert.deepEqual(sheet.frozen, { rows: 1, cols: 1 });

  const { docs, id } = open(bytes, 'ods');
  const m = docs.model({ id });
  assert.deepEqual(m.frozen, { rows: 1, cols: 1 });
  assert.deepEqual(m.columns.slice(0, 2).map((c) => c.name), ['A', 'C'], 'column B hidden');
  assert.ok(m.columns[0].width > 100, 'column A as wide as it was: ' + m.columns[0].width);
  assert.deepEqual(m.rows.slice(0, 2).map((r) => r.index), [0, 2], 'row 2 hidden');
});
