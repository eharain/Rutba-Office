// Word: Table Layout — the contextual tab's operations, through the view and
// through the document service as the window sends them.
//
// Header rows (w:tblHeader) set from the top through a row and cleared from
// one; a cell's words at its top, centre or bottom (w:vAlign in its schema
// place, a tracked change's record left alone); columns shared out evenly;
// cells merged across a selection and split again; rows and columns put in
// and taken out — each named by `kind`, since `op` is the apply call's own.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const grid = [['Region', 'Q1', 'Q2'], ['North', '120', '135'], ['South', '140', '150'], ['East', '160', '170']];
const doc = () => openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Before' }, { table: { rows: grid } }, { text: 'After' }] }));
const xmlOf = (view) => view.doc.doc.xml;
const blocks = (view) => view.render({ pages: false }).blocks;
const at = (view, text) => blocks(view).findIndex((b) => b.text === text);

test('Repeat Header Rows: the rows from the top through the caret\'s are the header, and cleared from a row the header ends above it', () => {
  const view = doc();
  view.setSelection({ block: at(view, 'North'), offset: 0 });
  view.tableOp('headerRows', { on: true });
  assert.equal((xmlOf(view).match(/<w:tblHeader\/>/g) || []).length, 2, 'the first two rows');
  assert.deepEqual(['Region', 'Q1', 'North', '120', 'South'].map((t) => Boolean(blocks(view)[at(view, t)].rowHeader)), [true, true, true, true, false]);
  // A header row is laid out as one: the paginator repeats it.
  view.setSelection({ block: at(view, 'North'), offset: 0 });
  view.tableOp('headerRows', { on: false });
  assert.equal((xmlOf(view).match(/<w:tblHeader\/>/g) || []).length, 1, 'only the first row');
  view.setSelection({ block: at(view, 'Region'), offset: 0 });
  view.tableOp('headerRows', { on: false });
  assert.doesNotMatch(xmlOf(view), /tblHeader/);
  // A row's height stays before the header mark, as the schema orders trPr.
  view.tableOp('headerRows', { on: true });
  view.setTableRowHeight({ table: Number(/^t(\d+)/.exec(blocks(view)[at(view, 'Region')].container)[1]), row: 0, twips: 600 });
  assert.match(xmlOf(view), /<w:trPr><w:trHeight w:val="600" w:hRule="atLeast"\/><w:tblHeader\/><\/w:trPr>/);
});

test('Cell Alignment: each selected cell\'s words at its centre or bottom, top writing nothing, in the schema\'s place', () => {
  const view = doc();
  view.setSelection({ block: at(view, '120'), offset: 0 }, { block: at(view, '150'), offset: 0 });
  view.tableOp('cellVAlign', { v: 'center' });
  const aligned = blocks(view);
  assert.deepEqual(['120', '135', '140', '150', 'North', 'Q1'].map((t) => aligned[at(view, t)].cellVAlign ?? 'top'), ['center', 'center', 'center', 'center', 'top', 'top']);
  assert.equal((xmlOf(view).match(/<w:vAlign w:val="center"\/>/g) || []).length, 4);
  view.setSelection({ block: at(view, '120'), offset: 0 });
  view.tableOp('cellVAlign', { v: 'bottom' });
  assert.equal(blocks(view)[at(view, '120')].cellVAlign, 'bottom');
  view.tableOp('cellVAlign', { v: 'top' });
  assert.equal(blocks(view)[at(view, '120')].cellVAlign, undefined);
  assert.equal((xmlOf(view).match(/<w:vAlign\b/g) || []).length, 3);
  assert.throws(() => view.tableOp('cellVAlign', { v: 'sideways' }), /top, center or bottom/);
});

test('a cell\'s alignment goes before hideMark and leaves a tracked change\'s record of the old properties alone', () => {
  const view = doc();
  const engine = view.doc.doc;
  const tc = '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:shd w:val="clear" w:fill="EEEEEE"/><w:hideMark/><w:tcPrChange w:id="9" w:author="A"><w:tcPr><w:vAlign w:val="bottom"/></w:tcPr></w:tcPrChange></w:tcPr><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>';
  engine.xml = engine.xml.replace(/<w:tc>(?:(?!<w:tc>)[\s\S])*?North[\s\S]*?<\/w:tc>/, tc);
  const table = Number(/^t(\d+)/.exec(blocks(view)[at(view, 'North')].container)[1]);
  engine.setTableCellVAlign(table, 1, 0, 'center');
  assert.match(engine.xml, /<w:shd w:val="clear" w:fill="EEEEEE"\/><w:vAlign w:val="center"\/><w:hideMark\/><w:tcPrChange w:id="9" w:author="A"><w:tcPr><w:vAlign w:val="bottom"\/><\/w:tcPr><\/w:tcPrChange><\/w:tcPr>/);
  engine.setTableCellVAlign(table, 1, 0, 'top');
  assert.match(engine.xml, /<w:shd w:val="clear" w:fill="EEEEEE"\/><w:hideMark\/><w:tcPrChange w:id="9" w:author="A"><w:tcPr><w:vAlign w:val="bottom"\/><\/w:tcPr><\/w:tcPrChange><\/w:tcPr>/);
});

test('Distribute Columns: every column the same width, the table as wide as it was', () => {
  const view = doc();
  const table = Number(/^t(\d+)/.exec(blocks(view)[at(view, 'Region')].container)[1]);
  view.setTableColumnWidths({ table, widths: { 0: 3000, 1: 1500, 2: 1500 } });
  const before = blocks(view)[at(view, 'Region')].gridPx.reduce((a, b) => a + b, 0);
  view.setSelection({ block: at(view, 'Q1'), offset: 0 });
  view.tableOp('distributeColumns');
  const after = blocks(view)[at(view, 'Region')].gridPx;
  assert.ok(after.every((w) => Math.abs(w - after[0]) < 0.5), after.join(', '));
  assert.ok(Math.abs(after.reduce((a, b) => a + b, 0) - before) < 1.5, `${before} → ${after.reduce((a, b) => a + b, 0)}`);
});

test('through the document service, each table operation is named by kind: rows, columns, merging a selection and splitting it', () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-table-layout-')), 'grid.docx');
  fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [{ text: 'Before' }, { table: { rows: grid } }, { text: 'After' }] }));
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const s = docs.open({ path: file });
  const model = () => docs.model({ id: s.id });
  const index = (t) => model().blocks.findIndex((b) => b.text === t);
  const rows = () => new Set(model().blocks.map((b) => /:r(\d+):/.exec(b.container || '')?.[1]).filter(Boolean)).size;
  const caret = (t) => ({ op: 'setSelection', anchor: { block: index(t), offset: 0 }, focus: { block: index(t), offset: 0 } });

  docs.apply({ id: s.id, ops: [caret('North'), { op: 'tableOp', kind: 'insertRowBelow' }] });
  assert.equal(rows(), 5, 'a row under North');
  docs.apply({ id: s.id, ops: [caret('North'), { op: 'tableOp', kind: 'deleteRow' }] });
  assert.equal(rows(), 4, 'North\'s row gone');
  assert.equal(index('North'), -1);
  docs.apply({ id: s.id, ops: [caret('Q1'), { op: 'tableOp', kind: 'insertColumnRight' }] });
  assert.equal(new Set(model().blocks.filter((b) => /:r0:/.test(b.container || '')).map((b) => b.container)).size, 4, 'four columns');
  docs.apply({ id: s.id, ops: [caret('Q1'), { op: 'tableOp', kind: 'deleteColumn' }] });
  assert.equal(index('Q1'), -1);

  // Merge from South down to East: one cell over two rows, its words kept.
  docs.apply({ id: s.id, ops: [{ op: 'setSelection', anchor: { block: index('South'), offset: 0 }, focus: { block: index('East'), offset: 0 } }, { op: 'tableOp', kind: 'mergeCells' }] });
  const merged = model().blocks;
  assert.ok(merged.some((b) => b.hiddenCell), 'the lower place a continuation');
  assert.ok(merged.some((b) => b.text === 'South') && merged.some((b) => b.text === 'East'), 'both words kept');
  docs.apply({ id: s.id, ops: [caret('South'), { op: 'tableOp', kind: 'splitCell' }] });
  assert.ok(!model().blocks.some((b) => b.hiddenCell), 'split again');

  docs.apply({ id: s.id, ops: [caret('Region'), { op: 'tableOp', kind: 'headerRows', arg: { on: true } }] });
  assert.equal(model().blocks[index('Region')].rowHeader, true);
  assert.throws(() => docs.apply({ id: s.id, ops: [caret('Before'), { op: 'tableOp', kind: 'insertRowBelow' }] }), /not in a table/);
});

test('a table\'s lines are its style\'s along basedOn, its own sides over them; a cell\'s own border and shading ride its paragraphs', () => {
  const view = doc();
  const engine = view.doc.doc;
  const stylesXml = engine.pkg.text('word/styles.xml').replace('</w:styles>',
    '<w:style w:type="table" w:styleId="Base"><w:name w:val="Base"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="8" w:color="000000"/><w:insideH w:val="single" w:sz="4" w:color="808080"/></w:tblBorders></w:tblPr><w:tcPr><w:shd w:val="clear" w:fill="EEEEEE"/></w:tcPr></w:style>'
    + '<w:style w:type="table" w:styleId="Child"><w:name w:val="Child"/><w:basedOn w:val="Base"/><w:tblPr><w:tblBorders><w:insideH w:val="double" w:sz="4" w:color="0000FF"/></w:tblBorders></w:tblPr>'
    + '<w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:val="clear" w:fill="FF0000"/></w:tcPr></w:tblStylePr></w:style></w:styles>');
  engine.pkg.write_('word/styles.xml', stylesXml);
  assert.deepEqual(engine._tableStyleLook('Child'), {
    borders: { top: { style: 'single', widthPx: 4 / 3, colour: '#000000' }, insideH: { style: 'double', widthPx: 2 / 3, colour: '#0000ff' } },
    fill: '#eeeeee',
  }, 'the child\'s inside line over its base\'s, the base\'s top and shading, a conditional part left out');
  // The table names the style, sets its own bottom, and one cell its own border and shading.
  engine.xml = engine.xml
    .replace(/<w:tblPr>/, '<w:tblPr><w:tblStyle w:val="Child"/>')
    .replace(/<w:tblBorders>[\s\S]*?<\/w:tblBorders>/, '<w:tblBorders><w:bottom w:val="single" w:sz="16" w:color="00FF00"/></w:tblBorders>')
    .replace(/(<w:tc><w:tcPr>)((?:(?!<\/w:tcPr>)[\s\S])*?)(<\/w:tcPr>(?:(?!<\/w:tc>)[\s\S])*?Q1)/, '$1$2<w:tcBorders><w:left w:val="dotted" w:sz="8" w:color="123456"/></w:tcBorders><w:shd w:val="clear" w:fill="ABCDEF"/>$3');
  const blocks = engine.editParagraphs();
  const region = blocks.find((p) => p.text === 'Region');
  const q1 = blocks.find((p) => p.text === 'Q1');
  assert.equal(region.tableBorders.insideH.style, 'double', 'the style\'s inside line');
  assert.equal(region.tableBorders.top.colour, '#000000', 'the base style\'s top');
  assert.ok(region.tableBorders.bottom, 'the table\'s own bottom');
  assert.equal(region.cellFill, '#eeeeee', 'the style\'s shading where the cell gives none');
  assert.deepEqual(q1.cellBorders, { left: { style: 'dotted', widthPx: 4 / 3, colour: '#123456' } });
  assert.equal(q1.cellFill, '#abcdef');
  // And the view passes them on to the page.
  const shown = view.render({ pages: false }).blocks.find((b) => b.text === 'Q1');
  assert.equal(shown.cellFill, '#abcdef');
  assert.ok(shown.tableBorders && shown.cellBorders);
});
