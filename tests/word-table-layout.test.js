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
import { tint, tableStyleXml } from '@rutba/ooxml/table-styles';

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
    parts: { firstRow: { fill: '#ff0000' } },
    rowBand: 1,
    colBand: 1,
  }, 'the child\'s inside line over its base\'s, the base\'s top and shading, the header row\'s part on its own');
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
  assert.equal(region.cellFill, undefined, 'a cell\'s fill is its own');
  assert.equal(region.tableStyle.fill, '#eeeeee', 'the style\'s shading, for the page to lay under the cell');
  assert.deepEqual(region.tableStyle.parts.firstRow, { fill: '#ff0000' });
  assert.deepEqual(region.tableStyle.look, { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true }, 'tblLook\'s switches');
  assert.deepEqual(q1.cellBorders, { left: { style: 'dotted', widthPx: 4 / 3, colour: '#123456' } });
  assert.equal(q1.cellFill, '#abcdef');
  // And the view passes them on to the page.
  const shown = view.render({ pages: false }).blocks.find((b) => b.text === 'Q1');
  assert.equal(shown.cellFill, '#abcdef');
  assert.ok(shown.tableBorders && shown.cellBorders);
});

test('Table Design: a built-in style written in the theme\'s colours, its options as tblLook, a cell shaded, the style taken away', () => {
  // Word tints a theme colour in HSL: Office 2013's accent 1 and Office 2023's give Word's own values.
  assert.deepEqual([tint('4472C4', 0.6), tint('4472C4', 0.2), tint('156082', 0.6), tint('156082', 0.2)], ['8EAADB', 'D9E2F3', '45B0E1', 'C1E4F5']);
  assert.equal(tableStyleXml('NoSuchStyle'), null);
  const view = doc();
  view.setSelection({ block: at(view, 'North'), offset: 0 });
  view.tableOp('style', { id: 'GridTable4-Accent1' });
  const engine = view.doc.doc;
  const styles = engine.pkg.text('word/styles.xml');
  assert.match(styles, /<w:style w:type="table" w:styleId="GridTable4-Accent1">[\s\S]*?<w:tblStylePr w:type="firstRow"><w:rPr><w:b\/><w:bCs\/><w:color w:val="FFFFFF"\/><\/w:rPr><w:tblPr\/><w:tcPr><w:tcBorders>[\s\S]*?<w:shd w:val="clear" w:color="auto" w:fill="4472C4"\/>/);
  assert.match(styles, /w:styleId="TableNormal"/, 'Normal Table, which it is based on');
  assert.match(xmlOf(view), /<w:tblPr><w:tblStyle w:val="GridTable4-Accent1"\/>/);
  assert.doesNotMatch(/<w:tblPr>[\s\S]*?<\/w:tblPr>/.exec(xmlOf(view))[0], /tblBorders/, 'its own lines give way to the style\'s');
  assert.match(xmlOf(view), /<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"\/>/);
  const north = blocks(view)[at(view, 'North')];
  assert.deepEqual([north.tableStyle.id, north.tableStyle.parts.firstRow.fill, north.tableStyle.parts.band1Horz.fill], ['GridTable4-Accent1', '#4472c4', '#d9e2f3']);
  // Applied twice, the definition is written once.
  view.tableOp('style', { id: 'GridTable4-Accent1' });
  assert.equal((engine.pkg.text('word/styles.xml').match(/w:styleId="GridTable4-Accent1"/g) || []).length, 1);
  view.tableOp('styleOptions', { look: { firstRow: true, lastRow: true, firstColumn: false, lastColumn: false, noHBand: true, noVBand: true } });
  assert.match(xmlOf(view), /<w:tblLook w:val="0660" w:firstRow="1" w:lastRow="1" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"\/>/);
  view.setSelection({ block: at(view, '120'), offset: 0 }, { block: at(view, '150'), offset: 0 });
  view.tableOp('shading', { fill: 'FFF2CC' });
  assert.equal((xmlOf(view).match(/<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"\/>/g) || []).length, 4, 'every selected cell');
  view.tableOp('shading', { fill: null });
  assert.doesNotMatch(xmlOf(view), /FFF2CC/);
  assert.throws(() => view.tableOp('shading', { fill: 'blue' }), /colour like/);
  view.tableOp('style', { id: null });
  assert.doesNotMatch(xmlOf(view), /<w:tblStyle\b/);
});

test('Table Design → Borders: outside, inside, all and none over the selected cells, in the pen\'s line, the other sides left alone', () => {
  const view = doc();
  const engine = view.doc.doc;
  const sides = (text) => {
    const cell = new RegExp(`<w:tc>((?:(?!<w:tc>)[\\s\\S])*?)<w:t[^>]*>${text}<`).exec(engine.xml)?.[1] || '';
    return Object.fromEntries([...(/<w:tcBorders>([\s\S]*?)<\/w:tcBorders>/.exec(cell)?.[1] || '').matchAll(/<w:(\w+) w:val="(\w+)"(?: w:sz="(\d+)")?/g)].map((m) => [m[1], m[3] ? `${m[2]} ${m[3]}` : m[2]]));
  };
  const pen = { val: 'double', sz: 12, color: 'C00000' };
  view.setSelection({ block: at(view, '120'), offset: 0 }, { block: at(view, '150'), offset: 0 });
  view.tableOp('borders', { kind: 'outside', pen });
  assert.deepEqual([sides('120'), sides('135'), sides('140'), sides('150')], [
    { top: 'double 12', left: 'double 12' }, { top: 'double 12', right: 'double 12' }, { left: 'double 12', bottom: 'double 12' }, { bottom: 'double 12', right: 'double 12' },
  ]);
  view.tableOp('borders', { kind: 'inside', pen: { val: 'dotted', sz: 4 } });
  assert.deepEqual(sides('120'), { top: 'double 12', left: 'double 12', bottom: 'dotted 4', right: 'dotted 4' }, 'the outside kept, the inside added');
  view.setSelection({ block: at(view, 'North'), offset: 0 });
  view.tableOp('borders', { kind: 'none' });
  assert.deepEqual(sides('North'), { top: 'nil', left: 'nil', bottom: 'nil', right: 'nil' }, 'none is written so the table\'s own lines do not show through');
  view.tableOp('borders', { kind: 'bottom', pen });
  assert.deepEqual(sides('North'), { top: 'nil', left: 'nil', bottom: 'double 12', right: 'nil' });
  assert.throws(() => view.tableOp('borders', { kind: 'diagonal', pen }), /unknown border choice/);
  // The page draws them as the cell's own lines.
  assert.equal(blocks(view)[at(view, '120')].cellBorders.top.style, 'double');
});

test('Table Layout → Sort: rows by the caret\'s column, as numbers or as words, rising or falling, a header row kept on top', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ table: { rows: [['Region', 'Sales'], ['north', '1,200'], ['East', '95'], ['South', ''], ['west', '1000']] } }] }));
  const column = (c) => { const rows = new Map(); for (const b of blocks(view)) { const m = /:r(\d+):c(\d+)$/.exec(b.container || ''); if (m && Number(m[2]) === c) rows.set(Number(m[1]), b.text); } return [...rows.keys()].sort((a, b) => a - b).map((r) => rows.get(r)); };
  view.setSelection({ block: at(view, '95'), offset: 0 });
  view.tableOp('sort', { descending: false, header: true });
  assert.deepEqual(column(1), ['Sales', '95', '1000', '1,200', ''], 'as numbers, the empty cell last, the header kept');
  assert.deepEqual(column(0), ['Region', 'East', 'west', 'north', 'South'], 'each row moved whole');
  view.setSelection({ block: at(view, 'East'), offset: 0 });
  view.tableOp('sort', { descending: true, header: true });
  assert.deepEqual(column(0), ['Region', 'west', 'South', 'north', 'East'], 'as words, in the language\'s order, falling');
  // A header row the table marks needs no telling.
  view.setSelection({ block: at(view, 'Region'), offset: 0 });
  view.tableOp('headerRows', { on: true });
  view.setSelection({ block: at(view, 'East'), offset: 0 });
  view.tableOp('sort', { descending: false });
  assert.deepEqual(column(0), ['Region', 'East', 'north', 'South', 'west']);
  // Merged down a column, the rows cannot be put in another order.
  view.setSelection({ block: at(view, 'north'), offset: 0 }, { block: at(view, 'South'), offset: 0 });
  view.tableOp('mergeCells');
  view.setSelection({ block: at(view, 'East'), offset: 0 });
  assert.throws(() => view.tableOp('sort', {}), /merged down a column/);
});

test('Table Layout → Convert to Text and Insert → Convert Text to Table: rows and paragraphs into each other, the words\' looks kept going out', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Before' }, { table: { rows: [['Region', 'Q1'], ['North', '120']] } }, { text: 'After' }] }));
  view.doc.doc.xml = view.doc.doc.xml.replace(/<w:r>(<w:t[^>]*>North<\/w:t>)/, '<w:r><w:rPr><w:b/></w:rPr>$1');
  view._invalidate?.();
  view.setSelection({ block: at(view, 'North'), offset: 0 });
  view.tableOp('toText', { separator: 'tab' });
  const texts = blocks(view).map((b) => b.text);
  assert.deepEqual(texts, ['Before', 'Region\tQ1', 'North\t120', 'After']);
  assert.ok(!blocks(view).some((b) => b.container), 'the table gone');
  assert.match(xmlOf(view), /<w:r><w:rPr><w:b\/><\/w:rPr><w:t[^>]*>North<\/w:t><\/w:r><w:r><w:tab\/><\/w:r>/, 'North still bold, a tab after it');
  // And back: the two paragraphs split at their tabs into a table of two columns.
  view.setSelection({ block: at(view, 'Region\tQ1'), offset: 0 }, { block: at(view, 'North\t120'), offset: 3 });
  view.textToTable({ separator: 'tab' });
  const cells = blocks(view).filter((b) => b.container).map((b) => `${/:r(\d+):c(\d+)$/.exec(b.container).slice(1).join('.')}=${b.text}`);
  assert.deepEqual(cells, ['0.0=Region', '0.1=Q1', '1.0=North', '1.1=120']);
  assert.deepEqual(blocks(view).filter((b) => !b.container).map((b) => b.text), ['Before', 'After']);
  // Commas, and each cell its own paragraph.
  view.setSelection({ block: at(view, 'Region'), offset: 0 });
  view.tableOp('toText', { separator: 'comma' });
  assert.deepEqual(blocks(view).map((b) => b.text), ['Before', 'Region, Q1', 'North, 120', 'After']);
  view.setSelection({ block: at(view, 'Region, Q1'), offset: 0 }, { block: at(view, 'North, 120'), offset: 1 });
  view.textToTable({ separator: 'comma' });
  view.setSelection({ block: at(view, 'Q1'), offset: 0 });
  view.tableOp('toText', { separator: 'paragraph' });
  assert.deepEqual(blocks(view).map((b) => b.text), ['Before', 'Region', 'Q1', 'North', '120', 'After']);
  view.setSelection({ block: at(view, 'Before'), offset: 0 });
  view.textToTable({});
  assert.throws(() => { view.setSelection({ block: at(view, 'Before'), offset: 0 }); view.textToTable({}); }, /in a table already/);
});

test('Table Layout → Split Table and AutoFit: two tables of the same make, and the table to its window, its contents or fixed widths', () => {
  const view = doc();
  view.setSelection({ block: at(view, 'South'), offset: 0 });
  view.tableOp('splitTable');
  const tables = new Set(blocks(view).filter((b) => b.container).map((b) => /^t(\d+)/.exec(b.container)[1]));
  assert.equal(tables.size, 2, 'two tables');
  assert.deepEqual(blocks(view).filter((b) => !b.container).map((b) => b.text), ['Before', '', 'After'], 'an empty paragraph between them, as Word leaves one');
  assert.equal((xmlOf(view).match(/<w:tblGrid>/g) || []).length, 2, 'each with the grid');
  assert.match(blocks(view)[at(view, 'South')].container, /:r0:c0$/, 'South heads the second');
  view.setSelection({ block: at(view, 'Region'), offset: 0 });
  assert.throws(() => view.tableOp('splitTable'), /not the first/);
  view.tableOp('autoFit', { mode: 'window' });
  assert.match(/<w:tblPr>[\s\S]*?<\/w:tblPr>/.exec(xmlOf(view))[0], /<w:tblW w:w="5000" w:type="pct"\/>/);
  view.tableOp('autoFit', { mode: 'fixed' });
  assert.match(/<w:tblPr>[\s\S]*?<\/w:tblPr>/.exec(xmlOf(view))[0], /<w:tblW w:w="0" w:type="auto"\/>[\s\S]*<w:tblLayout w:type="fixed"\/>/);
  view.tableOp('autoFit', { mode: 'contents' });
  assert.match(/<w:tblPr>[\s\S]*?<\/w:tblPr>/.exec(xmlOf(view))[0], /<w:tblLayout w:type="autofit"\/>/);
  assert.equal((/<w:tblPr>[\s\S]*?<\/w:tblPr>/.exec(xmlOf(view))[0].match(/<w:tblW\b/g) || []).length, 1, 'one width, rewritten');
});

test('Table Layout → Cell Margins, Text Direction and Align Table: written in their schema places and read back for the page', () => {
  const view = doc();
  view.setSelection({ block: at(view, 'North'), offset: 0 });
  view.tableOp('cellMargins', { margins: { top: 72, left: 216, bottom: 72, right: 216 } });
  const tblPr = () => /<w:tblPr>[\s\S]*?<\/w:tblPr>/.exec(xmlOf(view))[0];
  assert.match(tblPr(), /<w:tblCellMar><w:top w:w="72" w:type="dxa"\/><w:left w:w="216" w:type="dxa"\/><w:bottom w:w="72" w:type="dxa"\/><w:right w:w="216" w:type="dxa"\/><\/w:tblCellMar>/);
  assert.deepEqual(blocks(view)[at(view, 'North')].tableLook.cellMarginPx, { left: 14.4, right: 14.4, top: 4.8, bottom: 4.8 }, 'the page pads every cell by them');
  view.tableOp('align', { align: 'center' });
  assert.match(tblPr(), /<w:tblW\b[^>]*\/><w:jc w:val="center"\/>/, 'after the width, as the schema has it');
  assert.equal(blocks(view)[at(view, 'North')].tableLook.align, 'center');
  view.tableOp('align', { align: 'left' });
  assert.doesNotMatch(tblPr(), /<w:jc\b/);
  view.setSelection({ block: at(view, 'North'), offset: 0 }, { block: at(view, 'South'), offset: 0 });
  view.tableOp('textDirection', { dir: 'up' });
  assert.equal((xmlOf(view).match(/<w:textDirection w:val="btLr"\/>/g) || []).length, 2, 'both selected cells turned to read upwards');
  assert.deepEqual(['North', 'South', 'Region'].map((t) => blocks(view)[at(view, t)].cellDirection ?? null), ['up', 'up', null]);
  view.setSelection({ block: at(view, 'North'), offset: 0 });
  view.tableOp('textDirection', { dir: 'down' });
  assert.equal(blocks(view)[at(view, 'North')].cellDirection, 'down');
  view.tableOp('textDirection', { dir: null });
  assert.equal(blocks(view)[at(view, 'North')].cellDirection, undefined);
  assert.throws(() => view.tableOp('align', { align: 'middle' }), /left, the centre or the right/);
});

test('through the document service, Convert Text to Table is an operation of its own', () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-text-table-')), 'lines.docx');
  fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [{ text: 'a\tb' }, { text: 'c\td' }] }));
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const s = docs.open({ path: file });
  docs.apply({ id: s.id, ops: [{ op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 1, offset: 1 } }, { op: 'textToTable', separator: 'tab' }] });
  assert.deepEqual(docs.model({ id: s.id }).blocks.filter((b) => b.container).map((b) => b.text), ['a', 'b', 'c', 'd']);
});
