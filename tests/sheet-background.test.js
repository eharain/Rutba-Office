/**
 * Page Layout → Background: a picture tiled behind a sheet's cells, written
 * as Excel writes it — `<picture r:id>` in the sheet's tail, in schema
 * order, an image relationship and the image part — and never printed.
 * Delete Background takes all three away. Neither is undone, as in Excel.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';
import { printHtml } from '@rutba/sheet-view/print';

// A 2×2 PNG, red and white.
const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAFklEQVR4nGP4z8DAwMDAwMDAwMDAAAANAAH/kPjvAAAAAElFTkSuQmCC', 'base64');
const open = () => new SheetView(buildXlsx({ sheets: [{ name: 'Plan', rows: [['Week', 'Goal'], [1, 'Draft'], [2, 'Review']] }, { name: 'Other', rows: [['x']] }] }));

test('Background writes <picture r:id>, an image relationship and the image part, as Excel does', () => {
  const view = open();
  view.setCell(3, 0, 'kept');
  const part = view.setBackground({ contentType: 'image/png', data: PNG });
  assert.match(part, /^xl\/media\/image\d+\.png$/);
  const pkg = OoxmlPackage.read(view.save());
  const sheet = pkg.text('xl/worksheets/sheet1.xml');
  const m = /<picture r:id="(rId\d+)"[^>]*\/>/.exec(sheet);
  assert.ok(m, 'the picture element');
  assert.ok(sheet.indexOf('<picture') > sheet.indexOf('</sheetData>'), 'in the tail');
  const rel = pkg.rels('xl/worksheets/sheet1.xml').find((r) => r.Id === m[1]);
  assert.match(rel.Type, /\/relationships\/image$/);
  assert.equal(OoxmlPackage.resolveTarget('xl/worksheets/sheet1.xml', rel.Target), part);
  assert.deepEqual(pkg.read(part), PNG, 'the picture itself');
  assert.match(pkg.text('[Content_Types].xml'), /Extension="png"/);
  const reopened = new SheetView(view.save());
  assert.equal(reopened.sheetBackground(), part, 'read back');
  assert.equal(reopened.render().background, part, 'and in the frame');
  reopened.selectSheet('Other');
  assert.equal(reopened.render().background, null, 'the other sheet has none');
  assert.equal(view.displayValue(3, 0).text, 'kept', 'the cells are untouched');
});

test('the picture lands where the schema puts it — after the drawing, before the table parts', () => {
  const view = open();
  view.insertShape({ geometry: 'rect', text: '' });
  view.select(0, 0);
  view.select(2, 1, { extend: true });
  view.formatAsTable({ style: 'TableStyleLight9' });
  view.setBackground({ contentType: 'image/png', data: PNG });
  const sheet = OoxmlPackage.read(view.save()).text('xl/worksheets/sheet1.xml');
  const at = (tag) => sheet.indexOf('<' + tag);
  assert.ok(at('drawing') > 0 && at('picture') > at('drawing'), 'after the drawing');
  assert.ok(at('tableParts') > at('picture'), 'before the table parts');
});

test('a background is not printed', () => {
  const view = open();
  view.setBackground({ contentType: 'image/png', data: PNG });
  const html = printHtml(view, {});
  assert.doesNotMatch(html, /<img|background-image|url\(/, 'no picture on paper, as in Excel');
  assert.match(html, /Draft/);
});

test('a new background replaces the old one, and Delete Background takes it away altogether', () => {
  const view = open();
  const first = view.setBackground({ contentType: 'image/png', data: PNG });
  const second = view.setBackground({ contentType: 'image/png', data: PNG });
  let pkg = OoxmlPackage.read(view.save());
  assert.ok(!pkg.has(first) || first === second, 'the first image is gone');
  assert.equal((pkg.text('xl/worksheets/sheet1.xml').match(/<picture\b/g) || []).length, 1, 'one picture element');
  view.setCell(5, 5, 'typed');
  view.deleteBackground();
  assert.equal(view.canUndo, false, 'not undone, as in Excel');
  pkg = OoxmlPackage.read(view.save());
  assert.doesNotMatch(pkg.text('xl/worksheets/sheet1.xml'), /<picture\b/);
  assert.ok(!pkg.partNames().some((p) => p.startsWith('xl/media/')), 'the image part is gone');
  assert.doesNotMatch(pkg.text('xl/worksheets/_rels/sheet1.xml.rels'), /relationships\/image"/);
  assert.equal(view.sheetBackground(), null);
  assert.throws(() => view.setBackground({ contentType: 'text/plain', data: PNG }), /PNG, JPEG/);
});
