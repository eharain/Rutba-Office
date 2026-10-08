// What XML 1.0 has no place for is left out of what is written: a cell or a
// paragraph holding a control character saves as a file Excel and Word open.
import test from 'node:test';
import assert from 'node:assert/strict';
import { esc, OoxmlPackage, buildXlsx, buildDocx } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

const bad = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/;

test('the escaper keeps tab, line feed and return and leaves the rest of the control characters out', () => {
  assert.equal(esc('a\u0007b\tc\nd\re\u0000f<'), 'ab\tc\nd\ref&lt;');
});

test('a cell and a paragraph with a control character save as valid XML', () => {
  const view = SheetView.open(buildXlsx({ sheets: [{ name: 'Sheet1', rows: [] }] }), { viewportWidth: 600, viewportHeight: 300 });
  view.setCell?.(0, 0, 'bell\u0007here') ?? view.select(0, 0);
  if (!view.setCell) { view.beginEdit(); view.typeText?.('bell\u0007here'); view.commitEdit?.(); }
  const sheet = OoxmlPackage.read(Buffer.from(view.serialize())).text('xl/worksheets/sheet1.xml');
  assert.doesNotMatch(sheet, bad);
  assert.match(sheet, /bellhere/);
  const doc = openDocx(buildDocx({ paragraphs: [{ text: 'Start' }] }));
  doc.setSelection({ block: 0, offset: 5 });
  doc.insertText(' \u0001end');
  const body = OoxmlPackage.read(doc.save()).text('word/document.xml');
  assert.doesNotMatch(body, bad);
  assert.match(body, /Start end/);
});
