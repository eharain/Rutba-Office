// Tables in Documents — edges the layout tests leave alone: a table Word made
// (merged cells, a table inside a cell, a tracked change of the table's own
// properties) through formulas, styles and an edit elsewhere.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx, OoxmlPackage } from '@rutba/ooxml';

const WHO = 'w:author="Ann" w:date="2026-01-01T00:00:00Z"';
const cell = (t, extra = '') => `<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/>${extra}</w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
const fromBody = (xml) => {
  const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: [{ text: 'PLACEHOLDER' }, { text: 'Second paragraph' }] }));
  pkg.write_('word/document.xml', pkg.text('word/document.xml').replace(/<w:p\b(?:(?!<\/w:p>)[\s\S])*?PLACEHOLDER[\s\S]*?<\/w:p>/, () => xml));
  return pkg.write();
};
const tableOf = (bytes) => /<w:tbl>[\s\S]*<\/w:tbl>/.exec(OoxmlPackage.read(bytes).text('word/document.xml'))[0];

const nestedSum = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid><w:tr>${cell('7')}</w:tr><w:tr>${cell('8')}</w:tr><w:tr><w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p><w:fldSimple w:instr=" =SUM(ABOVE) "><w:r><w:t>0</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr></w:tbl>`;
const merged =
  `<w:tbl><w:tblPr><w:tblW w:w="4500" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>` +
  `<w:tr>${cell('A', '<w:gridSpan w:val="2"/>')}${cell('B')}</w:tr>` +
  `<w:tr>${cell('C', '<w:vMerge w:val="restart"/>')}${cell('1')}${cell('2')}</w:tr>` +
  `<w:tr>${cell('', '<w:vMerge/>')}${cell('3')}<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>${nestedSum}<w:p/></w:tc></w:tr></w:tbl>`;

test('a Word table with gridSpan, vMerge and a nested table saves byte-for-byte untouched, and an edit elsewhere leaves the table exactly as it was', () => {
  const bytes = fromBody(merged);
  const view = openDocx(bytes);
  assert.equal(Buffer.compare(Buffer.from(view.save()), Buffer.from(bytes)), 0, 'nothing touched, nothing rewritten');

  const last = view.render({ pages: false }).blocks.findIndex((b) => b.text === 'Second paragraph');
  view.setSelection({ block: last, offset: 0 });
  view.insertText('>');
  assert.equal(tableOf(view.save()), tableOf(bytes));
});

test('Update all formulas works a nested table\'s sum from the nested table\'s own cells, not the table around it', () => {
  const view = openDocx(fromBody(merged));
  assert.equal(view.updateTableFormulas(), 1);
  const xml = tableOf(view.save());
  assert.match(xml, /<w:fldSimple w:instr=" =SUM\(ABOVE\) "><w:r><w:t xml:space="preserve">15<\/w:t>/, '7 + 8, not the 2 above the outer cell');
  // The merged cells around it are as they were.
  assert.match(xml, /<w:gridSpan w:val="2"\/>/);
  assert.match(xml, /<w:vMerge w:val="restart"\/>/);
});

test('Table Design leaves a tracked change of the table\'s properties holding the old ones whole', () => {
  const table = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="4500" w:type="dxa"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>` +
    `<w:tblPrChange w:id="3" ${WHO}><w:tblPr><w:tblStyle w:val="OldStyle"/><w:tblW w:w="3000" w:type="dxa"/></w:tblPr></w:tblPrChange></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid><w:tr>${cell('A')}</w:tr></w:tbl>`;
  const view = openDocx(fromBody(table));
  const at = view.render({ pages: false }).blocks.findIndex((b) => b.text === 'A');
  view.setSelection({ block: at, offset: 0 });
  view.tableOp('style', { id: 'GridTable4-Accent1' });
  view.tableOp('tableWidth', { type: 'pct', value: 50 });
  view.tableOp('styleOptions', { look: { firstRow: false } });
  const xml = tableOf(view.save());
  assert.match(xml, /^<w:tbl><w:tblPr><w:tblStyle w:val="GridTable4-Accent1"\/><w:tblW w:w="2500" w:type="pct"\/>/);
  assert.match(xml, /<w:tblPrChange w:id="3"[^>]*><w:tblPr><w:tblStyle w:val="OldStyle"\/><w:tblW w:w="3000" w:type="dxa"\/><\/w:tblPr><\/w:tblPrChange><\/w:tblPr>/, 'the record still names the old style and width');
  assert.equal(xml.match(/<w:tblLook\b/g).length, 1, 'one look, the table\'s own');
});
