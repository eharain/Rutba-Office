// Worksheets' Home, Ctrl+Home and Ctrl+End, as Excel's: the row's first
// column, A1 (or the first cell past frozen panes), and the last cell in use,
// Shift stretching the selection there. None of the three did anything.
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';

const sheet = () => {
  const rows = [['Item', 'Region', 'Units']];
  for (let i = 1; i <= 50; i++) rows.push([`Item ${i}`, i % 2 ? 'East' : 'West', i]);
  return SheetView.open(buildXlsx({ sheets: [{ name: 'Data', rows }] }), { viewportWidth: 900, viewportHeight: 500 });
};
const at = (v) => ({ row: v.selection.active.row, col: v.selection.active.col });

test('Ctrl+End goes to the last cell in use, Ctrl+Home back to A1, Home to the row\'s first column', () => {
  const v = sheet();
  v.select(3, 1);
  v.homeEnd('end');
  assert.deepEqual(at(v), { row: 50, col: 2 });
  v.homeEnd('rowStart');
  assert.deepEqual(at(v), { row: 50, col: 0 });
  v.homeEnd('start');
  assert.deepEqual(at(v), { row: 0, col: 0 });
});

test('with Shift the selection stretches there; with frozen panes Ctrl+Home stops past them', () => {
  const v = sheet();
  v.select(2, 1);
  v.homeEnd('end', { extend: true });
  assert.equal(v.selection.toString(), 'B3:C51');
  v.freezePanes(1, 1);
  v.select(20, 2);
  v.homeEnd('start');
  assert.deepEqual(at(v), { row: 1, col: 1 }, 'the first cell that scrolls');
  v.select(20, 2);
  v.homeEnd('rowStart');
  assert.deepEqual(at(v), { row: 20, col: 1 });
  assert.throws(() => v.homeEnd('middle'), /not a place/);
});
