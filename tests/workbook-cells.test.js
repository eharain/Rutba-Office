/**
 * Walking a sheet's cells, and the row cache behind it.
 *
 * `Workbook.cells(sheet)` is what the calc model is built from. It used to
 * be `getCell` per cell, and `getCell` re-parsed the whole row per call — so
 * loading was quadratic in a sheet's width, and a tender workbook of 2.4
 * million mostly-blank formatted cells took eleven minutes to open. The
 * cache is keyed on the row's XML string, so an edit is seen at once.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';

const book = () => Workbook.open(buildXlsx({
  sheets: [{ name: 'S', rows: [[1, 'two', null, '=A1*2', 3.5], [], ['last']] }],
}));

test('cells() yields every cell that holds something, decoded, in row order', () => {
  const wb = book();
  const cells = [...wb.cells('S')];
  assert.deepEqual(cells.map((c) => c.ref), ['A1', 'B1', 'D1', 'E1', 'A3']);
  assert.deepEqual(cells.map((c) => [c.row, c.col]), [[0, 0], [0, 1], [0, 3], [0, 4], [2, 0]]);
  assert.deepEqual(cells.map((c) => c.value), [1, 'two', '=A1*2', 3.5, 'last']);
  // The same answers `getCell` gives, one at a time.
  for (const c of cells) assert.equal(wb.getCell('S', c.ref), c.value);
  assert.equal(wb.getCell('S', 'C1'), null, 'the gap is a gap');
});

test('a row is parsed once per version of its XML, and an edit refreshes it', () => {
  const wb = book();
  const { part } = wb._sheetPart('S');
  const row = part.rows[0];
  const first = part.constructor.cellsByCol(row);
  assert.strictEqual(part.constructor.cellsByCol(row), first, 'the same map while the XML stands');
  wb.setCell('S', 'C1', 'filled');
  assert.notStrictEqual(part.constructor.cellsByCol(row), first, 'a new map after the row changed');
  assert.equal(wb.getCell('S', 'C1'), 'filled');
  assert.deepEqual([...wb.cells('S')].map((c) => c.ref), ['A1', 'B1', 'C1', 'D1', 'E1', 'A3']);
});

test('a wide sheet loads in linear time, not quadratic', () => {
  // 60 rows of every cell present, 200 columns wide and then 400. Linear
  // walking takes about twice as long for twice the width; quadratic
  // parsing, which once made the wide one take seconds, takes four times.
  // Timed side by side in the same run, so a machine busy with other work
  // slows both alike — a fixed limit in milliseconds failed on one.
  const walk = (cols) => {
    const rows = Array.from({ length: 60 }, (_, r) => Array.from({ length: cols }, (_, c) => (c % 7 === 0 ? r * cols + c : `t${c}`)));
    const wb = Workbook.open(buildXlsx({ sheets: [{ name: 'Wide', rows }] }));
    let best = Infinity;
    let count = 0;
    for (let i = 0; i < 3; i++) {
      const fresh = i === 0 ? wb : Workbook.open(buildXlsx({ sheets: [{ name: 'Wide', rows }] }));
      const started = process.hrtime.bigint();
      count = [...fresh.cells('Wide')].length;
      best = Math.min(best, Number(process.hrtime.bigint() - started) / 1e6);
    }
    return { count, ms: best };
  };
  const narrow = walk(200);
  const wide = walk(400);
  assert.equal(narrow.count, 12000);
  assert.equal(wide.count, 24000);
  assert.ok(wide.ms < narrow.ms * 3 + 40, `24,000 cells walked in ${wide.ms.toFixed(0)} ms against ${narrow.ms.toFixed(0)} ms for 12,000`);
});

test('text holding $&, $\', $` or $$ is written as typed, into a cell, a sheet name and a formula', () => {
  // In a replacement string those four are instructions, and the engine
  // spliced what a person typed in as one: "cost $' and $&" over a cell left
  // it holding the old value, "a $$ b" lost a dollar, and a sheet renamed
  // "Sales $'000" took the workbook part apart.
  const typed = ["cost $' and $&", 'a $$ b $` c', 'US$1,000'];
  const wb = Workbook.open(buildXlsx({ sheets: [{ name: 'S', rows: [['x', 'y'], ['p']] }] }));
  wb.setCell('S', 'A1', typed[0]);
  wb.setCell('S', 'B2', typed[1]);
  wb.setCell('S', 'C1', typed[2]);
  const back = Workbook.open(wb.save());
  assert.deepEqual(['A1', 'B2', 'C1'].map((r) => back.getCell('S', r)), typed);
  assert.equal(back.getCell('S', 'B1'), 'y', 'the cell beside the edited one is untouched');

  const renamed = Workbook.open(buildXlsx({ sheets: [{ name: 'S', rows: [['x']] }, { name: 'T', rows: [['=S!A1']] }] }));
  renamed.renameSheet('S', "Sales $'000");
  const reread = Workbook.open(renamed.save());
  assert.deepEqual(reread.sheetNames(), ["Sales $'000", 'T']);

  const view = SheetView.open(buildXlsx({ sheets: [{ name: 'S', rows: [['x']] }] }));
  view.setCell(0, 0, '=CONCATENATE("a$$b","-$&")');
  const saved = SheetView.open(view.save());
  assert.equal(saved.workbook.getCell('S', 'A1'), '=CONCATENATE("a$$b","-$&")', 'the formula is saved as typed');
  assert.equal(saved.displayValue(0, 0).text, 'a$$b-$&', 'and so is the value it caches');
  // The reader is forgiving enough to find the formula inside a broken cell,
  // so the cell itself is read: caching the value had nested a second A1
  // inside the first.
  const sheetXml = saved.workbook.pkg.text(saved.workbook.partNameFor('S'));
  assert.equal(sheetXml.match(/<c\b[^>]*\br="A1"[\s\S]*?<\/c>/)?.[0],
    '<c r="A1" t="str"><f>CONCATENATE(&quot;a$$b&quot;,&quot;-$&amp;&quot;)</f><v>a$$b-$&amp;</v></c>');
});
