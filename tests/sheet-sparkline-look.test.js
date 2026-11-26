/**
 * Sparklines: the axis, the colours and the points picked out. Before 2.28
 * every sparkline was drawn from nought to its own highest value, in one
 * colour and with no markers, so sales of 21 to 24 a month drew a flat line
 * near the top of the cell.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { Workbook } from '@rutba/ooxml/workbook';
import { SheetView } from '@rutba/sheet-view';
import { sparklineShape } from '@rutba/sheet-view/sparkline';

const open = () => {
  const view = SheetView.open(buildXlsx({
    sheets: [{
      name: 'Sales',
      rows: [
        ['Item', 'Jan', 'Feb', 'Mar', 'Apr', ''],
        ['Pens', 21, 24, 22, 23, ''],
        ['Ink', 2, 60, -9, 30, ''],
      ],
    }],
  }));
  view.addSparklines({ type: 'line', data: 'B2:E3', at: 'F2:F3' });
  return view;
};
const group = (view, i = 0) => view.sparklineGroups('Sales')[i];
const sheetXml = (view) => {
  const wb = Workbook.open(view.save());
  return wb.pkg.text(wb.partNameFor('Sales'));
};

// ---- how one is drawn -------------------------------------------------------------

test('a fresh group: each sparkline runs from its own lowest to its own highest, with no points picked out', () => {
  const g = group(open());
  assert.equal(g.axis, 'each');
  assert.deepEqual(g.marks, { high: false, low: false, first: false, last: false, negative: false, markers: false });
  const shape = sparklineShape(g, [21, 24, 22, 23]);
  assert.deepEqual(shape, { floor: 21, ceil: 24, points: [] }, '21 to 24 fills the cell top to bottom');
});

test('the axis: started at nought, or one pair of ends for the whole group', () => {
  const view = open();
  assert.equal(view.setSparklineGroup('F2', { axis: 'zero' }), 1);
  assert.equal(group(view).axis, 'zero');
  assert.deepEqual(sparklineShape(group(view), [21, 24, 22, 23]), { floor: 0, ceil: 24, points: [] });
  view.setSparklineGroup('F3', { axis: 'same' });
  const g = group(view);
  assert.equal(g.axis, 'same');
  const all = [21, 24, 22, 23, 2, 60, -9, 30];
  assert.deepEqual(sparklineShape(g, [21, 24, 22, 23], all), { floor: -9, ceil: 60, points: [] });
  view.setSparklineGroup('F2', { axis: 'each' });
  assert.equal(group(view).axis, 'each');
  assert.doesNotMatch(sheetXml(view), /minAxisType|manualMin/);
});

test('columns stand on nought when it is in sight, and the shortest keeps some height when it is not', () => {
  const view = open();
  view.setSparklineGroup('F2', { type: 'column' });
  const g = group(view);
  assert.equal(g.type, 'column');
  const high = sparklineShape(g, [21, 24, 22, 23]);
  assert.ok(high.floor > 0 && high.floor < 21, 'the floor is under the shortest column: ' + high.floor);
  assert.equal(high.ceil, 24);
  const mixed = sparklineShape(g, [2, 60, -9, 30]);
  assert.deepEqual([mixed.floor, mixed.ceil], [-9, 60]);
  const small = sparklineShape(g, [1, 9]);
  assert.equal(small.floor, 0, 'never below nought for numbers above it');
  view.setSparklineGroup('F2', { axis: 'zero' });
  assert.equal(sparklineShape(group(view), [21, 24, 22, 23]).floor, 0);
});

test('the points picked out: high, low, first, last, negative and every point, each in its colour', () => {
  const view = open();
  view.setSparklineGroup('F3', { high: true, low: true, markColour: '00B050' });
  let g = group(view);
  assert.equal(g.marks.high, true);
  assert.equal(g.colours.high, '00B050');
  assert.deepEqual(sparklineShape(g, [2, 60, -9, 30]).points, [
    { i: 1, kind: 'high', colour: '00B050' },
    { i: 2, kind: 'low', colour: '00B050' },
  ]);
  view.setSparklineGroup('F3', { high: false, low: false, first: true, last: true, negative: true, markers: true });
  g = group(view);
  assert.deepEqual(sparklineShape(g, [2, null, -9, 30, null]).points.map((p) => p.i + ':' + p.kind), ['0:first', '2:negative', '3:last'],
    'a blank is no point, and the last point is the last number');
  assert.deepEqual(sparklineShape(g, [2, 60, 9, 30]).points.map((p) => p.kind), ['first', 'markers', 'markers', 'last']);
  // A column has no markers of its own; its picked columns change colour.
  view.setSparklineGroup('F3', { type: 'column' });
  assert.deepEqual(sparklineShape(group(view), [2, 60, 9, 30]).points.map((p) => p.i + ':' + p.kind), ['0:first', '3:last']);
});

// ---- the file ---------------------------------------------------------------------

test('the look is written on the group as Excel writes it, and is there when the file is opened again', () => {
  const view = open();
  view.setSparklineGroup('F2', { type: 'column', colour: '7030a0', axis: 'same', high: true, negative: true, markColour: '#FFC000' });
  const xml = sheetXml(view);
  const head = /<x14:sparklineGroup\b([^>]*)>/.exec(xml)[1];
  assert.match(head, / type="column"/);
  assert.match(head, / high="1"/);
  assert.match(head, / negative="1"/);
  assert.match(head, / minAxisType="group"/);
  assert.match(head, / maxAxisType="group"/);
  assert.match(head, / displayEmptyCellsAs="gap"/, 'what was on the group stays');
  assert.match(xml, /<x14:colorSeries rgb="FF7030A0"\/><x14:colorNegative rgb="FFFFC000"\/><x14:colorAxis rgb="FF000000"\/><x14:colorMarkers rgb="FFFFC000"\/><x14:colorFirst rgb="FFFFC000"\/><x14:colorLast rgb="FFFFC000"\/><x14:colorHigh rgb="FFFFC000"\/><x14:colorLow rgb="FFFFC000"\/><x14:sparklines>/,
    'the colours stay in the order the format asks for');
  const again = SheetView.open(view.save());
  const g = again.sparklineGroups('Sales')[0];
  assert.equal(g.type, 'column');
  assert.equal(g.colour, '7030A0');
  assert.equal(g.axis, 'same');
  assert.deepEqual(g.marks, { high: true, low: false, first: false, last: false, negative: true, markers: false });
  assert.equal(g.sparklines.length, 2, 'the sparklines themselves are untouched');
});

test('an axis typed in Excel as two figures is read and drawn, and other attributes ride through a change', () => {
  const wb = Workbook.open(open().save());
  const part = wb.partNameFor('Sales');
  const xml = wb.pkg.text(part);
  assert.match(xml, /<x14:sparklineGroup /);
  wb.pkg.write_(part, xml.replace('<x14:sparklineGroup ', '<x14:sparklineGroup manualMax="50" manualMin="10" lineWeight="1.5" minAxisType="custom" maxAxisType="custom" '));
  const typed = SheetView.open(wb.save());
  const g = typed.sparklineGroups('Sales')[0];
  assert.equal(g.axis, 'custom');
  assert.deepEqual([g.manualMin, g.manualMax], [10, 50]);
  assert.deepEqual(sparklineShape(g, [21, 24, 22, 23]), { floor: 10, ceil: 50, points: [] });
  typed.setSparklineGroup('F2', { low: true });
  const after = Workbook.open(typed.save());
  const head = /<x14:sparklineGroup\b([^>]*)>/.exec(after.pkg.text(part))[1];
  assert.match(head, / lineWeight="1.5"/);
  assert.match(head, / manualMin="10"/, 'a change that does not name the axis leaves it');
  assert.match(head, / low="1"/);
});

test('only the group under the cell changes; one undo puts the look back; a protected sheet and a wrong word are refused', () => {
  const view = open();
  view.setCell(4, 0, 'Paper');
  view.setCell(4, 1, 5);
  view.setCell(4, 2, 8);
  view.addSparklines({ type: 'column', data: 'B5:C5', at: 'F5' });
  assert.equal(view.setSparklineGroup('F5', { colour: 'C00000', first: true }), 1);
  assert.equal(group(view, 0).colour, '376092');
  assert.equal(group(view, 1).colour, 'C00000');
  assert.equal(group(view, 1).marks.first, true);
  view.undo();
  assert.equal(group(view, 1).colour, '376092');
  assert.equal(group(view, 1).marks.first, false);
  view.redo();
  assert.equal(group(view, 1).colour, 'C00000');
  assert.equal(view.setSparklineGroup('F2:F5', { last: true }), 2, 'a range reaches every group in it');
  assert.equal(view.setSparklineGroup('A1', { last: false }), 0, 'no sparkline there');
  assert.throws(() => view.setSparklineGroup('F2', { colour: 'blue' }), /six hex digits/);
  assert.throws(() => view.setSparklineGroup('F2', { type: 'pie' }), /line.*column/);
  assert.throws(() => view.setSparklineGroup('F2', { axis: 'log' }), /each.*same.*zero/);
  view.protect();
  assert.throws(() => view.setSparklineGroup('F2', { high: true }), /protected/);
});
