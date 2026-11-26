// A typed entry read as Excel reads one.
//
// A date, a time, a percentage, an amount of money and a number written
// with commas go in as numbers, each with the number format that shows it
// the way it was typed, so SUM counts them and the fill handle steps a date
// by days. A cell that already wears a format keeps its own. Under the Text
// format the words stay as typed, a leading 0 with them. The formula bar
// shows a date as a date, never the day count under it.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView, readTyped, typedText } from '@rutba/sheet-view';
import { readDatesAs } from '@rutba/formula';

const BOOK = () => buildXlsx({ sheets: [{ name: 'S', rows: [['Typed']] }] });
const NOV_9 = 46335;

const type = (view, row, col, text) => {
  view.select(row, col);
  view.beginEdit({ replace: true, initial: '' });
  view.updateDraft(text);
  view.commitEdit({ move: 'none' });
};
const input = (view, row, col) => view.calc.getInput('S', row, col);
const shown = (view, row, col) => view.displayValue(row, col).text;

test.beforeEach(() => readDatesAs({ dayFirst: true, year: 2026 }));
test.after(() => readDatesAs());

test('a date typed three ways is the same day, shown as it was typed', () => {
  const view = new SheetView(BOOK());
  type(view, 1, 0, '09/11/2026');
  type(view, 2, 0, '9 Nov 2026');
  type(view, 3, 0, '2026-11-09');
  for (const row of [1, 2, 3]) assert.equal(input(view, row, 0), NOV_9, 'row ' + row + ' holds the day count');
  assert.equal(shown(view, 1, 0), '09/11/2026');
  assert.equal(shown(view, 2, 0), '9 Nov 2026');
  assert.equal(shown(view, 3, 0), '09/11/2026', 'a year-first date is shown as a short date, as Excel shows it');
  assert.equal(view.formatFor(1, 0), 'dd/mm/yyyy');
  assert.equal(view.formatFor(2, 0), 'd mmm yyyy');
});

test('the month comes first where the computer writes it first', () => {
  readDatesAs({ dayFirst: false, year: 2026 });
  const view = new SheetView(BOOK());
  type(view, 1, 0, '11/09/2026');
  assert.equal(input(view, 1, 0), NOV_9);
  assert.equal(view.formatFor(1, 0), 'mm/dd/yyyy');
  assert.equal(shown(view, 1, 0), '11/09/2026');
});

test('a time, a percentage, money and a number with commas are numbers that SUM counts', () => {
  const view = new SheetView(BOOK());
  type(view, 1, 0, '9:30');
  type(view, 2, 0, '15%');
  type(view, 3, 0, '£12.60');
  type(view, 4, 0, '1,250');
  assert.equal(input(view, 1, 0), 9.5 / 24);
  assert.equal(input(view, 2, 0), 0.15);
  assert.equal(input(view, 3, 0), 12.6);
  assert.equal(input(view, 4, 0), 1250);
  assert.deepEqual([1, 2, 3, 4].map((r) => shown(view, r, 0)), ['9:30', '15%', '£12.60', '1,250']);
  assert.deepEqual([1, 2, 3, 4].map((r) => view.formatFor(r, 0)), ['h:mm', '0%', '"£"#,##0.00', '#,##0']);
  view.setCell(5, 0, '=SUM(A3:A5)');
  assert.equal(view.calc.getValue('S', 5, 0), 0.15 + 12.6 + 1250, 'words would have been left out');
});

test('the forms a typed entry takes, and the ones that stay words', () => {
  const read = (t) => { const r = readTyped(t); return r ? [r.value, r.format] : null; };
  assert.deepEqual(read('12.5%'), [0.125, '0.00%']);
  assert.deepEqual(read('-£5'), [-5, '"£"#,##0']);
  assert.deepEqual(read('($1,250.50)'), [-1250.5, '"$"#,##0.00']);
  assert.deepEqual(read('1,250.5'), [1250.5, '#,##0.00']);
  assert.deepEqual(read('9:30 pm'), [21.5 / 24, 'h:mm AM/PM']);
  assert.deepEqual(read('9:30:15'), [(9 * 3600 + 30 * 60 + 15) / 86400, 'h:mm:ss']);
  assert.deepEqual(read('25:30'), [25.5 / 24, '[h]:mm'], 'hours past a day are counted on');
  assert.deepEqual(read('09/11/2026 14:30'), [NOV_9 + 14.5 / 24, 'dd/mm/yyyy hh:mm']);
  assert.deepEqual(read('9 Nov'), [NOV_9, 'd mmm'], 'no year: this year');
  assert.deepEqual(read('Nov 2026'), [NOV_9 - 8, 'mmm yyyy'], 'a month and a year: its first day');
  assert.deepEqual(read('9 November 2026'), [NOV_9, 'd mmm yyyy']);
  for (const words of ['3 Marbles 26', '10.5.20', '1/2', '1,25', 'Item 1', '12:99', '31/02/2026', 'Tea', '=1+1', "'15%", '250', '£']) {
    assert.equal(readTyped(words), null, words + ' is left as it is');
  }
});

test('a cell that wears a format of its own keeps it', () => {
  const view = new SheetView(BOOK());
  view.select(1, 0);
  view.setFormat({ numberFormat: '0.00' });
  type(view, 1, 0, '15%');
  assert.equal(input(view, 1, 0), 0.15);
  assert.equal(view.formatFor(1, 0), '0.00');
  assert.equal(shown(view, 1, 0), '0.15');
});

test('the Text format keeps what was typed, a leading 0 with it', () => {
  const view = new SheetView(BOOK());
  view.select(1, 0);
  view.setFormat({ numberFormat: '@' });
  type(view, 1, 0, '0123');
  type(view, 2, 0, '0123');
  assert.equal(input(view, 1, 0), '0123');
  assert.equal(shown(view, 1, 0), '0123');
  assert.equal(input(view, 2, 0), 123, 'without the Text format the 0 goes, as in Excel');
  view.select(3, 0);
  view.setFormat({ numberFormat: '@' });
  type(view, 3, 0, '09/11/2026');
  assert.equal(input(view, 3, 0), '09/11/2026', 'a date typed into a Text cell stays words');

  const again = SheetView.open(view.save());
  assert.equal(again.calc.getInput('S', 1, 0), '0123', 'and it is words in the saved file');
  assert.match(again.pkg.text(again.workbook.partNameFor('S')), /<c r="A2"[^>]*t="(inlineStr|s)"/);
});

test('the formula bar shows a date, a time and a percentage as they are typed', () => {
  const view = new SheetView(BOOK());
  type(view, 1, 0, '9 Nov 2026');
  type(view, 2, 0, '9:30');
  type(view, 3, 0, '5.7%');
  type(view, 4, 0, '09/11/2026 14:30');
  type(view, 5, 0, '1,250');
  view.select(1, 0);
  assert.equal(view.selectionFrame().formulaBar, '09/11/2026');
  assert.equal(view.barValue(2, 0), '9:30:00');
  assert.equal(view.barValue(3, 0), '5.7%');
  assert.equal(view.barValue(4, 0), '09/11/2026 14:30:00');
  assert.equal(view.barValue(5, 0), '1250');
  assert.equal(typedText(46335, 'General'), null);
  assert.equal(typedText(0.0569, '0.0%'), '5.69%', 'the whole number, not the rounded one shown');

  // Amended from what the bar shows: still a date, in the cell's own format.
  view.select(1, 0);
  view.beginEdit();
  assert.equal(view.editing.draft, '09/11/2026');
  view.updateDraft('10/11/2026');
  view.commitEdit({ move: 'none' });
  assert.equal(input(view, 1, 0), NOV_9 + 1);
  assert.equal(shown(view, 1, 0), '10 Nov 2026');

  // Opened and closed with no change, nothing is written.
  const steps = view.history.past.length;
  view.select(2, 0);
  view.beginEdit();
  view.commitEdit({ move: 'none' });
  assert.equal(view.history.past.length, steps);
});

test('one undo takes the number and its format away again; the file keeps both', () => {
  const view = new SheetView(BOOK());
  type(view, 1, 0, '09/11/2026');
  const again = SheetView.open(view.save());
  assert.equal(again.calc.getInput('S', 1, 0), NOV_9);
  assert.equal(again.displayValue(1, 0).text, '09/11/2026');
  assert.match(again.pkg.text(again.workbook.partNameFor('S')), /<c r="A2" s="\d+"><v>46335<\/v><\/c>/, 'a number with a style, as Excel writes a date');

  view.undo();
  assert.equal(view.editValue(1, 0), '');
  assert.equal(view.formatFor(1, 0), 'General');
  view.redo();
  assert.equal(input(view, 1, 0), NOV_9);
  assert.equal(shown(view, 1, 0), '09/11/2026');
});

test('the fill handle steps a typed date by days, not its year', () => {
  const view = new SheetView(BOOK());
  type(view, 1, 0, '9 Nov 2026');
  view.select(1, 0);
  view.fill({ top: 1, left: 0, bottom: 4, right: 0 });
  assert.deepEqual([2, 3, 4].map((r) => shown(view, r, 0)), ['10 Nov 2026', '11 Nov 2026', '12 Nov 2026']);
});

test('the 1904 date system is honoured', () => {
  assert.equal(readTyped('09/11/2026', { date1904: true }).value, NOV_9 - 1462);
  assert.equal(typedText(NOV_9 - 1462, 'dd/mm/yyyy', { date1904: true }), '09/11/2026');
  assert.equal(typedText(NOV_9, 'dd/mm/yyyy'), '09/11/2026');
});
