/**
 * Number, text and date filters. Before 2.28 a filter arrow's menu could
 * only tick values: "customers who spent over a figure" could not be asked.
 * A rule is now picked from the menu, written into the file as Excel writes
 * it (customFilters, top10), and the rows it leaves out are hidden.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetView } from '@rutba/sheet-view';
import { Workbook } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';
import { ruleToFile, fileToRule, rulePasses, topThreshold, ruleFromForm, FILTER_OPS } from '@rutba/sheet-view/filter-rules';

const rows = [
  ['Customer', 'Town', 'Spent', 'Joined'],
  ['Asif', 'Lahore', 1200, null],
  ['Bina', 'Karachi', 450, null],
  ['Chand', 'Lahore Cantt', 3000, null],
  ['Dua', 'Multan', 800, null],
  ['Ejaz', 'Karachi', 1200, null],
  ['Fiza', '', 90, null],
];
const JOINED = ['2026-01-15', '2026-03-02', '2026-03-20', '2026-05-01', '2026-05-01', '2026-07-09'];

/** The list with a sheet filter on, dates typed in as dates. */
const filtered = () => {
  const view = SheetView.open(buildXlsx({ sheets: [{ name: 'Data', rows }] }));
  JOINED.forEach((d, i) => view.setCell(i + 1, 3, d));
  view.select(1, 0);
  view.toggleAutoFilter();
  return view;
};
const showing = (view) => [1, 2, 3, 4, 5, 6].filter((r) => view.geo.rowHeight(r) > 0).map((r) => view.calc.getValue('Data', r, 0));
const sheetXml = (view) => Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml');

test('a rule becomes the file\'s form and comes back as the rule', () => {
  const there = (rule, type) => fileToRule(ruleToFile(rule), type);
  for (const op of FILTER_OPS.number.filter((o) => o !== 'between')) {
    assert.deepEqual(there({ kind: 'number', op, value: 100 }, 'number'), { kind: 'number', op, value: 100 }, op);
  }
  assert.deepEqual(there({ kind: 'number', op: 'between', value: 500, value2: 100 }, 'number'), { kind: 'number', op: 'between', value: 100, value2: 500 }, 'the smaller first');
  for (const op of FILTER_OPS.text) {
    assert.deepEqual(there({ kind: 'text', op, value: 'tea' }, 'text'), { kind: 'text', op, value: 'tea' }, op);
  }
  for (const op of ['equals', 'before', 'after']) {
    assert.deepEqual(there({ kind: 'date', op, value: 46000 }, 'date'), { kind: 'date', op, value: 46000 }, op);
  }
  assert.deepEqual(there({ kind: 'date', op: 'between', value: 46000, value2: 46030 }, 'date'), { kind: 'date', op: 'between', value: 46000, value2: 46030 });
  assert.deepEqual(there({ kind: 'top', bottom: true, count: 3, percent: false }), { kind: 'top', bottom: true, count: 3, percent: false });
  // What Excel writes for "greater than 100" and for "contains tea".
  assert.deepEqual(ruleToFile({ kind: 'number', op: 'greater', value: 100 }), { custom: { and: false, tests: [{ operator: 'greaterThan', val: '100' }] } });
  assert.deepEqual(ruleToFile({ kind: 'text', op: 'contains', value: 'tea' }), { custom: { and: false, tests: [{ operator: 'equal', val: '*tea*' }] } });
});

test('a rule that cannot be one is refused in a plain sentence', () => {
  assert.throws(() => ruleToFile({ kind: 'number', op: 'greater', value: '' }), /Type a number/);
  assert.throws(() => ruleToFile({ kind: 'number', op: 'between', value: 1, value2: '' }), /both numbers/);
  assert.throws(() => ruleToFile({ kind: 'date', op: 'before', value: NaN }), /Type a date/);
  assert.throws(() => ruleToFile({ kind: 'text', op: 'contains', value: '' }), /Type the text/);
  assert.throws(() => ruleToFile({ kind: 'top', count: 0 }), /1 or more/);
  assert.throws(() => ruleToFile({ kind: 'top', count: 150, percent: true }), /100 or less/);
});

test('a cell passes or fails a rule by its value, and a text by what it shows, without regard to case', () => {
  const gt = ruleToFile({ kind: 'number', op: 'greater', value: 1000 });
  assert.equal(rulePasses(gt, 1200, '1,200'), true);
  assert.equal(rulePasses(gt, 1000, '1,000'), false);
  assert.equal(rulePasses(gt, 'n/a', 'n/a'), false, 'a text is not greater than a number');
  assert.equal(rulePasses(gt, null, ''), false);
  const has = ruleToFile({ kind: 'text', op: 'contains', value: 'LAH' });
  assert.equal(rulePasses(has, 'Lahore Cantt', 'Lahore Cantt'), true);
  assert.equal(rulePasses(has, 'Multan', 'Multan'), false);
  const not = ruleToFile({ kind: 'text', op: 'notContains', value: 'lah' });
  assert.equal(rulePasses(not, 'Multan', 'Multan'), true);
  assert.equal(rulePasses(not, '', ''), true, 'a blank does not contain it');
  const begins = ruleToFile({ kind: 'text', op: 'begins', value: 'kar' });
  assert.equal(rulePasses(begins, 'Karachi', 'Karachi'), true);
  assert.equal(rulePasses(begins, 'Sukkur', 'Sukkur'), false);
  // A wildcard typed by hand works as Excel's does; a tilde makes it a plain character.
  assert.equal(rulePasses({ custom: { and: false, tests: [{ operator: 'equal', val: 'K?rachi' }] } }, 'Karachi', 'Karachi'), true);
  assert.equal(rulePasses({ custom: { and: false, tests: [{ operator: 'equal', val: 'What~?' }] } }, 'What?', 'What?'), true);
  assert.equal(rulePasses({ custom: { and: false, tests: [{ operator: 'equal', val: 'What~?' }] } }, 'Whats', 'Whats'), false);
  // Either of two tests, when the file says "or".
  const either = { custom: { and: false, tests: [{ operator: 'lessThan', val: '100' }, { operator: 'greaterThan', val: '2000' }] } };
  assert.deepEqual([90, 800, 3000].map((v) => rulePasses(either, v, String(v))), [true, false, true]);
});

test('the top and the bottom few cut at a figure, and ties with it stay', () => {
  const spent = [1200, 450, 3000, 800, 1200, 90];
  assert.equal(topThreshold(spent, { top: true, val: 2 }), 1200);
  assert.equal(topThreshold(spent, { top: false, val: 2 }), 450);
  assert.equal(topThreshold(spent, { top: true, percent: true, val: 50 }), 1200, 'half of six is three');
  assert.equal(topThreshold(spent, { top: true, val: 99 }), 90, 'more than there are keeps all');
  assert.equal(topThreshold(['a', null], { top: true, val: 3 }), null);
});

test('the menu says what kind of column it is', () => {
  const view = filtered();
  assert.equal(view.filterOptions('#sheet', 'Spent').type, 'number');
  assert.equal(view.filterOptions('#sheet', 'Town').type, 'text');
  assert.equal(view.filterOptions('#sheet', 'Joined').type, 'date');
  assert.equal(view.filterOptions('#sheet', 'Spent').rule, null);
  assert.equal(view.filterOptions('#sheet', 'Spent').filtered, false);
});

test('a number filter hides the rows it leaves out, is written as Excel writes it, and reads back', () => {
  const view = filtered();
  view.openFilterPanel('#sheet', 'Spent');
  view.applyFilterRule('#sheet', 'Spent', { kind: 'number', op: 'greater', value: '1,000' });
  assert.equal(view.render().filterPanel, null, 'applying closes the menu');
  assert.deepEqual(showing(view), ['Asif', 'Chand', 'Ejaz']);
  assert.deepEqual(view.render().tables.find((t) => t.name === '#sheet').filtered, [2], 'the arrow shows it filters');
  const xml = sheetXml(view);
  assert.match(xml, /<filterColumn colId="2"><customFilters><customFilter operator="greaterThan" val="1000"\/><\/customFilters><\/filterColumn>/);
  assert.match(xml, /<row[^>]*r="3"[^>]*hidden="1"/);

  const again = SheetView.open(view.save());
  const panel = again.filterOptions('#sheet', 'Spent');
  assert.deepEqual(panel.rule, { kind: 'number', op: 'greater', value: 1000 });
  assert.equal(panel.filtered, true);
  assert.deepEqual(showing(again), ['Asif', 'Chand', 'Ejaz']);

  view.undo();
  assert.deepEqual(showing(view), ['Asif', 'Bina', 'Chand', 'Dua', 'Ejaz', 'Fiza'], 'one undo shows every row again');
  assert.doesNotMatch(sheetXml(view), /customFilters/);
});

test('between two numbers, and the top two with a tie kept', () => {
  const view = filtered();
  view.applyFilterRule('#sheet', 'Spent', { kind: 'number', op: 'between', value: 400, value2: 1200 });
  assert.deepEqual(showing(view), ['Asif', 'Bina', 'Dua', 'Ejaz']);
  assert.match(sheetXml(view), /<customFilters and="1"><customFilter operator="greaterThanOrEqual" val="400"\/><customFilter operator="lessThanOrEqual" val="1200"\/><\/customFilters>/);

  view.applyFilterRule('#sheet', 'Spent', { kind: 'top', count: 2 });
  assert.deepEqual(showing(view), ['Asif', 'Chand', 'Ejaz'], 'the two at 1,200 both stay');
  assert.match(sheetXml(view), /<filterColumn colId="2"><top10 val="2" filterVal="1200"\/><\/filterColumn>/);
  assert.deepEqual(view.filterOptions('#sheet', 'Spent').rule, { kind: 'top', bottom: false, count: 2, percent: false });

  view.applyFilterRule('#sheet', 'Spent', { kind: 'top', bottom: true, count: 1 });
  assert.deepEqual(showing(view), ['Fiza']);
  assert.match(sheetXml(view), /<top10 top="0" val="1" filterVal="90"\/>/);
});

test('a text filter: contains, begins with, does not equal', () => {
  const view = filtered();
  view.applyFilterRule('#sheet', 'Town', { kind: 'text', op: 'contains', value: 'lahore' });
  assert.deepEqual(showing(view), ['Asif', 'Chand']);
  assert.deepEqual(view.filterOptions('#sheet', 'Town').rule, { kind: 'text', op: 'contains', value: 'lahore' });
  view.applyFilterRule('#sheet', 'Town', { kind: 'text', op: 'begins', value: 'K' });
  assert.deepEqual(showing(view), ['Bina', 'Ejaz']);
  view.applyFilterRule('#sheet', 'Town', { kind: 'text', op: 'notEquals', value: 'Karachi' });
  assert.deepEqual(showing(view), ['Asif', 'Chand', 'Dua', 'Fiza']);
});

test('a date filter: after a day, between two days, on a day; typed as a date', () => {
  const view = filtered();
  view.applyFilterRule('#sheet', 'Joined', { kind: 'date', op: 'after', value: '2026-03-02' });
  assert.deepEqual(showing(view), ['Chand', 'Dua', 'Ejaz', 'Fiza']);
  assert.deepEqual(view.filterOptions('#sheet', 'Joined').rule, { kind: 'date', op: 'after', value: '2026-03-02' }, 'the menu gets the day back as a day');

  view.applyFilterRule('#sheet', 'Joined', { kind: 'date', op: 'between', value: '2026-03-01', value2: '2026-05-01' });
  assert.deepEqual(showing(view), ['Bina', 'Chand', 'Dua', 'Ejaz'], 'the last day is inside');
  assert.deepEqual(view.filterOptions('#sheet', 'Joined').rule, { kind: 'date', op: 'between', value: '2026-03-01', value2: '2026-05-01' });

  view.applyFilterRule('#sheet', 'Joined', { kind: 'date', op: 'equals', value: '1 May 2026' });
  assert.deepEqual(showing(view), ['Dua', 'Ejaz']);
  view.applyFilterRule('#sheet', 'Joined', { kind: 'date', op: 'before', value: '2026-03-02' });
  assert.deepEqual(showing(view), ['Asif']);
  // The file holds the days as their numbers, which is what Excel compares.
  assert.match(sheetXml(view), /<customFilter operator="lessThan" val="\d+"\/>/);
});

test('a rule on one column and ticked values on another both hold, and each is cleared by itself', () => {
  const view = filtered();
  view.applyFilterRule('#sheet', 'Spent', { kind: 'number', op: 'greaterOrEqual', value: 800 });
  view.applyFilter('#sheet', 'Town', ['Karachi', 'Multan']);
  assert.deepEqual(showing(view), ['Dua', 'Ejaz']);
  // Ticked values on the ruled column take the rule's place.
  view.applyFilter('#sheet', 'Spent', null);
  assert.deepEqual(showing(view), ['Bina', 'Dua', 'Ejaz']);
  assert.doesNotMatch(sheetXml(view), /customFilters/);
  view.applyFilter('#sheet', 'Town', null);
  assert.deepEqual(showing(view), ['Asif', 'Bina', 'Chand', 'Dua', 'Ejaz', 'Fiza']);
});

test('a sort from the menu keeps a rule\'s rows the ones hidden', () => {
  const view = filtered();
  view.applyFilterRule('#sheet', 'Spent', { kind: 'number', op: 'less', value: 1000 });
  view.sortFilterColumn('#sheet', 'Spent', false);
  assert.deepEqual(showing(view), ['Dua', 'Bina', 'Fiza']);
});

test('a table\'s column takes a rule too, in the table\'s own part', () => {
  const view = SheetView.open(buildXlsx({ sheets: [{ name: 'Data', rows }] }));
  view.select(1, 0);
  view.formatAsTable();
  const name = view.render().tables.find((t) => t.name !== '#sheet').name;
  view.applyFilterRule(name, 'Spent', { kind: 'number', op: 'lessOrEqual', value: 450 });
  assert.deepEqual(showing(view), ['Bina', 'Fiza']);
  const pkg = Workbook.open(view.save()).pkg;
  assert.match(pkg.text('xl/tables/table1.xml'), /<filterColumn colId="2"><customFilters><customFilter operator="lessThanOrEqual" val="450"\/><\/customFilters><\/filterColumn>/);
  assert.deepEqual(SheetView.open(view.save()).filterOptions(name, 'Spent').rule, { kind: 'number', op: 'lessOrEqual', value: 450 });
});

test('a filter of a kind made elsewhere is kept, shown as a filter, and can be cleared', () => {
  const view = filtered();
  view.applyFilterRule('#sheet', 'Spent', { kind: 'number', op: 'greater', value: 1000 });
  // An "above average" filter, as Excel writes one: not read here, kept as it is.
  const wb = Workbook.open(view.save());
  const sheet = wb._sheetPart('Data').part;
  sheet.setAutoFilter(sheet.autoFilterXml().replace(/<customFilters>.*?<\/customFilters>/, '<dynamicFilter type="aboveAverage" val="1123"/>'));
  const again = SheetView.open(wb.save());
  const panel = again.filterOptions('#sheet', 'Spent');
  assert.equal(panel.filtered, true);
  assert.equal(panel.rule, null);
  // Another column's filter does not undo the rows the kept filter hid, nor the filter.
  again.applyFilter('#sheet', 'Town', ['Lahore', 'Karachi']);
  assert.match(sheetXml(again), /<dynamicFilter type="aboveAverage" val="1123"\/>/);
  again.applyFilter('#sheet', 'Spent', null);
  assert.doesNotMatch(sheetXml(again), /dynamicFilter/);
});

test('the menu\'s form gives a rule only when it is filled', () => {
  assert.equal(ruleFromForm('number', { op: 'greater', a: ' ' }), null);
  assert.deepEqual(ruleFromForm('number', { op: 'greater', a: ' 1,000 ' }), { kind: 'number', op: 'greater', value: '1,000' });
  assert.equal(ruleFromForm('number', { op: 'between', a: '1', b: '' }), null);
  assert.deepEqual(ruleFromForm('date', { op: 'between', a: '2026-03-01', b: '2026-05-01' }), { kind: 'date', op: 'between', value: '2026-03-01', value2: '2026-05-01' });
  assert.deepEqual(ruleFromForm('number', { op: 'top', count: '5', percent: true }), { kind: 'top', bottom: false, count: 5, percent: true });
  assert.equal(ruleFromForm('number', { op: 'bottom', count: '0' }), null);
  assert.equal(ruleFromForm('number', { op: 'top', count: '150', percent: true }), null);
  assert.equal(ruleFromForm('text', { op: 'contains', a: '' }), null);
  // Equals nothing asks for the blanks.
  assert.deepEqual(ruleFromForm('text', { op: 'equals', a: '' }), { kind: 'text', op: 'equals', value: '' });
  // What the form gives, the view takes.
  const view = filtered();
  view.applyFilterRule('#sheet', 'Spent', ruleFromForm('number', { op: 'greater', a: '1,000' }));
  assert.deepEqual(showing(view), ['Asif', 'Chand', 'Ejaz']);
  view.applyFilterRule('#sheet', 'Town', ruleFromForm('text', { op: 'equals', a: '' }));
  assert.deepEqual(showing(view), []);
});
