// Power Query's Fill Down and Up, Unpivot Other Columns, Merge Columns and
// Extract, as the Query Editor's Transform tab gives them.
import test from 'node:test';
import assert from 'node:assert/strict';
import { runSteps, applyStep, describeStep } from '@rutba/sheet-view/queries';

test('Fill Down gives a blank cell the value above it, Fill Up the value below', () => {
  const report = { columns: ['Region', 'Rep', 'Sales'], rows: [['East', 'Kim', 3], [null, 'Lee', 2], ['', 'Ann', 5], ['West', 'Bo', 1], [null, 'Cy', 4]] };
  assert.deepEqual(runSteps(report, [{ kind: 'fillDown', columns: ['Region'] }]).rows.map((r) => r[0]), ['East', 'East', 'East', 'West', 'West']);
  assert.deepEqual(runSteps(report, [{ kind: 'fillUp', columns: ['Region'] }]).rows.map((r) => r[0]), ['East', 'West', 'West', 'West', null], 'nothing below the last: left blank');
  assert.equal(describeStep({ kind: 'fillDown', columns: ['Region'] }), '"Region" filled down');
});

test('Unpivot Other Columns makes each month\'s cell a row of its own, a blank giving none', () => {
  const wide = { columns: ['Product', 'Jan', 'Feb', 'Mar'], rows: [['Pens', 10, 12, null], ['Ink', 5, null, 7]] };
  const typed = runSteps(wide, [{ kind: 'changeType', column: 'Jan', type: 'number' }, { kind: 'changeType', column: 'Feb', type: 'number' }, { kind: 'changeType', column: 'Mar', type: 'number' }]);
  const long = applyStep(typed, { kind: 'unpivotOthers', columns: ['Product'] });
  assert.deepEqual(long.columns, ['Product', 'Attribute', 'Value']);
  assert.deepEqual(long.rows, [['Pens', 'Jan', 10], ['Pens', 'Feb', 12], ['Ink', 'Jan', 5], ['Ink', 'Mar', 7]]);
  assert.deepEqual(long.types, { Attribute: 'text', Value: 'number' }, 'the values keep the type their columns shared');
  assert.deepEqual(applyStep(wide, { kind: 'unpivotOthers', columns: ['Product'] }).types, { Attribute: 'text' }, 'untyped values stay untyped');
});

test('Merge Columns joins columns with a separator where the first was, and Extract keeps part of a text', () => {
  const people = { columns: ['First', 'Code', 'Last'], rows: [['Ada', 'UK-001', 'Lovelace'], ['Alan', 'UK-002', null]] };
  const merged = applyStep(people, { kind: 'mergeColumns', columns: ['First', 'Last'], separator: ' ', name: 'Name' });
  assert.deepEqual(merged.columns, ['Name', 'Code']);
  assert.deepEqual(merged.rows, [['Ada Lovelace', 'UK-001'], ['Alan ', 'UK-002']]);
  assert.deepEqual(merged.types, { Name: 'text' });
  assert.throws(() => applyStep(people, { kind: 'mergeColumns', columns: ['First'] }), /two columns/);
  const ex = (how, extra) => applyStep(people, { kind: 'extractText', column: 'Code', how, ...extra }).rows.map((r) => r[1]);
  assert.deepEqual(ex('first', { count: 2 }), ['UK', 'UK']);
  assert.deepEqual(ex('last', { count: 3 }), ['001', '002'], 'the leading zeros kept: text');
  assert.deepEqual(ex('before', { delimiter: '-' }), ['UK', 'UK']);
  assert.deepEqual(ex('after', { delimiter: '-' }), ['001', '002']);
  assert.throws(() => applyStep(people, { kind: 'extractText', column: 'Code', how: 'after', delimiter: '' }), /before or after something/);
  assert.match(describeStep({ kind: 'extractText', column: 'Code', how: 'after', delimiter: '-' }), /the text after "-"/);
});
