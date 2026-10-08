// Power Query keeps the zeros a file quoted: a quoted "007" in a CSV is the
// text 007, and Split Column leaves a part with leading zeros as text.
import test from 'node:test';
import assert from 'node:assert/strict';
import { parseDelimited, applyStep } from '../packages/sheet-view/src/queries.js';

test('a quoted value with leading zeros stays text; any other that reads as a number is a number', () => {
  assert.deepEqual(parseDelimited('Code,Qty\n"007",5\n008,"6"\n"",1\n'), [['Code', 'Qty'], ['007', 5], [8, 6], [null, 1]]);
  assert.deepEqual(parseDelimited('a,"b ""c"""\n'), [['a', 'b "c"']]);
});

test('Split Column keeps a part with leading zeros as text, and makes the rest numbers', () => {
  const t = { columns: ['Ref'], rows: [['A-007'], ['B-12'], ['C-0']] };
  const out = applyStep(t, { kind: 'splitColumn', column: 'Ref', delimiter: '-' });
  assert.deepEqual(out.rows, [['A', '007'], ['B', 12], ['C', 0]]);
});
