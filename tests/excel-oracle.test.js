// The engine held to Excel's own answers (tools/excel-oracle.mjs wrote them).
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Spreadsheet, isError } from '@rutba/formula';
import { formatValue } from '../packages/formula/src/numfmt.js';

const oracle = JSON.parse(fs.readFileSync(new URL('./fixtures/excel-oracle.json', import.meta.url), 'utf8'));
const NOW = () => new Date('2026-09-25T12:00:00Z');

function book() {
  const s = new Spreadsheet({ now: NOW });
  for (const [sheet, cells] of Object.entries(oracle.sheets)) {
    s.addSheet?.(sheet);
    for (const [a1, v] of Object.entries(cells)) s.setByRef(sheet, a1, v);
  }
  return s;
}

function shown(v) {
  if (isError(v)) return { error: v.type };
  return { value: v === undefined || v === '' ? v ?? null : v };
}

function same(mine, excel) {
  if (excel.error) return mine.error === excel.error;
  if (mine.error) return false;
  const a = mine.value; const b = excel.value;
  if ((a === null || a === undefined || a === '') && (b === null || b === '')) return true;
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(b));
  return a === b;
}

export function faults() {
  const out = [];
  const s = book();
  for (const { formula, answer, refused } of oracle.formulas) {
    if (refused) continue;
    s.setByRef('Sheet1', 'J20', formula);
    s.recalculate();
    const mine = shown(s.getValue('Sheet1', 19, 9));
    if (!same(mine, answer)) out.push(`${formula}  gives ${JSON.stringify(mine)}, Excel ${JSON.stringify(answer)}`);
  }
  return out;
}

test('every formula comes to what Excel says', () => {
  const wrong = faults();
  assert.deepEqual(wrong, [], `${wrong.length} of ${oracle.formulas.length}`);
});

test('every number format prints what Excel prints', () => {
  const wrong = [];
  for (const { value, format, text } of oracle.formats) {
    let mine;
    try {
      mine = formatValue(value, format);
      mine = typeof mine === 'string' ? mine : mine?.text;
    } catch (e) { mine = `threw ${e.message}`; }
    // A * fills the cell's width in Excel; here it is left to the grid.
    const same = format.includes('*') ? String(mine).replace(/ /g, '') === text.replace(/ /g, '') : String(mine).trimEnd() === text.trimEnd();
    if (!same) wrong.push(`${value} as ${format}  prints ${JSON.stringify(mine)}, Excel ${JSON.stringify(text)}`);
  }
  assert.deepEqual(wrong, [], `${wrong.length} of ${oracle.formats.length}`);
});

test('a workbook on the 1904 date system comes to what Excel says', () => {
  const s = new Spreadsheet({ now: NOW, date1904: true });
  for (const [sheet, cells] of Object.entries(oracle.sheets)) {
    s.addSheet?.(sheet);
    for (const [a1, v] of Object.entries(cells)) s.setByRef(sheet, a1, v);
  }
  const wrong = [];
  for (const { formula, answer } of oracle.on1904.formulas) {
    s.setByRef('Sheet1', 'J20', formula);
    s.recalculate();
    const mine = shown(s.getValue('Sheet1', 19, 9));
    if (!same(mine, answer)) wrong.push(`${formula}  gives ${JSON.stringify(mine)}, Excel ${JSON.stringify(answer)}`);
  }
  for (const { value, format, text } of oracle.on1904.formats) {
    const mine = formatValue(value, format, { date1904: true }).text;
    if (mine !== text) wrong.push(`${value} as ${format}  prints ${JSON.stringify(mine)}, Excel ${JSON.stringify(text)}`);
  }
  assert.deepEqual(wrong, []);
});
