// Three workbooks of formulas stored with no array mark, opened and worked out
// here, against what Excel 365 answered for the very same files. The formulas
// sit in row 2 from column F; the data is in A1:D10. Excel's own reading of
// each formula, with the @ it shows, is kept beside its answer.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../packages/ooxml/src/workbook.js';
import { toSpreadsheet } from '../packages/ooxml/src/recalc.js';

const dir = new URL('./fixtures/excel-answers/', import.meta.url);

// Where this program still differs, and why. Anything else that differs is a fault.
const NOT_YET = [
  [/^=@?LET\(/, 'LET in a file with no array mark: Excel narrows its names, here they are not'],
  [/[A-Z]\d+:[A-Z]\d+ [A-Z]\d+/, 'the space between two ranges (where they cross) is not read'],
  [/\(\([A-Z]\d+:[A-Z]\d+,/, 'two ranges in brackets as one reference is not read'],
  [/^=@?(?:SUM\()?(?:FREQUENCY|MMULT|MINVERSE|TREND|LINEST|GROWTH|NUMBERVALUE|ISREF|ISFORMULA|FORMULATEXT|CELL)\(/, 'a function not built yet'],
  [/^=ISOMITTED\(/, 'only inside LAMBDA'],
  [/^=DOLLAR\(/, 'the currency sign follows the computer Excel ran on'],
  [/^=@2:2$/, 'its own row: a circular reference, which Excel shows as 0 with a warning'],
];

// A number may be written 1E15 or 1000000000000000: the same formula.
const spelled = (f) => String(f ?? '').replace(/_xlfn\.|_xlws\./g, '').replace(/\b\d+(?:\.\d+)?(?:E[+-]?\d+)?\b/gi, (n) => String(Number(n)));
const shown = (v) => (v && typeof v === 'object' ? v.type : v);
const same = (mine, { value, text }) => mine === value || String(mine) === String(text)
  || (typeof mine === 'number' && typeof value === 'number' && Math.abs(mine - value) < 1e-9 * Math.max(1, Math.abs(value)));

for (const n of [1, 2, 3]) {
  test(`plain formulas, set ${n}: every answer is Excel's, apart from the listed gaps`, () => {
    const answers = JSON.parse(fs.readFileSync(new URL(`plain-formulas-${n}.json`, dir), 'utf8'));
    const sheet = toSpreadsheet(Workbook.open(fs.readFileSync(new URL(`plain-formulas-${n}.xlsx`, dir)))).sheet;
    sheet.recalculate();
    const wrong = [];
    let gaps = 0;
    for (const a of answers) {
      const mine = shown(sheet.getValue('S', 1, a.col));
      if (same(mine, a)) continue;
      if (NOT_YET.some(([re]) => re.test(a.excel))) { gaps += 1; continue; }
      wrong.push(`${a.excel}: here ${JSON.stringify(mine)}, Excel ${JSON.stringify(a.value)}`);
    }
    assert.deepEqual(wrong, []);
    assert.ok(gaps < answers.length / 5, `${gaps} gaps of ${answers.length}`);
  });
  test(`plain formulas, set ${n}: each is shown with the @ Excel shows`, () => {
    const answers = JSON.parse(fs.readFileSync(new URL(`plain-formulas-${n}.json`, dir), 'utf8'));
    const sheet = toSpreadsheet(Workbook.open(fs.readFileSync(new URL(`plain-formulas-${n}.xlsx`, dir)))).sheet;
    const wrong = answers.filter((a) => !NOT_YET.some(([re]) => re.test(a.excel)))
      .map((a) => [a.excel, sheet.cell('S', 1, a.col)?.input])
      .filter(([excel, mine]) => spelled(excel) !== spelled(mine));
    assert.deepEqual(wrong, []);
  });
}

test('a format with no place before the point still shows the whole number, and 0 in # ?/? is 0', async () => {
  const { formatValue } = await import('../packages/formula/src/numfmt.js');
  assert.equal(formatValue(1234.567, '.00').text, '1234.57');
  assert.equal(formatValue(-1.5, '.00').text, '-1.50');
  assert.equal(formatValue(0.5, '.00').text, '.50');
  assert.equal(formatValue(0, '# ?/?').text, '0    ');
  assert.equal(formatValue(0.5, '# ?/?').text, ' 1/2', 'the whole part gone, its space kept');
});

test('Evaluate Formula takes @ and what it narrows as one step', async () => {
  const { Spreadsheet } = await import('../packages/formula/src/engine.js');
  const { FormulaStepper } = await import('../packages/formula/src/evaluate-steps.js');
  const s = new Spreadsheet();
  s.addSheet('S');
  for (const [r, v] of [[0, 10], [1, 20], [2, 30]]) s.setCell('S', r, 0, v);
  s.setCell('S', 1, 2, '=@A1:A3*2');
  s.recalculate?.();
  assert.equal(s.getValue('S', 1, 2), 40);
  const st = new FormulaStepper('=@A1:A3*2', s.resolver(), { sheet: 'S', row: 1, col: 2 });
  const seen = [st.view().text];
  for (let i = 0; i < 6 && !st.done; i++) { st.evaluate(); seen.push(st.view().text); }
  assert.ok(seen.some((t) => /^=?20\*2$/.test(t)), JSON.stringify(seen));
  assert.ok(seen.every((t) => !/\{/.test(t)), 'never the whole range as an array first');
});
