/**
 * Formulas → Evaluate Formula, in the engine: a formula worked out one part
 * at a time, in the order Excel's dialog shows — references read first, a
 * function's arguments left to right and the function last, IF's untaken
 * branch left as written — with Step In to a referenced cell's own formula
 * and Step Out back to the formula it was stepped into from.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  Spreadsheet, FormulaStepper, FormulaEvaluation, evaluationText, parse,
} from '@rutba/formula';

/** A workbook in the engine, and the formula text of a cell, for Step In. */
function book(cells) {
  const s = new Spreadsheet({ now: () => new Date('2026-09-25T12:00:00Z') });
  for (const [a1, v] of Object.entries(cells)) s.setByRef('Sheet1', a1, v);
  s.recalculate();
  const getFormula = (sheet, row, col) => s.cell(sheet, row, col)?.formula ?? null;
  return { s, resolver: s.resolver(), getFormula };
}

/** Every text the formula passes through, Evaluate after Evaluate, underline in [brackets]. */
function walk(formula, cells, at = { sheet: 'Sheet1', row: 9, col: 9 }) {
  const { resolver } = book(cells);
  const st = new FormulaStepper(formula, resolver, at);
  const seen = [];
  const mark = () => {
    const v = st.view();
    seen.push(v.underline ? v.text.slice(0, v.underline[0]) + '[' + v.text.slice(...v.underline) + ']' + v.text.slice(v.underline[1]) : v.text);
  };
  mark();
  let guard = 0;
  while (!st.done && guard++ < 50) { st.evaluate(); mark(); }
  return { seen, value: st.value };
}

test('the parser gives each node its place in the text when asked, and nothing extra otherwise', () => {
  const src = 'SUM(A1:A3, (2+3) )*-B2';
  const ast = parse('=' + src, { spans: true });
  assert.deepEqual(ast.span, [0, src.length]);
  assert.equal(src.slice(...ast.left.span), 'SUM(A1:A3, (2+3) )');
  assert.equal(src.slice(...ast.left.args[0].span), 'A1:A3');
  assert.equal(src.slice(...ast.left.args[1].span), '2+3');
  assert.equal(src.slice(...ast.left.args[1].outer), '(2+3)');
  assert.equal(parse('=1+2').span, undefined, 'a plain parse is unchanged');
});

test('references first, then the operators, the most recent result shown last', () => {
  const { seen, value } = walk('=A1+B1*2', { A1: 5, B1: 3 });
  assert.deepEqual(seen, ['=[A1]+B1*2', '=5+[B1]*2', '=5+[3*2]', '=[5+6]', '=11']);
  assert.equal(value, 11);
});

test('a range is shown as Excel shows an array, and the function works on it', () => {
  const { seen } = walk('=SUM(A1:A3)*2', { A1: 1, A2: 2, A3: 3 });
  assert.deepEqual(seen, ['=SUM([A1:A3])*2', '=[SUM({1;2;3})]*2', '=[6*2]', '=12']);
  const across = walk('=SUM(A1:C1)', { A1: 1, B1: 2.5, C1: 'x' }).seen;
  assert.equal(across[1], '=[SUM({1,2.5,"x"})]');
});

test('an IF shows the branch it takes and leaves the other as written', () => {
  const yes = walk('=IF(A1>0,B1,C1/0)', { A1: 4, B1: 'big', C1: 1 }).seen;
  assert.deepEqual(yes, ['=IF([A1]>0,B1,C1/0)', '=IF([4>0],B1,C1/0)', '=IF(TRUE,[B1],C1/0)', '=[IF(TRUE,"big",C1/0)]', '="big"']);
  const no = walk('=IF(A1>0,B1,C1)', { A1: -1, B1: 'big', C1: 'small' }).seen;
  assert.deepEqual(no.slice(2), ['=IF(FALSE,B1,[C1])', '=[IF(FALSE,B1,"small")]', '="small"']);
  // IFERROR looks at its fallback only when the value is an error.
  const fine = walk('=IFERROR(A1/B1,"none")', { A1: 6, B1: 3 }).seen;
  assert.deepEqual(fine, ['=IFERROR([A1]/B1,"none")', '=IFERROR(6/[B1],"none")', '=IFERROR([6/3],"none")', '=[IFERROR(2,"none")]', '=2']);
});

test('an error is shown where it arose and carried to the result', () => {
  const { seen, value } = walk('=A1/B1+1', { A1: 1 });
  assert.deepEqual(seen, ['=[A1]/B1+1', '=1/[B1]+1', '=[1/0]+1', '=[#DIV/0!+1]', '=#DIV/0!']);
  assert.equal(value.type, '#DIV/0!');
});

test('constants are not steps; a formula of constants is one step', () => {
  assert.deepEqual(walk('=-2+3', {}).seen, ['=[-2+3]', '=1']);
  assert.deepEqual(walk('=ROUND(2.345, 2)', {}).seen, ['=[ROUND(2.345, 2)]', '=2.35']);
  assert.deepEqual(walk('=(1+2)*3', {}).seen, ['=([1+2])*3', '=[(3)*3]', '=9']);
});

test('the text the formula was typed in is kept, spaces and case', () => {
  const { seen } = walk('=sum( A1 , 2 )', { A1: 1 });
  assert.deepEqual(seen, ['=sum( [A1] , 2 )', '=[sum( 1 , 2 )]', '=3']);
});

test('values are written as the dialog writes them', () => {
  assert.equal(evaluationText(0.1 + 0.2), '0.3');
  assert.equal(evaluationText('say "hi"'), '"say ""hi"""');
  assert.equal(evaluationText(true), 'TRUE');
  assert.equal(evaluationText(''), '0');
  assert.equal(evaluationText([[1, 2], [3, '']]), '{1,2;3,0}');
});

test('Step In opens a referenced cell\'s own formula; Step Out puts its value back', () => {
  const { resolver, getFormula } = book({ A1: 2, B1: '=A1*10', C1: '=B1+1' });
  const run = new FormulaEvaluation(resolver, { sheet: 'Sheet1', row: 0, col: 2 }, { getFormula });
  let st = run.state();
  assert.equal(st.levels.length, 1);
  assert.equal(st.levels[0].ref, 'Sheet1!$C$1');
  assert.equal(st.canStepIn, true, 'B1 holds a formula');
  run.stepIn();
  st = run.state();
  assert.equal(st.levels.length, 2);
  assert.equal(st.levels[1].ref, 'Sheet1!$B$1');
  assert.equal(st.levels[1].text, '=A1*10');
  assert.equal(st.canStepOut, true);
  assert.equal(st.canStepIn, true, 'A1 holds a constant, which Step In shows');
  run.stepIn();
  st = run.state();
  assert.equal(st.levels[2].constant, true);
  assert.equal(st.levels[2].text, '2');
  assert.match(st.message, /contains a constant/);
  run.stepOut();
  assert.equal(run.state().levels[1].text, '=2*10');
  run.evaluate(); // 2*10
  assert.equal(run.state().levels[1].text, '=20');
  run.evaluate(); // a finished inner level: Evaluate steps out
  st = run.state();
  assert.equal(st.levels.length, 1);
  assert.equal(st.levels[0].text, '=20+1');
  run.evaluate();
  st = run.state();
  assert.equal(st.levels[0].text, '=21');
  assert.equal(st.done, true);
  assert.equal(st.canEvaluate, false);
});

test('a list of presses replays to the same state, Restart included', () => {
  const { resolver, getFormula } = book({ A1: 2, B1: '=A1*10', C1: '=B1+1' });
  const at = { sheet: 'Sheet1', row: 0, col: 2 };
  const a = FormulaEvaluation.replay(resolver, at, ['stepIn', 'evaluate', 'evaluate'], { getFormula }).state();
  assert.equal(a.levels.length, 2);
  assert.equal(a.levels[1].text, '=20');
  const b = FormulaEvaluation.replay(resolver, at, ['evaluate', 'restart'], { getFormula }).state();
  assert.equal(b.levels[0].text, '=B1+1');
  assert.deepEqual(b.levels[0].underline, [1, 3]);
});

test('an empty cell cannot be stepped into, and a cell on another sheet names its sheet', () => {
  const s = new Spreadsheet();
  s.setByRef('Data', 'A1', 7);
  s.setByRef('My Sheet', 'A1', "=Data!A1+Z9");
  s.recalculate();
  const getFormula = (sheet, row, col) => s.cell(sheet, row, col)?.formula ?? null;
  const run = new FormulaEvaluation(s.resolver(), { sheet: 'My Sheet', row: 0, col: 0 }, { getFormula });
  assert.equal(run.state().levels[0].ref, "'My Sheet'!$A$1");
  assert.equal(run.state().canStepIn, true, 'Data!A1 holds 7');
  run.evaluate();
  assert.equal(run.state().canStepIn, false, 'Z9 is empty');
  assert.equal(run.state().levels[0].text, '=7+Z9');
});
