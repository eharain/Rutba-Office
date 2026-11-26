// IPMT, PPMT, CUMIPMT and CUMPRINC: a loan payment split into interest and
// capital, as a loan schedule shows it.
//
// The engine had PMT, PV, FV, NPER and RATE and none of these four, so a
// schedule could not be written the way Excel's is. The answers below are
// the ones Excel's own help pages give for the same formulas.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { FUNCTION_CATALOG } from '@rutba/formula';

const view = new SheetView(buildXlsx({ sheets: [{ name: 'S', rows: [['x']] }] }));
let at = 1;
const calc = (formula) => {
  view.setCell(at, 0, formula);
  return view.calc.getValue('S', at++, 0);
};
const near = (formula, want, places = 2) => {
  const got = calc(formula);
  assert.equal(typeof got, 'number', formula + ' gave ' + JSON.stringify(got));
  assert.ok(Math.abs(got - want) < 0.6 * 10 ** -places, formula + ' gave ' + got + ', Excel gives ' + want);
};
const code = (formula) => {
  const got = calc(formula);
  return got && typeof got === 'object' ? got.code ?? got.error ?? String(got) : got;
};

test('IPMT is the interest part of a payment', () => {
  near('=IPMT(0.1/12, 1, 3*12, 8000)', -66.67);
  near('=IPMT(0.1, 3, 3, 8000)', -292.45);
  near('=IPMT(0.1/12, 1, 36, 8000, 0, 1)', 0, 6);
});

test('PPMT is the capital part, and the two together are PMT', () => {
  near('=PPMT(0.1/12, 1, 2*12, 2000)', -75.62);
  near('=PPMT(0.08, 10, 10, 200000)', -27598.05);
  for (const per of [1, 7, 24]) {
    near(`=IPMT(0.1/12, ${per}, 24, 2000)+PPMT(0.1/12, ${per}, 24, 2000)-PMT(0.1/12, 24, 2000)`, 0, 8);
    near(`=IPMT(0.1/12, ${per}, 24, 2000, 0, 1)+PPMT(0.1/12, ${per}, 24, 2000, 0, 1)-PMT(0.1/12, 24, 2000, 0, 1)`, 0, 8);
  }
  // With no interest the whole payment is capital.
  near('=PPMT(0, 3, 10, 1000)', -100, 8);
  near('=IPMT(0, 3, 10, 1000)', 0, 8);
});

test('the capital parts of every payment repay the loan', () => {
  let repaid = 0;
  for (let per = 1; per <= 12; per++) repaid += calc(`=PPMT(0.06/12, ${per}, 12, 5000)`);
  assert.ok(Math.abs(repaid + 5000) < 1e-6, 'twelve capital parts came to ' + repaid);
});

test('CUMIPMT and CUMPRINC add the parts up over a run of periods', () => {
  near('=CUMIPMT(0.09/12, 30*12, 125000, 13, 24, 0)', -11135.23213, 4);
  near('=CUMIPMT(0.09/12, 30*12, 125000, 1, 1, 0)', -937.5, 4);
  near('=CUMPRINC(0.09/12, 30*12, 125000, 13, 24, 0)', -934.1071234, 4);
  near('=CUMPRINC(0.09/12, 30*12, 125000, 1, 1, 0)', -68.27827118, 4);
  near('=CUMPRINC(0.09/12, 360, 125000, 1, 360, 0)', -125000, 4);
  near('=CUMIPMT(0.09/12, 360, 125000, 1, 360, 1)+CUMPRINC(0.09/12, 360, 125000, 1, 360, 1)-360*PMT(0.09/12, 360, 125000, 0, 1)', 0, 4);
});

test('what Excel refuses is refused', () => {
  assert.match(String(code('=IPMT(0.1, 0, 3, 8000)')), /NUM/);
  assert.match(String(code('=PPMT(0.1, 4, 3, 8000)')), /NUM/);
  assert.match(String(code('=CUMIPMT(0, 12, 1000, 1, 12, 0)')), /NUM/);
  assert.match(String(code('=CUMPRINC(0.01, 12, 1000, 5, 2, 0)')), /NUM/);
  assert.match(String(code('=CUMPRINC(0.01, 12, 1000, 1, 13, 0)')), /NUM/);
  assert.match(String(code('=CUMIPMT(0.01, 12, 1000, 1, 12, 2)')), /NUM/);
  assert.match(String(code('=IPMT("words", 1, 3, 8000)')), /VALUE/);
});

test('the four are in the function list, so Insert Function and the hints know them', () => {
  const names = new Set(FUNCTION_CATALOG.map((f) => f.name));
  for (const name of ['IPMT', 'PPMT', 'CUMIPMT', 'CUMPRINC']) assert.ok(names.has(name), name);
});
