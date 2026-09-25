/**
 * FORECAST.LINEAR, FORECAST.ETS, FORECAST.ETS.CONFINT and
 * FORECAST.ETS.SEASONALITY.
 *
 * Excel's exponential smoothing is its AAA model (additive error, trend and
 * season); Microsoft does not publish how it picks the smoothing constants,
 * so these tests hold the results to being sensible — a straight line goes
 * on straight, a season is found and carried forward, the interval widens
 * the further out it looks — and to being the same every time. The method
 * is written out at the head of packages/formula/src/forecast.js.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { fitAAA, forecastAt, confidenceAt, detectSeasonality, prepareSeries, normalQuantile } from '../packages/formula/src/forecast.js';

/** A sheet with a timeline in A and values in B from row 2, and formulas to evaluate in D. */
function sheetWith(times, values, formulas) {
  const rows = [['Month', 'Sales', '', 'Result']];
  const n = Math.max(times.length, formulas.length);
  for (let i = 0; i < n; i++) rows.push([times[i] ?? '', values[i] ?? '', '', formulas[i] ?? '']);
  const view = new SheetView(buildXlsx({ sheets: [{ name: 'Data', rows }] }));
  return (i) => view.calc.getValue('Data', i + 1, 3);
}

const seasonal = (n, { period = 12, base = 100, slope = 2, amp = 30, noise = 0 } = {}) =>
  Array.from({ length: n }, (_, i) => base + slope * i + amp * Math.sin((2 * Math.PI * i) / period) + noise * Math.sin(i * 7.3));

test('FORECAST.LINEAR and FORECAST are the least-squares line, as Excel\'s are', () => {
  const get = sheetWith([1, 2, 3, 4, 5], [6, 5, 11, 7, 5], [
    '=FORECAST.LINEAR(30,B2:B6,A2:A6)', '=FORECAST(30,B2:B6,A2:A6)', '=FORECAST.LINEAR(1,B2:B3,C2:C3)',
  ]);
  // Excel's own documented example: x 20,28,31,38,40 / y 6,7,9,15,21 gives 10.607253 at 30.
  const doc = sheetWith([20, 28, 31, 38, 40], [6, 7, 9, 15, 21], ['=FORECAST.LINEAR(30,B2:B6,A2:A6)']);
  assert.ok(Math.abs(doc(0) - 10.607253) < 1e-6, String(doc(0)));
  assert.equal(get(0), get(1), 'FORECAST is FORECAST.LINEAR');
  assert.equal(get(2)?.type, '#DIV/0!', 'no pairs of numbers');
});

test('a straight line goes on straight: ETS forecasts a trend without a season', () => {
  const times = Array.from({ length: 24 }, (_, i) => i + 1);
  const values = times.map((t) => 50 + 3 * t);
  const series = prepareSeries(values, times);
  assert.equal(detectSeasonality(series.ys), 0, 'no season in a line');
  const fit = fitAAA(series.ys, 0);
  for (const h of [1, 6, 12]) assert.ok(Math.abs(forecastAt(fit, h) - (50 + 3 * (24 + h))) < 0.5, `h ${h}: ${forecastAt(fit, h)}`);
  const get = sheetWith(times, values, ['=FORECAST.ETS(30,B2:B25,A2:A25)', '=FORECAST.ETS.SEASONALITY(B2:B25,A2:A25)']);
  assert.ok(Math.abs(get(0) - 140) < 0.5, String(get(0)));
  assert.equal(get(1), 0);
});

test('a seasonal series: the season is found, carried forward, and the forecast follows the pattern', () => {
  const values = seasonal(48, { noise: 1.5 });
  const times = values.map((_, i) => i + 1);
  const get = sheetWith(times, values, [
    '=FORECAST.ETS.SEASONALITY(B2:B49,A2:A49)',
    '=FORECAST.ETS(52,B2:B49,A2:A49)',
    '=FORECAST.ETS(58,B2:B49,A2:A49)',
    '=FORECAST.ETS(52,B2:B49,A2:A49,0)',
    '=FORECAST.ETS(52,B2:B49,A2:A49,12)',
  ]);
  assert.equal(get(0), 12, 'a twelve-step season');
  const truth = (t) => 100 + 2 * (t - 1) + 30 * Math.sin((2 * Math.PI * (t - 1)) / 12);
  assert.ok(Math.abs(get(1) - truth(52)) < 6, `t=52: ${get(1)} against ${truth(52)}`);
  assert.ok(Math.abs(get(2) - truth(58)) < 8, `t=58: ${get(2)} against ${truth(58)}`);
  assert.ok(get(1) > get(2), 'the peak and the trough of the season, in their places');
  assert.ok(Math.abs(get(4) - get(1)) < 1e-9, 'seasonality 12 set by hand is the one detected');
  assert.ok(Math.abs(get(3) - truth(52)) > Math.abs(get(1) - truth(52)), 'without a season the forecast misses the pattern');
});

test('the confidence interval is positive, widens the further out it looks, and follows the level asked for', () => {
  const values = seasonal(36, { noise: 4 });
  const times = values.map((_, i) => i + 1);
  const get = sheetWith(times, values, [
    '=FORECAST.ETS.CONFINT(37,B2:B37,A2:A37)',
    '=FORECAST.ETS.CONFINT(48,B2:B37,A2:A37)',
    '=FORECAST.ETS.CONFINT(48,B2:B37,A2:A37,0.8)',
    '=FORECAST.ETS.CONFINT(48,B2:B37,A2:A37,1)',
  ]);
  assert.ok(get(0) > 0);
  assert.ok(get(1) > get(0), 'wider twelve steps out than one');
  assert.ok(get(2) < get(1), 'an 80% interval is narrower than a 95% one');
  assert.ok(Math.abs(get(2) / get(1) - normalQuantile(0.9) / normalQuantile(0.975)) < 1e-9, 'by the ratio of the normal quantiles');
  assert.equal(get(3)?.type, '#NUM!', 'a confidence of 1 is refused');
  assert.ok(Math.abs(normalQuantile(0.975) - 1.959963985) < 1e-8);
});

test('the same inputs give the same forecast, every time, cached or not', () => {
  const values = seasonal(40, { period: 7, noise: 3 });
  const series = prepareSeries(values, values.map((_, i) => i));
  const a = fitAAA(series.ys, 7);
  const b = fitAAA([...series.ys], 7);
  assert.equal(a, b, 'the fit is kept for the next cell that asks');
  const fresh = fitAAA(series.ys.map((y) => y + 0), 7);
  assert.deepEqual([fresh.alpha, fresh.beta, fresh.gamma], [a.alpha, a.beta, a.gamma]);
  assert.equal(forecastAt(a, 3.5), (forecastAt(a, 3) + forecastAt(a, 4)) / 2, 'a date between two steps is interpolated');
  assert.ok(confidenceAt(a, 5) > 0);
});

test('Excel\'s refusals: a timeline without a steady step, a target before its end, sizes that differ; dates by the month; duplicates aggregated; gaps filled', () => {
  const get = sheetWith([1, 2, 3.5, 4, 5, 6], [5, 6, 7, 8, 9, 10], [
    '=FORECAST.ETS(9,B2:B7,A2:A7)',
    '=FORECAST.ETS(3,B2:B6,A2:A6)',
    '=FORECAST.ETS(9,B2:B7,A2:A6)',
  ]);
  assert.equal(get(0)?.type, '#NUM!', 'an inconsistent step');
  assert.equal(get(2)?.type, '#N/A', 'values and timeline of different sizes');
  const ok = sheetWith([1, 2, 3, 4, 5, 6], [5, 6, 7, 8, 9, 10], ['=FORECAST.ETS(3,B2:B7,A2:A7)']);
  assert.equal(ok(0)?.type, '#NUM!', 'a target before the end of the timeline');

  // Month-start dates: 28 to 31 days apart, one step each.
  const months = Array.from({ length: 24 }, (_, i) => Date.UTC(2024, i, 1) / 86400000 + 25569);
  const monthly = prepareSeries(months.map((_, i) => 10 + i), months);
  assert.equal(monthly.ys.length, 24);
  // A repeated time is averaged by default and summed with aggregation 7; a missing one is filled.
  const dup = prepareSeries([4, 6, 10, 14], [1, 1, 2, 4]);
  assert.deepEqual(dup.ys, [5, 10, 12, 14]);
  assert.deepEqual(prepareSeries([4, 6, 10, 14], [1, 1, 2, 4], { aggregation: 7 }).ys, [10, 10, 12, 14]);
  assert.deepEqual(prepareSeries([4, 6, 10, 14], [1, 1, 2, 4], { completion: 0 }).ys, [5, 10, 0, 14], 'zeros when asked');
});
