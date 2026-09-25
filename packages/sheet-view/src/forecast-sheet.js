/**
 * Data → Forecast Sheet: a timeline and its values, carried forward on a new
 * sheet as Excel's Forecast Sheet does — a table of the history with, past
 * its last point, FORECAST.ETS formulas and the confidence bounds either
 * side from FORECAST.ETS.CONFINT, and a chart of all of it.
 *
 * The table is Excel's: Timeline, Values, "Forecast(Values)", "Lower
 * Confidence Bound(Values)", "Upper Confidence Bound(Values)" — the column
 * names taken from the source's headers when it has them — with the last
 * actual point repeated in the three forecast columns, so the lines meet,
 * formatted as a table. The new sheet goes in front of the data's sheet, as
 * Excel puts it. The chart is a line chart (or columns, without the
 * bounds), its series written in Excel's colours with no markers.
 */
import { colName, ref } from './selection.js';
import { parseConsolidateRef, quoteSheet } from './consolidate.js';
import { prepareSeries, ForecastError, fitAAA, seasonFor, forecastAt, confidenceAt } from '@rutba/formula/forecast';
import { formatValue } from '@rutba/formula/numfmt';
import { buildChart, renderSvg } from '@rutba/drawing';
import { drawingAnchorXml, chartPartXml } from '@rutba/ooxml/build';

/** Excel's aggregation choices for duplicate times, by the code the formulas take. */
export const FORECAST_AGGREGATES = [
  [1, 'Average'], [2, 'Count'], [3, 'CountA'], [4, 'Max'], [5, 'Median'], [6, 'Min'], [7, 'Sum'],
];

/** A date serial's calendar parts, and back — Excel's 1900 system. */
const toDate = (serial) => new Date(Math.round((serial - 25569) * 86400000));
const toSerial = (d) => d.getTime() / 86400000 + 25569;

/**
 * The forecast's dates: a whole number of months or years on from the last
 * point when the timeline steps by months or years (month-start dates stay
 * month starts), else the step added again and again.
 */
function nextTimes(last, step, end, isDate) {
  const out = [];
  const monthly = isDate && step >= 27.5 && step <= 31.5;
  const yearly = isDate && step >= 364 && step <= 366.5;
  const quarterly = isDate && step >= 89 && step <= 92.5;
  const months = monthly ? 1 : quarterly ? 3 : yearly ? 12 : 0;
  for (let k = 1; k <= 5000; k++) {
    let t;
    if (months) {
      const d = toDate(last);
      const next = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + months * k, 1));
      const lastDay = new Date(Date.UTC(next.getUTCFullYear(), next.getUTCMonth() + 1, 0)).getUTCDate();
      next.setUTCDate(Math.min(d.getUTCDate(), lastDay));
      t = Math.round(toSerial(next));
    } else {
      t = last + k * step;
      if (isDate) t = Math.round(t * 1e6) / 1e6;
    }
    if (t > end + 1e-9) break;
    out.push(t);
  }
  return out;
}

/**
 * What the dialog starts from: the timeline and values the selection (or
 * the block round the active cell) holds — the first column the dates, the
 * next the numbers — whether they have headers, and a sensible end: a third
 * as far again as the history reaches.
 */
export function forecastDefaults(view) {
  const sheet = view.activeSheet;
  const r0 = view.selection.range;
  const r = r0.top === r0.bottom && r0.left === r0.right ? view._currentRegion(r0.top, r0.left) : r0;
  if (r.right - r.left < 1) return null;
  const value = (row, col) => view.calc.getValue(sheet, row, col);
  const header = typeof value(r.top, r.left) === 'string' && typeof value(r.top, r.left + 1) === 'string';
  const top = r.top + (header ? 1 : 0);
  const box = (col) => ({ sheet, top, left: col, bottom: r.bottom, right: col });
  const text = (b) => quoteSheet(b.sheet) + '!$' + colName(b.left) + '$' + (b.top + 1) + ':$' + colName(b.right) + '$' + (b.bottom + 1);
  const times = [];
  const values = [];
  for (let row = top; row <= r.bottom; row++) { times.push(value(row, r.left)); values.push(value(row, r.left + 1)); }
  let end = null;
  let last = null;
  let step = null;
  try {
    const s = prepareSeries(values, times);
    last = s.last;
    step = s.step;
    end = s.last + Math.max(1, Math.round((s.ys.length - 1) / 3)) * s.step;
  } catch { /* the dialog says what is wrong when it is used */ }
  const fmt = view.styleFor(top, r.left)?.numFmt ?? null;
  return {
    timeline: text(box(r.left)),
    values: text(box(r.left + 1)),
    header,
    last,
    step,
    end,
    isDate: /[dmy]/i.test(String(fmt ?? '')) && !/General/i.test(String(fmt)),
  };
}

/**
 * Create the forecast sheet. `timeline` and `values` are references of one
 * column each and the same length; `end` the last date (or number) to
 * forecast to; `kind` 'line' or 'column'; `confidence` a level between 0 and
 * 1, or null for none; `seasonality` 'auto' or a length (0 for none);
 * `completion` 1 to fill missing points by interpolation, 0 with zeros;
 * `aggregation` Excel's code for duplicate times. One undo step. Answers the
 * new sheet's name.
 */
export function createForecastSheet(view, spec = {}) {
  const plan = planForecast(view, spec);
  return writeForecastSheet(view, plan);
}

/**
 * The dialog's preview: the forecast worked out as the sheet's formulas
 * will work it out, drawn as the sheet's chart will be — or why it cannot be.
 */
export function forecastPreview(view, spec = {}, { width = 580, height = 250, mode = view.mode } = {}) {
  let plan;
  try {
    plan = planForecast(view, spec);
  } catch (err) {
    return { svg: '', error: String(err?.message || err) };
  }
  const { pairs, series, future, confidence, season, names, kind } = plan;
  const fit = fitAAA(series.ys, seasonFor(series.ys, season));
  const n = pairs.length + future.length;
  const col = () => new Array(n).fill(null);
  const actual = col();
  const fc = col();
  const low = col();
  const high = col();
  pairs.forEach((p, i) => { actual[i] = p.v === '' ? null : p.v; });
  const lastValue = actual[pairs.length - 1];
  fc[pairs.length - 1] = lastValue;
  low[pairs.length - 1] = lastValue;
  high[pairs.length - 1] = lastValue;
  future.forEach((t, k) => {
    const h = series.after(t);
    const v = forecastAt(fit, h);
    const at = pairs.length + k;
    fc[at] = v;
    if (confidence !== null) {
      const c = confidenceAt(fit, h, confidence);
      low[at] = v - c;
      high[at] = v + c;
    }
  });
  const code = plan.timeFormat || 'General';
  const categories = [...pairs.map((p) => p.t), ...future].map((t) => formatValue(t, code).text);
  const chartSeries = [{ name: names[1], values: actual, markers: false }, { name: names[2], values: fc, markers: false }];
  if (kind === 'line' && confidence !== null) chartSeries.push({ name: names[3], values: low, markers: false }, { name: names[4], values: high, markers: false });
  const svg = renderSvg(buildChart({ type: kind, categories, series: chartSeries, width, height, mode }));
  return { svg, error: '', points: future.length, season: fit.season >= 2 ? fit.season : 0 };
}

/** Everything the sheet and the preview are made from, checked (see createForecastSheet). */
function planForecast(view, {
  timeline, values, end, kind = 'line', confidence = 0.95, seasonality = 'auto', completion = 1, aggregation = 1,
} = {}) {
  if (!['line', 'column'].includes(kind)) throw new Error('A forecast is drawn as a line chart or a column chart.');
  if (confidence !== null && !(confidence > 0 && confidence < 1)) throw new Error('The confidence interval is a percentage between 0 and 100.');
  const season = seasonality === 'auto' ? 1 : Math.trunc(Number(seasonality));
  if (!(season === 0 || season === 1 || season >= 2)) throw new Error('Seasonality is detected automatically, or set to a whole number of points.');
  const tBox = parseConsolidateRef(timeline, view);
  const vBox = parseConsolidateRef(values, view);
  const rows = tBox.bottom - tBox.top + 1;
  if (tBox.left !== tBox.right || vBox.left !== vBox.right) throw new Error('The timeline and the values are one column each.');
  if (vBox.bottom - vBox.top + 1 !== rows) throw new Error('The timeline and the values must be the same size.');
  if (view.workbookProtection().structure) throw new Error('Workbook is protected and cannot be changed.');

  const src = view.calc;
  const pairs = [];
  for (let i = 0; i < rows; i++) {
    const t = src.getValue(tBox.sheet, tBox.top + i, tBox.left);
    const v = src.getValue(vBox.sheet, vBox.top + i, vBox.left);
    if (t === '' || t === null || t === undefined) continue;
    if (typeof t !== 'number') throw new Error('The timeline must be dates or numbers — ' + ref(tBox.top + i, tBox.left) + ' is not.');
    pairs.push({ t, v: typeof v === 'number' ? v : '' });
  }
  pairs.sort((a, b) => a.t - b.t);
  let series;
  try {
    series = prepareSeries(pairs.map((p) => p.v), pairs.map((p) => p.t), { completion, aggregation });
  } catch (err) {
    if (err instanceof ForecastError) throw new Error('The forecast cannot be made: ' + err.message + '.');
    throw err;
  }
  const stop = Number(end);
  if (!(stop > series.last)) throw new Error('The forecast end must come after the last point of the timeline.');
  const timeStyle = view._styleIndexAt(tBox.sheet, tBox.top, tBox.left);
  const valueStyle = view._styleIndexAt(vBox.sheet, vBox.top, vBox.left);
  const timeFormat = view.styleFor(tBox.top, tBox.left)?.numFmt;
  const isDate = pairs.every((p) => Number.isInteger(p.t)) && series.step >= 1 && /[dmy]/i.test(String(timeFormat ?? 'd'));
  const future = nextTimes(series.last, series.step, stop, isDate);
  if (!future.length) throw new Error('The forecast end must be at least one step after the last point.');

  // Column names: the source's headers when there is a row above the data.
  const headOf = (box, fallback) => {
    if (box.top === 0) return fallback;
    const h = src.getValue(box.sheet, box.top - 1, box.left);
    return typeof h === 'string' && h.trim() ? h.trim() : fallback;
  };
  const tName = headOf(tBox, 'Timeline');
  const vName = headOf(vBox, 'Values');
  const names = [tName, vName, 'Forecast(' + vName + ')'];
  if (confidence !== null) names.push('Lower Confidence Bound(' + vName + ')', 'Upper Confidence Bound(' + vName + ')');
  return {
    tBox, vBox, pairs, series, future, names, timeStyle, valueStyle, timeFormat,
    kind, confidence, season, completion, aggregation,
  };
}

/** The sheet, its table and its chart, as one undo step. */
function writeForecastSheet(view, {
  tBox, pairs, future, names, timeStyle, valueStyle, kind, confidence, season, completion, aggregation,
}) {
  const wb = view.workbook;
  const sourceSheet = tBox.sheet;
  const parts = [...new Set([...wb.sheets().map((s) => s.part), wb.mainPart, 'xl/_rels/workbook.xml.rels', '[Content_Types].xml'])];
  let made = null;
  // A failure part-way leaves nothing behind: the step it had begun is undone.
  const steps = view.history.past.length;
  try {
  view._edit('forecast sheet', null, [], () => {
    view._flushPendingEdits();
    view.dirtyCells.clear();
    view.styledCells.clear();
    const name = view._freshSheetName();
    wb.addSheet(name);
    wb.moveSheet(name, wb.sheetNames().indexOf(sourceSheet));
    const A1 = (r, c) => colName(c) + (r + 1);
    const put = (r, c, v) => wb.setCell(name, A1(r, c), v);
    const style = (r, c, s) => { if (s !== null && s !== undefined) wb.setCellStyle(name, A1(r, c), s); };
    names.forEach((h, c) => { put(0, c, h); style(0, c, view._boldIndex(null)); });
    const first = 1;
    const lastHist = first + pairs.length - 1;
    pairs.forEach((p, i) => {
      put(first + i, 0, p.t);
      style(first + i, 0, timeStyle);
      if (p.v !== '') put(first + i, 1, p.v);
      style(first + i, 1, valueStyle);
    });
    // The last actual point in the forecast columns, so the lines meet.
    for (let c = 2; c < names.length; c++) { put(lastHist, c, '=B' + (lastHist + 1)); style(lastHist, c, valueStyle); }
    const histV = '$B$' + (first + 1) + ':$B$' + (lastHist + 1);
    const histT = '$A$' + (first + 1) + ':$A$' + (lastHist + 1);
    const tail = ',' + season + ',' + completion + (aggregation !== 1 ? ',' + aggregation : '');
    const lastRow = lastHist + future.length;
    future.forEach((t, k) => {
      const r = lastHist + 1 + k;
      put(r, 0, t);
      style(r, 0, timeStyle);
      put(r, 2, '=FORECAST.ETS(A' + (r + 1) + ',' + histV + ',' + histT + tail + ')');
      style(r, 2, valueStyle);
      if (confidence !== null) {
        const ci = 'FORECAST.ETS.CONFINT(A' + (r + 1) + ',' + histV + ',' + histT + ',' + confidence + tail + ')';
        put(r, 3, '=C' + (r + 1) + '-' + ci);
        put(r, 4, '=C' + (r + 1) + '+' + ci);
        style(r, 3, valueStyle);
        style(r, 4, valueStyle);
      }
    });
    // Wide enough for the column names, as Excel sizes them.
    names.forEach((h, c) => wb.setColWidthChars(name, c, Math.max(11, Math.min(40, h.length + 3))));
    const tableRef = 'A1:' + A1(lastRow, names.length - 1);
    wb.addTable(name, tableRef, { style: 'TableStyleMedium2', headerNames: names });

    // The values, worked out, for the chart's caches.
    view._rebuildDerivedState();
    view.activeSheet = name;
    const q = quoteSheet(name);
    const col = (c) => q + '!$' + colName(c) + '$' + (first + 1) + ':$' + colName(c) + '$' + (lastRow + 1);
    const nums = (c) => {
      const out = [];
      for (let r = first; r <= lastRow; r++) {
        const v = view.calc.getValue(name, r, c);
        out.push(typeof v === 'number' && Number.isFinite(v) ? v : null);
      }
      return out;
    };
    const cats = [];
    for (let r = first; r <= lastRow; r++) cats.push(view.displayValue(r, 0).text);
    const colours = ['4472C4', 'ED7D31', 'ED7D31', 'ED7D31'];
    const drawn = kind === 'column' ? 3 : names.length;
    const chartSeries = [];
    for (let c = 1; c < drawn; c++) {
      chartSeries.push({
        name: names[c],
        nameRef: q + '!$' + colName(c) + '$1',
        ref: col(c),
        values: nums(c),
        ...(kind === 'line' ? { line: { colour: colours[c - 1], width: c >= 3 ? 12700 : 28575 }, marker: false } : {}),
      });
    }
    const drawingPart = wb.ensureSheetDrawing(name);
    const n = view.pkg.nextPartNumber('xl/charts/', 'chart');
    const chartPart = 'xl/charts/chart' + n + '.xml';
    view.pkg.addPart(chartPart,
      chartPartXml({ kind, categories: { ref: col(0), values: cats }, series: chartSeries }),
      'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
    const relId = view.pkg.addRelationshipTo(drawingPart,
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
      '../charts/chart' + n + '.xml');
    wb.appendDrawingAnchor(drawingPart, (id) => drawingAnchorXml({
      kind: 'chart', id, name: 'Forecast chart', from: { row: 1, col: names.length + 1 }, to: { row: 20, col: names.length + 9 },
    }, () => relId));
    view.drawings.set(name, view._readDrawings(wb.partNameFor(name)));
    // The new sheet's cells into its part now, so undo and redo carry the sheet as it was made.
    const madePart = wb.partNameFor(name);
    view.pkg.write_(madePart, wb.snapshotParts([madePart])[madePart]);
    view._structuralDirty = true;
    made = { sheet: name, rows: lastRow + 1, forecast: future.length, table: tableRef };
  }, { parts, structural: true, tracksNewParts: true, sheetGate: false });
  } catch (err) {
    if (view.history.past.length > steps) view.undo();
    throw err;
  }
  view.selectSheet(made.sheet);
  return made;
}
