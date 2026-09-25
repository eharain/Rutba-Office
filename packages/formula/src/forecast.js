/**
 * Forecasting: FORECAST.LINEAR and the exponential-smoothing family —
 * FORECAST.ETS, FORECAST.ETS.CONFINT and FORECAST.ETS.SEASONALITY.
 *
 * Excel's ETS functions are its "AAA" version of exponential triple
 * smoothing: additive error, additive trend, additive seasonality (Holt and
 * Winters' method). Microsoft documents the model and the arguments but not
 * how it chooses the smoothing constants, so the numbers here are sensible
 * and deterministic rather than Excel's to the last digit. The method, in
 * full:
 *
 *   1. The series. Values are paired with the timeline; duplicate times are
 *      aggregated (1 average, 2 count, 3 counta, 4 max, 5 median, 6 min,
 *      7 sum); the step is the typical gap between times, refined so the
 *      whole span is a whole number of steps (month-start dates are one
 *      step apart though a month is 28 to 31 days); a time more than a third
 *      of a step off the grid is an inconsistent timeline (#NUM!). Missing
 *      points are filled by straight lines between their neighbours
 *      (data completion 1) or with zeros (0).
 *   2. Seasonality, when asked to detect it: the series less its straight-
 *      line trend, correlated with itself at every lag from 2 to half its
 *      length; the season is the lag whose correlation is the highest local
 *      peak above 0.3 (and above 2/√n), else there is none.
 *   3. The fit. The level and trend start from the first two seasons (or
 *      the first points' straight line), the seasonal indices from the
 *      first season's distance from that line. α, β and γ minimise the sum
 *      of squared one-step errors: a coarse grid, then coordinate descent
 *      with halving steps — the same search every time, so the same inputs
 *      give the same forecast.
 *   4. The forecast h steps past the last point is level + h·trend + the
 *      season's index; a target between two steps is interpolated.
 *   5. The confidence interval is z·σ·√(1 + Σ c_j²) for j = 1…h−1, with
 *      c_j = α(1 + jβ) + γ when j is a whole number of seasons (Hyndman and
 *      others' variance for the additive model), σ² the fit's mean squared
 *      one-step error and z the normal quantile of (1 + confidence) / 2.
 */

/** Excel's aggregation codes, for times that repeat. */
const AGGREGATE = {
  1: (xs) => xs.reduce((a, b) => a + b, 0) / xs.length,
  2: (xs) => xs.length,
  3: (xs) => xs.length,
  4: (xs) => Math.max(...xs),
  5: (xs) => {
    const s = [...xs].sort((a, b) => a - b);
    const mid = s.length >> 1;
    return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
  },
  6: (xs) => Math.min(...xs),
  7: (xs) => xs.reduce((a, b) => a + b, 0),
};

/** A plain failure, carried as the Excel error it becomes. */
class ForecastError extends Error {
  constructor(type, message) {
    super(message);
    this.type = type;
  }
}

const num = (v) => typeof v === 'number' && Number.isFinite(v);
const blank = (v) => v === null || v === undefined || v === '';

/**
 * The series on an even grid: `ys` from the first time to the last, one a
 * step, filled; `t0`, `step`. Throws a ForecastError for Excel's refusals.
 */
export function prepareSeries(values, timeline, { completion = 1, aggregation = 1 } = {}) {
  if (values.length !== timeline.length) throw new ForecastError('#N/A', 'the values and the timeline are not the same size');
  if (!AGGREGATE[aggregation]) throw new ForecastError('#NUM!', 'aggregation is 1 to 7');
  if (completion !== 0 && completion !== 1) throw new ForecastError('#NUM!', 'data completion is 0 or 1');
  const byTime = new Map();
  for (let i = 0; i < timeline.length; i++) {
    const t = timeline[i];
    const v = values[i];
    if (blank(t)) continue;
    if (!num(t)) throw new ForecastError('#VALUE!', 'the timeline holds something that is not a number or a date');
    if (!blank(v) && !num(v)) throw new ForecastError('#VALUE!', 'the values hold something that is not a number');
    if (!byTime.has(t)) byTime.set(t, []);
    if (num(v)) byTime.get(t).push(v);
  }
  const times = [...byTime.keys()].sort((a, b) => a - b);
  if (times.length < 3) throw new ForecastError('#NUM!', 'a forecast needs at least three points in time');
  const diffs = [];
  for (let i = 1; i < times.length; i++) diffs.push(times[i] - times[i - 1]);
  const span = times[times.length - 1] - times[0];
  const t0 = times[0];
  // Where each time sits on the grid, in steps from the first. Gaps that
  // differ by little (a month is 28 to 31 days) are one step each; dates a
  // whole number of months apart are counted in months, gaps and all;
  // otherwise the commonest gap is the step (the smaller of two as common),
  // so a gap left by a missing point does not stretch it.
  const lo = Math.min(...diffs);
  const hi = Math.max(...diffs);
  let positionOf;
  let steps;
  const dayOf = (t) => new Date(Math.round((t - 25569) * 86400000));
  const monthIndex = (t) => { const d = dayOf(t); return d.getUTCFullYear() * 12 + d.getUTCMonth(); };
  const monthly = times.every((t) => Number.isInteger(t)) && lo >= 28 && diffs.every((d) => {
    const months = Math.round(d / 30.44);
    return months >= 1 && Math.abs(d - months * 30.44) <= 3.5;
  }) && new Set(times.map((t) => Math.min(dayOf(t).getUTCDate(), 28))).size === 1;
  if (!(lo > 0)) throw new ForecastError('#NUM!', 'the timeline has no step');
  if (hi / lo <= 1.12) {
    steps = times.length - 1;
    positionOf = (t) => (t - t0) / (span / steps);
  } else if (monthly) {
    steps = monthIndex(times[times.length - 1]) - monthIndex(t0);
    positionOf = (t) => monthIndex(t) - monthIndex(t0);
  } else {
    const counts = new Map();
    for (const d of diffs) { const k = Math.round(d * 1e9) / 1e9; counts.set(k, (counts.get(k) || 0) + 1); }
    let typical = 0;
    let most = 0;
    for (const [d, c] of [...counts].sort((a, b) => a[0] - b[0])) if (c > most) { most = c; typical = d; }
    steps = Math.max(1, Math.round(span / typical));
    positionOf = (t) => (t - t0) / (span / steps);
  }
  const step = span / steps;
  const known = new Map();
  for (const t of times) {
    const pos = positionOf(t);
    const at = Math.round(pos);
    if (Math.abs(pos - at) > 0.34) throw new ForecastError('#NUM!', 'the timeline does not keep a constant step');
    const got = byTime.get(t);
    if (!got.length) continue;
    const value = aggregation === 2 || aggregation === 3 ? got.length : AGGREGATE[aggregation](got);
    known.set(at, known.has(at) ? (known.get(at) + value) / 2 : value);
  }
  if (known.size < 2) throw new ForecastError('#NUM!', 'a forecast needs at least two values');
  const ys = new Array(steps + 1);
  const at = [...known.keys()].sort((a, b) => a - b);
  for (let i = 0; i <= steps; i++) {
    if (known.has(i)) { ys[i] = known.get(i); continue; }
    if (completion === 0) { ys[i] = 0; continue; }
    // A straight line between the neighbours; past the ends, the nearest.
    let lo = null;
    let hi = null;
    for (const k of at) { if (k < i) lo = k; else if (k > i && hi === null) hi = k; }
    if (lo === null) ys[i] = known.get(hi);
    else if (hi === null) ys[i] = known.get(lo);
    else ys[i] = known.get(lo) + ((known.get(hi) - known.get(lo)) * (i - lo)) / (hi - lo);
  }
  const last = t0 + steps * step;
  // How many steps a target lies past the last point: in months for a
  // monthly timeline (so the first of next month is one step on), else by
  // the step.
  const after = monthly
    ? (t) => (Number.isInteger(t) && dayOf(t).getUTCDate() === dayOf(last).getUTCDate() ? monthIndex(t) - monthIndex(last) : (t - last) / step)
    : (t) => (t - last) / step;
  return { ys, t0, step, last, after };
}

/** The straight line through the points (x = 0, 1, 2…): { slope, intercept }. */
function lineThrough(ys) {
  const n = ys.length;
  const mx = (n - 1) / 2;
  const my = ys.reduce((a, b) => a + b, 0) / n;
  let sxy = 0;
  let sxx = 0;
  for (let i = 0; i < n; i++) { sxy += (i - mx) * (ys[i] - my); sxx += (i - mx) ** 2; }
  const slope = sxx ? sxy / sxx : 0;
  return { slope, intercept: my - slope * mx };
}

/** The season length the series repeats over, or 0 when it has none (step 2 above). */
export function detectSeasonality(ys) {
  const n = ys.length;
  if (n < 6) return 0;
  const { slope, intercept } = lineThrough(ys);
  const r = ys.map((y, i) => y - (intercept + slope * i));
  const denom = r.reduce((a, b) => a + b * b, 0);
  if (!(denom > 1e-12)) return 0;
  const maxLag = Math.min(8760, Math.floor(n / 2));
  const acf = [1];
  for (let k = 1; k <= maxLag + 1; k++) {
    let s = 0;
    for (let i = k; i < n; i++) s += r[i] * r[i - k];
    acf[k] = k < n ? s / denom : 0;
  }
  const floor = Math.max(0.3, 2 / Math.sqrt(n));
  let best = 0;
  let bestValue = -Infinity;
  for (let k = 2; k <= maxLag; k++) {
    const peak = acf[k] > acf[k - 1] && acf[k] >= (acf[k + 1] ?? -Infinity);
    if (!peak || acf[k] <= floor) continue;
    if (acf[k] > bestValue + 1e-9) { best = k; bestValue = acf[k]; }
  }
  return best;
}

/** One pass of the additive Holt-Winters recursion; the state at the end and the squared errors. */
function runAAA(ys, m, alpha, beta, gamma, init) {
  const seasonal = m >= 2;
  let level = init.level;
  let trend = init.trend;
  const season = seasonal ? [...init.season] : null;
  let sse = 0;
  let count = 0;
  const skip = seasonal ? m : 2;
  for (let t = 0; t < ys.length; t++) {
    const s = seasonal ? season[t % m] : 0;
    const err = ys[t] - (level + trend + s);
    if (t >= skip) { sse += err * err; count += 1; }
    const next = alpha * (ys[t] - s) + (1 - alpha) * (level + trend);
    trend = beta * (next - level) + (1 - beta) * trend;
    if (seasonal) season[t % m] = gamma * (ys[t] - next) + (1 - gamma) * s;
    level = next;
  }
  return { level, trend, season, sse, count };
}

/** Where the recursion starts (step 3 above). */
function initialState(ys, m) {
  if (m >= 2) {
    const mean = (a, b) => { let s = 0; for (let i = a; i < b; i++) s += ys[i]; return s / (b - a); };
    const first = mean(0, m);
    const second = mean(m, 2 * m);
    const trend = (second - first) / m;
    const lineAt = (i) => first + (i - (m - 1) / 2) * trend;
    let season = Array.from({ length: m }, (_, i) => ys[i] - lineAt(i));
    const avg = season.reduce((a, b) => a + b, 0) / m;
    season = season.map((s) => s - avg);
    return { level: lineAt(0) - trend, trend, season };
  }
  const head = ys.slice(0, Math.min(ys.length, 10));
  const { slope, intercept } = lineThrough(head);
  return { level: intercept - slope, trend: slope, season: null };
}

const fits = new Map();

/**
 * The fitted model for a series and a season length (0 or 1: none): the
 * smoothing constants, the state after the last point and σ². Kept for the
 * next cell that asks about the same series — a forecast sheet asks the same
 * question in every row.
 */
export function fitAAA(ys, m) {
  const season = m >= 2 ? m : 1;
  const key = season + '|' + ys.join(',');
  const hit = fits.get(key);
  if (hit) return hit;
  const init = initialState(ys, season);
  const seasonal = season >= 2;
  const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
  const cost = (p) => {
    const r = runAAA(ys, season, p[0], p[1], p[2], init);
    return r.count ? r.sse / r.count : 0;
  };
  // A coarse grid, then coordinate descent with halving steps.
  let best = null;
  let bestCost = Infinity;
  for (const a of [0.1, 0.3, 0.5, 0.7, 0.9]) {
    for (const b of [0.01, 0.05, 0.15, 0.3]) {
      for (const g of seasonal ? [0.01, 0.1, 0.3, 0.6] : [0]) {
        const c = cost([a, b, g]);
        if (c < bestCost - 1e-12) { bestCost = c; best = [a, b, g]; }
      }
    }
  }
  const bounds = [[0.001, 0.999], [0.001, 0.999], [seasonal ? 0.001 : 0, seasonal ? 0.999 : 0]];
  for (let stepSize = 0.1; stepSize > 0.002; stepSize /= 2) {
    for (let round = 0; round < 2; round++) {
      for (let k = 0; k < 3; k++) {
        if (k === 2 && !seasonal) continue;
        for (const dir of [-1, 1]) {
          const trial = [...best];
          trial[k] = clamp(trial[k] + dir * stepSize, bounds[k][0], bounds[k][1]);
          const c = cost(trial);
          if (c < bestCost - 1e-12) { bestCost = c; best = trial; }
        }
      }
    }
  }
  const [alpha, beta, gamma] = best;
  const end = runAAA(ys, season, alpha, beta, gamma, init);
  const fit = {
    season, alpha, beta, gamma, n: ys.length,
    level: end.level, trend: end.trend, indices: end.season,
    sigma2: end.count ? end.sse / end.count : 0,
  };
  if (fits.size > 48) fits.delete(fits.keys().next().value);
  fits.set(key, fit);
  return fit;
}

/** The fitted value h steps past the last point (h may be fractional; 0 is the last point smoothed). */
export function forecastAt(fit, h) {
  const at = (k) => {
    const s = fit.season >= 2 ? fit.indices[((fit.n + k - 1) % fit.season + fit.season) % fit.season] : 0;
    return fit.level + k * fit.trend + s;
  };
  if (h <= 0) return at(0);
  const lo = Math.floor(h);
  const hi = Math.ceil(h);
  if (lo === hi) return at(lo);
  return at(lo) + (at(hi) - at(lo)) * (h - lo);
}

/** Half the width of the interval round the forecast h steps ahead (step 5 above). */
export function confidenceAt(fit, h, level = 0.95) {
  const steps = Math.max(1, Math.ceil(h - 1e-9));
  let sum = 1;
  for (let j = 1; j < steps; j++) {
    const c = fit.alpha * (1 + j * fit.beta) + (fit.season >= 2 && j % fit.season === 0 ? fit.gamma : 0);
    sum += c * c;
  }
  return normalQuantile((1 + level) / 2) * Math.sqrt(fit.sigma2 * sum);
}

/** The standard normal's quantile — Acklam's rational approximation (relative error under 1.2e-9). */
export function normalQuantile(p) {
  if (!(p > 0 && p < 1)) return NaN;
  const a = [-3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2, 1.383577518672690e2, -3.066479806614716e1, 2.506628277459239];
  const b = [-5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2, 6.680131188771972e1, -1.328068155288572e1];
  const c = [-7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
  const d = [7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996, 3.754408661907416];
  const lowP = 0.02425;
  let x;
  if (p < lowP) {
    const q = Math.sqrt(-2 * Math.log(p));
    x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  } else if (p <= 1 - lowP) {
    const q = p - 0.5;
    const r = q * q;
    x = ((((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q) / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
  } else {
    const q = Math.sqrt(-2 * Math.log(1 - p));
    x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  }
  return x;
}

/**
 * The season length to use: `seasonality` as Excel takes it — 1 (or none
 * given) to detect it, 0 for none, 2 and up for that length — checked
 * against the series (two seasons of data at least).
 */
export function seasonFor(ys, seasonality = 1) {
  if (!Number.isFinite(seasonality) || seasonality < 0 || seasonality > 8760) throw new ForecastError('#NUM!', 'seasonality is 0, 1 or a length up to 8760');
  const s = Math.trunc(seasonality);
  if (s === 1) return detectSeasonality(ys);
  if (s === 0) return 0;
  if (ys.length < 2 * s) throw new ForecastError('#NUM!', 'a season of ' + s + ' needs at least two seasons of data');
  return s;
}

/** Everything a FORECAST.ETS call needs, or the Excel error it answers with. */
export function etsQuestion({ target, values, timeline, seasonality = 1, completion = 1, aggregation = 1 }) {
  const series = prepareSeries(values, timeline, { completion, aggregation });
  if (!num(target)) throw new ForecastError('#VALUE!', 'the target date is not a number or a date');
  const h = series.after(target);
  if (h < -1e-9) throw new ForecastError('#NUM!', 'the target date is before the end of the timeline');
  const m = seasonFor(series.ys, seasonality);
  return { fit: fitAAA(series.ys, m), h: Math.max(0, h), series };
}

/** FORECAST.LINEAR's line: the value at x of the least-squares line through the pairs. */
export function linearForecast(x, ys, xs) {
  if (ys.length !== xs.length) throw new ForecastError('#N/A', 'known_y\'s and known_x\'s are not the same size');
  const pairs = [];
  for (let i = 0; i < ys.length; i++) if (num(ys[i]) && num(xs[i])) pairs.push([xs[i], ys[i]]);
  if (!pairs.length) throw new ForecastError('#DIV/0!', 'no pairs of numbers');
  const mx = pairs.reduce((a, p) => a + p[0], 0) / pairs.length;
  const my = pairs.reduce((a, p) => a + p[1], 0) / pairs.length;
  let sxx = 0;
  let sxy = 0;
  for (const [px, py] of pairs) { sxx += (px - mx) ** 2; sxy += (px - mx) * (py - my); }
  if (sxx === 0) throw new ForecastError('#DIV/0!', 'the known x values do not vary');
  const b = sxy / sxx;
  return my - b * mx + b * x;
}

export { ForecastError };
