/**
 * Chart layers — bars, columns, lines, areas, pies and donuts.
 *
 * WHAT MAKES CHARTS LOOK HARD IS NOT THE DRAWING. A bar chart is rectangles and
 * a donut is arcs; both are less code than the word-art warps already here.
 * What is hard is where the numbers come from, and that is a separate question
 * (T4, data binding) with its own design to settle. So this takes its data from
 * the patch: a chart you type four numbers into is useful on the day it ships,
 * and it is the same painter a bound chart will use once there is something to
 * bind to. Nothing here needs revisiting when that lands — a binding resolves
 * to `series` and everything below is unchanged.
 *
 * SINGLE SERIES FOR BARS, MULTIPLE FOR LINES, and that asymmetry is deliberate
 * rather than unfinished. Two lines on one axis is a comparison anybody reads
 * instantly; two *bars* per category is a grouped bar chart, which needs its
 * own spacing model, its own legend placement and its own label strategy. Pies
 * are single-series by nature. When grouped bars are wanted they are a real
 * feature, not a loop — and pretending otherwise would ship something that
 * looks right with two series and falls apart with four.
 *
 * AXES ARE DRAWN, NOT LAID OUT. There is no tick engine: the scale is a nice
 * round maximum and a fixed number of gridlines, because a creative is read at
 * a glance from a phone and a chart with eleven labelled ticks on it is a
 * spreadsheet screenshot. `niceMax` is what keeps the top of the axis a number
 * a person would have chosen.
 *
 * Everything is fractional (fx/fy/fw/fh of the frame) like every other layer.
 */

import { registerLayerType } from './index.js';

const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
const num = (v, fallback) => (Number.isFinite(Number(v)) ? Number(v) : fallback);
const TAU = Math.PI * 2;

const FAMILY = '"Segoe UI", system-ui, Roboto, "Helvetica Neue", Arial, sans-serif';

/** The chart kinds this layer can take. */
export const CHART_KINDS = ['bar', 'column', 'line', 'area', 'pie', 'donut'];

export const CHART_LABELS = {
    bar: 'Bars', column: 'Columns', line: 'Line', area: 'Area', pie: 'Pie', donut: 'Donut',
};

/**
 * The default palette — THE SUITE'S, not this file's.
 *
 * These are `@rutba/drawing`'s validated categorical slots (CVD-checked,
 * contrast-checked, the validator invocation recorded in its
 * `src/palette.js` header), VENDORED here because this package is zero-dep
 * by contract — the render worker runs it unchanged, and an import across
 * the shelf would end that. The commons rule the values serve: the palette
 * is the one thing every chart surface MUST share, so a chart made here
 * matches one made in Sheets. `tests/charts.test.mjs` asserts this copy
 * equals the source, so the two cannot drift silently.
 *
 * SLOT ORDER IS THE COLOUR-BLIND-SAFETY MECHANISM, not decoration — hues are
 * assigned in fixed order and never reordered or cycled by design intent
 * (past the eighth, `colorAt` wraps rather than inventing a hue, which is
 * this renderer's cheaper answer to "fold into Other").
 *
 * A fixed set rather than a ramp derived from the theme, because a chart's
 * colours have to be TELLABLE APART first and on-brand second — a ramp of one
 * hue makes the third and fourth slice a guess. A brand kit overrides it, which
 * is the point at which "on brand" becomes somebody's actual decision.
 */
export const CHART_PALETTES = {
    light: ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#4a3aa7', '#e34948'],
    dark: ['#3987e5', '#d95926', '#199e70', '#c98500', '#d55181', '#008300', '#9085e9', '#e66767'],
};

/** The canonical set — what swatch rows show. Painting picks by theme. */
export const CHART_PALETTE = CHART_PALETTES.light;

// ── scale ────────────────────────────────────────────────────────────────────

/**
 * Round a maximum up to a number a person would have chosen.
 *
 * 47 → 50, 230 → 250, 1.7 → 2. Without this the top gridline reads "47" and
 * the chart looks like it was measured rather than designed.
 */
export function niceMax(value) {
    const v = Math.abs(num(value, 0));
    if (v <= 0) return 1;
    const mag = 10 ** Math.floor(Math.log10(v));
    const norm = v / mag;
    const step = norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10;
    return step * mag;
}

/** Every value across every series, for the axis. */
function extent(series) {
    let max = 0;
    let min = 0;
    for (const s of series) {
        for (const v of s.values) {
            if (!Number.isFinite(v)) continue;
            if (v > max) max = v;
            if (v < min) min = v;
        }
    }
    return { min, max };
}

// ── data ─────────────────────────────────────────────────────────────────────

/**
 * Normalise whatever the patch carries into `{ labels, series }`.
 *
 * Accepts the shorthand a template or a person is likely to write — a bare
 * array of numbers — as well as the full form, because refusing the short one
 * would mean the simplest possible chart is the one that needs the most typing.
 */
export function readData(patch) {
    // `categories` is accepted as an alias of `labels` — it is the word
    // `@rutba/drawing` uses for the same axis in Sheets, and one bound
    // dataset should feed a sheet chart and a creative chart unchanged.
    // `labels` wins when both are present, because it is this layer's own
    // spelling and the one the inspector writes.
    const rawLabels = Array.isArray(patch.labels) ? patch.labels
        : (Array.isArray(patch.categories) ? patch.categories : []);
    const labels = rawLabels.map((l) => String(l ?? ''));

    let series = [];
    if (Array.isArray(patch.series) && patch.series.length && typeof patch.series[0] === 'object') {
        series = patch.series.map((s, i) => ({
            name: String(s.name ?? `Series ${i + 1}`),
            values: (Array.isArray(s.values) ? s.values : []).map((v) => num(v, 0)),
            color: s.color || null,
        }));
    } else if (Array.isArray(patch.values)) {
        series = [{ name: patch.name || 'Series 1', values: patch.values.map((v) => num(v, 0)), color: null }];
    }

    const count = series.reduce((m, s) => Math.max(m, s.values.length), 0);
    // Labels and values are edited separately and drift; the chart draws what
    // it can rather than throwing away the whole thing over a missing label.
    const padded = Array.from({ length: count }, (_, i) => labels[i] ?? '');
    return { labels: padded, series, count };
}

/** The colour for slice/series `i`. */
const colorAt = (palette, i, override) => override || palette[i % palette.length];

// ── painting ─────────────────────────────────────────────────────────────────

function paintChart(ctx, plan, layer) {
    const { labels, series, count } = readData(layer);
    if (!count || !series.length) return;

    const w = plan.W * (layer.fw == null ? 0.7 : layer.fw);
    const h = layer.fh ? plan.H * layer.fh : w * 0.62;
    if (w <= 0 || h <= 0) return;

    const sizePx = Math.max(8, Math.round(plan.W * num(layer.sizeFrac, 0.022)));
    // The theme decides which validated variant paints — each is tuned for
    // its surface (the light set warns below 3:1 on three slots; the dark
    // set clears every check). An explicit layer palette still wins: that is
    // a brand kit or a person having decided.
    const palette = Array.isArray(layer.palette) && layer.palette.length
        ? layer.palette
        : (CHART_PALETTES[plan.opts?.theme] || CHART_PALETTE);
    const ink = layer.ink || plan.theme.text;
    const dim = layer.dim || plan.theme.dim;

    ctx.save();
    ctx.translate(layer.x - w / 2, layer.y - h / 2);
    ctx.font = `500 ${sizePx}px ${FAMILY}`;
    ctx.textBaseline = 'middle';

    const kind = CHART_KINDS.includes(layer.chart) ? layer.chart : 'bar';
    const args = { ctx, plan, layer, w, h, sizePx, palette, ink, dim, labels, series, count };

    if (kind === 'pie' || kind === 'donut') paintPie(args, kind === 'donut');
    else if (kind === 'line' || kind === 'area') paintLine(args, kind === 'area');
    else paintBars(args, kind === 'bar');

    ctx.restore();
}

/**
 * Bars (horizontal) and columns (vertical).
 *
 * One function because they differ in which axis carries the category — the
 * spacing, the labels and the value placement are the same decisions twice.
 */
function paintBars({ ctx, layer, w, h, sizePx, palette, ink, dim, labels, series, count }, horizontal) {
    const values = series[0].values;
    const showLabels = layer.showLabels !== false && labels.some(Boolean);
    const showValues = layer.showValues !== false;
    const showGrid = layer.showGrid !== false;

    const { max } = extent(series);
    const top = niceMax(layer.max != null ? layer.max : max);
    const gap = clamp(num(layer.gap, 0.28), 0, 0.8);

    // Room for the category labels, along whichever axis carries them.
    const labelSpace = showLabels ? (horizontal ? Math.min(w * 0.34, longest(ctx, labels) + sizePx * 0.6) : sizePx * 1.8) : 0;
    const valueSpace = showValues ? sizePx * (horizontal ? 2.4 : 1.6) : 0;

    const plotX = horizontal ? labelSpace : 0;
    const plotY = 0;
    const plotW = horizontal ? w - labelSpace - valueSpace : w;
    const plotH = horizontal ? h : h - labelSpace - valueSpace;
    if (plotW <= 0 || plotH <= 0) return;

    if (showGrid) drawGrid(ctx, plotX, plotY, plotW, plotH, dim, horizontal);

    const band = (horizontal ? plotH : plotW) / count;
    const thick = band * (1 - gap);

    for (let i = 0; i < count; i++) {
        const v = num(values[i], 0);
        const frac = top > 0 ? clamp(v / top, 0, 1) : 0;
        const color = colorAt(palette, i, series[0].color);
        const at = (horizontal ? plotY : plotX) + band * i + (band - thick) / 2;

        ctx.fillStyle = color;
        const r = Math.min(thick * 0.18, num(layer.radius, 0.18) * thick);
        if (horizontal) {
            roundBar(ctx, plotX, at, plotW * frac, thick, r, 'right');
        } else {
            const barH = plotH * frac;
            roundBar(ctx, at, plotY + plotH - barH, thick, barH, r, 'up');
        }

        if (showLabels && labels[i]) {
            ctx.fillStyle = dim;
            if (horizontal) {
                ctx.textAlign = 'right';
                ctx.fillText(labels[i], plotX - sizePx * 0.4, at + thick / 2);
            } else {
                ctx.textAlign = 'center';
                ctx.fillText(labels[i], at + thick / 2, plotY + plotH + sizePx * 0.9);
            }
        }

        if (showValues) {
            ctx.fillStyle = ink;
            const text = formatValue(v, layer);
            if (horizontal) {
                ctx.textAlign = 'left';
                ctx.fillText(text, plotX + plotW * frac + sizePx * 0.4, at + thick / 2);
            } else {
                ctx.textAlign = 'center';
                ctx.fillText(text, at + thick / 2, plotY + plotH - plotH * frac - sizePx * 0.8);
            }
        }
    }
}

/** Lines and areas — several series over one axis. */
function paintLine({ ctx, layer, w, h, sizePx, palette, ink, dim, labels, series, count }, filled) {
    const showLabels = layer.showLabels !== false && labels.some(Boolean);
    const showGrid = layer.showGrid !== false;
    const showValues = layer.showValues === true; // off by default: a line with
    // a number over every point is unreadable, which is the opposite of a bar
    // chart where the number IS the point.

    const { max } = extent(series);
    const top = niceMax(layer.max != null ? layer.max : max);
    const labelSpace = showLabels ? sizePx * 1.8 : 0;
    const plotH = h - labelSpace;
    if (plotH <= 0 || count < 2) return;

    if (showGrid) drawGrid(ctx, 0, 0, w, plotH, dim, false);

    const stepX = count > 1 ? w / (count - 1) : w;
    const yOf = (v) => plotH - (top > 0 ? clamp(num(v, 0) / top, 0, 1) : 0) * plotH;

    series.forEach((s, si) => {
        const color = colorAt(palette, si, s.color);
        const pts = s.values.slice(0, count).map((v, i) => [i * stepX, yOf(v)]);
        if (pts.length < 2) return;

        if (filled) {
            ctx.beginPath();
            ctx.moveTo(pts[0][0], plotH);
            for (const [x, y] of pts) ctx.lineTo(x, y);
            ctx.lineTo(pts[pts.length - 1][0], plotH);
            ctx.closePath();
            // A translucent fill so overlapping series stay readable — an opaque
            // area chart with two series hides one of them completely.
            ctx.globalAlpha = clamp(num(layer.fillOpacity, 0.35), 0, 1);
            ctx.fillStyle = color;
            ctx.fill();
            ctx.globalAlpha = 1;
        }

        ctx.beginPath();
        for (const [i, [x, y]] of pts.entries()) { if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y); }
        ctx.strokeStyle = color;
        ctx.lineWidth = Math.max(1, num(layer.lineWidth, 0.006) * (w / 0.7));
        ctx.lineJoin = 'round';
        ctx.lineCap = 'round';
        ctx.stroke();

        if (layer.showPoints !== false) {
            ctx.fillStyle = color;
            const r = Math.max(2, sizePx * 0.22);
            for (const [x, y] of pts) { ctx.beginPath(); ctx.ellipse(x, y, r, r, 0, 0, TAU); ctx.fill(); }
        }

        if (showValues) {
            ctx.fillStyle = ink;
            ctx.textAlign = 'center';
            for (const [i, [x, y]] of pts.entries()) ctx.fillText(formatValue(s.values[i], layer), x, y - sizePx);
        }
    });

    if (showLabels) {
        ctx.fillStyle = dim;
        for (let i = 0; i < count; i++) {
            if (!labels[i]) continue;
            // The end labels are pulled inside the plot, or the first and last
            // hang off the edges of the layer's own box.
            ctx.textAlign = i === 0 ? 'left' : i === count - 1 ? 'right' : 'center';
            ctx.fillText(labels[i], i * stepX, plotH + sizePx * 0.9);
        }
    }
}

/** Pies and donuts. */
function paintPie({ ctx, layer, w, h, sizePx, palette, ink, labels, series, count }, donut) {
    const values = series[0].values.slice(0, count).map((v) => Math.max(0, num(v, 0)));
    const total = values.reduce((a, b) => a + b, 0);
    if (total <= 0) return;

    const showLegend = layer.showLegend !== false && labels.some(Boolean);
    const legendW = showLegend ? Math.min(w * 0.42, longest(ctx, labels) + sizePx * 3.2) : 0;
    const plotW = w - legendW;
    const r = Math.min(plotW, h) / 2;
    const cx = r;
    const cy = h / 2;

    // From twelve o'clock, clockwise — the direction every pie chart is read in.
    let a0 = -Math.PI / 2;
    values.forEach((v, i) => {
        const sweep = (v / total) * TAU;
        ctx.beginPath();
        ctx.moveTo(cx, cy);
        ctx.arc(cx, cy, r, a0, a0 + sweep);
        ctx.closePath();
        ctx.fillStyle = colorAt(palette, i, null);
        ctx.fill();
        a0 += sweep;
    });

    if (donut) {
        // The hole is punched with a second fill in the layer's own background
        // rather than with even-odd, because a slice and the hole are separate
        // paths drawn at different times — there is no single path to punch.
        const inner = clamp(num(layer.innerRatio, 0.58), 0.05, 0.95) * r;
        ctx.save();
        ctx.globalCompositeOperation = 'destination-out';
        ctx.beginPath();
        ctx.ellipse(cx, cy, inner, inner, 0, 0, TAU);
        ctx.fill();
        ctx.restore();

        if (layer.centerText) {
            ctx.fillStyle = ink;
            ctx.textAlign = 'center';
            ctx.font = `700 ${Math.round(sizePx * 1.7)}px ${FAMILY}`;
            ctx.fillText(String(layer.centerText), cx, cy);
            ctx.font = `500 ${sizePx}px ${FAMILY}`;
        }
    }

    if (showLegend) {
        ctx.textAlign = 'left';
        const rowH = sizePx * 1.7;
        const startY = cy - (count * rowH) / 2 + rowH / 2;
        const chip = sizePx * 0.8;
        for (let i = 0; i < count; i++) {
            const y = startY + i * rowH;
            ctx.fillStyle = colorAt(palette, i, null);
            ctx.beginPath();
            if (typeof ctx.roundRect === 'function') ctx.roundRect(plotW + sizePx * 0.4, y - chip / 2, chip, chip, chip * 0.3);
            else ctx.rect(plotW + sizePx * 0.4, y - chip / 2, chip, chip);
            ctx.fill();
            ctx.fillStyle = ink;
            const pct = layer.showValues !== false ? `  ${Math.round((values[i] / total) * 100)}%` : '';
            ctx.fillText(`${labels[i] || ''}${pct}`, plotW + sizePx * 1.7, y);
        }
    }
}

// ── small helpers ────────────────────────────────────────────────────────────

const longest = (ctx, list) => list.reduce((m, s) => Math.max(m, ctx.measureText(String(s ?? '')).width), 0);

function formatValue(v, layer) {
    const n = num(v, 0);
    const dp = clamp(Math.round(num(layer.decimals, 0)), 0, 4);
    const body = n.toFixed(dp);
    return `${layer.prefix || ''}${body}${layer.suffix || ''}`;
}

/** A few faint gridlines, along the value axis only. */
function drawGrid(ctx, x, y, w, h, color, horizontal) {
    ctx.save();
    ctx.strokeStyle = color;
    ctx.globalAlpha = 0.25;
    ctx.lineWidth = 1;
    const N = 4;
    ctx.beginPath();
    for (let i = 0; i <= N; i++) {
        if (horizontal) {
            const gx = x + (w / N) * i;
            ctx.moveTo(gx, y); ctx.lineTo(gx, y + h);
        } else {
            const gy = y + (h / N) * i;
            ctx.moveTo(x, gy); ctx.lineTo(x + w, gy);
        }
    }
    ctx.stroke();
    ctx.restore();
}

/**
 * A bar with its FAR end rounded and its base square.
 *
 * Rounding both ends makes a short bar a lozenge floating off its axis; the
 * base has to stay flat or the chart stops reading as measured from a line.
 */
function roundBar(ctx, x, y, w, h, r, towards) {
    const rad = Math.max(0, Math.min(r, towards === 'up' ? h : w, (towards === 'up' ? w : h) / 2));
    ctx.beginPath();
    if (rad <= 0.5) { ctx.rect(x, y, Math.max(0, w), Math.max(0, h)); ctx.fill(); return; }
    if (towards === 'up') {
        ctx.moveTo(x, y + h);
        ctx.lineTo(x, y + rad);
        ctx.quadraticCurveTo(x, y, x + rad, y);
        ctx.lineTo(x + w - rad, y);
        ctx.quadraticCurveTo(x + w, y, x + w, y + rad);
        ctx.lineTo(x + w, y + h);
    } else {
        ctx.moveTo(x, y);
        ctx.lineTo(x + w - rad, y);
        ctx.quadraticCurveTo(x + w, y, x + w, y + rad);
        ctx.lineTo(x + w, y + h - rad);
        ctx.quadraticCurveTo(x + w, y + h, x + w - rad, y + h);
        ctx.lineTo(x, y + h);
    }
    ctx.closePath();
    ctx.fill();
}

// ── the stored form ──────────────────────────────────────────────────────────

function chartBounds(ctx, plan, layer) {
    const w = plan.W * (layer.fw == null ? 0.7 : layer.fw);
    const h = layer.fh ? plan.H * layer.fh : w * 0.62;
    return { x: layer.x - w / 2, y: layer.y - h / 2, w, h };
}

function compileChart(patch, cx) {
    const { W, H } = cx;
    const env = cx.envelope(patch.anim);

    return {
        name: patch.name || 'Chart',
        chart: CHART_KINDS.includes(patch.chart) ? patch.chart : 'bar',
        // Carried through verbatim; readData normalises at paint time so a
        // template that stored the shorthand keeps working.
        labels: patch.labels,
        series: patch.series,
        values: patch.values,
        visible: patch.visible !== false,
        timing: patch.timing || null,
        enter: patch.enter || env.enter,
        exit: patch.exit || env.exit,
        z: patch.z == null ? cx.nextZ() : patch.z,
        x: W * (patch.fx == null ? 0.5 : patch.fx),
        y: H * (patch.fy == null ? 0.5 : patch.fy),
        fx: patch.fx == null ? 0.5 : patch.fx,
        fy: patch.fy == null ? 0.5 : patch.fy,
        fw: patch.fw == null ? 0.7 : patch.fw,
        fh: patch.fh == null ? null : num(patch.fh, 0.3),
        sizeFrac: num(patch.sizeFrac, 0.022),
        palette: Array.isArray(patch.palette) && patch.palette.length ? patch.palette : null,
        ink: patch.ink || null,
        dim: patch.dim || null,
        showLabels: patch.showLabels !== false,
        showValues: patch.showValues,
        showGrid: patch.showGrid !== false,
        showLegend: patch.showLegend !== false,
        showPoints: patch.showPoints !== false,
        gap: num(patch.gap, 0.28),
        radius: num(patch.radius, 0.18),
        innerRatio: num(patch.innerRatio, 0.58),
        lineWidth: num(patch.lineWidth, 0.006),
        fillOpacity: num(patch.fillOpacity, 0.35),
        max: patch.max == null ? null : num(patch.max, null),
        decimals: num(patch.decimals, 0),
        prefix: patch.prefix || '',
        suffix: patch.suffix || '',
        centerText: patch.centerText || '',
        rot: patch.rot || 0,
        opacity: patch.opacity == null ? 1 : patch.opacity,
        blend: patch.blend || null,
        mask: patch.mask || null,
        anim: patch.anim || 'none',
        keys: cx.keys(patch.keys),
    };
}

export const CHART = registerLayerType('chart', {
    paint: paintChart,
    compile: compileChart,
    bounds: chartBounds,
});
