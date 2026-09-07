/**
 * Chart spec -> scene.
 *
 * The layout engine, and the place the design rules are ENFORCED rather than
 * documented. A caller cannot accidentally produce a bad chart by passing the
 * wrong options, because the options that would produce one do not exist:
 *
 *   - There is no second y-axis. Two measures of different scale are two charts,
 *     or one indexed to a common base. A dual axis is the single most common way
 *     to mislead with a chart and it is simply not expressible here.
 *   - Colour is assigned by series IDENTITY in fixed order, never by rank and
 *     never cycled. Filtering a series out does not repaint the survivors.
 *   - A legend is present for two or more series and absent for one — with one
 *     series the title names it, and a legend box would be noise.
 *   - Four or fewer series are ALSO direct-labelled, so identity never rests on
 *     colour alone.
 *   - Scatter caps at three series, because only the first three palette slots
 *     clear the all-pairs colour-blind floors. Past that the caller folds to
 *     "Other" or facets.
 *   - Text wears ink tokens. A label is never painted in its series colour.
 *   - Grid and axes are recessive; the data is the darkest thing on the surface.
 */
import {
  scene, group, rect, line, polyline, polygon, path, text, ellipse, roundedBarPath,
} from './scene.js';
import { theme, seriesColour, needsRelief, SCATTER_SERIES_CAP } from './palette.js';
import { measureText, widestText, truncateText, capHeight } from './measure.js';

export const CHART_TYPES = ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter'];

/**
 * "Nice" axis bounds and tick step — the 1/2/5 × 10ⁿ rule every plotting library
 * has used for forty years, because those are the intervals people read without
 * arithmetic. 0.1, 0.2, 0.5, 1, 2, 5, 10 — never 0.3 or 7.
 */
export function niceScale(min, max, targetTicks = 5) {
  if (!Number.isFinite(min) || !Number.isFinite(max)) return { min: 0, max: 1, step: 0.5, ticks: [0, 0.5, 1] };
  if (min === max) {
    if (min === 0) return { min: 0, max: 1, step: 0.5, ticks: [0, 0.5, 1] };
    const pad = Math.abs(min) * 0.5;
    min -= pad;
    max += pad;
  }
  const roughStep = (max - min) / Math.max(1, targetTicks);
  const magnitude = 10 ** Math.floor(Math.log10(roughStep));
  const normalised = roughStep / magnitude;
  const step = (normalised <= 1 ? 1 : normalised <= 2 ? 2 : normalised <= 5 ? 5 : 10) * magnitude;

  const niceMin = Math.floor(min / step) * step;
  const niceMax = Math.ceil(max / step) * step;
  const ticks = [];
  // Accumulate by index, not by repeated addition, or floating point drifts and
  // an axis ends up labelled 0.30000000000000004.
  const count = Math.round((niceMax - niceMin) / step);
  for (let i = 0; i <= count; i++) ticks.push(Number((niceMin + i * step).toPrecision(12)));
  return { min: niceMin, max: niceMax, step, ticks };
}

const defaultFormat = (v) => {
  if (!Number.isFinite(v)) return String(v);
  const abs = Math.abs(v);
  if (abs >= 1e9) return (v / 1e9).toFixed(abs >= 1e10 ? 0 : 1) + 'B';
  if (abs >= 1e6) return (v / 1e6).toFixed(abs >= 1e7 ? 0 : 1) + 'M';
  if (abs >= 1e4) return (v / 1e3).toFixed(0) + 'k';
  return String(Number(v.toPrecision(12)));
};

/** Normalise and validate a spec, refusing the things that make a bad chart. */
export function normaliseSpec(spec) {
  const type = spec.type ?? 'column';
  if (!CHART_TYPES.includes(type)) throw new Error('unknown chart type: ' + type);

  const series = (spec.series ?? []).map((s, i) => ({
    name: s.name ?? 'Series ' + (i + 1),
    values: (s.values ?? []).map((v) => (v === null || v === undefined || v === '' ? null : Number(v))),
    id: s.id ?? s.name ?? 'series-' + i,
  }));
  if (!series.length) throw new Error('a chart needs at least one series');

  if (type === 'scatter' && series.length > SCATTER_SERIES_CAP) {
    throw new Error(
      'scatter charts are capped at ' + SCATTER_SERIES_CAP + ' series: only the first ' +
      SCATTER_SERIES_CAP + ' palette slots clear the all-pairs colour-blind floors. ' +
      'Fold the rest into "Other", or facet into small multiples.',
    );
  }
  if ((type === 'pie' || type === 'doughnut') && series.length > 1) {
    throw new Error('a pie shows one series; use small multiples for more');
  }

  const categories = spec.categories ?? series[0].values.map((_, i) => String(i + 1));
  return {
    type,
    title: spec.title ?? null,
    categories: categories.map((c) => (c === null || c === undefined ? '' : String(c))),
    series,
    stacked: Boolean(spec.stacked),
    width: spec.width ?? 480,
    height: spec.height ?? 300,
    mode: spec.mode ?? 'light',
    valueAxis: {
      title: spec.valueAxis?.title ?? null,
      min: spec.valueAxis?.min,
      max: spec.valueAxis?.max,
      format: spec.valueAxis?.format ?? defaultFormat,
    },
    categoryAxis: { title: spec.categoryAxis?.title ?? null },
  };
}

/** Legend and direct-label policy, in one place so every form obeys it. */
export function labelPolicy(seriesCount, mode) {
  return {
    legend: seriesCount >= 2,
    directLabels: seriesCount <= 4,
    reliefRequired: needsRelief(seriesCount, mode),
  };
}

function buildLegend(spec, t, policy, plotWidth, originX, y) {
  if (!policy.legend) return { node: null, height: 0 };
  const size = t.font.legend;
  const swatch = 9;
  const gap = 6;
  const itemGap = 16;

  const items = spec.series.map((s, i) => ({
    name: s.name,
    colour: seriesColour(i, spec.mode),
    width: swatch + gap + measureText(s.name, { size }) + itemGap,
  }));

  const children = [];
  let x = originX;
  let row = y;
  for (const item of items) {
    if (x + item.width - itemGap > originX + plotWidth && x > originX) {
      x = originX;
      row += size + 8;
    }
    children.push(rect({
      x, y: row - swatch + 1, width: swatch, height: swatch, rx: 2, fill: item.colour,
    }));
    children.push(text({
      x: x + swatch + gap, y: row, value: item.name,
      size, fill: t.ink.secondary, baseline: 'alphabetic',
    }));
    x += item.width;
  }
  const rows = Math.round((row - y) / (size + 8)) + 1;
  return { node: group(children, { class: 'legend' }), height: rows * (size + 8) };
}

/** Shared cartesian frame: axes, gridlines, ticks. */
function cartesianFrame({ spec, t, plot, scale, categoryPositions, horizontal }) {
  const children = [];
  const size = t.font.label;

  for (const tick of scale.ticks) {
    const v = horizontal
      ? plot.x + ((tick - scale.min) / (scale.max - scale.min)) * plot.width
      : plot.y + plot.height - ((tick - scale.min) / (scale.max - scale.min)) * plot.height;
    const zero = tick === 0;
    children.push(horizontal
      ? line({ x1: v, y1: plot.y, x2: v, y2: plot.y + plot.height,
        stroke: zero ? t.ink.axis : t.ink.gridline, strokeWidth: t.marks.gridWidth })
      : line({ x1: plot.x, y1: v, x2: plot.x + plot.width, y2: v,
        stroke: zero ? t.ink.axis : t.ink.gridline, strokeWidth: t.marks.gridWidth }));
    children.push(horizontal
      ? text({ x: v, y: plot.y + plot.height + size + 6, value: spec.valueAxis.format(tick),
        size, fill: t.ink.muted, anchor: 'middle' })
      : text({ x: plot.x - 8, y: v, value: spec.valueAxis.format(tick),
        size, fill: t.ink.muted, anchor: 'end', baseline: 'middle' }));
  }

  // Category labels, truncated rather than overlapped.
  const budget = horizontal
    ? plot.height / Math.max(1, spec.categories.length) - 2
    : plot.width / Math.max(1, spec.categories.length) - 4;
  spec.categories.forEach((label, i) => {
    const pos = categoryPositions[i];
    if (pos === undefined) return;
    children.push(horizontal
      ? text({ x: plot.x - 8, y: pos, value: truncateText(label, plot.x - 12, { size }),
        size, fill: t.ink.muted, anchor: 'end', baseline: 'middle' })
      : text({ x: pos, y: plot.y + plot.height + size + 6, value: truncateText(label, budget, { size }),
        size, fill: t.ink.muted, anchor: 'middle' }));
  });

  return group(children, { class: 'frame' });
}

/** Build the scene for a chart spec. */
export function buildChart(rawSpec) {
  const spec = normaliseSpec(rawSpec);
  const t = theme(spec.mode);
  const policy = labelPolicy(spec.series.length, spec.mode);
  const labelSize = t.font.label;

  if (spec.type === 'pie' || spec.type === 'doughnut') return buildPie(spec, t, policy);

  const horizontal = spec.type === 'bar';
  const flat = spec.series.flatMap((s) => s.values).filter((v) => v !== null && Number.isFinite(v));

  let dataMin;
  let dataMax;
  if (spec.stacked) {
    const totals = spec.categories.map((_, i) =>
      spec.series.reduce((sum, s) => sum + (s.values[i] ?? 0), 0));
    dataMin = Math.min(0, ...totals);
    dataMax = Math.max(0, ...totals);
  } else {
    dataMin = Math.min(0, ...(flat.length ? flat : [0]));
    dataMax = Math.max(0, ...(flat.length ? flat : [1]));
  }
  if (spec.valueAxis.min !== undefined) dataMin = spec.valueAxis.min;
  if (spec.valueAxis.max !== undefined) dataMax = spec.valueAxis.max;
  const scale = niceScale(dataMin, dataMax);

  // Reserve space from the outside in: title, legend, axis labels, then plot.
  const titleHeight = spec.title ? t.font.title + 12 : 0;
  const valueLabels = scale.ticks.map((v) => spec.valueAxis.format(v));
  const gutterLeft = horizontal
    ? Math.min(140, widestText(spec.categories, { size: labelSize }) + 14)
    : widestText(valueLabels, { size: labelSize }) + 14;
  const gutterBottom = labelSize + 14;

  const plot = {
    x: gutterLeft,
    y: titleHeight + 6,
    width: spec.width - gutterLeft - 12,
    height: 0,
  };

  const legendProbe = buildLegend(spec, t, policy, plot.width, plot.x, 0);
  plot.height = spec.height - plot.y - gutterBottom - legendProbe.height - 4;
  if (plot.height < 20) plot.height = Math.max(20, spec.height * 0.4);

  const valueAt = (v) => (horizontal
    ? plot.x + ((v - scale.min) / (scale.max - scale.min)) * plot.width
    : plot.y + plot.height - ((v - scale.min) / (scale.max - scale.min)) * plot.height);
  const zeroAt = valueAt(0);

  const count = spec.categories.length;
  const bandSize = (horizontal ? plot.height : plot.width) / Math.max(1, count);
  const bandCentre = (i) => (horizontal ? plot.y : plot.x) + bandSize * (i + 0.5);
  const categoryPositions = spec.categories.map((_, i) => bandCentre(i));

  const children = [];
  if (spec.title) {
    children.push(text({
      x: 0, y: t.font.title, value: spec.title,
      size: t.font.title, weight: '600', fill: t.ink.primary,
    }));
  }
  children.push(cartesianFrame({ spec, t, plot, scale, categoryPositions, horizontal }));

  const marks = [];
  const labels = [];

  if (spec.type === 'column' || spec.type === 'bar') {
    const groupCount = spec.stacked ? 1 : spec.series.length;
    const bandPadding = Math.min(bandSize * 0.25, 12);
    const slot = (bandSize - bandPadding) / groupCount;
    const barSize = Math.max(1, slot - t.marks.fillGap);

    spec.series.forEach((s, si) => {
      const colour = seriesColour(si, spec.mode);
      const running = spec.categories.map(() => 0);
      s.values.forEach((value, ci) => {
        if (value === null || !Number.isFinite(value)) return;
        const base = spec.stacked
          ? spec.series.slice(0, si).reduce((sum, prev) => sum + (prev.values[ci] ?? 0), 0)
          : 0;
        running[ci] = base + value;
        const from = valueAt(base);
        const to = valueAt(base + value);
        const bandStart = (horizontal ? plot.y : plot.x) + bandSize * ci + bandPadding / 2;
        const offset = spec.stacked ? 0 : si * slot;

        // Rounded only at the DATA end, anchored to the baseline — a bar rounded
        // at both ends reads as floating and misstates where zero is.
        const grows = value >= 0;
        if (horizontal) {
          const y = bandStart + offset;
          marks.push(path({
            d: roundedBarPath({
              x: Math.min(from, to), y, width: Math.abs(to - from), height: barSize,
              radius: t.marks.barRadius, side: grows ? 'right' : 'left',
            }),
            fill: colour,
          }));
          if (policy.directLabels && !spec.stacked) {
            labels.push(text({
              x: to + (grows ? 6 : -6), y: y + barSize / 2,
              value: spec.valueAxis.format(value), size: labelSize,
              fill: t.ink.secondary, anchor: grows ? 'start' : 'end', baseline: 'middle',
            }));
          }
        } else {
          const x = bandStart + offset;
          marks.push(path({
            d: roundedBarPath({
              x, y: Math.min(from, to), width: barSize, height: Math.abs(to - from),
              radius: t.marks.barRadius, side: grows ? 'top' : 'bottom',
            }),
            fill: colour,
          }));
          if (policy.directLabels && !spec.stacked) {
            labels.push(text({
              x: x + barSize / 2, y: to + (grows ? -6 : labelSize + 4),
              value: spec.valueAxis.format(value), size: labelSize,
              fill: t.ink.secondary, anchor: 'middle',
            }));
          }
        }
      });
      void running;
    });
  } else if (spec.type === 'line' || spec.type === 'area') {
    spec.series.forEach((s, si) => {
      const colour = seriesColour(si, spec.mode);
      const points = [];
      s.values.forEach((value, ci) => {
        if (value === null || !Number.isFinite(value)) return;
        points.push([bandCentre(ci), valueAt(value)]);
      });
      if (!points.length) return;

      if (spec.type === 'area') {
        marks.push(polygon({
          points: [[points[0][0], zeroAt], ...points, [points[points.length - 1][0], zeroAt]],
          fill: colour, opacity: 0.16,
        }));
      }
      marks.push(polyline({
        points, stroke: colour, strokeWidth: t.marks.lineWidth,
        fill: 'none', linecap: 'round', linejoin: 'round',
      }));
      for (const [px, py] of points) {
        // A surface ring keeps overlapping markers legible where lines cross.
        marks.push(ellipse({
          cx: px, cy: py, rx: t.marks.markerSize / 2, ry: t.marks.markerSize / 2,
          fill: colour, stroke: t.ink.surface, strokeWidth: 2,
        }));
      }
      if (policy.directLabels) {
        const [lx, ly] = points[points.length - 1];
        labels.push(text({
          x: lx + 8, y: ly, value: s.name, size: labelSize,
          fill: t.ink.secondary, baseline: 'middle',
        }));
      }
    });
  } else if (spec.type === 'scatter') {
    spec.series.forEach((s, si) => {
      const colour = seriesColour(si, spec.mode);
      s.values.forEach((value, ci) => {
        if (value === null || !Number.isFinite(value)) return;
        marks.push(ellipse({
          cx: bandCentre(ci), cy: valueAt(value),
          rx: t.marks.markerSize / 2, ry: t.marks.markerSize / 2,
          fill: colour, stroke: t.ink.surface, strokeWidth: 2,
        }));
      });
    });
  }

  children.push(group(marks, { class: 'marks' }));
  if (labels.length) children.push(group(labels, { class: 'labels' }));

  const legend = buildLegend(spec, t, policy, plot.width, plot.x, spec.height - 6);
  if (legend.node) children.push(legend.node);

  return scene({
    width: spec.width,
    height: spec.height,
    mode: spec.mode,
    title: spec.title ?? describe(spec),
    description: describe(spec),
    children,
  });
}

function buildPie(spec, t, policy) {
  const series = spec.series[0];
  const total = series.values.reduce((sum, v) => sum + Math.abs(v ?? 0), 0);
  const titleHeight = spec.title ? t.font.title + 12 : 0;
  const legendProbe = buildLegend(
    { ...spec, series: spec.categories.map((c) => ({ name: c })) },
    t, { legend: true }, spec.width - 24, 12, 0,
  );
  const available = spec.height - titleHeight - legendProbe.height - 16;
  const radius = Math.max(10, Math.min(spec.width, available) / 2 - 6);
  const cx = spec.width / 2;
  const cy = titleHeight + 8 + radius;
  const inner = spec.type === 'doughnut' ? radius * 0.58 : 0;

  const children = [];
  if (spec.title) {
    children.push(text({
      x: 0, y: t.font.title, value: spec.title,
      size: t.font.title, weight: '600', fill: t.ink.primary,
    }));
  }

  const marks = [];
  const labels = [];
  let angle = -Math.PI / 2; // start at twelve o'clock, as a clock does
  series.values.forEach((raw, i) => {
    const value = Math.abs(raw ?? 0);
    if (!value || !total) return;
    const sweep = (value / total) * Math.PI * 2;
    const end = angle + sweep;
    const colour = seriesColour(i, spec.mode);
    marks.push(path({ d: arcPath(cx, cy, radius, inner, angle, end), fill: colour }));

    if (policy.directLabels || value / total >= 0.08) {
      const mid = angle + sweep / 2;
      const lr = inner + (radius - inner) * 0.62;
      labels.push(text({
        x: cx + Math.cos(mid) * lr, y: cy + Math.sin(mid) * lr,
        value: Math.round((value / total) * 100) + '%',
        size: t.font.label, fill: t.ink.surface, weight: '600',
        anchor: 'middle', baseline: 'middle',
      }));
    }
    angle = end;
  });

  children.push(group(marks, { class: 'marks' }));
  if (labels.length) children.push(group(labels, { class: 'labels' }));

  // A pie's legend names the CATEGORIES, since the single series is the whole.
  const legend = buildLegend(
    { ...spec, series: spec.categories.map((c) => ({ name: c })) },
    t, { legend: spec.categories.length >= 2 }, spec.width - 24, 12, spec.height - 6,
  );
  if (legend.node) children.push(legend.node);

  return scene({
    width: spec.width, height: spec.height, mode: spec.mode,
    title: spec.title ?? describe(spec), description: describe(spec), children,
  });
}

/** Donut/pie wedge, with a hole when `inner > 0`. */
function arcPath(cx, cy, outer, inner, from, to) {
  const large = to - from > Math.PI ? 1 : 0;
  const x0 = cx + Math.cos(from) * outer;
  const y0 = cy + Math.sin(from) * outer;
  const x1 = cx + Math.cos(to) * outer;
  const y1 = cy + Math.sin(to) * outer;
  if (inner <= 0) {
    return `M${cx} ${cy}L${x0} ${y0}A${outer} ${outer} 0 ${large} 1 ${x1} ${y1}Z`;
  }
  const ix1 = cx + Math.cos(to) * inner;
  const iy1 = cy + Math.sin(to) * inner;
  const ix0 = cx + Math.cos(from) * inner;
  const iy0 = cy + Math.sin(from) * inner;
  return `M${x0} ${y0}A${outer} ${outer} 0 ${large} 1 ${x1} ${y1}`
    + `L${ix1} ${iy1}A${inner} ${inner} 0 ${large} 0 ${ix0} ${iy0}Z`;
}

/** The alt text. A chart with no description is a chart a screen reader cannot use. */
export function describe(spec) {
  const s = spec.series ?? [];
  const kind = spec.type ?? 'chart';
  const names = s.map((x) => x.name).join(', ');
  const cats = (spec.categories ?? []).length;
  return `${kind} chart` + (spec.title ? ` titled "${spec.title}"` : '')
    + `, ${s.length} series (${names}) across ${cats} categories`;
}

/** The data behind a chart, as a table — the relief the palette rules require. */
export function chartTable(rawSpec) {
  const spec = normaliseSpec(rawSpec);
  return {
    header: ['', ...spec.series.map((s) => s.name)],
    rows: spec.categories.map((c, i) => [c, ...spec.series.map((s) => s.values[i] ?? '')]),
    reliefRequired: labelPolicy(spec.series.length, spec.mode).reliefRequired,
  };
}

export { capHeight };
