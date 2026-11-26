/**
 * Chart Design: a chart's axis titles, data labels, legend and series
 * colours, read from its chart part and written into it the way Excel writes
 * each, so a chart formatted here opens formatted there and the other way.
 *
 * Pure: XML in, XML out. The elements go where the schema orders them,
 * since Excel refuses a chart part whose children are out of order.
 *
 * A legend and data labels are each a choice the part either makes or leaves
 * open. A part written here before either was a choice says nothing of them,
 * and its chart is drawn by the house rule (a legend from two series, values
 * on a few bars). The first time someone chooses, that rule is written down
 * as the chart stood (`settle`), so nothing else on the chart moves, and from
 * then on the part is read as it is written, the way a chart from Excel is.
 */

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const unesc = (s) => String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/**
 * The kinds of plot a chart is read as, in the order the reading looks for
 * them (the first one a part holds is the one drawn), and whether a series
 * there is a filled shape, a line or a ring of slices.
 */
const PLOTS = {
  barChart: 'fill', bar3DChart: 'fill',
  lineChart: 'line', line3DChart: 'line',
  areaChart: 'fill', area3DChart: 'fill',
  pieChart: 'slices', pie3DChart: 'slices', doughnutChart: 'slices',
  scatterChart: 'line', bubbleChart: 'fill',
};

export const LEGEND_PLACES = ['right', 'top', 'left', 'bottom'];
const LEGEND_WRITE = { right: 'r', left: 'l', top: 't', bottom: 'b' };
const LEGEND_READ = { r: 'right', tr: 'right', l: 'left', t: 'top', b: 'bottom' };

/** The prefix chart elements wear in this part: `c:` in every file Excel writes, none in some others. */
function prefixOf(xml) {
  const m = /<(\w+:)?chartSpace\b/.exec(String(xml ?? ''));
  if (!m) throw new Error('this chart could not be read');
  return m[1] ?? '';
}

/** Every element called `name`, with what it holds: `{ start, end, xml }`, in the order of the file. */
function elementsOf(xml, p, name) {
  const re = new RegExp('<' + p + name + '(?:\\s[^>]*)?(?:/>|>[\\s\\S]*?</' + p + name + '>)', 'g');
  return [...xml.matchAll(re)].map((m) => ({ start: m.index, end: m.index + m[0].length, xml: m[0] }));
}

/** The words of a title: every run's text, joined. */
const wordsOf = (xml) => [...xml.matchAll(/<(?:\w+:)?t>([\s\S]*?)<\/(?:\w+:)?t>/g)].map((m) => unesc(m[1])).join('');

/** The `a:` prefix the words and colours are in, declared where the chart space has none. */
function withDrawingNs(xml) {
  if (/xmlns:a="/.test(xml)) return xml;
  return xml.replace(/<(\w+:)?chartSpace\b/, (m) => m + ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"');
}

/**
 * The two axes a chart names: the one its categories lie along and the one
 * its values are measured on. A scatter chart has two value axes, the first
 * along the foot. A pie has neither.
 */
function axesOf(xml, p) {
  const list = [];
  for (const name of ['catAx', 'dateAx', 'valAx']) {
    for (const el of elementsOf(xml, p, name)) list.push({ ...el, name });
  }
  list.sort((a, b) => a.start - b.start);
  const cats = list.filter((a) => a.name !== 'valAx');
  const vals = list.filter((a) => a.name === 'valAx');
  return {
    category: cats[0] ?? (vals.length > 1 ? vals[0] : null),
    value: cats.length ? vals[0] ?? null : vals[vals.length - 1] ?? null,
  };
}

/** The plot that is drawn: the first of the kinds above that the part holds. */
function plotOf(xml, p) {
  for (const [name, mark] of Object.entries(PLOTS)) {
    const el = elementsOf(xml, p, name)[0];
    if (el) return { ...el, name, mark };
  }
  return null;
}

/** The series of the plot that is drawn, in the order of the file: the order the chart's reading lists them in. */
function seriesOf(xml, p) {
  const plot = plotOf(xml, p);
  if (!plot) return [];
  return elementsOf(plot.xml, p, 'ser').map((ser) => ({
    start: plot.start + ser.start, end: plot.start + ser.end, xml: ser.xml, mark: plot.mark,
  }));
}

/**
 * Where a series' own children end and the parts with shapes of their own
 * begin: its marker, its single points, its labels, its values. A `spPr`
 * before this is the series' own; one after belongs to one of those.
 */
function ownHead(ser, p) {
  let end = ser.length;
  for (const name of ['marker', 'dPt', 'dLbls', 'trendline', 'errBars', 'cat', 'val', 'xVal', 'yVal']) {
    const at = ser.search(new RegExp('<' + p + name + '[\\s>/]'));
    if (at >= 0 && at < end) end = at;
  }
  return end;
}

const hexIn = (xml) => {
  const m = xml ? /<a:solidFill>\s*<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/.exec(xml) : null;
  return m ? '#' + m[1].toUpperCase() : null;
};

/** The colour a series was given by hand, or null where it takes the theme's. */
function colourOf(s, p) {
  const head = s.xml.slice(0, ownHead(s.xml, p));
  const own = elementsOf(head, p, 'spPr')[0]?.xml ?? null;
  if (own) {
    const ln = /<a:ln\b[\s\S]*?<\/a:ln>/.exec(own)?.[0] ?? null;
    const fill = hexIn(ln ? own.replace(ln, '') : own);
    if (s.mark !== 'line' && fill) return fill;
    if (s.mark === 'line' && ln && !/<a:noFill\/>/.test(ln) && hexIn(ln)) return hexIn(ln);
    if (fill) return fill;
  }
  const marker = elementsOf(s.xml, p, 'marker')[0]?.xml ?? null;
  return marker ? hexIn(marker) : null;
}

const settledIn = (xml, p) => new RegExp('<' + p + 'dispBlanksAs\\b').test(xml);
const shownIn = (labels, p) => new RegExp('<' + p + 'show(Val|Percent)\\s+val="(1|true)"').test(labels)
  && !new RegExp('<' + p + 'delete\\s+val="(1|true)"').test(labels);

/**
 * What a chart part says about its look: the titles of its two axes (null
 * where an axis has none, and `axes: false` for a chart with no axes),
 * whether its values are written beside its marks, where its legend sits
 * and each series' own colour.
 *
 * `legend` is a place, `none`, or null where the part leaves it open;
 * `dataLabels` is true, false, or null the same way.
 */
export function chartLook(chartXml) {
  const xml = String(chartXml ?? '');
  const p = prefixOf(xml);
  const axes = axesOf(xml, p);
  const titleOf = (ax) => {
    const t = ax ? elementsOf(ax.xml, p, 'title')[0] : null;
    return t ? wordsOf(t.xml) || null : null;
  };
  const settled = settledIn(xml, p);
  const shown = elementsOf(xml, p, 'dLbls').some((l) => shownIn(l.xml, p));
  const legend = elementsOf(xml, p, 'legend')[0] ?? null;
  const pos = legend ? new RegExp('<' + p + 'legendPos\\s+val="(\\w+)"').exec(legend.xml)?.[1] ?? 'r' : null;
  const plot = plotOf(xml, p);
  return {
    axes: Boolean(axes.category || axes.value),
    categoryAxisTitle: titleOf(axes.category),
    valueAxisTitle: titleOf(axes.value),
    dataLabels: shown ? true : settled ? false : null,
    legend: legend ? LEGEND_READ[pos] ?? 'right' : settled ? 'none' : null,
    slices: plot?.mark === 'slices',
    // A pie's colours are its slices', one each: its one series has none of its own.
    seriesColours: plot?.mark === 'slices' ? [] : seriesOf(xml, p).map((s) => colourOf(s, p)),
  };
}

/** The labels written once for a kind of plot: each mark's value, or each slice's share. */
function labelsFor(p, mark) {
  const flag = (name, val) => '<' + p + name + ' val="' + val + '"/>';
  return '<' + p + 'dLbls>' + flag('showLegendKey', 0) + flag('showVal', mark === 'slices' ? 0 : 1) + flag('showCatName', 0) + flag('showSerName', 0)
    + flag('showPercent', mark === 'slices' ? 1 : 0) + flag('showBubbleSize', 0) + '</' + p + 'dLbls>';
}

/** Labels put after the series of every kind of plot the part holds. */
function writeLabels(xml, p) {
  let next = xml;
  const close = '</' + p + 'ser>';
  const plots = Object.entries(PLOTS)
    .flatMap(([name, mark]) => elementsOf(next, p, name).map((el) => ({ ...el, mark })))
    .sort((a, b) => b.start - a.start);
  for (const plot of plots) {
    const at = plot.xml.lastIndexOf(close);
    if (at < 0) continue;
    const where = plot.start + at + close.length;
    next = next.slice(0, where) + labelsFor(p, plot.mark) + next.slice(where);
  }
  return next;
}

/** A legend put where the schema has it, after the plot area; one the part has is moved. */
function writeLegend(xml, p, place) {
  const had = elementsOf(xml, p, 'legend')[0] ?? null;
  const pos = '<' + p + 'legendPos val="' + LEGEND_WRITE[place] + '"/>';
  if (had && !had.xml.endsWith('/>')) {
    let inner = had.xml
      .replace(new RegExp('<' + p + 'legendPos\\b[^>]*/>'), '')
      .replace(new RegExp('<' + p + 'layout>[\\s\\S]*?</' + p + 'layout>'), '')
      .replace(new RegExp('<' + p + 'layout\\s*/>'), '');
    inner = inner.replace(new RegExp('^<' + p + 'legend(\\s[^>]*)?>'), (m) => m + pos);
    return xml.slice(0, had.start) + inner + xml.slice(had.end);
  }
  const fresh = '<' + p + 'legend>' + pos + '<' + p + 'overlay val="0"/></' + p + 'legend>';
  if (had) return xml.slice(0, had.start) + fresh + xml.slice(had.end);
  const plot = '</' + p + 'plotArea>';
  const at = xml.indexOf(plot);
  if (at < 0) throw new Error('this chart could not be read');
  return xml.slice(0, at + plot.length) + fresh + xml.slice(at + plot.length);
}

/**
 * The house rule written down, once: a part that leaves its legend and its
 * labels open is given the ones its chart is drawn with (a legend at the foot
 * from two series or two slices; values on the bars of up to four series
 * side by side; shares on a pie), and marked as saying what it means.
 */
function settle(xml, p) {
  if (settledIn(xml, p)) return xml;
  let next = xml;
  const plot = plotOf(next, p);
  if (plot) {
    const series = elementsOf(plot.xml, p, 'ser');
    const stacked = new RegExp('<' + p + 'grouping\\s+val="(stacked|percentStacked)"').test(plot.xml);
    const bars = plot.name === 'barChart' || plot.name === 'bar3DChart';
    if (!elementsOf(next, p, 'dLbls').length && ((bars && !stacked && series.length <= 4) || plot.mark === 'slices')) {
      next = writeLabels(next, p);
    }
    if (!elementsOf(next, p, 'legend').length) {
      const slices = Number(new RegExp('<' + p + 'ptCount\\s+val="(\\d+)"').exec(series[0]?.xml ?? '')?.[1] ?? 0);
      if (plot.mark === 'slices' ? slices >= 2 : series.length >= 2) next = writeLegend(next, p, 'bottom');
    }
  }
  const mark = '<' + p + 'dispBlanksAs val="gap"/>';
  for (const name of ['plotVisOnly', 'legend']) {
    const found = elementsOf(next, p, name)[0];
    if (found) return next.slice(0, found.end) + mark + next.slice(found.end);
  }
  const close = '</' + p + 'plotArea>';
  const at = next.indexOf(close);
  if (at < 0) throw new Error('this chart could not be read');
  return next.slice(0, at + close.length) + mark + next.slice(at + close.length);
}

/**
 * An axis given a title, or its title taken away with no words. `axis` is
 * `category` or `value`. The title goes after the axis's gridlines, where
 * the schema has it, and reads upward on an axis that stands at the side.
 */
export function setAxisTitle(chartXml, axis, words) {
  const xml = String(chartXml ?? '');
  const p = prefixOf(xml);
  const ax = axesOf(xml, p)[axis === 'category' ? 'category' : 'value'];
  if (!ax) throw new Error('This chart has no axes to name.');
  const text = String(words ?? '').trim();
  let inner = ax.xml.replace(new RegExp('<' + p + 'title>[\\s\\S]*?</' + p + 'title>'), '');
  if (text) {
    const upright = new RegExp('<' + p + 'axPos\\s+val="[lr]"').test(inner);
    const title = '<' + p + 'title><' + p + 'tx><' + p + 'rich><a:bodyPr' + (upright ? ' rot="-5400000" vert="horz"' : '') + '/><a:lstStyle/>'
      + '<a:p><a:r><a:t>' + esc(text) + '</a:t></a:r></a:p></' + p + 'rich></' + p + 'tx><' + p + 'overlay val="0"/></' + p + 'title>';
    let at = -1;
    for (const name of ['minorGridlines', 'majorGridlines', 'axPos', 'delete', 'scaling', 'axId']) {
      const found = elementsOf(inner, p, name);
      if (found.length) { at = found[found.length - 1].end; break; }
    }
    if (at < 0) throw new Error('this chart could not be read');
    inner = inner.slice(0, at) + title + inner.slice(at);
  }
  const next = xml.slice(0, ax.start) + inner + xml.slice(ax.end);
  return text ? withDrawingNs(next) : next;
}

/**
 * Data labels on or off: each mark's value written beside it (each slice's
 * share, on a pie). On, the chart's own labels are written once for every
 * series; off, every label the part holds goes.
 */
export function setDataLabels(chartXml, on) {
  const p = prefixOf(chartXml);
  const xml = settle(String(chartXml), p);
  const bare = xml.replace(new RegExp('<' + p + 'dLbls>[\\s\\S]*?</' + p + 'dLbls>', 'g'), '');
  return on ? writeLabels(bare, p) : bare;
}

/**
 * The legend put at the right, the top, the left or the bottom, or taken
 * away (`none`). A legend the part already has keeps its entries and its
 * look and loses a hand-set layout, as Excel drops one when the legend moves.
 */
export function setLegend(chartXml, place) {
  if (place !== 'none' && !LEGEND_WRITE[place]) throw new Error('A legend goes at the right, the top, the left or the bottom, or is taken away.');
  const p = prefixOf(chartXml);
  const xml = settle(String(chartXml), p);
  if (place !== 'none') return writeLegend(xml, p, place);
  const had = elementsOf(xml, p, 'legend')[0] ?? null;
  return had ? xml.slice(0, had.start) + xml.slice(had.end) : xml;
}

/**
 * One series given a colour of its own (`#RRGGBB`), or given back to the
 * theme with null. A column, a bar or an area is filled with it; a line is
 * drawn in it, its markers with it; a scatter of markers alone keeps its
 * line off and colours the markers. A pie's slices are single points, each
 * with a colour of its own, and are not coloured as a series.
 */
export function setSeriesColour(chartXml, index, colour) {
  const xml = String(chartXml ?? '');
  const p = prefixOf(xml);
  const s = seriesOf(xml, p)[Number(index)];
  if (!s) throw new Error('This chart has no series ' + (Number(index) + 1) + '.');
  if (s.mark === 'slices') throw new Error('A pie has one series, and each slice a colour of its own.');
  const hex = colour === null || colour === undefined || colour === '' ? null : /^#?([0-9A-Fa-f]{6})$/.exec(String(colour))?.[1]?.toUpperCase();
  if (hex === undefined) throw new Error('A colour is written like #1F6FB2.');
  const fill = hex ? '<a:solidFill><a:srgbClr val="' + hex + '"/></a:solidFill>' : '';

  let ser = s.xml;
  const headEnd = ownHead(ser, p);
  const had = elementsOf(ser.slice(0, headEnd), p, 'spPr')[0] ?? null;
  const hadLine = had ? /<a:ln\b[^>]*>[\s\S]*?<\/a:ln>|<a:ln\b[^>]*\/>/.exec(had.xml)?.[0] ?? null : null;
  const lineOff = Boolean(hadLine && /<a:noFill\/>/.test(hadLine));

  // The series' own shape.
  let own = '';
  if (s.mark === 'line') {
    if (lineOff) own = had.xml;
    else if (hex) {
      const width = hadLine ? /\bw="(\d+)"/.exec(hadLine)?.[1] ?? '28575' : '28575';
      const dash = hadLine ? /<a:prstDash\b[^>]*\/>/.exec(hadLine)?.[0] ?? '' : '';
      own = '<' + p + 'spPr><a:ln w="' + width + '" cap="rnd">' + fill + dash + '<a:round/></a:ln></' + p + 'spPr>';
    }
  } else if (hex) own = '<' + p + 'spPr>' + fill + '</' + p + 'spPr>';
  if (had) ser = ser.slice(0, had.start) + own + ser.slice(had.end);
  else if (own) {
    // After the series' name, else after its order: where the schema has it.
    let at = -1;
    for (const name of ['tx', 'order', 'idx']) {
      const found = elementsOf(ser.slice(0, headEnd), p, name);
      if (found.length) { at = found[found.length - 1].end; break; }
    }
    if (at < 0) throw new Error('this chart could not be read');
    ser = ser.slice(0, at) + own + ser.slice(at);
  }

  // A line's markers wear the line's colour; a marker switched off stays off.
  if (s.mark === 'line') {
    const marker = elementsOf(ser, p, 'marker')[0] ?? null;
    const none = Boolean(marker && new RegExp('<' + p + 'symbol\\s+val="none"').test(marker.xml));
    const look = hex ? '<' + p + 'spPr>' + fill + '<a:ln w="9525">' + fill + '</a:ln></' + p + 'spPr>' : '';
    const open = '<' + p + 'marker>';
    const shut = '</' + p + 'marker>';
    if (marker && !none && !marker.xml.endsWith('/>')) {
      // Its symbol and its size stay; its shape goes last, where the schema has it.
      let made = marker.xml.replace(new RegExp('<' + p + 'spPr>[\\s\\S]*?</' + p + 'spPr>'), '');
      made = made.slice(0, made.length - shut.length) + look + shut;
      if (made === open + shut) made = '';
      ser = ser.slice(0, marker.start) + made + ser.slice(marker.end);
    } else if (!marker && hex) {
      // No symbol or size of its own: whether markers show stays the chart's choice.
      const at = ownHead(ser, p);
      ser = ser.slice(0, at) + open + look + shut + ser.slice(at);
    }
  }
  const next = xml.slice(0, s.start) + ser + xml.slice(s.end);
  return hex ? withDrawingNs(next) : next;
}
