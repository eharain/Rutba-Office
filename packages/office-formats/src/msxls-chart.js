// Excel's charts in a binary workbook (Excel 5.0/95 and 97-2003): a chart's
// records — its own substream, after the object that places it on a sheet —
// read into a plain model, and that model written as the DrawingML chart
// part an .xlsx keeps.
//
// A chart's records nest between BEGIN and END. Inside the chart come its
// series, each its name, values and categories as formulas, with a format
// for the series and for any point of it; then each axes set, with its axes
// and its chart groups (the kind of chart each series is drawn as: bars,
// lines, a pie...); and the texts — the title, the axes' titles, the data
// labels — each tied to what it labels by an object link. After the chart
// Excel keeps the values the series last had, cell by cell.
//
// Excel 2007 and later also keep each part's DrawingML shape and text
// properties as XML, in records of their own; where they are there they are
// written as they were, so a chart keeps its theme colours, line widths and
// fonts. Before that a part's colour is its own, or — automatic — the one
// its place gives it in the workbook's palette.
//
// The layouts follow LibreOffice's chart import (sc/source/filter/excel/xichart.cxx).

const u8 = (b, at) => b[at] ?? 0;
const u16 = (b, at) => (b[at] ?? 0) | ((b[at + 1] ?? 0) << 8);
const i16 = (b, at) => (u16(b, at) << 16) >> 16;
const u32 = (b, at) => (u16(b, at) | (u16(b, at + 2) << 16)) >>> 0;
const i32 = (b, at) => u32(b, at) | 0;
const f64 = (b, at) => (at + 8 <= b.length ? new DataView(b.buffer, b.byteOffset + at, 8).getFloat64(0, true) : 0);
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const C = {
  BEGIN: 0x1033, END: 0x1034, CHART: 0x1002, SERIES: 0x1003, DATAFORMAT: 0x1006, LINEFORMAT: 0x1007, MARKERFORMAT: 0x1009,
  AREAFORMAT: 0x100a, PIEFORMAT: 0x100b, ATTACHEDLABEL: 0x100c, SERIESTEXT: 0x100d, CHARTFORMAT: 0x1014, LEGEND: 0x1015,
  BAR: 0x1017, LINE: 0x1018, PIE: 0x1019, AREA: 0x101a, SCATTER: 0x101b, AXIS: 0x101d, TICK: 0x101e, VALUERANGE: 0x101f,
  CATSERRANGE: 0x1020, AXISLINEFORMAT: 0x1021, TEXT: 0x1025, FONTX: 0x1026, OBJECTLINK: 0x1027, FRAME: 0x1032, PLOTAREA: 0x1035,
  CHART3D: 0x103a, RADAR: 0x103e, SURF: 0x103f, RADARAREA: 0x1040, AXISPARENT: 0x1041, SHTPROPS: 0x1044, SERTOCRT: 0x1045,
  SERPARENT: 0x104a, AI: 0x1051, SERFMT: 0x105d, BOPPOP: 0x1061, SIINDEX: 0x1065, GELFRAME: 0x1066,
  LABELPROPS: 0x086b, SHAPEPROPS: 0x08a4, TEXTPROPS: 0x08a5, CONTINUEFRT12: 0x0875,
  NUMBER: 0x0203, LABEL: 0x0204, BOOLERR: 0x0205,
};

// The automatic colours, by a series' (or a pie slice's) format index: places in the palette.
const range = (a, b) => Array.from({ length: b - a }, (_, i) => a + i);
const FILL_AUTO = [...range(24, 64), ...range(8, 24)];
const LINE_AUTO = [...range(32, 63), ...range(8, 32), 63];

const DASHES = ['solid', 'dash', 'sysDot', 'dashDot', 'sysDashDotDot'];
const WEIGHTS = { [-1]: 3175, 0: 12700, 1: 25400, 2: 38100 };
const SYMBOLS = ['none', 'square', 'diamond', 'triangle', 'x', 'star', 'dash', 'dash', 'circle', 'plus'];
const TICKS = ['none', 'in', 'out', 'cross'];
const LABEL_SPOTS = ['', 'none', 'low', 'high', 'nextTo'];
const LABEL_PLACES = [null, 'outEnd', 'inEnd', 'ctr', 'inBase', 't', 'b', 'l', 'r', 'bestFit'];
// Where each kind of chart lets a data label sit; anywhere else Excel refuses the file.
const PLACES_FOR = {
  bar: ['outEnd', 'inEnd', 'ctr', 'inBase'], pie: ['outEnd', 'inEnd', 'ctr', 'bestFit'],
  line: ['t', 'b', 'l', 'r', 'ctr'], scatter: ['t', 'b', 'l', 'r', 'ctr'],
};

/** The records as a tree: each record that a BEGIN follows holds what is between it and its END. */
function tree(records) {
  const root = { kids: [] };
  const stack = [root];
  let last = null;
  for (const r of records) {
    if (r.id === C.BEGIN) { if (last) { last.kids = []; stack.push(last); } last = null; continue; }
    if (r.id === C.END) { if (stack.length > 1) stack.pop(); last = null; continue; }
    const node = { id: r.id, data: r.data, kids: null };
    stack[stack.length - 1].kids.push(node);
    last = node;
  }
  return root.kids;
}
const kid = (node, id) => node?.kids?.find((k) => k.id === id) ?? null;

/**
 * Excel 2007's DrawingML for a part: its shape properties (spPr) or text
 * properties (txPr), the root's insides. Empty shape properties are a part
 * left automatic — its colours the theme's — which Excel marks so beside
 * records holding only the palette's nearest colours: '' then, and null
 * where there are none.
 */
function frtXml(list, id) {
  for (let i = 0; i < (list?.length ?? 0); i++) {
    if (list[i].id !== id) continue;
    const d = list[i].data;
    const at = id === C.SHAPEPROPS ? 24 : 20;
    const cb = u32(d, at - 4);
    if (!cb) { if (id === C.SHAPEPROPS) return ''; continue; }
    const parts = [d.subarray(at)];
    for (let j = i + 1; j < list.length && list[j].id === C.CONTINUEFRT12; j++) parts.push(list[j].data.subarray(12));
    const bytes = Buffer.concat(parts.map((p) => Buffer.from(p))).subarray(0, cb);
    const xml = bytes.toString('utf8').replace(/^<\?xml[^>]*\?>\s*/, '');
    const m = /^<a:(spPr|txPr)\b[^>]*?(?:\/>|>([\s\S]*)<\/a:\1>)\s*$/.exec(xml);
    if (m) return m[2] ?? '';
  }
  return null;
}

/** A colour Office Art keeps in a GELFRAME: its fill, when it says there is one. */
function gelFill(node) {
  if (!node) return undefined;
  const d = node.data;
  const props = new Map();
  if (u16(d, 2) !== 0xf00b) return undefined;
  const n = u16(d, 0) >> 4;
  for (let k = 0; k < n; k++) {
    const id = u16(d, 8 + k * 6);
    props.set(id & 0x3fff, u32(d, 8 + k * 6 + 2));
  }
  const bools = props.get(0x01bf);
  if (bools != null && (bools & 0x100000) && !(bools & 0x10)) return 'none';
  const c = props.get(0x0181);
  if (c == null || (c >>> 24) & 0x18) return undefined;
  return [0, 8, 16].map((s) => ((c >>> s) & 0xff).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/**
 * Read one chart's records into its model.
 *
 * @param {Array<{ id: number, data: Uint8Array }>} records the chart's substream, BOF to EOF
 * @param {{ biff: number, text: Function, formula: Function, palette: string[], font: Function }} ctx
 *   text(bytes, at, countBytes) reads a string as the version writes one; formula(rgce) reads one into text
 */
export function readChart(records, ctx) {
  const top = tree(records);
  const root = top.find((n) => n.id === C.CHART);
  if (!root?.kids) return null;
  const b8 = ctx.biff === 8;
  const colour = (rgbAt, icvAt, d) => (b8 && d.length > icvAt + 1 ? ctx.palette[u16(d, icvAt)] ?? null : rgb(d, rgbAt));
  const rgb = (d, at) => [0, 1, 2].map((c) => u8(d, at + c).toString(16).padStart(2, '0')).join('').toUpperCase();

  /**
   * A part's look: its kept DrawingML, or its fill and line as the records
   * give them (null: automatic). A series or a point Excel 2007 left
   * automatic takes the theme's colour for its place; any other part left so
   * keeps the colours recorded beside it, since automatic grid lines and
   * frames without a chart style are not what Excel drew.
   */
  const look = (node, auto = {}, data = false) => {
    const kids = node?.kids ?? [];
    let xml = frtXml(kids, C.SHAPEPROPS);
    if (xml === '' && !data) xml = null;
    const out = { xml, fill: null, line: null };
    const area = kid(node, C.AREAFORMAT);
    if (area) {
      const d = area.data;
      const pattern = u16(d, 8);
      const isAuto = u16(d, 10) & 1;
      if (!pattern) out.fill = 'none';
      else if (!isAuto) out.fill = colour(0, 12, d);
      else if (auto.fill != null) out.fill = ctx.palette[auto.fill] ?? null;
    }
    const gel = gelFill(kid(node, C.GELFRAME));
    if (gel !== undefined && out.fill !== 'none') out.fill = gel;
    const lineRec = kid(node, C.LINEFORMAT);
    if (lineRec) out.line = lineOf(lineRec.data, auto.line);
    return out;
  };
  /** A LINEFORMAT's line: none, its own colour, width and dashes, or (automatic) the palette's or null. */
  const lineOf = (d, autoIndex = null) => {
    const pattern = u16(d, 4);
    const isAuto = u16(d, 8) & 1;
    if (pattern === 5) return 'none';
    if (isAuto && autoIndex == null) return null;
    return {
      colour: isAuto ? ctx.palette[autoIndex] ?? null : colour(0, 10, d),
      width: WEIGHTS[i16(d, 6)] ?? 12700,
      dash: DASHES[pattern] ?? 'solid',
    };
  };

  /** A text's words, its font and its look: a title, an axis title, the legend's or a label's. */
  const textOf = (node) => {
    const d = node.data;
    const kids = node.kids ?? [];
    const link = kid(node, C.OBJECTLINK);
    const out = {
      target: link ? u16(link.data, 0) : 0,
      series: link ? u16(link.data, 2) : 0,
      point: link ? u16(link.data, 4) : 0,
      flags: u16(d, 24),
      place: b8 ? u16(d, 28) & 0x0f : 0,
      rotation: b8 ? u16(d, 30) : 0,
      text: null, ref: null,
      font: null, txPr: frtXml(kids, C.TEXTPROPS),
      look: look(kid(node, C.FRAME)),
      shows: null,
    };
    const fx = kid(node, C.FONTX);
    if (fx) out.font = ctx.font(u16(fx.data, 0));
    for (let i = 0; i < kids.length; i++) {
      const k = kids[i];
      if (k.id === C.SERIESTEXT) out.text = ctx.text(k.data, 2, 1);
      if (k.id === C.AI && u8(k.data, 1) === 2) out.ref = ctx.formula(k.data.subarray(8, 8 + u16(k.data, 6))) || null;
      if (k.id === C.LABELPROPS) out.shows = u16(k.data, 12);
    }
    return out;
  };

  const chart = {
    series: [], groups: [], axes: [], texts: [],
    title: null, legend: null, look: look(kid(root, C.FRAME)),
    plot: null, visibleOnly: true, blanks: 'gap', threeD: null,
  };

  // The series, each its source links and formats.
  for (const node of root.kids.filter((k) => k.id === C.SERIES)) {
    const index = chart.series.length;
    const s = {
      index, group: 0, child: false, name: null, nameRef: null,
      values: null, categories: null, bubbles: null,
      categoriesText: u16(node.data, 0) === 3,
      count: u16(node.data, 6),
      look: null, formatIndex: index, points: [], marker: null, smooth: false, explosion: 0, labels: null,
    };
    const kids = node.kids ?? [];
    for (let i = 0; i < kids.length; i++) {
      const k = kids[i];
      if (k.id === C.AI) {
        const id = u8(k.data, 0);
        const ref = u8(k.data, 1) === 2 ? ctx.formula(k.data.subarray(8, 8 + u16(k.data, 6))) || null : null;
        if (id === 0) {
          s.nameRef = ref;
          if (kids[i + 1]?.id === C.SERIESTEXT) s.name = ctx.text(kids[i + 1].data, 2, 1);
        } else if (id === 1) s.values = ref;
        else if (id === 2) s.categories = ref;
        else if (id === 3) s.bubbles = ref;
      } else if (k.id === C.DATAFORMAT) {
        const point = u16(k.data, 0);
        const formatIndex = u16(k.data, 4);
        if (point === 0xffff) {
          s.formatIndex = formatIndex;
          s.format = k;
        } else s.points.push({ index: point, format: k, formatIndex });
      } else if (k.id === C.SERTOCRT) s.group = u16(k.data, 0);
      else if (k.id === C.SERPARENT) s.child = true;
    }
    chart.series.push(s);
  }

  // The axes sets: their axes, the chart groups drawn on them, the axes' titles.
  for (const set of root.kids.filter((k) => k.id === C.AXISPARENT)) {
    const axesSet = u16(set.data, 0);
    for (const k of set.kids ?? []) {
      if (k.id === C.AXIS) {
        const type = u16(k.data, 0);
        const ax = { set: axesSet, type, title: null, major: false, minor: false, deleted: false, reverse: false, log: false, min: null, max: null, majorUnit: null, minorUnit: null, crossesMax: false, crossesAt: null, between: true, tickMajor: 'out', tickMinor: 'none', labels: 'nextTo', lineXml: null, gridXml: null, minorXml: null, lineNone: false, txPr: null, font: null };
        const kids = k.kids ?? [];
        const lineKinds = [];
        for (let i = 0; i < kids.length; i++) {
          const a = kids[i];
          const d = a.data;
          if (a.id === C.VALUERANGE) {
            const flags = u16(d, 40);
            if (!(flags & 1)) ax.min = f64(d, 0);
            if (!(flags & 2)) ax.max = f64(d, 8);
            if (!(flags & 4)) ax.majorUnit = f64(d, 16);
            if (!(flags & 8)) ax.minorUnit = f64(d, 24);
            if (!(flags & 0x10)) ax.crossesAt = f64(d, 32);
            ax.log = Boolean(flags & 0x20);
            ax.reverse = Boolean(flags & 0x40);
            ax.crossesMax = Boolean(flags & 0x80);
          } else if (a.id === C.CATSERRANGE) {
            const flags = u16(d, 6);
            ax.between = Boolean(flags & 1);
            ax.crossesMax = Boolean(flags & 2);
            ax.reverse = Boolean(flags & 4);
          } else if (a.id === C.TICK) {
            ax.tickMajor = TICKS[u8(d, 0)] ?? 'out';
            ax.tickMinor = TICKS[u8(d, 1)] ?? 'none';
            ax.labels = LABEL_SPOTS[u8(d, 2) + 1] ?? 'nextTo';
          } else if (a.id === C.AXISLINEFORMAT) {
            const which = u16(d, 0);
            lineKinds.push(which);
            const line = kids[i + 1]?.id === C.LINEFORMAT ? kids[i + 1].data : null;
            if (which === 1) ax.major = true;
            if (which === 2) ax.minor = true;
            if (which === 0 && line && u16(line, 4) === 5) ax.lineNone = true;
            // An axis whose line lacks "show the axis" is not drawn at all (a second axes set's categories, often).
            if (which === 0 && line && !(u16(line, 8) & 4)) ax.deleted = true;
            if (line && which <= 2) ax[['line', 'grid', 'minor'][which] + 'Line'] = lineOf(line);
          } else if (a.id === C.FONTX) ax.font = ctx.font(u16(d, 0));
        }
        // Excel 2007's shape properties, one for each axis line or grid line in turn.
        let n = 0;
        for (const a of kids) {
          if (a.id !== C.SHAPEPROPS) continue;
          const which = lineKinds[n++];
          const xml = frtXml([a], C.SHAPEPROPS) || null;
          if (which === 0) ax.lineXml = xml;
          else if (which === 1) ax.gridXml = xml;
          else if (which === 2) ax.minorXml = xml;
        }
        ax.txPr = frtXml(kids, C.TEXTPROPS);
        // An axis whose labels and line are both gone is a deleted one.
        if (ax.labels === 'none' && ax.lineNone && ax.tickMajor === 'none') ax.deleted = true;
        chart.axes.push(ax);
      } else if (k.id === C.TEXT) {
        chart.texts.push({ ...textOf(k), set: axesSet });
      } else if (k.id === C.PLOTAREA) {
        chart.plotArea = true;
      } else if (k.id === C.FRAME && chart.plotArea && !chart.plot) {
        chart.plot = look(k);
      } else if (k.id === C.CHARTFORMAT) {
        const g = { index: u16(k.data, 18), set: axesSet, varied: Boolean(u16(k.data, 16) & 1), kind: 'bar', horizontal: false, grouping: 'clustered', gap: 150, overlap: 0, hole: 0, firstSlice: 0, bubbles: false, filled: false, threeD: null };
        for (const t of k.kids ?? []) {
          const d = t.data;
          if (t.id === C.BAR) {
            g.kind = 'bar';
            g.overlap = -i16(d, 0);
            g.gap = i16(d, 2);
            const flags = u16(d, 4);
            g.horizontal = Boolean(flags & 1);
            g.grouping = flags & 4 ? 'percentStacked' : flags & 2 ? 'stacked' : 'clustered';
          } else if (t.id === C.LINE || t.id === C.AREA) {
            g.kind = t.id === C.LINE ? 'line' : 'area';
            const flags = u16(d, 0);
            g.grouping = flags & 2 ? 'percentStacked' : flags & 1 ? 'stacked' : 'standard';
          } else if (t.id === C.PIE || t.id === C.BOPPOP) {
            g.kind = 'pie';
            g.firstSlice = t.id === C.PIE ? u16(d, 0) % 360 : 0;
            g.hole = t.id === C.PIE ? u16(d, 2) : 0;
          } else if (t.id === C.SCATTER) {
            g.kind = 'scatter';
            g.bubbles = b8 && Boolean(u16(d, 4) & 1);
          } else if (t.id === C.RADAR || t.id === C.RADARAREA) {
            g.kind = 'radar';
            g.filled = t.id === C.RADARAREA;
          } else if (t.id === C.SURF) {
            g.kind = 'surface';
            g.filled = Boolean(u16(d, 0) & 1);
          } else if (t.id === C.CHART3D) {
            const flags = u16(d, 12);
            g.threeD = { rotation: u16(d, 0), elevation: i16(d, 2), distance: u16(d, 4), height: u16(d, 6), depth: u16(d, 8), gap: u16(d, 10), perspective: Boolean(flags & 1), cluster: Boolean(flags & 2) };
          } else if (t.id === C.LEGEND) {
            const dock = u8(d, 16);
            let pos = ({ 0: 'b', 1: 'tr', 2: 't', 3: 'r', 4: 'l' })[dock];
            if (!pos) {
              // Not docked: the side its box sits nearest.
              const x = i32(d, 0) / 4000; const y = i32(d, 4) / 4000; const w = i32(d, 8) / 4000; const h = i32(d, 12) / 4000;
              const cx = x + w / 2; const cy = y + h / 2;
              pos = !w && !h ? 'r' : [['t', cy], ['b', 1 - cy], ['l', cx], ['r', 1 - cx]].sort((a, b2) => a[1] - b2[1])[0][0];
            }
            const text = kid(t, C.TEXT);
            chart.legend = { pos, look: look(kid(t, C.FRAME)), txPr: frtXml(t.kids, C.TEXTPROPS) ?? (text ? frtXml(text.kids, C.TEXTPROPS) : null), font: text && kid(text, C.FONTX) ? ctx.font(u16(kid(text, C.FONTX).data, 0)) : null };
          }
        }
        if (g.threeD) chart.threeD = g.threeD;
        chart.groups.push(g);
      }
    }
  }
  delete chart.plotArea;

  // The chart's own texts — its title, its data labels — and how it plots.
  for (const k of root.kids) {
    if (k.id === C.TEXT && kid(k, C.OBJECTLINK)) chart.texts.push({ ...textOf(k), set: 0 });
    if (k.id === C.SHTPROPS) {
      chart.visibleOnly = Boolean(u16(k.data, 0) & 2);
      chart.blanks = ['gap', 'zero', 'span'][u8(k.data, 2)] ?? 'gap';
    }
  }
  if (!chart.groups.length) chart.groups.push({ index: 0, set: 0, varied: false, kind: 'bar', horizontal: false, grouping: 'clustered', gap: 150, overlap: 0, hole: 0, firstSlice: 0, threeD: null });

  // Each series' look now that its group's kind is known: lines for a line, a scatter or a radar; fills otherwise.
  const groupOf = (s) => chart.groups.find((g) => g.index === s.group) ?? chart.groups[0];
  for (const s of chart.series) {
    const g = groupOf(s);
    const lined = g.kind === 'line' || (g.kind === 'scatter' && !g.bubbles) || (g.kind === 'radar' && !g.filled);
    const auto = (fi) => (lined ? { line: LINE_AUTO[fi % LINE_AUTO.length] } : { fill: FILL_AUTO[fi % FILL_AUTO.length] });
    s.look = s.format ? look(s.format, g.varied ? {} : auto(s.formatIndex), true) : { xml: null, fill: null, line: null };
    if (s.format) {
      const m = kid(s.format, C.MARKERFORMAT);
      if (m && !(u16(m.data, 10) & 1)) {
        const d = m.data;
        const flags = u16(d, 10);
        s.marker = {
          symbol: SYMBOLS[u16(d, 8)] ?? 'none',
          size: b8 && d.length >= 20 ? Math.max(2, Math.min(72, Math.round(u32(d, 16) / 20))) : 5,
          fill: flags & 0x10 ? 'none' : colour(4, 14, d),
          line: flags & 0x20 ? 'none' : colour(0, 12, d),
        };
      }
      // A series Excel 2007 left automatic has automatic markers too: their shape kept, not the palette's colours.
      if (s.marker && s.look.xml === '') { s.marker.fill = null; s.marker.line = null; }
      const pie = kid(s.format, C.PIEFORMAT);
      if (pie) s.explosion = u16(pie.data, 0);
      const fmt = kid(s.format, C.SERFMT);
      if (fmt) s.smooth = Boolean(u16(fmt.data, 0) & 1);
      const label = kid(s.format, C.ATTACHEDLABEL);
      if (label) {
        const f = u16(label.data, 0);
        if (f) s.labels = { value: Boolean(f & 1), percent: Boolean(f & 6), category: Boolean(f & 0x14), name: false, bubble: Boolean(f & 0x20), place: null, txPr: null, look: null };
      }
    }
    s.points = s.points.map((p) => ({ index: p.index, look: look(p.format, g.varied ? { fill: FILL_AUTO[p.formatIndex % FILL_AUTO.length] } : auto(p.formatIndex), true), explosion: kid(p.format, C.PIEFORMAT) ? u16(kid(p.format, C.PIEFORMAT).data, 0) : 0 }))
      .filter((p) => p.look.xml != null || p.look.fill || p.look.line || p.explosion);
  }

  // The texts by what they label.
  for (const t of chart.texts) {
    if (t.target === 1) chart.title = t;
    else if (t.target === 4 && t.point === 0xffff && chart.series[t.series]) {
      const f = t.flags;
      const shows = t.shows;
      const s = chart.series[t.series];
      s.labels = (f & 0x40) ? null : {
        value: shows != null ? Boolean(shows & 4) : Boolean(f & 4),
        percent: shows != null ? Boolean(shows & 8) : Boolean(f & 0x1800),
        category: shows != null ? Boolean(shows & 2) : Boolean(f & 0x4800),
        name: shows != null ? Boolean(shows & 1) : false,
        bubble: shows != null ? Boolean(shows & 0x10) : Boolean(f & 0x2000),
        place: LABEL_PLACES[t.place] ?? null,
        txPr: t.txPr, look: t.look,
      };
    } else if (t.target === 2 || t.target === 3 || t.target === 7) {
      const type = { 2: 1, 3: 0, 7: 2 }[t.target];
      const ax = chart.axes.find((a) => a.set === t.set && a.type === type);
      if (ax) ax.title = t;
    }
  }

  // The values Excel last plotted, kept after the chart: categories, values and bubble sizes, series by series.
  const cache = { 1: [], 2: [], 3: [] };
  let into = 0;
  for (const k of top) {
    if (k.id === C.SIINDEX) into = u16(k.data, 0);
    else if (cache[into] && (k.id === C.NUMBER || k.id === C.LABEL || k.id === C.BOOLERR)) {
      const point = u16(k.data, 0);
      const series = u16(k.data, 2);
      const v = k.id === C.NUMBER ? f64(k.data, 6) : k.id === C.LABEL ? ctx.text(k.data, 6, 2) : u8(k.data, 6);
      (cache[into][series] ??= [])[point] = v;
    }
  }
  for (const s of chart.series) {
    s.cachedValues = cache[1][s.index] ?? null;
    s.cachedCategories = cache[2][s.index] ?? null;
    s.cachedBubbles = cache[3][s.index] ?? null;
    delete s.format;
  }
  chart.series = chart.series.filter((s) => !s.child);
  return chart;
}

/* ── writing ───────────────────────────────────────────────────────────── */

const fillXml = (fill) => (fill === 'none' ? '<a:noFill/>' : fill ? '<a:solidFill><a:srgbClr val="' + fill + '"/></a:solidFill>' : '');
const lineXml = (line) => (line === 'none' ? '<a:ln><a:noFill/></a:ln>'
  : line ? '<a:ln w="' + line.width + '">' + (line.colour ? fillXml(line.colour) : '') + (line.dash && line.dash !== 'solid' ? '<a:prstDash val="' + line.dash + '"/>' : '') + '</a:ln>' : '');
/** A part's c:spPr: its kept DrawingML, else its fill and line. */
function spPr(lk, lined = false) {
  if (!lk) return '';
  if (lk.xml != null) return '<c:spPr>' + lk.xml + '</c:spPr>';
  const inner = (lined ? '' : fillXml(lk.fill)) + lineXml(lk.line);
  return inner ? '<c:spPr>' + inner + '</c:spPr>' : '';
}
/** A line's c:spPr: its kept DrawingML, else the line the record gives. */
const lineSpPr = (xml, line) => (xml != null ? '<c:spPr>' + xml + '</c:spPr>' : line ? '<c:spPr>' + lineXml(line) + '</c:spPr>' : '');
/** Text properties: kept DrawingML, else the font the record names. */
function txPr(kept, font, rotation = null) {
  if (kept != null) return '<c:txPr>' + kept + '</c:txPr>';
  if (!font && rotation == null) return '';
  return '<c:txPr><a:bodyPr' + (rotation != null ? ' rot="' + rotation + '" vert="horz"' : '') + '/><a:lstStyle/><a:p><a:pPr>' + defRPr(font) + '</a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>';
}
function defRPr(font) {
  if (!font) return '<a:defRPr/>';
  return '<a:defRPr sz="' + Math.round((font.height || 200) / 20 * 100) + '" b="' + (font.bold ? 1 : 0) + '" i="' + (font.italic ? 1 : 0) + '">'
    + (font.colour ? fillXml(font.colour) : '') + (font.name ? '<a:latin typeface="' + esc(font.name) + '"/>' : '') + '</a:defRPr>';
}
/** A title's words as rich text, in its kept text properties or its font. */
function richXml(text, kept, font, rotation = null) {
  const run = '<a:r><a:rPr lang="en-US"/><a:t>' + esc(text) + '</a:t></a:r>';
  if (kept != null && /<a:p>/.test(kept)) {
    const inner = /<a:endParaRPr\b/.test(kept) ? kept.replace(/<a:endParaRPr\b/, () => run + '<a:endParaRPr') : kept.replace(/<\/a:p>/, () => run + '</a:p>');
    return '<c:rich>' + inner + '</c:rich>';
  }
  return '<c:rich><a:bodyPr' + (rotation != null ? ' rot="' + rotation + '" vert="horz"' : '') + '/><a:lstStyle/><a:p><a:pPr>' + defRPr(font) + '</a:pPr>' + run + '</a:p></c:rich>';
}
function titleXml(t, text, rotation = null) {
  if (text == null || text === '') return '';
  const tx = t.ref
    ? '<c:tx><c:strRef><c:f>' + esc(t.ref) + '</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + esc(text) + '</c:v></c:pt></c:strCache></c:strRef></c:tx>'
    : '<c:tx>' + richXml(text, t.txPr, t.font, rotation) + '</c:tx>';
  return '<c:title>' + tx + '<c:overlay val="0"/>' + spPr(t.look) + (t.ref ? txPr(t.txPr, t.font, rotation) : '') + '</c:title>';
}

const pts = (list) => list.map((v, i) => (v == null || v === '' ? '' : '<c:pt idx="' + i + '"><c:v>' + esc(typeof v === 'number' ? String(v) : v) + '</c:v></c:pt>')).join('');
function numData(ref, list) {
  const nums = (list ?? []).map((v) => (typeof v === 'number' && Number.isFinite(v) ? v : typeof v === 'boolean' ? Number(v) : null));
  const cacheXml = '<c:formatCode>General</c:formatCode><c:ptCount val="' + nums.length + '"/>' + pts(nums);
  return ref ? '<c:numRef><c:f>' + esc(ref) + '</c:f><c:numCache>' + cacheXml + '</c:numCache></c:numRef>' : '<c:numLit>' + cacheXml + '</c:numLit>';
}
function strData(ref, list) {
  const words = (list ?? []).map((v) => (v == null ? null : String(v)));
  const cacheXml = '<c:ptCount val="' + words.length + '"/>' + pts(words);
  return ref ? '<c:strRef><c:f>' + esc(ref) + '</c:f><c:strCache>' + cacheXml + '</c:strCache></c:strRef>' : '<c:strLit>' + cacheXml + '</c:strLit>';
}
/** Categories as words when any is, as numbers when all are. */
function catData(ref, list) {
  return (list ?? []).some((v) => typeof v === 'string') || !(list ?? []).length ? strData(ref, list) : numData(ref, list);
}

/**
 * The DrawingML chart part for a chart read by readChart.
 *
 * @param {object} chart readChart's model
 * @param {(ref: string) => Array<number|string|boolean|null>|null} resolve a range's values from the workbook
 * @returns {string|null} the part, or null for a chart with no series to draw
 */
export function chartXml(chart, resolve = () => null) {
  const valuesOf = (ref, cached) => (ref ? resolve(ref) : null) ?? cached ?? [];
  const axisIds = (set) => ({ cat: 50010 + set * 10, val: 50011 + set * 10, ser: 50012 + set * 10 });
  const sets = new Map();
  const groupsXml = [];
  const threeD = chart.threeD;
  const first = chart.series[0];

  for (const g of chart.groups) {
    const list = chart.series.filter((s) => (chart.groups.some((x) => x.index === s.group) ? s.group : chart.groups[0].index) === g.index);
    if (!list.length) continue;
    const kind = g.kind === 'pie' && g.hole ? 'doughnut' : g.kind;
    const deep = Boolean(g.threeD) && ['bar', 'line', 'area', 'pie', 'surface'].includes(kind);
    const lined = kind === 'line' || (kind === 'scatter' && !g.bubbles) || (kind === 'radar' && !g.filled);
    const ids = axisIds(g.set);
    const needsSeriesAxis = kind === 'surface' || (deep && (kind === 'line' || (kind === 'bar' && !g.threeD.cluster) || (kind === 'area' && g.grouping === 'standard')));
    if (kind !== 'pie' && kind !== 'doughnut') {
      const s = sets.get(g.set) ?? { scatter: false, horizontal: false, series: false };
      s.scatter ||= kind === 'scatter';
      s.horizontal ||= kind === 'bar' && g.horizontal;
      s.series ||= needsSeriesAxis;
      sets.set(g.set, s);
    }
    const sers = list.map((s) => {
      const name = s.name ?? (s.nameRef ? (resolve(s.nameRef) ?? [])[0] : null);
      const tx = s.nameRef
        ? '<c:tx>' + strData(s.nameRef, [name == null ? '' : String(name)]) + '</c:tx>'
        : name != null ? '<c:tx><c:v>' + esc(name) + '</c:v></c:tx>' : '';
      const values = valuesOf(s.values, s.cachedValues);
      const categories = s.categories ? valuesOf(s.categories, s.cachedCategories) : s.cachedCategories;
      const isBar = kind === 'bar';
      const dPts = s.points.map((p) => '<c:dPt><c:idx val="' + p.index + '"/>' + (isBar ? '<c:invertIfNegative val="0"/>' : '')
        + (kind === 'pie' || kind === 'doughnut' ? '<c:bubble3D val="0"/>' + (p.explosion ? '<c:explosion val="' + p.explosion + '"/>' : '') : '')
        + spPr(p.look, lined) + '</c:dPt>').join('');
      const lb = s.labels;
      const allowed = PLACES_FOR[kind === 'doughnut' ? 'none' : kind] ?? [];
      const dLbls = lb ? '<c:dLbls>' + spPr(lb.look) + txPr(lb.txPr, null)
        + (lb.place && allowed.includes(lb.place) ? '<c:dLblPos val="' + lb.place + '"/>' : '')
        + '<c:showLegendKey val="0"/><c:showVal val="' + (lb.value ? 1 : 0) + '"/><c:showCatName val="' + (lb.category ? 1 : 0) + '"/>'
        + '<c:showSerName val="' + (lb.name ? 1 : 0) + '"/><c:showPercent val="' + (lb.percent && (kind === 'pie' || kind === 'doughnut') ? 1 : 0) + '"/>'
        + '<c:showBubbleSize val="' + (lb.bubble ? 1 : 0) + '"/>' + (kind === 'pie' ? '<c:showLeaderLines val="1"/>' : '') + '</c:dLbls>' : '';
      const marker = (lined || kind === 'radar') && s.marker
        ? '<c:marker><c:symbol val="' + s.marker.symbol + '"/>' + (s.marker.symbol !== 'none' ? '<c:size val="' + s.marker.size + '"/>'
          + '<c:spPr>' + fillXml(s.marker.fill) + (s.marker.line === 'none' ? '<a:ln><a:noFill/></a:ln>' : s.marker.line ? '<a:ln>' + fillXml(s.marker.line) + '</a:ln>' : '') + '</c:spPr>' : '') + '</c:marker>'
        : '';
      const head = '<c:idx val="' + s.index + '"/><c:order val="' + s.index + '"/>' + tx + spPr(s.look, lined);
      if (kind === 'scatter') {
        const xs = s.categories ? valuesOf(s.categories, s.cachedCategories) : null;
        const x = xs ? '<c:xVal>' + catData(s.categories, xs) + '</c:xVal>' : '';
        if (g.bubbles) {
          return '<c:ser>' + head + '<c:invertIfNegative val="0"/>' + dPts + dLbls + x + '<c:yVal>' + numData(s.values, values) + '</c:yVal>'
            + '<c:bubbleSize>' + numData(s.bubbles, valuesOf(s.bubbles, s.cachedBubbles)) + '</c:bubbleSize><c:bubble3D val="0"/></c:ser>';
        }
        return '<c:ser>' + head + marker + dPts + dLbls + x + '<c:yVal>' + numData(s.values, values) + '</c:yVal><c:smooth val="' + (s.smooth ? 1 : 0) + '"/></c:ser>';
      }
      const cat = categories?.length || s.categories ? '<c:cat>' + catData(s.categories, categories ?? []) + '</c:cat>' : '';
      const val = '<c:val>' + numData(s.values, values) + '</c:val>';
      if (kind === 'pie' || kind === 'doughnut') return '<c:ser>' + head + (s.explosion ? '<c:explosion val="' + s.explosion + '"/>' : '') + dPts + dLbls + cat + val + '</c:ser>';
      if (kind === 'surface') return '<c:ser>' + head + cat + val + '</c:ser>';
      if (kind === 'bar') return '<c:ser>' + head + '<c:invertIfNegative val="0"/>' + dPts + dLbls + cat + val + (deep ? '<c:shape val="box"/>' : '') + '</c:ser>';
      if (kind === 'area') return '<c:ser>' + head + dPts + dLbls + cat + val + '</c:ser>';
      return '<c:ser>' + head + marker + dPts + dLbls + cat + val + (kind === 'line' ? '<c:smooth val="' + (s.smooth ? 1 : 0) + '"/>' : '') + '</c:ser>';
    }).join('');

    const varied = '<c:varyColors val="' + (g.varied ? 1 : 0) + '"/>';
    const ax2 = '<c:axId val="' + ids.cat + '"/><c:axId val="' + ids.val + '"/>';
    const ax3 = ax2 + '<c:axId val="' + ids.ser + '"/>';
    const axes = needsSeriesAxis ? ax3 : (deep && (kind === 'bar' || kind === 'area')) ? ax3 : ax2;
    if (deep && (kind === 'bar' || kind === 'area')) sets.get(g.set).series = true;
    const grouping3d = deep && kind === 'bar' && !g.threeD.cluster ? 'standard' : g.grouping;
    let xml;
    if (kind === 'bar') {
      xml = '<c:' + (deep ? 'bar3DChart' : 'barChart') + '><c:barDir val="' + (g.horizontal ? 'bar' : 'col') + '"/><c:grouping val="' + grouping3d + '"/>' + varied + sers
        + '<c:gapWidth val="' + Math.max(0, Math.min(500, g.gap)) + '"/>'
        + (!deep && g.grouping !== 'clustered' ? '<c:overlap val="100"/>' : !deep && g.overlap ? '<c:overlap val="' + Math.max(-100, Math.min(100, g.overlap)) + '"/>' : '')
        + (deep ? '<c:shape val="box"/>' : '') + axes + '</c:' + (deep ? 'bar3DChart' : 'barChart') + '>';
    } else if (kind === 'line') {
      xml = deep
        ? '<c:line3DChart><c:grouping val="' + g.grouping + '"/>' + varied + sers + ax3 + '</c:line3DChart>'
        : '<c:lineChart><c:grouping val="' + g.grouping + '"/>' + varied + sers + '<c:marker val="1"/>' + ax2 + '</c:lineChart>';
    } else if (kind === 'area') {
      xml = '<c:' + (deep ? 'area3DChart' : 'areaChart') + '><c:grouping val="' + g.grouping + '"/>' + varied + sers + axes + '</c:' + (deep ? 'area3DChart' : 'areaChart') + '>';
    } else if (kind === 'pie') {
      xml = deep ? '<c:pie3DChart>' + varied + sers + '</c:pie3DChart>' : '<c:pieChart>' + varied + sers + '<c:firstSliceAng val="' + g.firstSlice + '"/></c:pieChart>';
    } else if (kind === 'doughnut') {
      xml = '<c:doughnutChart>' + varied + sers + '<c:firstSliceAng val="' + g.firstSlice + '"/><c:holeSize val="' + Math.max(10, Math.min(90, g.hole)) + '"/></c:doughnutChart>';
    } else if (kind === 'scatter') {
      xml = g.bubbles
        ? '<c:bubbleChart>' + varied + sers + '<c:bubbleScale val="100"/>' + ax2 + '</c:bubbleChart>'
        : '<c:scatterChart><c:scatterStyle val="lineMarker"/>' + varied + sers + ax2 + '</c:scatterChart>';
    } else if (kind === 'radar') {
      xml = '<c:radarChart><c:radarStyle val="' + (g.filled ? 'filled' : 'marker') + '"/>' + varied + sers + ax2 + '</c:radarChart>';
    } else {
      xml = '<c:' + (deep ? 'surface3DChart' : 'surfaceChart') + '><c:wireframe val="' + (g.filled ? 0 : 1) + '"/>' + sers + ax3 + '</c:' + (deep ? 'surface3DChart' : 'surfaceChart') + '>';
    }
    groupsXml.push(xml);
  }

  if (!groupsXml.length) return null;

  // The axes each set's groups are drawn on: categories (or X values) along, values up, series into the page.
  const axesXml = [];
  for (const [set, use] of sets) {
    const ids = axisIds(set);
    const find = (type) => chart.axes.find((a) => a.set === set && a.type === type) ?? null;
    const secondary = set === 1;
    const scale = (a) => '<c:scaling>' + (a?.log ? '<c:logBase val="10"/>' : '') + '<c:orientation val="' + (a?.reverse ? 'maxMin' : 'minMax') + '"/>'
      + (a?.max != null ? '<c:max val="' + a.max + '"/>' : '') + (a?.min != null ? '<c:min val="' + a.min + '"/>' : '') + '</c:scaling>';
    const common = (a, id, pos, cross, vertical) => {
      const titleText = a?.title ? (a.title.text ?? (a.title.ref ? (resolve(a.title.ref) ?? [])[0] : null)) : null;
      return '<c:axId val="' + id + '"/>' + scale(a) + '<c:delete val="' + (!a || a.deleted ? 1 : 0) + '"/><c:axPos val="' + pos + '"/>'
        + (a?.major ? '<c:majorGridlines>' + lineSpPr(a.gridXml, a.gridLine) + '</c:majorGridlines>' : '')
        + (a?.minor ? '<c:minorGridlines>' + lineSpPr(a.minorXml, a.minorLine) + '</c:minorGridlines>' : '')
        + (a?.title ? titleXml(a.title, titleText == null ? null : String(titleText), vertical ? -5400000 : null) : '')
        + '<c:numFmt formatCode="General" sourceLinked="1"/>'
        + '<c:majorTickMark val="' + (a?.tickMajor ?? 'out') + '"/><c:minorTickMark val="' + (a?.tickMinor ?? 'none') + '"/>'
        + '<c:tickLblPos val="' + (a?.labels ?? 'nextTo') + '"/>'
        + (a ? lineSpPr(a.lineXml, a.lineLine) : '')
        + txPr(a?.txPr ?? null, a?.font ?? null)
        + '<c:crossAx val="' + cross + '"/>';
    };
    const crosses = (other) => (other?.crossesAt != null && !other.log ? '<c:crossesAt val="' + other.crossesAt + '"/>' : '<c:crosses val="' + (other?.crossesMax ? 'max' : 'autoZero') + '"/>');
    const catAxis = find(0);
    const valAxis = find(1);
    const horizontal = use.horizontal;
    const catPos = horizontal ? (secondary ? 'r' : 'l') : (secondary ? 't' : 'b');
    const valPos = horizontal ? (secondary ? 't' : 'b') : (secondary ? 'r' : 'l');
    if (use.scatter) {
      axesXml.push('<c:valAx>' + common(catAxis, ids.cat, catPos, ids.val, false) + crosses(valAxis) + '<c:crossBetween val="midCat"/>'
        + (catAxis?.majorUnit ? '<c:majorUnit val="' + catAxis.majorUnit + '"/>' : '') + '</c:valAx>');
    } else {
      axesXml.push('<c:catAx>' + common(catAxis, ids.cat, catPos, ids.val, horizontal) + crosses(valAxis)
        + '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>');
    }
    axesXml.push('<c:valAx>' + common(valAxis, ids.val, valPos, ids.cat, !horizontal) + crosses(catAxis)
      + '<c:crossBetween val="' + (use.scatter || catAxis?.between === false ? 'midCat' : 'between') + '"/>'
      + (valAxis?.majorUnit ? '<c:majorUnit val="' + valAxis.majorUnit + '"/>' : '') + (valAxis?.minorUnit ? '<c:minorUnit val="' + valAxis.minorUnit + '"/>' : '') + '</c:valAx>');
    if (use.series) {
      const serAxis = find(2);
      axesXml.push('<c:serAx>' + common(serAxis, ids.ser, 'b', ids.val, false) + '<c:crosses val="autoZero"/></c:serAx>');
    }
  }

  // The title: its own words, a cell's, or — made by Excel from the one series — the series' name.
  let title = '';
  if (chart.title && !(chart.title.flags & 0x40)) {
    let text = chart.title.text ?? (chart.title.ref ? (resolve(chart.title.ref) ?? [])[0] : null);
    if (text == null && chart.series.length === 1) text = first.name ?? (first.nameRef ? (resolve(first.nameRef) ?? [])[0] : null);
    title = titleXml(chart.title, text == null ? null : String(text));
  }
  const view3D = threeD
    ? '<c:view3D><c:rotX val="' + Math.max(-90, Math.min(90, threeD.elevation)) + '"/>'
      + (chart.groups.some((g) => g.kind === 'pie') ? '' : '<c:rotY val="' + (threeD.rotation % 360) + '"/>')
      + (threeD.depth >= 20 && threeD.depth <= 2000 ? '<c:depthPercent val="' + threeD.depth + '"/>' : '')
      + '<c:rAngAx val="' + (threeD.perspective ? 0 : 1) + '"/>' + (threeD.perspective ? '<c:perspective val="' + Math.min(240, threeD.distance) + '"/>' : '') + '</c:view3D>'
    : '';
  const legend = chart.legend
    ? '<c:legend><c:legendPos val="' + chart.legend.pos + '"/><c:overlay val="0"/>' + spPr(chart.legend.look) + txPr(chart.legend.txPr, chart.legend.font) + '</c:legend>'
    : '';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/>'
    + '<c:chart>' + title + '<c:autoTitleDeleted val="' + (title ? 0 : 1) + '"/>' + view3D
    + '<c:plotArea><c:layout/>' + groupsXml.join('') + axesXml.join('') + spPr(chart.plot) + '</c:plotArea>'
    + legend + '<c:plotVisOnly val="' + (chart.visibleOnly ? 1 : 0) + '"/><c:dispBlanksAs val="' + chart.blanks + '"/></c:chart>'
    + spPr(chart.look) + '</c:chartSpace>';
}
