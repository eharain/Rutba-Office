/**
 * DrawingML adapter — the OOXML side, kept at the edge.
 *
 * Deliberately takes XML STRINGS and a part resolver rather than an OOXML
 * package object, so this file has no dependency on `@rutba/ooxml` and the whole
 * drawing package stays consumable by Studio and Mail, which have no packages to
 * hand it.
 *
 * Two things are read:
 *   c:chartSpace   -> a chart spec (categories, series, type)
 *   xdr:wsDr       -> anchors saying WHERE a drawing sits on a sheet
 *
 * What is deliberately not read: DrawingML's own styling — theme colours, fills,
 * effects, fonts. A chart from a customer's file is re-drawn in OUR palette
 * rather than approximated in theirs, which is both honest about what we can do
 * and produces a chart that matches the rest of the product. The original part is
 * preserved untouched either way, so saving does not overwrite their styling
 * with our interpretation of it.
 */
import { emuToPx } from './scene.js';
import { toDataUri, SUPPORTED_GEOMETRY } from './shapes.js';

const attrs = (tag) => {
  const out = {};
  const re = /([\w:.-]+)\s*=\s*"([^"]*)"/g;
  let m;
  while ((m = re.exec(tag))) out[m[1]] = m[2];
  return out;
};
const unesc = (s) =>
  String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/** Strip a namespace prefix: `c:barChart` -> `barChart`. */
const local = (name) => name.replace(/^[^:]*:/, '');

/** All immediate-ish occurrences of an element, by local name. */
function elements(xml, localName) {
  const re = new RegExp('<([\\w]+:)?' + localName + '\\b[^>]*?(/>|>[\\s\\S]*?</([\\w]+:)?' + localName + '>)', 'g');
  return [...String(xml).matchAll(re)].map((m) => m[0]);
}
function firstElement(xml, localName) {
  return elements(xml, localName)[0] ?? null;
}
/** All `<a:t>` text inside a fragment, concatenated — a shape's or title's words. */
function textOf(xml) {
  const re = /<([\w]+:)?t>([\s\S]*?)<\/([\w]+:)?t>/g;
  return [...String(xml).matchAll(re)].map((m) => unesc(m[2])).join('');
}

/** `<c:v>3</c:v>` -> "3" */
function values(xml) {
  return [...String(xml).matchAll(/<([\w]+:)?v>([\s\S]*?)<\/([\w]+:)?v>/g)].map((m) => unesc(m[2]));
}
/** Cached numeric/string points come with an index; honour it so gaps stay gaps. */
function pointList(xml) {
  const out = [];
  for (const pt of elements(xml, 'pt')) {
    const idx = Number(attrs(pt).idx ?? out.length);
    out[idx] = values(pt)[0] ?? null;
  }
  return [...out];
}

const CHART_ELEMENTS = {
  barChart: 'column',
  bar3DChart: 'column',
  lineChart: 'line',
  line3DChart: 'line',
  areaChart: 'area',
  area3DChart: 'area',
  pieChart: 'pie',
  pie3DChart: 'pie',
  doughnutChart: 'doughnut',
  scatterChart: 'scatter',
  bubbleChart: 'scatter',
};

/**
 * Parse a `c:chartSpace` part into a chart spec.
 *
 * @param {string} chartXml
 * @param {object} [opts]
 * @param {'light'|'dark'} [opts.mode]
 * @returns {object|null} a spec for buildChart, or null if unreadable
 */
export function parseChartXml(chartXml, { mode = 'light', width = 480, height = 300 } = {}) {
  if (!chartXml) return null;
  const xml = String(chartXml);

  let type = null;
  let plotXml = null;
  for (const [element, mapped] of Object.entries(CHART_ELEMENTS)) {
    const found = firstElement(xml, element);
    if (found) { type = mapped; plotXml = found; break; }
  }
  if (!type) return null;

  // A bar chart with `barDir=bar` is horizontal; `col` is vertical.
  if (type === 'column') {
    const dir = firstElement(plotXml, 'barDir');
    if (dir && /val="bar"/.test(dir)) type = 'bar';
  }
  const grouping = firstElement(plotXml, 'grouping');
  const stacked = Boolean(grouping && /val="(stacked|percentStacked)"/.test(grouping));

  let categories = null;
  const series = [];
  for (const ser of elements(plotXml, 'ser')) {
    const nameXml = firstElement(ser, 'tx');
    const name = nameXml ? (values(nameXml)[0] ?? null) : null;

    const catXml = firstElement(ser, 'cat');
    if (catXml && !categories) {
      const pts = pointList(catXml);
      if (pts.length) categories = pts.map((p) => (p === undefined ? '' : p));
    }
    const valXml = firstElement(ser, 'val') ?? firstElement(ser, 'yVal');
    const nums = valXml ? pointList(valXml) : [];
    series.push({
      name: name ?? 'Series ' + (series.length + 1),
      values: nums.map((v) => (v === undefined || v === null || v === '' ? null : Number(v))),
    });
  }
  if (!series.length) return null;

  const titleXml = firstElement(xml, 'title');
  const title = titleXml
    ? [...titleXml.matchAll(/<([\w]+:)?t>([\s\S]*?)<\/([\w]+:)?t>/g)].map((m) => unesc(m[2])).join('') || null
    : null;

  const longest = Math.max(...series.map((s) => s.values.length), 0);
  return {
    type,
    stacked,
    title,
    categories: categories ?? Array.from({ length: longest }, (_, i) => String(i + 1)),
    series,
    width,
    height,
    mode,
  };
}

/**
 * Parse `xdr:wsDr` anchors — where drawings sit on a sheet.
 *
 * A two-cell anchor pins both corners to cells (it resizes with them); a
 * one-cell anchor pins the top-left and carries its own extent. Both are
 * returned in cell coordinates plus an EMU offset, because the caller knows the
 * column widths and we do not.
 */
export function parseDrawingAnchors(drawingXml) {
  if (!drawingXml) return [];
  const out = [];

  const readMarker = (markerXml) => {
    if (!markerXml) return null;
    const num = (name) => {
      const el = firstElement(markerXml, name);
      if (!el) return 0;
      const m = />([\s\S]*?)</.exec(el);
      return Number(m ? m[1] : 0) || 0;
    };
    return {
      col: num('col'), colOffsetEmu: num('colOff'),
      row: num('row'), rowOffsetEmu: num('rowOff'),
    };
  };
  const relId = (xml) => {
    const chart = firstElement(xml, 'chart');
    if (chart) return attrs(chart)['r:id'] ?? attrs(chart).id ?? null;
    const blip = firstElement(xml, 'blip');
    if (blip) return attrs(blip)['r:embed'] ?? null;
    return null;
  };
  const kindOf = (xml) => {
    if (firstElement(xml, 'graphicFrame')) return 'chart';
    if (firstElement(xml, 'pic')) return 'image';
    if (firstElement(xml, 'sp')) return 'shape';
    return 'unknown';
  };

  // DOCUMENT ORDER, not grouped by anchor type: the order anchors appear in is
  // their z-order, so collecting all the two-cell ones first would put a shape
  // in front of a picture that was drawn over it.
  const anchorRe = /<([\w]+:)?(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[^>]*>[\s\S]*?<\/([\w]+:)?\2>/g;
  const ANCHOR_TYPES = { twoCellAnchor: 'twoCell', oneCellAnchor: 'oneCell', absoluteAnchor: 'absolute' };

  for (const match of String(drawingXml).matchAll(anchorRe)) {
    const anchor = match[0];
    const anchorType = ANCHOR_TYPES[match[2]];
    const ext = firstElement(anchor, 'ext');
    const extent = ext ? attrs(ext) : null;
    const to = readMarker(firstElement(anchor, 'to'));

    // An absolute anchor carries no cell reference at all — it is positioned in
    // EMUs from the sheet's origin. Converting it to a synthetic marker keeps it
    // visible instead of dropping it on the floor.
    let from = readMarker(firstElement(anchor, 'from'));
    if (!from && anchorType === 'absolute') {
      const pos = firstElement(anchor, 'pos');
      const at = pos ? attrs(pos) : {};
      from = { col: 0, colOffsetEmu: Number(at.x ?? 0), row: 0, rowOffsetEmu: Number(at.y ?? 0) };
    }

    out.push({
      kind: kindOf(anchor),
      relationshipId: relId(anchor),
      from,
      to,
      widthPx: extent?.cx ? emuToPx(extent.cx) : null,
      heightPx: extent?.cy ? emuToPx(extent.cy) : null,
      anchorType,
      xml: anchor,
      name: (() => {
        const pr = firstElement(anchor, 'cNvPr');
        return pr ? attrs(pr).name ?? null : null;
      })(),
    });
  }
  return out;
}

/** A colour reference: srgbClr, schemeClr or an explicit noFill. */
function readColour(xml) {
  if (!xml) return null;
  if (firstElement(xml, 'noFill')) return { type: 'none' };
  const srgb = firstElement(xml, 'srgbClr');
  if (srgb) return { type: 'srgb', value: attrs(srgb).val ?? '000000' };
  const scheme = firstElement(xml, 'schemeClr');
  if (scheme) return { type: 'scheme', value: attrs(scheme).val ?? 'accent1' };
  return null;
}

/**
 * Parse an `xdr:sp` into a shape descriptor.
 *
 * Reads the preset geometry, the fill, the outline and any text. Deliberately
 * NOT read: effects, gradients, 3-D, theme lookups. Those are preserved in the
 * part and simply not drawn — a flat shape in the right place and roughly the
 * right colour is far more useful than a blank.
 */
export function parseShapeXml(spXml) {
  if (!spXml) return null;
  const prstGeom = firstElement(spXml, 'prstGeom');
  const geometry = prstGeom ? (attrs(prstGeom).prst ?? 'rect') : 'rect';

  const spPr = firstElement(spXml, 'spPr') ?? spXml;
  // The outline's own fill must not be mistaken for the shape's fill.
  const lnXml = firstElement(spPr, 'ln');
  const fillSource = lnXml ? spPr.replace(lnXml, '') : spPr;

  const text = textOf(spXml);
  const rPr = firstElement(spXml, 'rPr');
  const nvPr = firstElement(spXml, 'cNvPr');

  return {
    kind: 'shape',
    geometry,
    name: nvPr ? (attrs(nvPr).name ?? null) : null,
    fill: readColour(fillSource),
    stroke: lnXml ? readColour(lnXml) : null,
    // Scene units are pixels, so the outline width converts here rather than
    // leaking EMUs into the renderer.
    strokeWidth: lnXml && attrs(lnXml).w ? Math.max(1, Math.round(emuToPx(attrs(lnXml).w))) : null,
    text: text || null,
    textBold: rPr ? attrs(rPr).b === '1' : false,
    textSize: rPr && attrs(rPr).sz ? Number(attrs(rPr).sz) / 100 : null,
    textColour: (() => {
      const body = firstElement(spXml, 'txBody');
      return body ? readColour(firstElement(body, 'rPr') ?? '') : null;
    })(),
    // False means we draw a box instead of the real outline. The caller can say
    // so; the part itself is preserved either way.
    supported: SUPPORTED_GEOMETRY.includes(geometry),
  };
}

/** Parse an `xdr:pic` into a picture descriptor, resolving the image bytes. */
export function parsePictureXml(picXml, { resolveRelationship, readPartBinary } = {}) {
  if (!picXml) return null;
  const blip = firstElement(picXml, 'blip');
  const relId = blip ? (attrs(blip)['r:embed'] ?? attrs(blip).embed ?? null) : null;
  const nvPr = firstElement(picXml, 'cNvPr');

  let href = null;
  let partName = null;
  if (relId && resolveRelationship && readPartBinary) {
    partName = resolveRelationship(relId);
    if (partName) href = toDataUri(readPartBinary(partName), partName);
  }
  return {
    kind: 'picture',
    name: nvPr ? (attrs(nvPr).name ?? null) : null,
    relationshipId: relId,
    part: partName,
    href,
    supported: Boolean(href),
  };
}

/**
 * Everything drawn on one sheet, ready to render.
 *
 * @param {object} args
 * @param {string} args.drawingXml            the xdr:wsDr part
 * @param {(relId: string) => string|null} args.resolveRelationship  relId -> part name
 * @param {(partName: string) => string|null} args.readPart          part name -> XML
 */
export function readSheetDrawings({
  drawingXml, resolveRelationship, readPart, readPartBinary, mode = 'light',
}) {
  const anchors = parseDrawingAnchors(drawingXml);
  return anchors.map((anchor) => {
    if (anchor.kind === 'chart') {
      if (!anchor.relationshipId) return { ...anchor, spec: null, descriptor: null };
      const partName = resolveRelationship(anchor.relationshipId);
      const chartXml = partName ? readPart(partName) : null;
      const spec = chartXml
        ? parseChartXml(chartXml, {
          mode,
          width: Math.max(160, Math.round(anchor.widthPx ?? 480)),
          height: Math.max(120, Math.round(anchor.heightPx ?? 300)),
        })
        : null;
      return { ...anchor, part: partName, spec, descriptor: null };
    }
    if (anchor.kind === 'shape') {
      return { ...anchor, spec: null, descriptor: parseShapeXml(anchor.xml) };
    }
    if (anchor.kind === 'image') {
      return {
        ...anchor,
        spec: null,
        descriptor: parsePictureXml(anchor.xml, { resolveRelationship, readPartBinary }),
      };
    }
    return { ...anchor, spec: null, descriptor: null };
  });
}

export { local, attrs as parseAttributes };
