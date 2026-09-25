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
    const entry = {
      name: name ?? 'Series ' + (series.length + 1),
      values: nums.map((v) => (v === undefined || v === null || v === '' ? null : Number(v))),
    };
    // A line series with its marker switched off (symbol "none").
    const marker = firstElement(ser, 'marker');
    if (type === 'line' && marker && /<([\w]+:)?symbol\s+val="none"/.test(marker)) entry.markers = false;
    // A scatter series plots against its own X values. Text X values (a
    // strRef) are not numbers, and Excel then plots the points 1, 2, 3…
    // across, which is what no `x` means to the drawing.
    if (type === 'scatter') {
      const xXml = firstElement(ser, 'xVal');
      if (xXml && !firstElement(xXml, 'strRef') && !firstElement(xXml, 'strLit')) {
        entry.x = pointList(xXml).map((v) => (v === undefined || v === null || v === '' ? null : Number(v)));
      }
      // The series' own line, not its marker's: switched off, it is markers only.
      const own = ser.replace(/<([\w]+:)?marker>[\s\S]*?<\/([\w]+:)?marker>/g, '').replace(/<([\w]+:)?dPt>[\s\S]*?<\/([\w]+:)?dPt>/g, '');
      const spPr = firstElement(own, 'spPr');
      const ln = spPr ? firstElement(spPr, 'ln') : null;
      entry.line = !(ln && firstElement(ln, 'noFill'));
      entry.smooth = /<([\w]+:)?smooth\s+val="(1|true)"/.test(own);
    }
    series.push(entry);
  }
  if (!series.length) return null;

  // Excel's three scatter looks, from what the series say: markers only
  // (every line off), smooth lines, or straight ones.
  let scatterStyle;
  if (type === 'scatter') {
    const style = attrs(firstElement(plotXml, 'scatterStyle') ?? '').val;
    scatterStyle = series.every((s) => s.line === false) || style === 'marker' || style === 'none'
      ? 'markers'
      : series.some((s) => s.smooth) ? 'smooth' : 'lines';
    for (const s of series) { delete s.line; delete s.smooth; }
  }

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
    ...(scatterStyle ? { scatterStyle } : {}),
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
    // A slicer's graphic frame sits inside mc:AlternateContent with a plain
    // rectangle as the fallback: the frame, not the rectangle, is what it is.
    if (/drawing\/2010\/slicer/.test(xml) && /<([\w]+:)?slicer\b/.test(xml)) return 'slicer';
    // A group: its own grpSp, whatever its members are.
    if (anchorBody(xml)?.name === 'grpSp') return 'group';
    if (firstElement(xml, 'graphicFrame')) return 'chart';
    if (firstElement(xml, 'pic')) return 'image';
    if (firstElement(xml, 'sp')) return 'shape';
    // A connector is a shape with a line for a body: the same reader, and
    // the line runs corner to corner of its box, flipped as the box says.
    if (firstElement(xml, 'cxnSp')) return 'shape';
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

    const kind = kindOf(anchor);
    out.push({
      kind,
      relationshipId: kind === 'slicer' ? null : relId(anchor),
      // A slicer panel is found by its name in the sheet's slicers part.
      ...(kind === 'slicer' ? { slicerName: unesc(attrs((/<([\w]+:)?slicer\b([^>]*?)\/?>/.exec(anchor) ?? [])[2] ?? '').name ?? '') } : {}),
      from,
      to,
      widthPx: extent?.cx ? emuToPx(extent.cx) : null,
      heightPx: extent?.cy ? emuToPx(extent.cy) : null,
      anchorType,
      xml: anchor,
      name: (() => {
        const pr = firstElement(anchor, 'cNvPr');
        return pr ? (attrs(pr.slice(0, pr.indexOf('>') + 1)).name !== undefined ? unesc(attrs(pr.slice(0, pr.indexOf('>') + 1)).name) : null) : null;
      })(),
      // The drawing's own id (cNvPr id), which stays put when the order
      // changes, and whether the Selection Pane has hidden it.
      nvId: (() => {
        const pr = firstElement(anchor, 'cNvPr');
        return pr ? Number(attrs(pr.slice(0, pr.indexOf('>') + 1)).id) || 0 : 0;
      })(),
      hidden: (() => {
        const pr = firstElement(anchor, 'cNvPr');
        return pr ? /^(1|true)$/.test(attrs(pr.slice(0, pr.indexOf('>') + 1)).hidden ?? '') : false;
      })(),
      editAs: attrs(/^<[^>]*>/.exec(anchor)[0]).editAs ?? null,
      // The size the drawing states for itself (its first non-empty xfrm),
      // which a "don't size with cells" anchor keeps.
      xfrmPx: (() => {
        const m = /<a:xfrm\b[^>]*>[\s\S]*?<a:ext\b([^>]*)\/>/.exec(anchor);
        const a = m ? attrs(m[1]) : null;
        return a && Number(a.cx) > 0 && Number(a.cy) > 0 ? { width: Math.round(emuToPx(Number(a.cx))), height: Math.round(emuToPx(Number(a.cy))) } : null;
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
  // The opening tag only: a scheme colour carries its modifiers as children
  // (<a:shade val="50000"/>), and the last val in the element is theirs.
  if (scheme) return { type: 'scheme', value: attrs(scheme.slice(0, scheme.indexOf('>'))).val ?? 'accent1' };

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

  // A shape Excel has just drawn states no fill and no line of its own: both
  // come from the theme through `<xdr:style>` — fillRef idx 1 in accent1, lnRef
  // idx 2 in accent1 shaded — and a reader that stops at spPr draws it empty.
  // Index 0 means "none", and is left alone.
  const style = firstElement(spXml, 'style');
  const ref = (name) => {
    const el = style ? firstElement(style, name) : null;
    if (!el || attrs(el).idx === '0') return null;
    return readColour(el);
  };
  const xfrm = firstElement(spPr, 'xfrm');

  return {
    kind: 'shape',
    geometry,
    name: nvPr ? (attrs(nvPr).name ?? null) : null,
    fill: readColour(fillSource) ?? ref('fillRef'),
    stroke: (lnXml ? readColour(lnXml) : null) ?? ref('lnRef'),
    connector: Boolean(firstElement(spXml, 'cxnSp')),
    flipH: xfrm ? attrs(xfrm).flipH === '1' : false,
    flipV: xfrm ? attrs(xfrm).flipV === '1' : false,
    // Degrees clockwise about the centre; the file keeps 60,000ths of one.
    rotation: xfrm && attrs(xfrm).rot ? Number(attrs(xfrm).rot) / 60000 : 0,


    // Scene units are pixels, so the outline width converts here rather than
    // leaking EMUs into the renderer.
    strokeWidth: lnXml && attrs(lnXml).w ? Math.max(1, Math.round(emuToPx(attrs(lnXml).w))) : null,
    text: text || null,
    textBold: rPr ? attrs(rPr).b === '1' : false,
    textSize: rPr && attrs(rPr).sz ? Number(attrs(rPr).sz) / 100 : null,
    textColour: (() => {
      const body = firstElement(spXml, 'txBody');
      return (body ? readColour(firstElement(body, 'rPr') ?? '') : null) ?? ref('fontRef');
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
    // A picture turns and flips about its centre, as its xfrm says.
    ...(() => {
      const x = xfrmOf(picXml);
      return x ? { rotation: x.rot, flipH: x.flipH, flipV: x.flipV } : {};
    })(),
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
    if (anchor.kind === 'group') {
      const body = anchorBody(anchor.xml);
      const own = xfrmOf(/<([\w]+:)?grpSpPr\b[\s\S]*?<\/([\w]+:)?grpSpPr>/.exec(body.xml)?.[0] ?? '') ?? {};
      return {
        ...anchor,
        spec: null,
        descriptor: {
          kind: 'group', rotation: own.rot || 0, flipH: Boolean(own.flipH), flipV: Boolean(own.flipV),
          children: readGroup(body.xml, { resolveRelationship, readPart, readPartBinary, mode }),
        },
      };
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


/**
 * The top-level elements of a fragment, by a depth count of its tags — what
 * a regex cannot do once groups nest. Each: its local name, where it starts
 * and ends, and its XML.
 */
export function childElements(xml) {
  const out = [];
  const re = /<(\/?)([\w]+:)?([\w]+)\b[^>]*?(\/?)>/g;
  let depth = 0;
  let open = null;
  let m;
  while ((m = re.exec(xml))) {
    if (m[0].startsWith('<?') || m[0].startsWith('<!')) continue;
    const closing = m[1] === '/';
    const selfClosing = m[4] === '/';
    if (closing) {
      depth -= 1;
      if (depth === 0 && open) {
        out.push({ name: open.name, start: open.start, end: m.index + m[0].length, xml: xml.slice(open.start, m.index + m[0].length) });
        open = null;
      }
    } else if (selfClosing) {
      if (depth === 0) out.push({ name: m[3], start: m.index, end: m.index + m[0].length, xml: m[0] });
    } else {
      if (depth === 0) open = { name: m[3], start: m.index };
      depth += 1;
    }
  }
  return out;
}

/** An element's own box, from its xfrm (`a:xfrm` or `xdr:xfrm`): EMU offset, extent, turn and flips. */
export function xfrmOf(elementXml) {
  const m = /<([\w]+:)?xfrm\b([^>]*)>([\s\S]*?)<\/([\w]+:)?xfrm>/.exec(elementXml);
  if (!m) return null;
  const a = attrs(m[2]);
  const off = attrs((/<([\w]+:)?off\b([^>]*?)\/>/.exec(m[3]) ?? [])[2] ?? '');
  const ext = attrs((/<([\w]+:)?ext\b([^>]*?)\/>/.exec(m[3]) ?? [])[2] ?? '');
  const chOff = /<([\w]+:)?chOff\b([^>]*?)\/>/.exec(m[3]);
  const chExt = /<([\w]+:)?chExt\b([^>]*?)\/>/.exec(m[3]);
  return {
    x: Number(off.x) || 0, y: Number(off.y) || 0, cx: Number(ext.cx) || 0, cy: Number(ext.cy) || 0,
    rot: a.rot ? Number(a.rot) / 60000 : 0, flipH: a.flipH === '1', flipV: a.flipV === '1',
    ...(chOff ? { chX: Number(attrs(chOff[2]).x) || 0, chY: Number(attrs(chOff[2]).y) || 0 } : {}),
    ...(chExt ? { chCx: Number(attrs(chExt[2]).cx) || 0, chCy: Number(attrs(chExt[2]).cy) || 0 } : {}),
  };
}

/** The drawing element an anchor holds (sp, pic, grpSp, graphicFrame, cxnSp), without its markers. */
export function anchorBody(anchorXml) {
  const inner = anchorXml.replace(/^<[^>]*>/, '').replace(/<\/[^>]*>$/, '');
  return childElements(inner).find((e) => /^(sp|pic|grpSp|graphicFrame|cxnSp|AlternateContent)$/.test(e.name)) ?? null;
}

/**
 * A group's members, each as the reader would read it on its own, with its
 * box inside the group as fractions of the group's child space — so the
 * caller lays them out in whatever pixels the group's anchor gives it.
 * Groups inside groups are flattened into the same space.
 */
export function readGroup(grpXml, { resolveRelationship, readPart, readPartBinary, mode = 'light' } = {}) {
  const own = xfrmOf(/<([\w]+:)?grpSpPr\b[\s\S]*?<\/([\w]+:)?grpSpPr>/.exec(grpXml)?.[0] ?? '') ?? { x: 0, y: 0, cx: 1, cy: 1 };
  const space = { x: own.chX ?? own.x, y: own.chY ?? own.y, cx: own.chCx || own.cx || 1, cy: own.chCy || own.cy || 1 };
  const inner = grpXml.replace(/^<[^>]*>/, '').replace(/<\/[^>]*>$/, '');
  const out = [];
  for (const el of childElements(inner)) {
    if (!/^(sp|pic|grpSp|graphicFrame|cxnSp)$/.test(el.name)) continue;
    const box = xfrmOf(el.name === 'grpSp' ? (/<([\w]+:)?grpSpPr\b[\s\S]*?<\/([\w]+:)?grpSpPr>/.exec(el.xml)?.[0] ?? '') : el.xml);
    if (!box) continue;
    const frac = {
      x: (box.x - space.x) / space.cx, y: (box.y - space.y) / space.cy,
      w: box.cx / space.cx, h: box.cy / space.cy,
    };
    if (el.name === 'grpSp') {
      for (const c of readGroup(el.xml, { resolveRelationship, readPart, readPartBinary, mode })) {
        out.push({ ...c, frac: { x: frac.x + c.frac.x * frac.w, y: frac.y + c.frac.y * frac.h, w: c.frac.w * frac.w, h: c.frac.h * frac.h } });
      }
      continue;
    }
    const name = attrs((/<([\w]+:)?cNvPr\b([^>]*?)\/?>/.exec(el.xml) ?? [])[2] ?? '').name ?? null;
    if (el.name === 'sp' || el.name === 'cxnSp') {
      out.push({ kind: 'shape', name, frac, descriptor: parseShapeXml(el.xml) });
    } else if (el.name === 'pic') {
      out.push({ kind: 'image', name, frac, descriptor: { ...parsePictureXml(el.xml, { resolveRelationship, readPartBinary }), rotation: box.rot, flipH: box.flipH, flipV: box.flipV } });
    } else {
      const chart = firstElement(el.xml, 'chart');
      const relId = chart ? (attrs(chart)['r:id'] ?? null) : null;
      const part = relId && resolveRelationship ? resolveRelationship(relId) : null;
      const xml = part && readPart ? readPart(part) : null;
      out.push({ kind: 'chart', name, frac, part, spec: xml ? parseChartXml(xml, { mode }) : null });
    }
  }
  return out;
}

export { local, attrs as parseAttributes };
