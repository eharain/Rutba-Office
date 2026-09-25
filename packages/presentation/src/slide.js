// A slide, read into a scene.
//
// The scene is deliberately format-neutral: positions in pixels, colours
// resolved to hex, text as paragraphs of runs. The renderer draws it, the
// editor edits it, and neither needs to know that `<a:off x="914400"/>` means
// one inch from the left.
//
// Three OOXML habits shape this file:
//   - A shape's geometry is often *not* on the shape. A placeholder inherits
//     position, size and text style from the matching placeholder on its
//     layout, and that from the master. Resolving that chain is most of the
//     work of rendering a real deck correctly.
//   - Colours are references. `<a:schemeClr val="accent1"/>` means "look it up
//     in the theme, through the master's colour map, then apply these
//     transforms".
//   - Text style cascades: run properties, then paragraph properties, then the
//     shape's list style, then the placeholder's inherited style, then the
//     master's text styles, then the theme default.

import { parse, kids, first, all, textOf } from '@rutba/office-formats/xml';
import { bulletGlyph } from '@rutba/drawing/glyphs';
import { parseChartXml } from '@rutba/drawing';
import { emuToPx, szToPt, rotToDeg, applyColorTransforms, PRESET_COLORS, pctOf } from './units.js';

const A = (n) => `a:${n}`;
const P = (n) => `p:${n}`;

/** Read `<a:xfrm>` into pixel geometry. */
function readXfrm(spPr) {
  const xfrm = spPr && kids(spPr, A('xfrm'))[0];
  if (!xfrm) return null;
  const off = kids(xfrm, A('off'))[0];
  const ext = kids(xfrm, A('ext'))[0];
  if (!off && !ext) return null;
  return {
    x: emuToPx(off?.attrs.x),
    y: emuToPx(off?.attrs.y),
    w: emuToPx(ext?.attrs.cx),
    h: emuToPx(ext?.attrs.cy),
    rot: rotToDeg(xfrm.attrs.rot),
    flipH: xfrm.attrs.flipH === '1',
    flipV: xfrm.attrs.flipV === '1',
  };
}

/** Resolve one colour element (the child of a fill or a line). */
function readColor(node, theme) {
  if (!node) return null;
  const transforms = kids(node).map((c) => ({ name: c.name.replace(/^a:/, ''), val: c.attrs.val }));
  let base = null;
  if (node.name === A('srgbClr')) base = `#${node.attrs.val}`;
  else if (node.name === A('sysClr')) base = `#${node.attrs.lastClr || '000000'}`;
  else if (node.name === A('prstClr')) base = PRESET_COLORS[node.attrs.val] || '#000000';
  else if (node.name === A('schemeClr')) base = theme?.color(node.attrs.val) || '#000000';
  else if (node.name === A('scrgbClr')) {
    const to255 = (v) => Math.round((parseFloat(v) / 100000) * 255);
    base = `#${[node.attrs.r, node.attrs.g, node.attrs.b].map((v) => to255(v).toString(16).padStart(2, '0')).join('')}`;
  }
  if (!base) return null;
  const out = applyColorTransforms(base, transforms);
  return typeof out === 'string' ? { hex: out, alpha: 1 } : out;
}

function colorChildOf(node, theme) {
  if (!node) return null;
  for (const c of kids(node)) {
    if (c.name.endsWith('Clr')) return readColor(c, theme);
  }
  return null;
}

/** Solid, gradient (every stop, and the angle), picture (its relationship id and whether it tiles) or none. */
function readFill(spPr, theme) {
  if (!spPr) return null;
  if (kids(spPr, A('noFill'))[0]) return { type: 'none' };
  const solid = kids(spPr, A('solidFill'))[0];
  if (solid) {
    const c = colorChildOf(solid, theme);
    return c ? { type: 'solid', color: c.hex, alpha: c.alpha } : null;
  }
  const grad = kids(spPr, A('gradFill'))[0];
  if (grad) {
    const stops = all(grad, A('gs'))
      .map((gs) => {
        const c = colorChildOf(gs, theme);
        return c ? { offset: pctOf(gs.attrs.pos, 0), color: c.hex, alpha: c.alpha } : null;
      })
      .filter(Boolean);
    const lin = first(grad, A('lin'));
    return stops.length ? { type: 'gradient', stops, angle: rotToDeg(lin?.attrs.ang) } : null;
  }
  const blip = kids(spPr, A('blipFill'))[0];
  if (blip) {
    const b = first(blip, A('blip'));
    return { type: 'picture', embed: b?.attrs['r:embed'] || null, tile: Boolean(kids(blip, A('tile'))[0]) };
  }
  const pattern = kids(spPr, A('pattFill'))[0];
  if (pattern) {
    const fg = colorChildOf(kids(pattern, A('fgClr'))[0], theme);
    return { type: 'solid', color: fg?.hex || '#888888', alpha: fg?.alpha ?? 1 };
  }
  return null;
}

/**
 * PowerPoint's three reflection gallery presets, by name — "tight",
 * "half" and "full" — as their own `<a:reflection>` attributes (EMU-ish
 * 60,000ths and thousandths, the way DrawingML already has blur, alpha and
 * angle elsewhere). Shared with `deck.js`, which writes the very same
 * numbers `reflectionKindOf` reads back, so the two can never drift apart.
 */
export const REFLECTION_PRESETS = {
  tight: { blurRad: 6350, stA: 50000, stPos: 0, endA: 300, endPos: 35000, dist: 0 },
  half: { blurRad: 6350, stA: 50000, stPos: 0, endA: 300, endPos: 55000, dist: 12700 },
  full: { blurRad: 6350, stA: 55000, stPos: 0, endA: 0, endPos: 90000, dist: 0 },
};

/** Which preset a `<a:reflection>` most resembles, by its own endPos and dist — the pane's own state, not a file another program might have hand-built. */
function reflectionKindOf(endPos, dist) {
  let best = 'half';
  let bestDiff = Infinity;
  for (const [kind, p] of Object.entries(REFLECTION_PRESETS)) {
    const diff = Math.abs(p.endPos - endPos) + Math.abs(p.dist - dist);
    if (diff < bestDiff) { bestDiff = diff; best = kind; }
  }
  return best;
}

/**
 * Effects on a shape: the outer shadow (in pixels and degrees — DrawingML's
 * dir runs clockwise from the right), a glow, soft edges and a reflection.
 * null when the shape states none of them; an empty effect list, which
 * turns a style's shadow off too, is `{ shadow: null }` — the one falsy
 * value every reader of this already checks for, so a new effect being
 * absent needs no `?? null` at every call site that only cares about the
 * shadow.
 */
function readEffects(spPr, theme) {
  const lst = spPr && kids(spPr, A('effectLst'))[0];
  if (!lst) return null;
  const out = { shadow: null };
  const sh = kids(lst, A('outerShdw'))[0];
  if (sh) {
    const c = colorChildOf(sh, theme);
    out.shadow = {
      blurPx: (Number(sh.attrs.blurRad || 0) / 12700) * (96 / 72),
      distPx: (Number(sh.attrs.dist || 0) / 12700) * (96 / 72),
      dir: Number(sh.attrs.dir || 0) / 60000,
      color: c?.hex || '#000000',
      alpha: c?.alpha ?? 1,
    };
  }
  const glow = kids(lst, A('glow'))[0];
  if (glow) {
    const c = colorChildOf(glow, theme);
    out.glow = { radiusPt: (Number(glow.attrs.rad || 0) / 12700), color: c?.hex || '#000000', alpha: c?.alpha ?? 1 };
  }
  const softEdge = kids(lst, A('softEdge'))[0];
  if (softEdge) out.softEdge = { radiusPt: Number(softEdge.attrs.rad || 0) / 12700 };
  const reflection = kids(lst, A('reflection'))[0];
  if (reflection) out.reflection = reflectionKindOf(Number(reflection.attrs.endPos || 0), Number(reflection.attrs.dist || 0));
  return out;
}

function readLine(spPr, theme) {
  const ln = spPr && kids(spPr, A('ln'))[0];
  if (!ln) return null;
  if (kids(ln, A('noFill'))[0]) return { type: 'none' };
  const solid = kids(ln, A('solidFill'))[0];
  const c = solid ? colorChildOf(solid, theme) : null;
  const dash = kids(ln, A('prstDash'))[0]?.attrs.val;
  return {
    width: ln.attrs.w ? Number(ln.attrs.w) / 12700 : 1,
    color: c?.hex || null,
    alpha: c?.alpha ?? 1,
    dash: dash && dash !== 'solid' ? dash : null,
  };
}

/** Run properties → the shape the renderer wants. */
function readRunProps(rPr, theme) {
  if (!rPr) return {};
  const fill = kids(rPr, A('solidFill'))[0];
  const color = fill ? colorChildOf(fill, theme) : null;
  const latin = kids(rPr, A('latin'))[0];
  const hlink = kids(rPr, A('hlinkClick'))[0];
  const out = {};
  if (rPr.attrs.sz) out.size = szToPt(rPr.attrs.sz);
  if (rPr.attrs.b != null) out.bold = rPr.attrs.b === '1';
  if (rPr.attrs.i != null) out.italic = rPr.attrs.i === '1';
  if (rPr.attrs.u && rPr.attrs.u !== 'none') out.underline = true;
  if (rPr.attrs.strike && rPr.attrs.strike !== 'noStrike') out.strike = true;
  if (rPr.attrs.baseline) out.baseline = Number(rPr.attrs.baseline) > 0 ? 'super' : 'sub';
  if (rPr.attrs.spc) out.spacing = Number(rPr.attrs.spc) / 100;
  if (rPr.attrs.cap && rPr.attrs.cap !== 'none') out.caps = rPr.attrs.cap;
  if (color) {
    out.color = color.hex;
    if (color.alpha < 1) out.colorAlpha = color.alpha;
  }
  const highlight = kids(rPr, A('highlight'))[0];
  if (highlight) out.highlight = colorChildOf(highlight, theme)?.hex || null;
  if (latin?.attrs.typeface) out.font = theme?.font(latin.attrs.typeface) || latin.attrs.typeface;
  if (hlink) out.link = hlink.attrs['r:id'] || true;
  return out;
}

function readParagraphProps(pPr, theme) {
  if (!pPr) return {};
  const out = {};
  if (pPr.attrs.lvl) out.level = Number(pPr.attrs.lvl);
  if (pPr.attrs.algn) out.align = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' }[pPr.attrs.algn] || 'left';
  if (pPr.attrs.marL) out.indent = emuToPx(pPr.attrs.marL);
  if (pPr.attrs.indent) out.hanging = emuToPx(pPr.attrs.indent);
  const lnSpc = kids(pPr, A('lnSpc'))[0];
  if (lnSpc) {
    const pct = first(lnSpc, A('spcPct'));
    const pts = first(lnSpc, A('spcPts'));
    if (pct) out.lineHeight = Number(pct.attrs.val) / 100000;
    else if (pts) out.lineHeightPt = Number(pts.attrs.val) / 100;
  }
  const spcBef = first(kids(pPr, A('spcBef'))[0], A('spcPts'));
  if (spcBef) out.spaceBefore = Number(spcBef.attrs.val) / 100;
  const spcAft = first(kids(pPr, A('spcAft'))[0], A('spcPts'));
  if (spcAft) out.spaceAfter = Number(spcAft.attrs.val) / 100;

  if (kids(pPr, A('buNone'))[0]) out.bullet = { type: 'none' };
  const buChar = kids(pPr, A('buChar'))[0];
  const buAuto = kids(pPr, A('buAutoNum'))[0];
  // The character is stored for a symbol font — "§" in Wingdings is a small
  // square — and drawn in whatever font the slide has, so it is mapped to the
  // Unicode character that looks the same everywhere.
  const buFont = kids(pPr, A('buFont'))[0]?.attrs.typeface;
  if (buChar) out.bullet = { type: 'char', char: bulletGlyph(buChar.attrs.char || '•', buFont) };

  if (buAuto) out.bullet = { type: 'number', scheme: buAuto.attrs.type || 'arabicPeriod', start: Number(buAuto.attrs.startAt || 1) };
  const buClr = kids(pPr, A('buClr'))[0];
  if (buClr && out.bullet) out.bullet.color = colorChildOf(buClr, theme)?.hex;
  return out;
}

/** `<p:txBody>` → paragraphs of runs. */
function readTextBody(txBody, theme) {
  if (!txBody) return null;
  const bodyPr = kids(txBody, A('bodyPr'))[0];
  const paragraphs = [];
  for (const p of kids(txBody, A('p'))) {
    const pPr = kids(p, A('pPr'))[0];
    const props = readParagraphProps(pPr, theme);
    const endProps = readRunProps(kids(p, A('endParaRPr'))[0], theme);
    const runs = [];
    for (const child of kids(p)) {
      if (child.name === A('r')) {
        const t = kids(child, A('t'))[0];
        runs.push({ text: textOf(t), ...readRunProps(kids(child, A('rPr'))[0], theme) });
      } else if (child.name === A('br')) {
        runs.push({ text: '\n', break: true });
      } else if (child.name === A('fld')) {
        // A field — slide number, date. Its cached text is what PowerPoint drew.
        const t = kids(child, A('t'))[0];
        runs.push({ text: textOf(t), field: child.attrs.type || 'field', ...readRunProps(kids(child, A('rPr'))[0], theme) });
      }
    }
    paragraphs.push({ ...props, runs, endProps });
  }
  const anchorMap = { t: 'top', ctr: 'middle', b: 'bottom' };
  return {
    paragraphs,
    anchor: anchorMap[bodyPr?.attrs.anchor] || 'top',
    // Which way the words run and how many columns they fill: PowerPoint's
    // vert (horz, vert = down, vert270 = up, eaVert = stacked) and numCol.
    vert: bodyPr?.attrs.vert || 'horz',
    columns: Number(bodyPr?.attrs.numCol || 1) || 1,
    wrap: bodyPr?.attrs.wrap !== 'none',
    autofit: Boolean(kids(bodyPr || { children: [] }, A('normAutofit'))[0]),
    insets: {
      l: bodyPr?.attrs.lIns != null ? emuToPx(bodyPr.attrs.lIns) : 7.2,
      t: bodyPr?.attrs.tIns != null ? emuToPx(bodyPr.attrs.tIns) : 3.6,
      r: bodyPr?.attrs.rIns != null ? emuToPx(bodyPr.attrs.rIns) : 7.2,
      b: bodyPr?.attrs.bIns != null ? emuToPx(bodyPr.attrs.bIns) : 3.6,
    },
  };
}

function placeholderOf(sp) {
  const nv = kids(sp, P('nvSpPr'))[0] || kids(sp, P('nvPicPr'))[0] || kids(sp, P('nvGraphicFramePr'))[0];
  const nvPr = nv && kids(nv, P('nvPr'))[0];
  const ph = nvPr && kids(nvPr, P('ph'))[0];
  if (!ph) return null;
  return { type: ph.attrs.type || 'body', idx: ph.attrs.idx != null ? String(ph.attrs.idx) : null };
}

function nameOf(sp) {
  const nv = kids(sp, P('nvSpPr'))[0] || kids(sp, P('nvPicPr'))[0] || kids(sp, P('nvGraphicFramePr'))[0] || kids(sp, P('nvCxnSpPr'))[0] || kids(sp, P('nvGrpSpPr'))[0];
  const cNv = nv && (kids(nv, P('cNvPr'))[0] || null);
  return { id: cNv?.attrs.id || null, name: cNv?.attrs.name || '', hidden: cNv?.attrs.hidden === '1' };
}

/**
 * A group's placement composed with one child's own local transform.
 *
 * `container` is the group's own box — `offX/offY/extX/extY` where it sits,
 * `chOffX/chOffY/chExtX/chExtY` the child coordinate window that box stands
 * for, `rot`/`flipH`/`flipV` the group's own — already resolved into
 * whatever outer frame `container` itself is expressed in (absolute slide
 * units for a top-level group, or another group's own frame for a nested
 * one: the same shape this function returns, so nesting composes by calling
 * it again with the result). `local` is the child's own off/ext/rot/flip
 * exactly as its own xfrm states them, in the group's child space.
 *
 * DrawingML's own order: map through chOff/chExt onto the group's box
 * (translate and scale — a resize after grouping is exactly a stretch of
 * this), then the group's own flip (mirrored about its centre, which also
 * flips the sense of any rotation under it), then the group's own rotation
 * (about the same centre). Units are whatever `container`'s are — pixels or
 * EMU both work, since every term is a ratio or a sum of like units.
 */
function composeGroupChild(container, local) {
  const sx = container.chExtX ? container.extX / container.chExtX : 1;
  const sy = container.chExtY ? container.extY / container.chExtY : 1;
  let x = container.offX + (local.offX - container.chOffX) * sx;
  let y = container.offY + (local.offY - container.chOffY) * sy;
  let w = local.extX * sx;
  let h = local.extY * sy;
  let rot = local.rot || 0;
  let flipH = Boolean(local.flipH);
  let flipV = Boolean(local.flipV);

  const cx = container.offX + container.extX / 2;
  const cy = container.offY + container.extY / 2;

  if (container.flipH) { x = 2 * cx - x - w; flipH = !flipH; rot = -rot; }
  if (container.flipV) { y = 2 * cy - y - h; flipV = !flipV; rot = -rot; }
  if (container.rot) {
    const bx = x + w / 2;
    const by = y + h / 2;
    const rad = (container.rot * Math.PI) / 180;
    const dx = bx - cx;
    const dy = by - cy;
    const nbx = cx + dx * Math.cos(rad) - dy * Math.sin(rad);
    const nby = cy + dx * Math.sin(rad) + dy * Math.cos(rad);
    x = nbx - w / 2;
    y = nby - h / 2;
    rot += container.rot;
  }
  rot = ((rot % 360) + 360) % 360;
  return { offX: x, offY: y, extX: w, extY: h, rot, flipH, flipV };
}

/** The identity container: no group wraps the node, so composing against it changes nothing. */
const IDENTITY_CONTAINER = { offX: 0, offY: 0, extX: 1, extY: 1, chOffX: 0, chOffY: 0, chExtX: 1, chExtY: 1, rot: 0, flipH: false, flipV: false };

/** A local box (px, deg) composed against a container, or null if there is no local box to place. */
function placeInContainer(container, local) {
  if (!local) return null;
  const abs = composeGroupChild(container, { offX: local.x, offY: local.y, extX: local.w, extY: local.h, rot: local.rot, flipH: local.flipH, flipV: local.flipV });
  return { x: abs.offX, y: abs.offY, w: abs.extX, h: abs.extY, rot: abs.rot, flipH: abs.flipH, flipV: abs.flipV };
}

/**
 * `<a:tbl>` → columns and rows, each cell with its own box — x, y, w, h in
 * px relative to the slide, computed from the grid's widths and the rows'
 * heights against the frame's own geometry — so the window can put an
 * editor exactly over the cell that was double-clicked, the way it puts one
 * over a text box.
 */
function readTable(graphicFrame, theme, geom) {
  const tbl = first(graphicFrame, A('tbl'));
  if (!tbl) return null;
  const grid = kids(first(tbl, A('tblGrid')) || { children: [] }, A('gridCol')).map((g) => emuToPx(g.attrs.w));
  const totalW = grid.reduce((a, b) => a + b, 0) || geom?.w || 0;
  const scaleX = geom && totalW ? geom.w / totalW : 1;
  let top = geom?.y || 0;
  const rows = kids(tbl, A('tr')).map((tr) => {
    const h = emuToPx(tr.attrs.h);
    const tcs = kids(tr, A('tc'));
    let left = geom?.x || 0;
    const cells = tcs.map((tc, ci) => {
      const span = Number(tc.attrs.gridSpan || 1);
      const ownW = (grid[ci] || totalW / tcs.length) * scaleX;
      const w = grid.length ? grid.slice(ci, ci + span).reduce((a, b) => a + b, 0) * scaleX || ownW : ownW * span;
      const box = { x: left, y: top, w, h };
      left += ownW;
      return {
        text: readTextBody(kids(tc, A('txBody'))[0], theme),
        fill: readFill(kids(tc, A('tcPr'))[0], theme),
        colspan: span,
        rowspan: Number(tc.attrs.rowSpan || 1),
        merged: tc.attrs.hMerge === '1' || tc.attrs.vMerge === '1',
        box,
      };
    });
    top += h;
    return { height: h, cells };
  });
  return { columns: grid, rows };
}

/**
 * Read one slide part into a scene.
 * @param {string} xml the slide's XML
 * @param {object} ctx { theme, inherit(ph) -> shape|null, rel(id) -> {target,type} }
 */
export function readSlideScene(xml, ctx = {}) {
  const root = parse(xml);
  const sld = first(root, P('sld')) || first(root, P('sldLayout')) || first(root, P('sldMaster')) || root;
  const cSld = first(sld, P('cSld'));
  const spTree = first(cSld || sld, P('spTree'));
  const shapes = [];

  const bgFill = (() => {
    const bg = first(cSld || sld, P('bg'));
    if (!bg) return null;
    const bgPr = first(bg, P('bgPr'));
    if (bgPr) return readFill(bgPr, ctx.theme);
    const ref = first(bg, P('bgRef'));
    if (ref) {
      const c = colorChildOf(ref, ctx.theme);
      return c ? { type: 'solid', color: c.hex, alpha: c.alpha } : null;
    }
    return null;
  })();

  const walkTree = (tree, container = IDENTITY_CONTAINER, groupId = null) => {
    for (const node of kids(tree)) {
      if (node.name === P('sp')) shapes.push(readShape(node, ctx, container, groupId));
      else if (node.name === P('pic')) shapes.push(readPicture(node, ctx, container, groupId));
      else if (node.name === P('graphicFrame')) {
        const xfrm = kids(node, P('xfrm'))[0];
        const local = xfrm
          ? {
              x: emuToPx(kids(xfrm, A('off'))[0]?.attrs.x),
              y: emuToPx(kids(xfrm, A('off'))[0]?.attrs.y),
              w: emuToPx(kids(xfrm, A('ext'))[0]?.attrs.cx),
              h: emuToPx(kids(xfrm, A('ext'))[0]?.attrs.cy),
              rot: rotToDeg(xfrm.attrs.rot),
              flipH: false,
              flipV: false,
            }
          : null;
        const geom = placeInContainer(container, local);
        const table = readTable(node, ctx.theme, geom);
        const meta = nameOf(node);
        if (table) shapes.push({ kind: 'table', ...meta, groupId, geometry: geom, table });
        else {
          // A chart. The frame names its part through a relationship, and
          // the part is read the way a worksheet's is, into the spec
          // @rutba/drawing draws — a chart on a slide and a chart on a
          // sheet are one drawing. The deck hands in `rel` and `readPart`;
          // a layout or master read without them keeps the frame as an
          // unsupported graphic, which is what it was before.
          const chartNode = all(node, 'c:chart')[0];
          const rid = chartNode?.attrs['r:id'];
          const r = rid && ctx.rel ? ctx.rel(rid) : null;
          const xml = r?.part && ctx.readPart ? ctx.readPart(r.part) : null;
          const spec = xml
            ? parseChartXml(xml, { width: Math.max(160, Math.round(geom?.w || 480)), height: Math.max(120, Math.round(geom?.h || 300)) })
            : null;
          if (spec) shapes.push({ kind: 'chart', ...meta, groupId, geometry: geom, spec, part: r.part });
          else shapes.push({ kind: 'unsupported', ...meta, groupId, geometry: geom, what: chartNode ? 'chart' : 'graphic' });
        }
      } else if (node.name === P('cxnSp')) {
        const spPr = kids(node, P('spPr'))[0];
        const meta = nameOf(node);
        shapes.push({
          kind: 'connector',
          ...meta,
          groupId,
          geometry: placeInContainer(container, readXfrm(spPr)),
          line: readLine(spPr, ctx.theme),
          preset: first(spPr, A('prstGeom'))?.attrs.prst || 'line',
        });
      } else if (node.name === P('grpSp')) {
        // A group has its own coordinate space: its children's own off/ext
        // sit inside its chOff/chExt window, mapped onto the group's own
        // off/ext (a scale, whenever a resize has made ext and chExt
        // differ), then carried through the group's own flip and rotation —
        // the same composition a nested group's own box goes through before
        // its children do. The group itself becomes one entry in the scene,
        // so the stage can select, drag and delete it as a whole; each
        // member keeps its absolute, already-composed geometry so drawing
        // and hit-testing need nothing more, and carries `groupId` so the
        // window knows which group it belongs to.
        const grpSpPr = kids(node, P('grpSpPr'))[0];
        const xfrm = kids(grpSpPr || { children: [] }, A('xfrm'))[0];
        const localBox = readXfrm(grpSpPr) || { x: 0, y: 0, w: 0, h: 0, rot: 0, flipH: false, flipV: false };
        const own = placeInContainer(container, localBox) || { x: 0, y: 0, w: 0, h: 0, rot: 0, flipH: false, flipV: false };
        const meta = nameOf(node);
        shapes.push({ kind: 'group', ...meta, groupId, geometry: own });
        const chOff = xfrm && kids(xfrm, A('chOff'))[0];
        const chExt = xfrm && kids(xfrm, A('chExt'))[0];
        const nextContainer = {
          offX: own.x, offY: own.y, extX: own.w, extY: own.h,
          chOffX: emuToPx(chOff?.attrs.x), chOffY: emuToPx(chOff?.attrs.y),
          chExtX: chExt ? emuToPx(chExt.attrs.cx) : own.w || 1,
          chExtY: chExt ? emuToPx(chExt.attrs.cy) : own.h || 1,
          rot: own.rot, flipH: own.flipH, flipV: own.flipV,
        };
        walkTree(node, nextContainer, meta.id);
      }
    }
  };
  if (spTree) walkTree(spTree);

  return { background: bgFill, shapes: shapes.filter(Boolean) };
}

function readShape(sp, ctx, container, groupId) {
  const spPr = kids(sp, P('spPr'))[0];
  const style = kids(sp, P('style'))[0];
  const ph = placeholderOf(sp);
  const meta = nameOf(sp);
  let geometry = placeInContainer(container, readXfrm(spPr));
  let fill = readFill(spPr, ctx.theme);
  let line = readLine(spPr, ctx.theme);
  const text = readTextBody(kids(sp, P('txBody'))[0], ctx.theme);

  // A picture fill's source, resolved through the slide's relationships the
  // same way a `<p:pic>`'s is — the renderer draws it the same way too.
  if (fill?.type === 'picture' && fill.embed && ctx.rel) fill = { ...fill, source: ctx.rel(fill.embed) };

  // A style reference gives the shape its theme fill and line when spPr is bare.
  if (!fill && style) {
    const fillRef = kids(style, A('fillRef'))[0];
    const c = fillRef ? colorChildOf(fillRef, ctx.theme) : null;
    if (c) fill = { type: 'solid', color: c.hex, alpha: c.alpha };
  }
  if (!line && style) {
    const lnRef = kids(style, A('lnRef'))[0];
    const c = lnRef ? colorChildOf(lnRef, ctx.theme) : null;
    if (c) line = { width: 1, color: c.hex, alpha: c.alpha };
  }

  // Placeholders inherit everything they did not state.
  const inherited = ph && ctx.inherit ? ctx.inherit(ph) : null;
  if (inherited) {
    if (!geometry) geometry = inherited.geometry || null;
    if (!fill) fill = inherited.fill || null;
    if (!line) line = inherited.line || null;
  }

  return {
    kind: 'shape',
    ...meta,
    groupId,
    placeholder: ph,
    geometry,
    fill,
    line,
    effects: readEffects(spPr, ctx.theme),
    preset: first(spPr, A('prstGeom'))?.attrs.prst || (kids(spPr || { children: [] }, A('custGeom'))[0] ? 'custom' : 'rect'),
    adjustments: readAdjustments(spPr),
    // A custom geometry's outline, as SVG path data in the path's own units
    // — a banner with a diagonal cut, a slash, a swoosh. Null for a preset.
    path: readCustomPath(spPr),
    text,
    inheritedText: inherited?.text || null,
  };
}

/**
 * `<a:custGeom>` → `{ w, h, d }`: the path's declared box and its commands
 * as SVG path data in that box's units. moveTo, lnTo, cubicBezTo, quadBezTo,
 * arcTo and close are what decks use; arcTo is turned into an SVG arc from
 * its sweep. The renderer scales the box onto the shape's geometry.
 */
function readCustomPath(spPr) {
  const cust = spPr ? kids(spPr, A('custGeom'))[0] : null;
  const list = cust ? kids(cust, A('pathLst'))[0] : null;
  const paths = list ? kids(list, A('path')) : [];
  if (!paths.length) return null;
  const pt = (node) => {
    const p = kids(node, A('pt'))[0];
    return p ? [Number(p.attrs.x) || 0, Number(p.attrs.y) || 0] : null;
  };
  const pts = (node) => kids(node, A('pt')).map((p) => [Number(p.attrs.x) || 0, Number(p.attrs.y) || 0]);
  const n = (v) => (Math.round(v * 100) / 100).toString();
  let w = 0;
  let h = 0;
  const d = [];
  let filled = false;
  for (const path of paths) {
    w = Math.max(w, Number(path.attrs.w) || 0);
    h = Math.max(h, Number(path.attrs.h) || 0);
    if (path.attrs.fill !== 'none') filled = true;
    let cx = 0;
    let cy = 0;
    for (const cmd of path.children || []) {
      const name = String(cmd.name || '').replace(/^a:/, '');
      if (name === 'moveTo') { const p = pt(cmd); if (p) { d.push(`M${n(p[0])} ${n(p[1])}`); [cx, cy] = p; } }
      else if (name === 'lnTo') { const p = pt(cmd); if (p) { d.push(`L${n(p[0])} ${n(p[1])}`); [cx, cy] = p; } }
      else if (name === 'cubicBezTo') { const p = pts(cmd); if (p.length === 3) { d.push(`C${p.map((q) => `${n(q[0])} ${n(q[1])}`).join(' ')}`); [cx, cy] = p[2]; } }
      else if (name === 'quadBezTo') { const p = pts(cmd); if (p.length === 2) { d.push(`Q${p.map((q) => `${n(q[0])} ${n(q[1])}`).join(' ')}`); [cx, cy] = p[1]; } }
      else if (name === 'arcTo') {
        // Radii and angles in DrawingML's units: 60,000ths of a degree, from
        // the current point, clockwise-positive on a y-down page.
        const wR = Number(cmd.attrs.wR) || 0;
        const hR = Number(cmd.attrs.hR) || 0;
        const st = ((Number(cmd.attrs.stAng) || 0) / 60000) * (Math.PI / 180);
        const sw = ((Number(cmd.attrs.swAng) || 0) / 60000) * (Math.PI / 180);
        const ox = cx - wR * Math.cos(st);
        const oy = cy - hR * Math.sin(st);
        const ex = ox + wR * Math.cos(st + sw);
        const ey = oy + hR * Math.sin(st + sw);
        d.push(`A${n(wR)} ${n(hR)} 0 ${Math.abs(sw) > Math.PI ? 1 : 0} ${sw > 0 ? 1 : 0} ${n(ex)} ${n(ey)}`);
        cx = ex;
        cy = ey;
      }
      else if (name === 'close') d.push('Z');
    }
  }
  if (!d.length) return null;
  return { w: w || 1, h: h || 1, d: d.join(' '), filled };
}

function readAdjustments(spPr) {
  const prst = first(spPr, A('prstGeom'));
  if (!prst) return {};
  const out = {};
  for (const gd of all(prst, A('gd'))) out[gd.attrs.name] = gd.attrs.fmla;
  return out;
}

function readPicture(pic, ctx, container, groupId) {
  const spPr = kids(pic, P('spPr'))[0];
  const blipFill = kids(pic, P('blipFill'))[0];
  const blip = blipFill && first(blipFill, A('blip'));
  const srcRect = blipFill && first(blipFill, A('srcRect'));
  const meta = nameOf(pic);
  const embed = blip?.attrs['r:embed'] || blip?.attrs['r:link'] || null;
  return {
    kind: 'picture',
    ...meta,
    groupId,
    placeholder: placeholderOf(pic),
    geometry: placeInContainer(container, readXfrm(spPr)),
    line: readLine(spPr, ctx.theme),
    embed,
    source: embed && ctx.rel ? ctx.rel(embed) : null,
    crop: srcRect
      ? {
          l: pctOf(srcRect.attrs.l, 0),
          t: pctOf(srcRect.attrs.t, 0),
          r: pctOf(srcRect.attrs.r, 0),
          b: pctOf(srcRect.attrs.b, 0),
        }
      : null,
    preset: first(spPr, A('prstGeom'))?.attrs.prst || 'rect',
  };
}

/** Plain text of a scene — for search, outlines and accessibility. */
export function sceneText(scene) {
  const out = [];
  for (const s of scene.shapes) {
    const body = s.text || s.inheritedText;
    if (body) {
      for (const p of body.paragraphs) {
        const line = p.runs.map((r) => r.text).join('');
        if (line.trim()) out.push(line);
      }
    }
    if (s.table) {
      for (const row of s.table.rows) {
        out.push(row.cells.map((c) => (c.text?.paragraphs || []).map((p) => p.runs.map((r) => r.text).join('')).join(' ')).join('\t'));
      }
    }
  }
  return out.join('\n');
}

export { readXfrm, readFill, readLine, readTextBody, placeholderOf, composeGroupChild };
