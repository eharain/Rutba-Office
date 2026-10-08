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
import { ommlToMathml, ommlToLinear, ommlInfo } from '@rutba/ooxml/math';

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
    // A path gradient runs out from its fillToRect's centre, the first stop there.
    const path = first(grad, A('path'));
    const to = path ? first(path, A('fillToRect')) : null;
    const center = to ? { x: (pctOf(to.attrs.l, 0) + 1 - pctOf(to.attrs.r, 0)) / 2, y: (pctOf(to.attrs.t, 0) + 1 - pctOf(to.attrs.b, 0)) / 2 } : null;
    return stops.length ? { type: 'gradient', stops, angle: rotToDeg(lin?.attrs.ang), ...(path ? { path: path.attrs.path || 'circle', center: center || { x: 0.5, y: 0.5 } } : {}) } : null;
  }
  const blip = kids(spPr, A('blipFill'))[0];
  if (blip) {
    const b = first(blip, A('blip'));
    return { type: 'picture', embed: b?.attrs['r:embed'] || null, tile: Boolean(kids(blip, A('tile'))[0]) };
  }
  // A pattern: its preset, its foreground (the colour a reader of one
  // colour takes for the shape's) and its background.
  const pattern = kids(spPr, A('pattFill'))[0];
  if (pattern) {
    const fg = colorChildOf(kids(pattern, A('fgClr'))[0], theme);
    const bg = colorChildOf(kids(pattern, A('bgClr'))[0], theme);
    return {
      type: 'pattern', preset: pattern.attrs.prst || 'pct5',
      color: fg?.hex || '#000000', alpha: fg?.alpha ?? 1,
      background: bg?.hex || '#FFFFFF', backgroundAlpha: bg?.alpha ?? 1,
    };
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

/**
 * Format Shape → 3-D Format and 3-D Rotation: the top bevel (preset, width
 * and height in points), the depth and its colour, and the camera preset —
 * null for a shape seen flat from the front.
 */
function read3d(spPr, theme) {
  const sp3d = spPr && kids(spPr, A('sp3d'))[0];
  const scene = spPr && kids(spPr, A('scene3d'))[0];
  if (!sp3d && !scene) return null;
  const bevel = sp3d && kids(sp3d, A('bevelT'))[0];
  const camera = scene && kids(scene, A('camera'))[0];
  const extrusion = sp3d && kids(sp3d, A('extrusionClr'))[0];
  const out = {
    bevel: bevel ? { prst: bevel.attrs.prst || 'circle', w: Number(bevel.attrs.w ?? 76200) / 12700, h: Number(bevel.attrs.h ?? 76200) / 12700 } : null,
    depth: sp3d?.attrs.extrusionH ? Number(sp3d.attrs.extrusionH) / 12700 : 0,
    depthColor: extrusion ? colorChildOf(extrusion, theme)?.hex ?? null : null,
    camera: camera?.attrs.prst || 'orthographicFront',
  };
  return out.bevel || out.depth > 0 || out.camera !== 'orthographicFront' ? out : null;
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
    // A pen's round end, or a highlighter's square one.
    cap: { rnd: 'round', sq: 'square' }[ln.attrs.cap] || null,
  };
}

/** A bare node holding one child, so a format-scheme entry reads the way an spPr does. */
const holder = (node) => ({ name: '#holder', attrs: {}, children: node ? [node] : [] });

/** The theme resolving `phClr` — a format-scheme entry's placeholder colour — to the reference's own. */
function withPlaceholderColour(theme, hex) {
  if (!theme || !hex) return theme;
  if (typeof theme.withPh === 'function') return theme.withPh(hex);
  const t = Object.create(theme);
  t.color = (name) => (name === 'phClr' ? hex : theme.color(name));
  return t;
}

/** `a:fillRef` → the theme's fill at that index (1–3 the fill styles, 1001– the backgrounds), or its colour when the theme has no list. */
function styleFill(fillRef, theme) {
  const idx = Number(fillRef.attrs.idx || 0);
  if (!idx) return { type: 'none' };
  const c = colorChildOf(fillRef, theme);
  const node = theme?.fillStyle?.(idx) || null;
  if (node) {
    const f = readFill(holder(node), withPlaceholderColour(theme, c?.hex || null));
    if (f) return f;
  }
  return c ? { type: 'solid', color: c.hex, alpha: c.alpha } : null;
}

/** `a:lnRef` → the theme's line at that index: its width, in the reference's colour. */
function styleLine(lnRef, theme) {
  const idx = Number(lnRef.attrs.idx || 0);
  if (!idx) return { type: 'none' };
  const c = colorChildOf(lnRef, theme);
  const node = theme?.lineStyle?.(idx) || null;
  if (node) {
    const l = readLine(holder(node), withPlaceholderColour(theme, c?.hex || null));
    if (l) return l.color || l.type === 'none' ? l : { ...l, color: c?.hex || null };
  }
  return c ? { width: 1, color: c.hex, alpha: c.alpha } : null;
}

/** `a:effectRef` → the theme's effect at that index (a shadow, a glow…), or none. */
function styleEffects(effectRef, theme) {
  const idx = Number(effectRef.attrs.idx || 0);
  if (!idx) return null;
  const node = theme?.effectStyle?.(idx) || null;
  if (!node) return null;
  const c = colorChildOf(effectRef, theme);
  const fx = readEffects(node, withPlaceholderColour(theme, c?.hex || null));
  return fx && (fx.shadow || fx.glow || fx.softEdge || fx.reflection) ? fx : null;
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
  // The language the run is proofed in, and whether it is proofed at all.
  if (rPr.attrs.lang) out.lang = rPr.attrs.lang;
  if (rPr.attrs.noProof === '1' || rPr.attrs.noProof === 'true') out.noProof = true;
  if (color) {
    out.color = color.hex;
    if (color.alpha < 1) out.colorAlpha = color.alpha;
  }
  const highlight = kids(rPr, A('highlight'))[0];
  if (highlight) out.highlight = colorChildOf(highlight, theme)?.hex || null;
  if (latin?.attrs.typeface) out.font = theme?.font(latin.attrs.typeface) || latin.attrs.typeface;
  if (hlink) out.link = hlink.attrs['r:id'] || true;
  // WordArt: the words' own outline, no fill (an outline alone), and a
  // shadow or glow round the words.
  const ln = kids(rPr, A('ln'))[0];
  if (ln) {
    const line = readLine(holder(ln), theme);
    if (line && line.type !== 'none' && line.color) out.outline = { width: line.width, color: line.color };
  }
  if (kids(rPr, A('noFill'))[0]) out.noFill = true;
  const fx = readEffects(rPr, theme);
  if (fx?.shadow || fx?.glow) out.textEffects = { ...(fx.shadow ? { shadow: fx.shadow } : {}), ...(fx.glow ? { glow: fx.glow } : {}) };
  return out;
}

function readParagraphProps(pPr, theme) {
  if (!pPr) return {};
  const out = {};
  if (pPr.attrs.lvl) out.level = Number(pPr.attrs.lvl);
  if (pPr.attrs.algn) out.align = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' }[pPr.attrs.algn] || 'left';
  // Right to left: an Arabic or Hebrew paragraph, its margin and indent
  // measured from the right; the alignment stays the side it names.
  if (pPr.attrs.rtl !== undefined) out.rtl = pPr.attrs.rtl === '1' || pPr.attrs.rtl === 'true';
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

/**
 * A list style — `<a:lstStyle>` on a text body, or a master's
 * `<p:titleStyle>`/`<p:bodyStyle>`/`<p:otherStyle>`, or the presentation's
 * `<p:defaultTextStyle>` — as nine levels of paragraph and run defaults:
 * alignment, indents and bullet from `a:lvlNpPr`, size, weight, colour and
 * face from its `a:defRPr`. `a:defPPr` fills every level. Only what the
 * style states is set, so styles merge by laying one over another.
 */
function readLevels(node, theme) {
  const out = Array.from({ length: 9 }, () => ({}));
  if (!node) return out;
  const one = (pPr) => {
    if (!pPr) return {};
    const { level: _level, ...para } = readParagraphProps(pPr, theme);
    const run = readRunProps(kids(pPr, A('defRPr'))[0], theme);
    const { link: _link, ...runProps } = run;
    return { ...para, ...runProps };
  };
  const every = one(kids(node, A('defPPr'))[0]);
  for (let i = 0; i < 9; i++) {
    const own = one(kids(node, A(`lvl${i + 1}pPr`))[0]);
    out[i] = mergeDefined(mergeDefined({}, every), own);
  }
  return out;
}

/** `into` with every key of `from` whose value is not undefined laid over it. */
function mergeDefined(into, from) {
  if (!from) return into;
  for (const [k, v] of Object.entries(from)) if (v !== undefined) into[k] = v;
  return into;
}

/** Nine levels each laid over the last list's, lowest first. */
function mergeLevels(...lists) {
  const out = Array.from({ length: 9 }, () => ({}));
  for (const list of lists) {
    if (!list) continue;
    for (let i = 0; i < 9; i++) mergeDefined(out[i], list[i]);
  }
  return out;
}

/** A parsed element back to XML, exactly enough to hand an equation's OMML to the math reader. */
function serializeNode(node) {
  if (typeof node === 'string') return node.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const attrs = Object.entries(node.attrs || {}).map(([k, v]) => ` ${k}="${String(v).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')}"`).join('');
  const inner = (node.children || []).map(serializeNode).join('');
  return inner ? `<${node.name}${attrs}>${inner}</${node.name}>` : `<${node.name}${attrs}/>`;
}

/** `<p:txBody>` → paragraphs of runs. */
function readTextBody(txBody, theme) {
  if (!txBody) return null;
  const bodyPr = kids(txBody, A('bodyPr'))[0];
  const lstStyle = kids(txBody, A('lstStyle'))[0];
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
      } else if (child.name === 'a14:m') {
        // An equation, as PowerPoint writes one in a text run: Office Math
        // (the same OMML Word writes) inside `a14:m`. Its XML is kept to be
        // written back and drawn as MathML; its linear form stands for its
        // words wherever plain text is wanted — search, the outline.
        const xml = kids(child).map(serializeNode).join('');
        let mathml = '';
        let linear = '';
        let info = { display: false, jc: null };
        try { mathml = ommlToMathml(xml); } catch { mathml = ''; }
        try { linear = ommlToLinear(xml); } catch { linear = textOf(child); }
        try { info = ommlInfo(xml); } catch { /* inline */ }
        runs.push({ text: linear || textOf(child), math: { xml, mathml, linear, display: Boolean(info.display), jc: info.jc || null } });
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
    // The body's own list style, when it states one — the lowest-but-one
    // layer of what a run without a size, colour or face falls back to.
    levels: lstStyle && kids(lstStyle).length ? readLevels(lstStyle, theme) : null,
    anchor: anchorMap[bodyPr?.attrs.anchor] || 'top',
    // Whether the body says where its words sit, or leaves it to the placeholder it fills.
    anchorStated: Boolean(bodyPr?.attrs.anchor),
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
  const out = { id: cNv?.attrs.id || null, name: cNv?.attrs.name || '', hidden: cNv?.attrs.hidden === '1' };
  // Insert → Action: what a click on the shape does in the show, as the
  // file says it; the deck resolves the relationship it may name.
  const click = cNv && kids(cNv, A('hlinkClick'))[0];
  if (click) out.click = { rId: click.attrs['r:id'] || null, action: click.attrs.action || null };
  return out;
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
/** The parts of a table a table style gives a look, in the order a later one lays its look over an earlier. */
const TABLE_PARTS = ['wholeTbl', 'band1V', 'band2V', 'band1H', 'band2H', 'firstCol', 'lastCol', 'firstRow', 'lastRow', 'nwCell', 'neCell', 'swCell', 'seCell'];
const EDGES = ['left', 'right', 'top', 'bottom', 'insideH', 'insideV'];

/**
 * A presentation's table styles (tableStyles.xml), by id: for each part of a
 * table — the whole, the bands, the first and last row and column, the
 * corners — the look its cells take: a fill, borders by edge, and the
 * words' bold, italic and colour. The style PowerPoint and the deck's own
 * tables use by default is known without being written out.
 */
export function readTableStyles(xml, theme) {
  const out = new Map();
  for (const source of [BUILT_IN_TABLE_STYLES, xml || '']) {
    for (const st of all(parse(source), A('tblStyle'))) out.set(st.attrs.styleId, readTableStyle(st, theme));
  }
  return out;
}

function readTableStyle(st, theme) {
  const parts = {};
  for (const name of TABLE_PARTS) {
    const node = kids(st, A(name))[0];
    if (!node) continue;
    const look = { edges: {} };
    const tx = kids(node, A('tcTxStyle'))[0];
    if (tx) {
      if (tx.attrs.b) look.bold = tx.attrs.b === 'on';
      if (tx.attrs.i) look.italic = tx.attrs.i === 'on';
      const c = colorChildOf(tx, theme);
      if (c) look.color = c.hex;
    }
    const tc = kids(node, A('tcStyle'))[0];
    const fill = tc && kids(tc, A('fill'))[0];
    if (fill) look.fill = readFill(fill, theme);
    const bdr = tc && kids(tc, A('tcBdr'))[0];
    for (const edge of EDGES) {
      const side = bdr && kids(bdr, A(edge))[0];
      if (side) look.edges[edge] = readLine(side, theme);
    }
    parts[name] = look;
  }
  return parts;
}

/**
 * A cell's look from its table's style: each part its place takes, the later
 * over the earlier — the whole table, the bands (not the first and last rows
 * and columns, when those are set apart), the first and last column and row,
 * a corner — each part's outer edges where the cell lies on them and its
 * inside ones elsewhere.
 */
function styledCell(style, flags, r, c, rows, cols) {
  const out = { edges: {} };
  if (!style) return out;
  const firstRow = flags.firstRow && r === 0;
  const lastRow = flags.lastRow && r === rows - 1;
  const firstCol = flags.firstCol && c === 0;
  const lastCol = flags.lastCol && c === cols - 1;
  const bandRow = r - (flags.firstRow ? 1 : 0);
  const bandCol = c - (flags.firstCol ? 1 : 0);
  // Each part with the rows and columns it covers.
  const whole = { top: 0, bottom: rows - 1, left: 0, right: cols - 1 };
  const applying = [['wholeTbl', whole]];
  if (flags.bandCol && !firstCol && !lastCol) applying.push([bandCol % 2 === 0 ? 'band1V' : 'band2V', { ...whole, left: c, right: c }]);
  if (flags.bandRow && !firstRow && !lastRow) applying.push([bandRow % 2 === 0 ? 'band1H' : 'band2H', { ...whole, top: r, bottom: r }]);
  if (firstCol) applying.push(['firstCol', { ...whole, left: 0, right: 0 }]);
  if (lastCol) applying.push(['lastCol', { ...whole, left: cols - 1, right: cols - 1 }]);
  if (firstRow) applying.push(['firstRow', { ...whole, top: 0, bottom: 0 }]);
  if (lastRow) applying.push(['lastRow', { ...whole, top: rows - 1, bottom: rows - 1 }]);
  if (firstRow && firstCol) applying.push(['nwCell', { top: 0, bottom: 0, left: 0, right: 0 }]);
  if (firstRow && lastCol) applying.push(['neCell', { top: 0, bottom: 0, left: c, right: c }]);
  if (lastRow && firstCol) applying.push(['swCell', { top: r, bottom: r, left: 0, right: 0 }]);
  if (lastRow && lastCol) applying.push(['seCell', { top: r, bottom: r, left: c, right: c }]);
  for (const [name, area] of applying) {
    const look = style[name];
    if (!look) continue;
    if (look.fill) out.fill = look.fill;
    if (look.bold != null) out.bold = look.bold;
    if (look.italic != null) out.italic = look.italic;
    if (look.color) out.color = look.color;
    const e = look.edges;
    const at = {
      left: c === area.left ? e.left : e.insideV,
      right: c === area.right ? e.right : e.insideV,
      top: r === area.top ? e.top : e.insideH,
      bottom: r === area.bottom ? e.bottom : e.insideH,
    };
    for (const [side, line] of Object.entries(at)) if (line) out.edges[side] = line;
  }
  return out;
}

/** A cell's own border on one side (tcPr's lnL, lnR, lnT or lnB), as a line. */
const cellEdge = (tcPr, name, theme) => {
  const ln = tcPr && kids(tcPr, A(name))[0];
  return ln ? readLine({ name: '#holder', attrs: {}, children: [{ ...ln, name: A('ln') }] }, theme) : null;
};

function readTable(graphicFrame, theme, geom, styles = null) {
  const tbl = first(graphicFrame, A('tbl'));
  if (!tbl) return null;
  // Which of its style's parts the table takes, and its style.
  const tblPr = kids(tbl, A('tblPr'))[0];
  const on = (k) => tblPr?.attrs[k] === '1' || tblPr?.attrs[k] === 'true';
  const flags = { firstRow: on('firstRow'), lastRow: on('lastRow'), firstCol: on('firstCol'), lastCol: on('lastCol'), bandRow: on('bandRow'), bandCol: on('bandCol') };
  const styleId = textOf(kids(tblPr, A('tableStyleId'))[0]).trim();
  const style = styleId ? styles?.get(styleId) ?? null : null;
  const rowCount = kids(tbl, A('tr')).length;
  const grid = kids(first(tbl, A('tblGrid')) || { children: [] }, A('gridCol')).map((g) => emuToPx(g.attrs.w));
  const totalW = grid.reduce((a, b) => a + b, 0) || geom?.w || 0;
  const scaleX = geom && totalW ? geom.w / totalW : 1;
  let top = geom?.y || 0;
  const colCount = grid.length || Math.max(1, ...kids(tbl, A('tr')).map((tr) => kids(tr, A('tc')).length));
  const rows = kids(tbl, A('tr')).map((tr, ri) => {
    const h = emuToPx(tr.attrs.h);
    const tcs = kids(tr, A('tc'));
    let left = geom?.x || 0;
    const cells = tcs.map((tc, ci) => {
      const span = Number(tc.attrs.gridSpan || 1);
      const rowspan = Number(tc.attrs.rowSpan || 1);
      const ownW = (grid[ci] || totalW / tcs.length) * scaleX;
      const w = grid.length ? grid.slice(ci, ci + span).reduce((a, b) => a + b, 0) * scaleX || ownW : ownW * span;
      const box = { x: left, y: top, w, h };
      left += ownW;
      // Its look: its own, over its table style's for where it is (a merged cell's far edges those of the last cell it covers).
      const tcPr = kids(tc, A('tcPr'))[0];
      const styled = styledCell(style, flags, ri, ci, rowCount, colCount);
      const far = styledCell(style, flags, ri + rowspan - 1, ci + span - 1, rowCount, colCount);
      const edges = {
        left: cellEdge(tcPr, 'lnL', theme) ?? styled.edges.left ?? null,
        right: cellEdge(tcPr, 'lnR', theme) ?? far.edges.right ?? null,
        top: cellEdge(tcPr, 'lnT', theme) ?? styled.edges.top ?? null,
        bottom: cellEdge(tcPr, 'lnB', theme) ?? far.edges.bottom ?? null,
      };
      const text = readTextBody(kids(tc, A('txBody'))[0], theme);
      // The style's bold, italic and colour for words that say none of their own.
      if (text && (styled.bold != null || styled.italic != null || styled.color)) {
        for (const p of text.paragraphs || []) {
          for (const run of p.runs || []) {
            if (run.bold == null && styled.bold != null) run.bold = styled.bold;
            if (run.italic == null && styled.italic != null) run.italic = styled.italic;
            if (run.color == null && styled.color) run.color = styled.color;
          }
        }
      }
      if (text && tcPr?.attrs.anchor) text.anchor = { t: 'top', ctr: 'middle', b: 'bottom' }[tcPr.attrs.anchor] || text.anchor;
      return {
        text,
        fill: readFill(tcPr, theme) ?? styled.fill ?? null,
        edges,
        styled: Boolean(style) || EDGE_NAMES.some((n) => kids(tcPr, A(n))[0]),
        colspan: span,
        rowspan,
        merged: tc.attrs.hMerge === '1' || tc.attrs.vMerge === '1',
        box,
      };
    });
    top += h;
    return { height: h, cells };
  });
  return { columns: grid, rows, styleId: styleId || null, flags };
}
const EDGE_NAMES = ['lnL', 'lnR', 'lnT', 'lnB'];

/**
 * The table style PowerPoint, and the deck's own tables, use where a file
 * names it without writing it out: Medium Style 2 in the first accent — a
 * fifth-tint body, two-fifths-tint bands, the first and last rows and columns
 * in the accent itself with bold light words, light borders between the
 * cells and thicker ones under the header and over the totals.
 */
const BUILT_IN_TABLE_STYLES = (() => {
  const ln = (w) => `<a:ln w="${w}" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln>`;
  const tint = (v) => `<a:fill><a:solidFill><a:schemeClr val="accent1">${v ? `<a:tint val="${v}"/>` : ''}</a:schemeClr></a:solidFill></a:fill>`;
  const strong = '<a:tcTxStyle b="on"><a:schemeClr val="lt1"/></a:tcTxStyle>';
  const edges = ['left', 'right', 'top', 'bottom', 'insideH', 'insideV'].map((e) => `<a:${e}>${ln(12700)}</a:${e}>`).join('');
  return '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    + '<a:tblStyle styleId="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" styleName="Medium Style 2 - Accent 1">'
    + `<a:wholeTbl><a:tcTxStyle><a:schemeClr val="dk1"/></a:tcTxStyle><a:tcStyle><a:tcBdr>${edges}</a:tcBdr>${tint(20000)}</a:tcStyle></a:wholeTbl>`
    + `<a:band1H><a:tcStyle><a:tcBdr/>${tint(40000)}</a:tcStyle></a:band1H><a:band1V><a:tcStyle><a:tcBdr/>${tint(40000)}</a:tcStyle></a:band1V>`
    + `<a:lastCol>${strong}<a:tcStyle><a:tcBdr/>${tint(0)}</a:tcStyle></a:lastCol><a:firstCol>${strong}<a:tcStyle><a:tcBdr/>${tint(0)}</a:tcStyle></a:firstCol>`
    + `<a:lastRow>${strong}<a:tcStyle><a:tcBdr><a:top>${ln(38100)}</a:top></a:tcBdr>${tint(0)}</a:tcStyle></a:lastRow>`
    + `<a:firstRow>${strong}<a:tcStyle><a:tcBdr><a:bottom>${ln(38100)}</a:bottom></a:tcBdr>${tint(0)}</a:tcStyle></a:firstRow>`
    + '</a:tblStyle></a:tblStyleLst>';
})();

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
      const f = styleFill(ref, ctx.theme);
      return f?.type === 'none' ? null : f;
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
        const table = readTable(node, ctx.theme, geom, ctx.tableStyles);
        const meta = nameOf(node);
        // Insert → Object: an embedded document, drawn as its picture (its
        // icon, or its first page), what it is and where it is kept beside.
        const ole = table ? null : first(node, P('oleObj'));
        const olePic = ole ? first(node, A('blip')) : null;
        if (table) shapes.push({ kind: 'table', ...meta, groupId, geometry: geom, table });
        else if (ole) {
          const embed = olePic?.attrs['r:embed'] || null;
          const rid = ole.attrs['r:id'] || null;
          shapes.push({
            kind: 'picture', ...meta, groupId, geometry: geom, line: null, preset: 'rect', crop: null,
            embed, source: embed && ctx.rel ? ctx.rel(embed) : null,
            object: { progId: ole.attrs.progId || null, name: ole.attrs.name || null, icon: ole.attrs.showAsIcon === '1', source: rid && ctx.rel ? ctx.rel(rid) : null },
          });
        } else {
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
      } else if (node.name === 'mc:AlternateContent') {
        // Markup compatibility: the first choice this reader understands —
        // an equation's text box (a14) — or else the fallback PowerPoint
        // wrote for readers like the ones that do not, a picture as a rule.
        const choice = kids(node, 'mc:Choice').find((c) => String(c.attrs.Requires || '').split(/\s+/).every((r) => r === 'a14'));
        const fallback = kids(node, 'mc:Fallback')[0] || null;
        const chosen = choice || fallback;
        if (!chosen) continue;
        const before = shapes.length;
        walkTree(chosen, container, groupId);
        const blip = choice && fallback ? first(fallback, A('blip')) : null;
        const embed = blip?.attrs['r:embed'] || null;
        for (let i = before; i < shapes.length; i++) {
          shapes[i] = { ...shapes[i], alt: Boolean(choice) };
          if (embed) shapes[i].fallback = { embed, source: ctx.rel ? ctx.rel(embed) : null };
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

  // A style reference gives the shape its theme fill, line and effect when
  // spPr states none: the index picks one of the theme's format scheme
  // entries (Design → Effects), drawn in the reference's own colour. Index
  // 0 is "none", as in PowerPoint.
  let effects = readEffects(spPr, ctx.theme);
  let styleText = null;
  if (style) {
    const ref = (name) => kids(style, A(name))[0] || null;
    const fillRef = ref('fillRef');
    const lnRef = ref('lnRef');
    const effectRef = ref('effectRef');
    const fontRef = ref('fontRef');
    if (!fill && fillRef) fill = styleFill(fillRef, ctx.theme);
    if (!line && lnRef) line = styleLine(lnRef, ctx.theme);
    if (!effects && effectRef) effects = styleEffects(effectRef, ctx.theme);
    if (fontRef) {
      const c = colorChildOf(fontRef, ctx.theme);
      const face = fontRef.attrs.idx === 'major' ? '+mj-lt' : fontRef.attrs.idx === 'minor' ? '+mn-lt' : null;
      styleText = { ...(c ? { color: c.hex } : {}), ...(face && ctx.theme?.font ? { font: ctx.theme.font(face) } : {}) };
    }
  }

  // Placeholders inherit everything they did not state.
  const inherited = ph && ctx.inherit ? ctx.inherit(ph) : null;
  if (inherited) {
    if (!geometry) geometry = inherited.geometry || null;
    if (!fill) fill = inherited.fill || null;
    if (!line) line = inherited.line || null;
    if (text && !text.anchorStated && inherited.anchor) text.anchor = inherited.anchor;
  }

  return {
    kind: 'shape',
    ...meta,
    groupId,
    placeholder: ph,
    geometry,
    fill,
    line,
    effects,
    shape3d: read3d(spPr, ctx.theme),
    styleText,
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
  // Each path by itself too, filled and stroked as it says — a smile's mouth is a line.
  const figures = [];
  for (const path of paths) {
    w = Math.max(w, Number(path.attrs.w) || 0);
    h = Math.max(h, Number(path.attrs.h) || 0);
    if (path.attrs.fill !== 'none') filled = true;
    const from = d.length;
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
    if (d.length > from) figures.push({ d: d.slice(from).join(' '), w: Number(path.attrs.w) || 0, h: Number(path.attrs.h) || 0, filled: path.attrs.fill !== 'none', stroked: path.attrs.stroke !== '0' && path.attrs.stroke !== 'false' });
  }
  if (!d.length) return null;
  return { w: w || 1, h: h || 1, d: d.join(' '), filled, ...(figures.length > 1 ? { figures } : {}) };
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
  // Insert → Video and Audio: a picture (the poster, or a speaker) whose
  // nvPr names the media — `a:videoFile`/`a:audioFile`, and PowerPoint
  // 2010's embedded copy in `p14:media`.
  const nvPr = first(kids(pic, P('nvPicPr'))[0] || pic, P('nvPr'));
  const file = nvPr && (first(nvPr, A('videoFile')) || first(nvPr, A('audioFile')));
  const p14 = nvPr && first(nvPr, 'p14:media');
  const mediaId = p14?.attrs['r:embed'] || file?.attrs['r:link'] || file?.attrs['r:embed'] || null;
  const media = file
    ? { kind: file.name === A('videoFile') ? 'video' : 'audio', id: mediaId, source: mediaId && ctx.rel ? ctx.rel(mediaId) : null }
    : null;
  return {
    kind: 'picture',
    ...meta,
    ...(media ? { media } : {}),
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

export { readXfrm, readFill, readLine, readTextBody, placeholderOf, composeGroupChild, readLevels, mergeLevels, mergeDefined };
