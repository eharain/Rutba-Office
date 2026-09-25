// A slide, drawn.
//
// The scene renders to SVG, which is the one output that serves every consumer
// at once: the editor canvas, the slide-sorter thumbnails, the PDF export path
// and the presenter view. No DOM is needed to produce it, so a thumbnail can be
// generated in the main process while the window is still opening.
//
// Text layout is approximate by necessity — real line breaking needs real font
// metrics, and we deliberately ship no font files. `@rutba/drawing`'s measure
// is calibrated against the metrics of the fonts these documents actually use,
// and the editor overlays live, contenteditable text on top of the SVG when you
// click into a shape, so what you edit is always exact even where what you
// preview is close.

import { wrapText, measureText, lineHeight } from '@rutba/drawing/measure';
import { buildChart, renderSvg } from '@rutba/drawing';
import { escapeXml } from '@rutba/office-formats/xml';

const DEFAULT_FONT = 'Segoe UI, system-ui, -apple-system, Roboto, Helvetica, Arial, sans-serif';

function fillAttr(fill, fallback = 'none') {
  if (!fill) return fallback;
  if (fill.type === 'none') return 'none';
  if (fill.type === 'solid') return fill.color;
  if (fill.type === 'gradient') return `url(#${fill._id})`;
  if (fill.type === 'picture') return fill._id ? `url(#${fill._id})` : fallback;
  return fallback;
}

function gradientDef(fill, id) {
  const stops = (fill.stops || [])
    .map((s) => `<stop offset="${Math.round(s.offset * 10000) / 100}%" stop-color="${s.color}"${s.alpha < 1 ? ` stop-opacity="${s.alpha}"` : ''}/>`)
    .join('');
  const angle = ((fill.angle || 0) * Math.PI) / 180;
  const x2 = (Math.cos(angle) * 0.5 + 0.5).toFixed(4);
  const y2 = (Math.sin(angle) * 0.5 + 0.5).toFixed(4);
  const x1 = (0.5 - Math.cos(angle) * 0.5).toFixed(4);
  const y1 = (0.5 - Math.sin(angle) * 0.5).toFixed(4);
  return `<linearGradient id="${id}" x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}">${stops}</linearGradient>`;
}

/** The preset geometries that actually appear in decks, as SVG. */
function shapePath(preset, { x, y, w, h }, custom = null, { simplify = false } = {}) {
  const r = Math.min(w, h);
  // A custom geometry: its path, scaled from the path's own box onto the
  // shape's, in the coordinates themselves — no transform attribute, so the
  // caller's rotation still applies cleanly.
  if (preset === 'custom' && custom?.d) {
    // A thumbnail of a vector map with a hundred thousand points is a
    // hundred thousand points in a 220-pixel box: half a second to draw and
    // invisible when drawn. The sidebar leaves such a shape out.
    if (simplify && custom.d.length > 20000) return null;

    const sx = w / (custom.w || 1);
    const sy = h / (custom.h || 1);
    const d = custom.d.replace(/([MLCQAZ])([^MLCQAZ]*)/g, (_, op, args) => {
      if (op === 'Z') return 'Z';
      const nums = args.trim().split(/[\s,]+/).filter((v) => v !== '').map(Number);
      if (op === 'A') {
        // rx ry rotation large sweep x y — radii scale by axis, flags stay.
        const out = [];
        for (let i = 0; i + 6 < nums.length + 1; i += 7) {
          out.push([nums[i] * sx, nums[i + 1] * sy, nums[i + 2], nums[i + 3], nums[i + 4], x + nums[i + 5] * sx, y + nums[i + 6] * sy].map((v) => (Math.round(v * 100) / 100)).join(' '));
        }
        return 'A' + out.join(' ');
      }
      const out = [];
      for (let i = 0; i + 1 < nums.length; i += 2) out.push(`${Math.round((x + nums[i] * sx) * 100) / 100} ${Math.round((y + nums[i + 1] * sy) * 100) / 100}`);
      return op + out.join(' ');
    });
    return `<path d="${d}"`;
  }
  switch (preset) {
    case 'ellipse':
    case 'circle':
      return `<ellipse cx="${x + w / 2}" cy="${y + h / 2}" rx="${w / 2}" ry="${h / 2}"`;
    case 'roundRect':
      return `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="${r * 0.16}" ry="${r * 0.16}"`;
    case 'triangle':
      return `<polygon points="${x + w / 2},${y} ${x + w},${y + h} ${x},${y + h}"`;
    case 'rtTriangle':
      return `<polygon points="${x},${y} ${x},${y + h} ${x + w},${y + h}"`;
    case 'diamond':
      return `<polygon points="${x + w / 2},${y} ${x + w},${y + h / 2} ${x + w / 2},${y + h} ${x},${y + h / 2}"`;
    case 'parallelogram':
      return `<polygon points="${x + w * 0.25},${y} ${x + w},${y} ${x + w * 0.75},${y + h} ${x},${y + h}"`;
    case 'trapezoid':
      return `<polygon points="${x + w * 0.25},${y} ${x + w * 0.75},${y} ${x + w},${y + h} ${x},${y + h}"`;
    case 'pentagon':
    case 'hexagon':
    case 'octagon': {
      const sides = { pentagon: 5, hexagon: 6, octagon: 8 }[preset];
      const pts = [];
      for (let i = 0; i < sides; i++) {
        const a = (Math.PI * 2 * i) / sides - Math.PI / 2;
        pts.push(`${x + w / 2 + (Math.cos(a) * w) / 2},${y + h / 2 + (Math.sin(a) * h) / 2}`);
      }
      return `<polygon points="${pts.join(' ')}"`;
    }
    case 'star5': {
      const pts = [];
      for (let i = 0; i < 10; i++) {
        const rad = i % 2 ? 0.4 : 1;
        const a = (Math.PI * i) / 5 - Math.PI / 2;
        pts.push(`${x + w / 2 + (Math.cos(a) * w * rad) / 2},${y + h / 2 + (Math.sin(a) * h * rad) / 2}`);
      }
      return `<polygon points="${pts.join(' ')}"`;
    }
    case 'line':
    case 'straightConnector1':
      return `<line x1="${x}" y1="${y}" x2="${x + w}" y2="${y + h}"`;
    case 'rightArrow':
      return `<polygon points="${x},${y + h * 0.3} ${x + w * 0.6},${y + h * 0.3} ${x + w * 0.6},${y} ${x + w},${y + h / 2} ${x + w * 0.6},${y + h} ${x + w * 0.6},${y + h * 0.7} ${x},${y + h * 0.7}"`;
    case 'chevron':
      return `<polygon points="${x},${y} ${x + w * 0.75},${y} ${x + w},${y + h / 2} ${x + w * 0.75},${y + h} ${x},${y + h} ${x + w * 0.25},${y + h / 2}"`;
    default:
      return `<rect x="${x}" y="${y}" width="${w}" height="${h}"`;
  }
}

const BULLET_CHARS = ['•', '–', '▪', '‣', '·'];

/** Faces with no lowercase; their text is set in capitals whatever the fallback has. */
const CAPS_ONLY_FONTS = new Set(['bebas neue', 'bebas', 'bebas kai', 'bebas neue pro', 'league gothic', 'anton', 'six caps', 'big shoulders display']);
/** What stands in for a condensed display face the machine does not have. */
const CONDENSED_FALLBACK = "'Bahnschrift SemiCondensed', 'Arial Narrow', Impact";

/** Faces with serifs: a machine without one falls back to another serif, not to the sans default. */
const SERIF_FACES = /georgia|times|palatino|cambria|constantia|garamond|book antiqua|baskerville|bodoni|didot|century schoolbook|rockwell|serif$/i;
/** The family list for a named face: the face, then others of its kind, then the suite's default. */
function fontStack(font) {
  const name = escapeXml(font);
  if (/^(segoe ui|system-ui)$/i.test(font)) return DEFAULT_FONT;
  if (CAPS_ONLY_FONTS.has(String(font).toLowerCase())) return `${name}, ${CONDENSED_FALLBACK}, ${DEFAULT_FONT}`;
  if (SERIF_FACES.test(font)) return `${name}, Georgia, 'Times New Roman', 'Liberation Serif', 'DejaVu Serif', serif`;
  return `${name}, ${DEFAULT_FONT}`;
}

/** Only the keys whose value is not undefined — so a run's own `bold: undefined` never hides the style's. */
function definedOnly(o) {
  const out = {};
  for (const [k, v] of Object.entries(o || {})) if (v !== undefined && v !== null) out[k] = v;
  return out;
}

/** Letter spacing in the file's hundredths of a point, as the pixels SVG wants. */
const spacingPx = (segment, scale) => (segment.spacing ? segment.spacing * (96 / 72) * scale : 0);


/** What a placeholder's text defaults to, when nothing states a size. */
function defaultSizeFor(placeholder) {
  const type = placeholder?.type;
  if (type === 'ctrTitle') return 40;
  if (type === 'title') return 32;
  if (type === 'subTitle') return 20;
  if (type === 'ftr' || type === 'sldNum' || type === 'dt') return 11;
  return 18;
}

/** Lay a text body out into positioned lines. Shared by SVG and the editor. */
export function layoutText(body, box, { scale = 1, baseSize = 18, levels = null } = {}) {
  if (!body || !box) return { lines: [], height: 0 };
  const insets = body.insets || { l: 7.2, t: 3.6, r: 7.2, b: 3.6 };
  const width = Math.max(8, box.w - insets.l - insets.r);
  const lines = [];
  let y = 0;

  for (const p of body.paragraphs || []) {
    const level = p.level || 0;
    // The level's inherited style — the master's text styles, the layout's
    // and the shape's list styles — under whatever the paragraph and its
    // runs state for themselves.
    const lv = levels?.[Math.min(8, level)] || {};
    const marL = p.indent ?? lv.indent;
    const hanging = p.hanging ?? lv.hanging;
    const indent = (marL ?? level * 24) * scale;
    const runLook = definedOnly({ size: lv.size, bold: lv.bold, italic: lv.italic, underline: lv.underline, color: lv.color, font: lv.font, caps: lv.caps, spacing: lv.spacing, colorAlpha: lv.colorAlpha });
    // A link without a colour of its own is drawn in the theme's hyperlink
    // colour, not the text colour the style would otherwise give it.
    const linkLook = { ...runLook, color: lv.linkColor || undefined };
    const ownRuns = lv && Object.keys(runLook).length ? p.runs.map((r) => (r.break ? r : { ...definedOnly(r.link ? linkLook : runLook), ...definedOnly(r) })) : p.runs;
    // A run that states no size inherits it — from the shape, then from the
    // placeholder it fills. A title that falls back to body size is the single
    // most obvious way a rendered deck looks wrong.
    const stated = ownRuns.find((r) => r.size)?.size || p.endProps?.size || lv.size;
    const size = (stated || Math.max(9, baseSize - level * 2)) * scale;
    const lineHeightPt = p.lineHeightPt ?? lv.lineHeightPt;
    const lh = lineHeightPt ? lineHeightPt * scale : lineHeight(size) * (p.lineHeight ?? lv.lineHeight ?? 1);
    const spaceBefore = p.spaceBefore ?? lv.spaceBefore;
    const spaceAfter = p.spaceAfter ?? lv.spaceAfter;
    if (spaceBefore) y += spaceBefore * scale;

    const bulletSpec = p.bullet || lv.bullet || null;
    const bullet =
      bulletSpec && bulletSpec.type !== 'none'
        ? bulletSpec.type === 'number'
          ? `${(bulletSpec.start || 1) + (lines.filter((l) => l.numbered && l.level === level).length)}.`
          : bulletSpec.char || BULLET_CHARS[Math.min(level, BULLET_CHARS.length - 1)]
        : null;
    // A hanging indent puts the bullet at the margin plus the (negative)
    // indent and every line of the words at the margin, as PowerPoint sets
    // a list; a bullet that would hang off the box's left edge, or a
    // paragraph with no hanging indent, keeps the bullet at the margin and
    // the words a bullet's width in.
    const hangs = bullet && hanging != null && hanging < 0 && indent + hanging * scale >= 0 && -hanging * scale >= size * 0.6;
    const bulletX = bullet ? (hangs ? indent + hanging * scale : indent) : null;
    const textX = bullet ? (hangs ? indent : indent + size * 0.9) : indent;

    // A face that has no lowercase — Bebas Neue is the one on every second
    // deck — shows its text in capitals on a machine that has it, and in the
    // fallback's lowercase on one that does not. The capitals are the design.
    const runs = ownRuns.map((r) => (r.text && CAPS_ONLY_FONTS.has(String(r.font || '').toLowerCase()) ? { ...r, text: r.text.toUpperCase() } : r));
    const text = runs.map((r) => r.text).join('');
    if (!text.trim()) {
      y += lh;
      continue;
    }


    // Wrapping is done on the paragraph's plain text, then runs are mapped back
    // onto the wrapped lines so formatting survives the break.
    const avail = width - (bullet ? textX : indent);
    const wrapped = wrapText(text, Math.max(20, avail), { size });
    let consumed = 0;
    wrapped.forEach((lineText, li) => {
      // The wrapped line is the source's words joined by single spaces. The
      // source may hold several ("7LP  UK", a tabbed address line) — so the
      // line is matched back onto the source word by word, and the segment
      // is the source span, not the first N characters of it. Counting
      // characters lost the last one of every line that had a double space.
      let pos = consumed;
      while (pos < text.length && /\s/.test(text[pos])) pos += 1;
      const start = pos;
      for (const word of lineText.split(' ')) {
        if (!word) continue;
        const at = text.indexOf(word, pos);
        if (at < 0) break;
        pos = at + word.length;
      }
      const end = Math.max(start, pos);

      const segments = [];
      let cursor = 0;
      for (const run of runs) {
        const runStart = cursor;
        const runEnd = cursor + (run.text?.length || 0);
        cursor = runEnd;
        if (runEnd <= start || runStart >= end) continue;
        const slice = (run.text || '').slice(Math.max(0, start - runStart), Math.min(runEnd, end) - runStart);
        if (slice) segments.push({ ...run, text: slice });
      }
      consumed = end;
      lines.push({
        y: y + size,
        x: hangs ? textX : indent + (bullet && li === 0 ? size * 0.9 : 0),
        bulletX: li === 0 ? bulletX : null,
        size,
        align: p.align || lv.align || 'left',
        level,
        bullet: li === 0 ? bullet : null,
        bulletColor: bulletSpec?.color,
        numbered: p.bullet?.type === 'number',
        segments: segments.length ? segments : [{ text: lineText }],
      });
      y += lh;
    });
    if (spaceAfter) y += spaceAfter * scale;
  }
  return { lines, height: y, insets };
}

function textSvg(body, box, opts) {
  // Words running up or down: laid out in the box turned on its side, then
  // the drawing turned back — PowerPoint's vert270 (up) and vert (down).
  const vert = body.vert === 'vert' || body.vert === 'vert270' ? body.vert : null;
  if (vert) {
    const cx = box.x + box.w / 2;
    const cy = box.y + box.h / 2;
    const turned = { x: cx - box.h / 2, y: cy - box.w / 2, w: box.h, h: box.w };
    const inner = textSvg({ ...body, vert: 'horz' }, turned, opts);
    return inner ? `<g transform="rotate(${vert === 'vert270' ? -90 : 90} ${cx.toFixed(2)} ${cy.toFixed(2)})">${inner}</g>` : '';
  }
  // Columns: the box split across, the words filling one column to its
  // foot before the next, as PowerPoint fills them.
  const columns = Math.max(1, Math.min(16, Number(body.columns) || 1));
  if (columns > 1) {
    const colW = box.w / columns;
    const laid = layoutText(body, { ...box, w: colW }, opts);
    if (!laid.lines.length) return '';
    const room = box.h - (laid.insets?.t || 0) - (laid.insets?.b || 0);
    let out = '';
    let i = 0;
    for (let c = 0; c < columns && i < laid.lines.length; c++) {
      const top = laid.lines[i].y - laid.lines[i].size;
      const slice = [];
      // The last column takes whatever is left, running on below the box as PowerPoint lets it.
      while (i < laid.lines.length && (!slice.length || c === columns - 1 || laid.lines[i].y - top <= room)) {
        slice.push({ ...laid.lines[i], y: laid.lines[i].y - top });
        i += 1;
      }
      out += drawLines({ ...body, anchor: 'top' }, { x: box.x + c * colW, y: box.y, w: colW, h: box.h }, opts, slice, slice[slice.length - 1].y, laid.insets);
    }
    return out;
  }
  const { lines, height, insets } = layoutText(body, box, opts);
  return drawLines(body, box, opts, lines, height, insets);
}

/** The laid-out lines of one text body, as SVG. */
function drawLines(body, box, opts, lines, height, insets) {
  if (!lines.length) return '';
  const anchor = body.anchor || 'top';
  const inner = box.h - (insets?.t || 0) - (insets?.b || 0);
  const dy = anchor === 'middle' ? Math.max(0, (inner - height) / 2) : anchor === 'bottom' ? Math.max(0, inner - height) : 0;
  const originX = box.x + (insets?.l || 0);
  const originY = box.y + (insets?.t || 0) + dy;

  const out = [];
  for (const line of lines) {
    const width = box.w - (insets?.l || 0) - (insets?.r || 0);
    const lineWidth = line.segments.reduce((n, s) => n + measureText(s.text, { size: s.size || line.size, weight: s.bold ? 'bold' : 'normal' }) + spacingPx(s, opts?.scale || 1) * s.text.length, 0);

    let x = originX + line.x;
    if (line.align === 'center') x = originX + Math.max(0, (width - lineWidth) / 2);
    else if (line.align === 'right') x = originX + Math.max(0, width - lineWidth);
    const y = originY + line.y;

    // A highlighted run: its colour behind the words, as PowerPoint draws it.
    let runX = x;
    for (const s of line.segments) {
      const w = measureText(s.text, { size: s.size || line.size, weight: s.bold ? 'bold' : 'normal' }) + spacingPx(s, opts?.scale || 1) * s.text.length;
      if (s.highlight && w > 0) {
        const h = (s.size || line.size) * 1.15;
        out.push(`<rect x="${runX.toFixed(2)}" y="${(y - h * 0.8).toFixed(2)}" width="${w.toFixed(2)}" height="${h.toFixed(2)}" fill="${s.highlight}"/>`);
      }
      runX += w;
    }
    if (line.bullet) {
      const bx = line.bulletX != null && line.align !== 'center' && line.align !== 'right' ? originX + line.bulletX : x - line.size * 0.9;
      out.push(
        `<text x="${bx.toFixed(2)}" y="${y.toFixed(2)}" font-size="${line.size.toFixed(2)}" ` +
        `fill="${line.bulletColor || line.segments[0]?.color || '#333'}" font-family="${DEFAULT_FONT}">${escapeXml(line.bullet)}</text>`
      );
    }
    const spans = line.segments
      .map((s) => {
        const attrs = [];
        if (s.bold) attrs.push('font-weight="700"');
        if (s.italic) attrs.push('font-style="italic"');
        // A link is underlined, in Office's link blue unless the run says otherwise.
        if (s.underline || s.strike || s.link) attrs.push(`text-decoration="${[s.underline || s.link ? 'underline' : '', s.strike ? 'line-through' : ''].filter(Boolean).join(' ')}"`);
        if (s.color) attrs.push(`fill="${s.color}"`);
        else if (s.link) attrs.push('fill="#0563C1"');
        if (s.size && s.size !== line.size) attrs.push(`font-size="${(s.size * (opts?.scale || 1)).toFixed(2)}"`);
        if (s.font) attrs.push(`font-family="${fontStack(s.font)}"`);
        if (s.colorAlpha != null && s.colorAlpha < 1) attrs.push(`fill-opacity="${s.colorAlpha}"`);
        if (s.spacing) attrs.push(`letter-spacing="${spacingPx(s, opts?.scale || 1).toFixed(2)}"`);
        return `<tspan ${attrs.join(' ')}>${escapeXml(s.text)}</tspan>`;

      })
      .join('');
    out.push(
      `<text x="${x.toFixed(2)}" y="${y.toFixed(2)}" font-size="${line.size.toFixed(2)}" ` +
      `font-family="${DEFAULT_FONT}" fill="${line.segments[0]?.color || '#1a1a1a'}" xml:space="preserve">${spans}</text>`
    );
  }
  return out.join('');
}

function tableSvg(shape, opts) {
  const { geometry: g, table } = shape;
  if (!g || !table) return '';
  const out = [];
  const totalW = table.columns.reduce((a, b) => a + b, 0) || g.w;
  const scaleX = g.w / totalW;
  let y = g.y;
  for (const row of table.rows) {
    let x = g.x;
    const h = row.height || g.h / table.rows.length;
    row.cells.forEach((cell, ci) => {
      const w = (table.columns[ci] || totalW / row.cells.length) * scaleX * (cell.colspan || 1);
      if (!cell.merged) {
        out.push(
          `<rect x="${x.toFixed(2)}" y="${y.toFixed(2)}" width="${w.toFixed(2)}" height="${h.toFixed(2)}" ` +
          `fill="${fillAttr(cell.fill, '#ffffff')}" stroke="#c9ccd1" stroke-width="1"/>`
        );
        if (cell.text) out.push(textSvg(cell.text, { x: x + 4, y: y + 2, w: w - 8, h: h - 4 }, opts));
      }
      x += w;
    });
    y += h;
  }
  return out.join('');
}

/**
 * Render a slide scene to SVG.
 * @param {object} slide from Deck#slide()
 * @param {{ width?: number, resolveImage?: (shape) => string|null, standalone?: boolean, selection?: string }} [opts]
 */
export function renderSlide(slide, opts = {}) {
  const { width: outWidth, resolveImage, standalone = true, simplify = false, tagShapes = false, idPrefix = '', placeholderFrames = false } = opts;
  // Ids for gradients, patterns and filters: prefixed when several drawings
  // share one page (a strip of thumbnails beside the stage, a printed deck),
  // since url(#g1) finds the first g1 in the document, not this drawing's.
  const ID = String(idPrefix || '').replace(/[^A-Za-z0-9_-]/g, '');

  const W = slide.size?.width || 960;
  const H = slide.size?.height || 540;
  const scale = outWidth ? outWidth / W : 1;

  const defs = [];
  let gradSeq = 0;
  let picSeq = 0;
  let effectSeq = 0;
  let reflSeq = 0;
  let shapeSeq = 0;
  // The shadow, the glow and the soft edge are one SVG filter — a shape can
  // carry all three at once, and an element takes only one `filter`. Order
  // inside it is layering, not the schema's: the glow sits furthest back,
  // the shadow in front of that, and the (softened, if asked) shape itself
  // on top — which is how PowerPoint draws the combination too.
  const registerEffects = (effects) => {
    if (!effects) return '';
    const { glow, shadow: sh, softEdge } = effects;
    if (!glow && !sh && !softEdge) return '';
    const id = `${ID}fx${++effectSeq}`;
    const parts = [];
    let top = 'SourceGraphic';
    if (softEdge) {
      // A Gaussian blur of a flat fill only visibly softens where it meets
      // transparency — its edge — which is exactly what soft edges are.
      parts.push(`<feGaussianBlur in="SourceGraphic" stdDeviation="${((softEdge.radiusPt ?? 2.5) * (96 / 72)).toFixed(2)}" result="softSrc"/>`);
      top = 'softSrc';
    }
    const layers = [];
    if (glow) {
      const std = Math.max(0, ((glow.radiusPt ?? 8) * (96 / 72)) / 2);
      parts.push(
        `<feGaussianBlur in="SourceAlpha" stdDeviation="${std.toFixed(2)}" result="glowBlur"/>` +
        `<feFlood flood-color="${glow.color || '#5B9BD5'}" flood-opacity="${(glow.alpha ?? 0.6).toFixed(2)}" result="glowFlood"/>` +
        `<feComposite in="glowFlood" in2="glowBlur" operator="in" result="glowShape"/>`
      );
      layers.push('glowShape');
    }
    if (sh) {
      const rad = ((sh.dir || 0) * Math.PI) / 180;
      const dx = (sh.distPx || 0) * Math.cos(rad);
      const dy = (sh.distPx || 0) * Math.sin(rad);
      const std = Math.max(0, (sh.blurPx || 0) / 2);
      parts.push(
        `<feOffset in="SourceAlpha" dx="${dx.toFixed(2)}" dy="${dy.toFixed(2)}" result="shOff"/>` +
        `<feGaussianBlur in="shOff" stdDeviation="${std.toFixed(2)}" result="shBlur"/>` +
        `<feFlood flood-color="${sh.color || '#000000'}" flood-opacity="${(sh.alpha ?? 1).toFixed(2)}" result="shFlood"/>` +
        `<feComposite in="shFlood" in2="shBlur" operator="in" result="shShape"/>`
      );
      layers.push('shShape');
    }
    layers.push(top);
    parts.push(`<feMerge>${layers.map((l) => `<feMergeNode in="${l}"/>`).join('')}</feMerge>`);
    defs.push(`<filter id="${id}" x="-60%" y="-60%" width="220%" height="220%">${parts.join('')}</filter>`);
    return ` filter="url(#${id})"`;
  };
  /**
   * Shape Effects → Reflection: a faded, mirrored copy of the shape's own
   * drawing (its outline, its fill, its words — everything the `<use>`
   * re-renders), flipped about its own bottom edge and masked with a
   * gradient that fades out sooner for "tight", later for "full" — the
   * same three gallery names `deck.js` writes.
   */
  const reflectionSvg = (shape, g, refId) => {
    const kind = shape.effects?.reflection;
    if (!kind || !refId) return '';
    const fadeAt = { tight: 55, half: 45, full: 90 }[kind] ?? 45;
    const gradId = `${ID}reflGrad${++reflSeq}`;
    const maskId = `${gradId}m`;
    defs.push(`<linearGradient id="${gradId}" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#fff" stop-opacity="0.55"/><stop offset="${fadeAt}%" stop-color="#fff" stop-opacity="0"/></linearGradient>`);
    defs.push(`<mask id="${maskId}" maskUnits="userSpaceOnUse" x="${(g.x - g.w).toFixed(2)}" y="${g.y.toFixed(2)}" width="${(g.w * 3).toFixed(2)}" height="${g.h.toFixed(2)}"><rect x="${(g.x - g.w).toFixed(2)}" y="${g.y.toFixed(2)}" width="${(g.w * 3).toFixed(2)}" height="${g.h.toFixed(2)}" fill="url(#${gradId})"/></mask>`);
    const bottom = g.y + g.h;
    return `<use href="#${refId}" transform="translate(0 ${(2 * bottom).toFixed(2)}) scale(1 -1)" mask="url(#${maskId})"/>`;
  };
  const registerFill = (fill, href) => {
    if (fill?.type === 'gradient') {
      fill._id = `${ID}g${++gradSeq}`;
      defs.push(gradientDef(fill, fill._id));
    } else if (fill?.type === 'picture' && href) {
      fill._id = `${ID}pic${++picSeq}`;
      if (fill.tile) {
        // The picture's own pixel size is not decoded here, so the tile is
        // an approximate, fixed size rather than the picture's true one.
        const tile = 96;
        defs.push(`<pattern id="${fill._id}" patternUnits="userSpaceOnUse" width="${tile}" height="${tile}"><image href="${href}" x="0" y="0" width="${tile}" height="${tile}" preserveAspectRatio="xMidYMid slice"/></pattern>`);
      } else {
        defs.push(`<pattern id="${fill._id}" patternUnits="objectBoundingBox" patternContentUnits="objectBoundingBox" width="1" height="1"><image href="${href}" x="0" y="0" width="1" height="1" preserveAspectRatio="none"/></pattern>`);
      }
    }
    return fill;
  };

  const body = [];
  const bg = registerFill(slide.background);
  body.push(`<rect x="0" y="0" width="${W}" height="${H}" fill="${fillAttr(bg, '#ffffff')}"/>`);

  // tagShapes: each shape's drawing wrapped in a <g data-shape="id">, with
  // the groups it sits in (outermost last) in data-groups, so the show can
  // reveal, hide and move one shape — or a whole group — on its own. Done
  // after the fact: whatever the loop below pushed for a shape is gathered
  // into its wrapper when the next shape starts.
  const parentOf = new Map((slide.shapes || []).map((s) => [String(s.id), s.groupId == null ? null : String(s.groupId)]));
  let pending = null;
  const flush = () => {
    if (!tagShapes || !pending) return;
    const parts = body.splice(pending.mark);
    if (parts.length) {
      const chain = [];
      for (let g = parentOf.get(pending.id); g != null && chain.length < 16; g = parentOf.get(g)) chain.push(g);
      body.push(`<g data-shape="${escapeXml(pending.id)}"${chain.length ? ` data-groups="${chain.map(escapeXml).join(' ')}"` : ''}>${parts.join('')}</g>`);
    }
    pending = null;
  };

  // The master's and the layout's own shapes first, under the slide's —
  // never tagged: the show animates the slide's shapes, not the design's.
  for (const shape of [...(slide.underlay || []), ...(slide.shapes || [])]) {
    flush();
    pending = shape.underlay ? null : { id: String(shape.id), mark: body.length };
    if (shape.hidden) continue;
    // A group is not drawn itself — it is only a handle on its members,
    // which are already in this same list with their own, already-composed
    // geometry (see slide.js's `composeGroupChild`). Drawing it too would
    // paint an extra, invisible rectangle for nothing.
    if (shape.kind === 'group') continue;
    const g = shape.geometry;
    if (!g) continue;
    // PowerPoint flips the geometry, not the words in it: a flipped shape's
    // outline and fill mirror, but its text is set the way it reads,
    // rotated the same as an unflipped shape's would be. So the shape gets
    // the full transform (flip, then rotation, both about its own centre)
    // and its text gets only the rotation part of the same transform.
    const cx = g.x + g.w / 2;
    const cy = g.y + g.h / 2;
    const rotatePart = g.rot ? `rotate(${g.rot.toFixed(2)} ${cx.toFixed(2)} ${cy.toFixed(2)}) ` : '';
    const flipPart = g.flipH || g.flipV
      ? `translate(${cx.toFixed(2)} ${cy.toFixed(2)}) scale(${g.flipH ? -1 : 1} ${g.flipV ? -1 : 1}) translate(${(-cx).toFixed(2)} ${(-cy).toFixed(2)})`
      : '';
    const transform = rotatePart || flipPart ? ` transform="${rotatePart}${flipPart}"` : '';
    const textTransform = g.rot ? ` transform="rotate(${g.rot.toFixed(2)} ${cx.toFixed(2)} ${cy.toFixed(2)})"` : '';

    if (shape.kind === 'picture') {
      const href = resolveImage ? resolveImage(shape) : null;
      if (href) {
        const clip = shape.crop
          ? ` preserveAspectRatio="none"`
          : ' preserveAspectRatio="xMidYMid slice"';
        body.push(
          `<image x="${g.x.toFixed(2)}" y="${g.y.toFixed(2)}" width="${g.w.toFixed(2)}" height="${g.h.toFixed(2)}" ` +
          `href="${href}"${clip}${transform}/>`
        );
      } else {
        body.push(
          `<g${transform}><rect x="${g.x.toFixed(2)}" y="${g.y.toFixed(2)}" width="${g.w.toFixed(2)}" height="${g.h.toFixed(2)}" ` +
          `fill="#eceef2" stroke="#c9ccd1" stroke-dasharray="4 3"/>` +
          `<text x="${(g.x + g.w / 2).toFixed(2)}" y="${(g.y + g.h / 2).toFixed(2)}" text-anchor="middle" ` +
          `font-size="12" font-family="${DEFAULT_FONT}" fill="#7a8090">Picture</text></g>`
        );
      }
      continue;
    }

    if (shape.kind === 'chart' && shape.spec) {
      // Drawn by the chart kit the worksheet uses, on a transparent ground,
      // and placed as a nested svg at the frame. A spec the kit refuses
      // leaves the frame empty rather than the slide unrendered.
      let inner = '';
      try {
        const chartScene = buildChart({ ...shape.spec, width: Math.max(160, Math.round(g.w)), height: Math.max(120, Math.round(g.h)), mode: 'light' });
        chartScene.background = 'none';
        inner = renderSvg(chartScene, { standalone: false });
      } catch {
        inner = '';
      }
      if (inner) {
        const rotate = g.rot ? ` rotate(${g.rot.toFixed(2)} ${(g.w / 2).toFixed(2)} ${(g.h / 2).toFixed(2)})` : '';
        body.push(`<g transform="translate(${g.x.toFixed(2)} ${g.y.toFixed(2)})${rotate}">${inner}</g>`);
        continue;
      }
    }

    if (shape.kind === 'table') {
      body.push(tableSvg(shape, { scale: 1 }));
      continue;
    }

    const fill = registerFill(shape.fill, shape.fill?.type === 'picture' && resolveImage ? resolveImage(shape) : null);
    const line = shape.line;
    const geom = shapePath(shape.preset, g, shape.path || null, { simplify });
    if (!geom) continue;

    const strokeBits = line && line.type !== 'none' && line.color
      ? ` stroke="${line.color}" stroke-width="${(line.width || 1).toFixed(2)}"${line.dash ? ` stroke-dasharray="${line.dash === 'dash' ? '6 4' : '2 3'}"` : ''}`
      : '';
    const fillValue = fillAttr(fill, shape.kind === 'connector' ? 'none' : 'none');
    const opacity = fill?.alpha != null && fill.alpha < 1 ? ` fill-opacity="${fill.alpha}"` : '';
    const text = shape.text && shape.text.paragraphs?.length ? shape.text : null;
    const shapeMarkup = `${geom} fill="${fillValue}"${opacity}${strokeBits}${registerEffects(shape.effects)}${transform}/>`;
    const textMarkup = text ? `<g${textTransform}>${textSvg(text, g, { scale: 1, baseSize: defaultSizeFor(shape.placeholder), levels: shape.textStyle || null })}</g>` : '';

    // A reflection is a mirrored copy of the shape and its words, drawn from
    // a `<use>` on a group that wraps both — which is only worth the extra
    // wrapper element when there is one to draw.
    if (shape.effects?.reflection) {
      const wrapId = `${ID}shpref${++shapeSeq}`;
      body.push(`<g id="${wrapId}">${shapeMarkup}${textMarkup}</g>`);
      body.push(reflectionSvg(shape, g, wrapId));
    } else {
      body.push(shapeMarkup + textMarkup);
    }
    // Slide Master view: each placeholder's box, dashed, the way PowerPoint marks them.
    if (placeholderFrames && shape.placeholder) {
      body.push(`<rect x="${g.x.toFixed(2)}" y="${g.y.toFixed(2)}" width="${g.w.toFixed(2)}" height="${g.h.toFixed(2)}" fill="none" stroke="#8a94a6" stroke-width="1.2" stroke-dasharray="6 4"${transform}/>`);
    }
  }

  flush();
  const inner = `${defs.length ? `<defs>${defs.join('')}</defs>` : ''}${body.join('')}`;
  if (!standalone) return inner;
  const w = outWidth || W;
  const h = outWidth ? H * scale : H;
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" width="${w.toFixed(0)}" height="${h.toFixed(0)}" ` +
    `viewBox="0 0 ${W.toFixed(0)} ${H.toFixed(0)}" role="img">${inner}</svg>`
  );
}

/** A thumbnail: the same drawing, smaller, with a border the sorter can show. */
let thumbSeq = 0;
export function renderThumbnail(slide, width = 240, opts = {}) {
  return renderSlide(slide, { simplify: true, idPrefix: `t${(++thumbSeq).toString(36)}_`, ...opts, width });
}

