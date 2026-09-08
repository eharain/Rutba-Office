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
import { escapeXml } from '@rutba/office-formats/xml';

const DEFAULT_FONT = 'Segoe UI, system-ui, -apple-system, Roboto, Helvetica, Arial, sans-serif';

function fillAttr(fill, fallback = 'none') {
  if (!fill) return fallback;
  if (fill.type === 'none') return 'none';
  if (fill.type === 'solid') return fill.color;
  if (fill.type === 'gradient') return `url(#${fill._id})`;
  if (fill.type === 'picture') return fallback;
  return fallback;
}

function gradientDef(fill, id) {
  const stops = (fill.stops || [])
    .map((s) => `<stop offset="${Math.round(s.offset * 100)}%" stop-color="${s.color}"${s.alpha < 1 ? ` stop-opacity="${s.alpha}"` : ''}/>`)
    .join('');
  const angle = ((fill.angle || 0) * Math.PI) / 180;
  const x2 = (Math.cos(angle) * 0.5 + 0.5).toFixed(4);
  const y2 = (Math.sin(angle) * 0.5 + 0.5).toFixed(4);
  const x1 = (0.5 - Math.cos(angle) * 0.5).toFixed(4);
  const y1 = (0.5 - Math.sin(angle) * 0.5).toFixed(4);
  return `<linearGradient id="${id}" x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}">${stops}</linearGradient>`;
}

/** The preset geometries that actually appear in decks, as SVG. */
function shapePath(preset, { x, y, w, h }, custom = null) {
  const r = Math.min(w, h);
  // A custom geometry: its path, scaled from the path's own box onto the
  // shape's, in the coordinates themselves — no transform attribute, so the
  // caller's rotation still applies cleanly.
  if (preset === 'custom' && custom?.d) {
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
export function layoutText(body, box, { scale = 1, baseSize = 18 } = {}) {
  if (!body || !box) return { lines: [], height: 0 };
  const insets = body.insets || { l: 7.2, t: 3.6, r: 7.2, b: 3.6 };
  const width = Math.max(8, box.w - insets.l - insets.r);
  const lines = [];
  let y = 0;

  for (const p of body.paragraphs || []) {
    const level = p.level || 0;
    const indent = (p.indent ?? level * 24) * scale;
    // A run that states no size inherits it — from the shape, then from the
    // placeholder it fills. A title that falls back to body size is the single
    // most obvious way a rendered deck looks wrong.
    const stated = p.runs.find((r) => r.size)?.size;
    const size = (stated || Math.max(9, baseSize - level * 2)) * scale;
    const lh = p.lineHeightPt ? p.lineHeightPt * scale : lineHeight(size) * (p.lineHeight || 1);
    if (p.spaceBefore) y += p.spaceBefore * scale;

    const bullet =
      p.bullet && p.bullet.type !== 'none'
        ? p.bullet.type === 'number'
          ? `${(p.bullet.start || 1) + (lines.filter((l) => l.numbered && l.level === level).length)}.`
          : p.bullet.char || BULLET_CHARS[Math.min(level, BULLET_CHARS.length - 1)]
        : null;

    // A face that has no lowercase — Bebas Neue is the one on every second
    // deck — shows its text in capitals on a machine that has it, and in the
    // fallback's lowercase on one that does not. The capitals are the design.
    const runs = p.runs.map((r) => (r.text && CAPS_ONLY_FONTS.has(String(r.font || '').toLowerCase()) ? { ...r, text: r.text.toUpperCase() } : r));
    const text = runs.map((r) => r.text).join('');
    if (!text.trim()) {
      y += lh;
      continue;
    }


    // Wrapping is done on the paragraph's plain text, then runs are mapped back
    // onto the wrapped lines so formatting survives the break.
    const avail = width - indent - (bullet ? size * 0.9 : 0);
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
        x: indent + (bullet && li === 0 ? size * 0.9 : 0),
        size,
        align: p.align || 'left',
        level,
        bullet: li === 0 ? bullet : null,
        bulletColor: p.bullet?.color,
        numbered: p.bullet?.type === 'number',
        segments: segments.length ? segments : [{ text: lineText }],
      });
      y += lh;
    });
    if (p.spaceAfter) y += p.spaceAfter * scale;
  }
  return { lines, height: y, insets };
}

function textSvg(body, box, opts) {
  const { lines, height, insets } = layoutText(body, box, opts);
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

    if (line.bullet) {
      out.push(
        `<text x="${(x - line.size * 0.9).toFixed(2)}" y="${y.toFixed(2)}" font-size="${line.size.toFixed(2)}" ` +
        `fill="${line.bulletColor || line.segments[0]?.color || '#333'}" font-family="${DEFAULT_FONT}">${escapeXml(line.bullet)}</text>`
      );
    }
    const spans = line.segments
      .map((s) => {
        const attrs = [];
        if (s.bold) attrs.push('font-weight="700"');
        if (s.italic) attrs.push('font-style="italic"');
        if (s.underline) attrs.push('text-decoration="underline"');
        if (s.color) attrs.push(`fill="${s.color}"`);
        if (s.size && s.size !== line.size) attrs.push(`font-size="${(s.size * (opts?.scale || 1)).toFixed(2)}"`);
        if (s.font) {
          const condensed = CAPS_ONLY_FONTS.has(String(s.font).toLowerCase()) ? `${CONDENSED_FALLBACK}, ` : '';
          attrs.push(`font-family="${escapeXml(s.font)}, ${condensed}${DEFAULT_FONT}"`);
        }
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
  const { width: outWidth, resolveImage, standalone = true } = opts;
  const W = slide.size?.width || 960;
  const H = slide.size?.height || 540;
  const scale = outWidth ? outWidth / W : 1;

  const defs = [];
  let gradSeq = 0;
  const registerFill = (fill) => {
    if (fill?.type === 'gradient') {
      fill._id = `g${++gradSeq}`;
      defs.push(gradientDef(fill, fill._id));
    }
    return fill;
  };

  const body = [];
  const bg = registerFill(slide.background);
  body.push(`<rect x="0" y="0" width="${W}" height="${H}" fill="${fillAttr(bg, '#ffffff')}"/>`);

  for (const shape of slide.shapes || []) {
    if (shape.hidden) continue;
    const g = shape.geometry;
    if (!g) continue;
    const transform = g.rot ? ` transform="rotate(${g.rot.toFixed(2)} ${(g.x + g.w / 2).toFixed(2)} ${(g.y + g.h / 2).toFixed(2)})"` : '';

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

    if (shape.kind === 'table') {
      body.push(tableSvg(shape, { scale: 1 }));
      continue;
    }

    const fill = registerFill(shape.fill);
    const line = shape.line;
    const geom = shapePath(shape.preset, g, shape.path || null);
    const strokeBits = line && line.type !== 'none' && line.color
      ? ` stroke="${line.color}" stroke-width="${(line.width || 1).toFixed(2)}"${line.dash ? ` stroke-dasharray="${line.dash === 'dash' ? '6 4' : '2 3'}"` : ''}`
      : '';
    const fillValue = fillAttr(fill, shape.kind === 'connector' ? 'none' : 'none');
    const opacity = fill?.alpha != null && fill.alpha < 1 ? ` fill-opacity="${fill.alpha}"` : '';
    body.push(`${geom} fill="${fillValue}"${opacity}${strokeBits}${transform}/>`);

    const text = shape.text && shape.text.paragraphs?.length ? shape.text : null;
    if (text) body.push(`<g${transform}>${textSvg(text, g, { scale: 1, baseSize: defaultSizeFor(shape.placeholder) })}</g>`);
  }

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
export function renderThumbnail(slide, width = 240, opts = {}) {
  return renderSlide(slide, { ...opts, width });
}
