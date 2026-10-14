'use strict';

/**
 * Comments and markup on a PDF's pages: read whoever wrote them, and add,
 * reply to, resolve and delete them as an incremental update.
 *
 * What is written is what Acrobat writes, so any reader shows it: text
 * markup (Highlight, Underline, StrikeOut) over the words by their
 * QuadPoints, sticky notes (Text), text boxes (FreeText), freehand ink (Ink),
 * rectangles (Square), ellipses (Circle), lines and arrows (Line with its
 * line endings), stamps, and marks for redaction (Redact). Every annotation
 * carries an appearance stream drawn here, so a reader that draws only
 * appearances (most do) shows it as it looks in this window.
 *
 * Replies are Text annotations "in reply to" the one they answer (/IRT), as
 * Acrobat threads them; a comment marked resolved has a hidden reply with
 * the review state Completed (/StateModel /Review /State /Completed), the
 * state Acrobat's "Set Status" writes. Deleting a comment takes it out of
 * its page's list with its replies and its pop-up, and frees their objects
 * in the update.
 *
 * Coordinates are the page's default user space, in points, as everything
 * in a PDF is: the window turns its pixels into these through the page's
 * own viewport before anything reaches here.
 */

const { Name, Ref, Dict, PdfString, Stream } = require('./reader');
const { textString } = require('./assemble');
const { PdfUpdate, UpdateError, copyDict, flateStream } = require('./update');
const { encode, literal } = require('./encoding');
const { widthOfBytes } = require('./metrics');
const { decodePng } = require('./images');
const zlib = require('zlib');

/** The kinds of annotation a person makes or reads as a comment; links, form widgets and pop-ups are not. */
const COMMENT_TYPES = new Set([
  'Text', 'FreeText', 'Line', 'Square', 'Circle', 'Polygon', 'PolyLine', 'Highlight', 'Underline', 'Squiggly',
  'StrikeOut', 'Stamp', 'Caret', 'Ink', 'FileAttachment', 'Sound', 'Redact',
]);

/** Acrobat's words for each kind, written as the comment's subject. */
const SUBJECTS = {
  Highlight: 'Highlight', Underline: 'Underline', StrikeOut: 'Strikethrough', Squiggly: 'Squiggly',
  Text: 'Sticky Note', FreeText: 'Text Box', Ink: 'Pencil', Square: 'Rectangle', Circle: 'Oval',
  Line: 'Line', Arrow: 'Arrow', Stamp: 'Stamp', Redact: 'Redact',
};

const DEFAULT_COLOURS = {
  Highlight: [1, 0.92, 0.23], Underline: [0.13, 0.55, 0.13], StrikeOut: [0.86, 0.16, 0.16], Text: [1, 0.85, 0.25],
  FreeText: [0, 0, 0], Ink: [0.11, 0.35, 0.85], Square: [0.86, 0.16, 0.16], Circle: [0.86, 0.16, 0.16],
  Line: [0.86, 0.16, 0.16], Stamp: [0.05, 0.15, 0.45], Redact: [0.86, 0.1, 0.1],
};

// Flags: 2 hidden, 4 print, 8 no zoom, 16 no rotate.
const FLAG_HIDDEN = 2;
const FLAG_PRINT = 4;

const n = (v) => {
  if (!Number.isFinite(v)) return '0';
  const s = (Math.round(v * 1000) / 1000).toString();
  return s === '-0' ? '0' : s;
};
const num = (v, fallback = 0) => (typeof v === 'number' && Number.isFinite(v) ? v : fallback);

function pdfDate(d = new Date()) {
  const p = (x) => String(x).padStart(2, '0');
  return `D:${d.getUTCFullYear()}${p(d.getUTCMonth() + 1)}${p(d.getUTCDate())}${p(d.getUTCHours())}${p(d.getUTCMinutes())}${p(d.getUTCSeconds())}Z`;
}

/** A PDF date ("D:20261014093000+05'00'") as an ISO string, or null. */
function parseDate(text) {
  const m = /^(?:D:)?(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?([Zz+-])?(\d{2})?'?(\d{2})?/.exec(String(text || '').trim());
  if (!m) return null;
  const [, y, mo = '01', d = '01', h = '00', mi = '00', s = '00', sign, oh = '00', om = '00'] = m;
  let ms = Date.UTC(+y, +mo - 1, +d, +h, +mi, +s);
  if (sign === '+' || sign === '-') ms -= (sign === '+' ? 1 : -1) * ((+oh * 60 + +om) * 60000);
  const date = new Date(ms);
  return Number.isNaN(date.getTime()) ? null : date.toISOString();
}

const textOf = (file, v) => {
  const s = file.resolve(v);
  return s instanceof PdfString ? s.text().replace(/\r\n?/g, '\n') : '';
};
const numbersOf = (file, v) => {
  const arr = file.resolve(v);
  return Array.isArray(arr) ? arr.map((x) => num(file.resolve(x))) : null;
};
const colourOf = (file, v) => {
  const c = numbersOf(file, v);
  if (!c || !c.length) return null;
  if (c.length === 1) return [c[0], c[0], c[0]];
  if (c.length === 3) return c;
  if (c.length === 4) return [(1 - c[0]) * (1 - c[3]), (1 - c[1]) * (1 - c[3]), (1 - c[2]) * (1 - c[3])];
  return null;
};
const normRect = (r) => (r && r.length >= 4 ? [Math.min(r[0], r[2]), Math.min(r[1], r[3]), Math.max(r[0], r[2]), Math.max(r[1], r[3])] : null);

/** The size a FreeText's default appearance string names, or null. */
function fontSizeOf(da) {
  const m = /([\d.]+)\s+Tf/.exec(da || '');
  return m ? Number(m[1]) : null;
}

/**
 * Every comment on every page, in page order and then the order its page
 * lists them: each as plain values a window can hold (see the fields below).
 * Replies come as annotations of their own with `replyTo`; `state` on an
 * annotation is its latest review state, from the hidden state replies.
 */
function readAnnotations(file) {
  const out = [];
  const pages = file.pages();
  pages.forEach((page, p) => {
    const list = file.resolve(page.dict.get('Annots'));
    if (!Array.isArray(list)) return;
    list.forEach((item, idx) => {
      let dict;
      try { dict = file.resolve(item); } catch { dict = null; }
      if (!(dict instanceof Dict)) return;
      const sub = dict.get('Subtype');
      const type = sub instanceof Name ? sub.name : '';
      if (!COMMENT_TYPES.has(type)) return;
      const irt = dict.get('IRT');
      const ap = file.resolve(dict.get('AP'));
      const bs = file.resolve(dict.get('BS'));
      const border = numbersOf(file, dict.get('Border'));
      const le = file.resolve(dict.get('LE'));
      const ink = file.resolve(dict.get('InkList'));
      const state = textOf(file, dict.get('State')) || (dict.get('State') instanceof Name ? dict.get('State').name : '');
      const model = dict.get('StateModel') instanceof Name ? dict.get('StateModel').name : textOf(file, dict.get('StateModel'));
      const flags = num(file.resolve(dict.get('F')));
      const quads = numbersOf(file, dict.get('QuadPoints'));
      const width = bs instanceof Dict && typeof file.resolve(bs.get('W')) === 'number' ? file.resolve(bs.get('W')) : border && border.length >= 3 ? border[2] : 1;
      const icon = dict.get('Name');
      out.push({
        id: item instanceof Ref ? String(item.num) : `d${p}.${idx}`,
        ref: item instanceof Ref ? item.num : null,
        page: p,
        type,
        arrow: type === 'Line' && Array.isArray(le) && le.some((e) => e instanceof Name && /Arrow/.test(e.name)),
        rect: normRect(numbersOf(file, dict.get('Rect'))),
        contents: textOf(file, dict.get('Contents')),
        author: textOf(file, dict.get('T')),
        subject: textOf(file, dict.get('Subj')),
        modified: parseDate(textOf(file, dict.get('M'))),
        created: parseDate(textOf(file, dict.get('CreationDate'))),
        name: textOf(file, dict.get('NM')),
        color: colourOf(file, dict.get('C')),
        interior: colourOf(file, dict.get('IC')),
        opacity: typeof file.resolve(dict.get('CA')) === 'number' ? file.resolve(dict.get('CA')) : 1,
        width,
        quads: quads && quads.length >= 8 ? chunk(quads, 8) : null,
        ink: Array.isArray(ink) ? ink.map((stroke) => (numbersOf(file, stroke) || [])) : null,
        line: numbersOf(file, dict.get('L')),
        replyTo: irt instanceof Ref ? String(irt.num) : null,
        stateModel: model || null,
        state: state || null,
        hidden: Boolean(flags & FLAG_HIDDEN),
        icon: icon instanceof Name ? icon.name : null,
        fontSize: type === 'FreeText' ? fontSizeOf(textOf(file, dict.get('DA'))) : null,
        overlay: type === 'Redact' ? textOf(file, dict.get('OverlayText')) : null,
        hasAppearance: ap instanceof Dict && Boolean(ap.get('N')),
      });
    });
  });
  // Each comment's latest review state, from the state replies to it.
  const byId = new Map(out.map((a) => [a.id, a]));
  const latest = new Map();
  for (const a of out) {
    if (!a.replyTo || !a.state || (a.stateModel && a.stateModel !== 'Review')) continue;
    const had = latest.get(a.replyTo);
    if (!had || String(a.modified || '') >= String(had.modified || '')) latest.set(a.replyTo, a);
  }
  for (const [id, s] of latest) {
    const a = byId.get(id);
    if (a) { a.status = s.state; a.statusBy = s.author; }
  }
  return out;
}

function chunk(list, size) {
  const out = [];
  for (let i = 0; i + size <= list.length; i += size) out.push(list.slice(i, i + size));
  return out;
}

// ── appearances ───────────────────────────────────────────────────────────

const rgb = (c) => (c || [0, 0, 0]).slice(0, 3).map((x) => n(Math.max(0, Math.min(1, x)))).join(' ');

function boundsOf(points, pad = 0) {
  let x0 = Infinity; let y0 = Infinity; let x1 = -Infinity; let y1 = -Infinity;
  for (let i = 0; i + 1 < points.length; i += 2) {
    x0 = Math.min(x0, points[i]); x1 = Math.max(x1, points[i]);
    y0 = Math.min(y0, points[i + 1]); y1 = Math.max(y1, points[i + 1]);
  }
  if (!Number.isFinite(x0)) return [0, 0, 0, 0];
  return [x0 - pad, y0 - pad, x1 + pad, y1 + pad];
}

const KAPPA = 0.5522847498;
function ellipsePath(x0, y0, x1, y1) {
  const cx = (x0 + x1) / 2; const cy = (y0 + y1) / 2;
  const rx = (x1 - x0) / 2; const ry = (y1 - y0) / 2;
  const kx = rx * KAPPA; const ky = ry * KAPPA;
  return [
    `${n(cx + rx)} ${n(cy)} m`,
    `${n(cx + rx)} ${n(cy + ky)} ${n(cx + kx)} ${n(cy + ry)} ${n(cx)} ${n(cy + ry)} c`,
    `${n(cx - kx)} ${n(cy + ry)} ${n(cx - rx)} ${n(cy + ky)} ${n(cx - rx)} ${n(cy)} c`,
    `${n(cx - rx)} ${n(cy - ky)} ${n(cx - kx)} ${n(cy - ry)} ${n(cx)} ${n(cy - ry)} c`,
    `${n(cx + kx)} ${n(cy - ry)} ${n(cx + rx)} ${n(cy - ky)} ${n(cx + rx)} ${n(cy)} c`,
  ].join('\n');
}

/** A quad's corners as (top-left, top-right, bottom-left, bottom-right), as Acrobat orders QuadPoints. */
const corners = (q) => ({ tl: [q[0], q[1]], tr: [q[2], q[3]], bl: [q[4], q[5]], br: [q[6], q[7]] });
const lerp = (a, b, t) => [a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t];
const dist = (a, b) => Math.hypot(b[0] - a[0], b[1] - a[1]);

/** Wrapped lines of `text` in Helvetica at `size` within `width` points. */
function wrapLines(text, size, width) {
  const out = [];
  for (const para of String(text || '').split(/\r\n|\r|\n/)) {
    const words = para.split(/(\s+)/).filter((w) => w.length);
    let line = '';
    for (const w of words) {
      const trial = line + w;
      if (line && widthOfBytes(encode(trial.trimEnd()), 'Helvetica', size) > width) {
        out.push(line.trimEnd());
        line = /^\s+$/.test(w) ? '' : w;
      } else line = trial;
    }
    out.push(line.trimEnd());
  }
  return out;
}

const helvetica = () => new Dict([
  ['Type', new Name('Font')], ['Subtype', new Name('Type1')], ['BaseFont', new Name('Helvetica')], ['Encoding', new Name('WinAnsiEncoding')],
]);

/**
 * The appearance of an annotation described by `spec`, and the Rect it
 * stands in. Returns { rect, content, resources } with the content in page
 * coordinates (the form's BBox is the Rect, its matrix the identity).
 */
function appearance(spec, extras = {}) {
  const type = spec.type;
  const colour = spec.color || DEFAULT_COLOURS[type] || [0, 0, 0];
  const width = num(spec.width, type === 'Ink' ? 2 : 1.5);
  const ops = [];
  const resources = new Dict();
  let rect = normRect(spec.rect);
  const opacity = num(spec.opacity, 1);

  const gs = (entries) => {
    const ext = new Dict([['GS0', new Dict(entries)]]);
    resources.set('ExtGState', ext);
    ops.push('/GS0 gs');
  };

  if (type === 'Highlight' || type === 'Underline' || type === 'StrikeOut' || type === 'Squiggly') {
    const quads = spec.quads || [];
    rect = boundsOf(quads.flat(), 1);
    if (type === 'Highlight') {
      gs([['BM', new Name('Multiply')], ['ca', opacity], ['CA', opacity]]);
      ops.push(`${rgb(colour)} rg`);
      for (const q of quads) {
        const c = corners(q);
        ops.push(`${n(c.tl[0])} ${n(c.tl[1])} m ${n(c.tr[0])} ${n(c.tr[1])} l ${n(c.br[0])} ${n(c.br[1])} l ${n(c.bl[0])} ${n(c.bl[1])} l h f`);
      }
    } else {
      if (opacity < 1) gs([['CA', opacity], ['ca', opacity]]);
      ops.push(`${rgb(colour)} RG`);
      for (const q of quads) {
        const c = corners(q);
        const h = dist(c.tl, c.bl);
        const lw = Math.max(0.5, h * 0.07);
        ops.push(`${n(lw)} w`);
        // Underline a little above the bottom edge, strikethrough through the middle.
        const t = type === 'StrikeOut' ? 0.52 : 1 - Math.min(0.12, (lw * 1.2) / Math.max(h, 1));
        const a = lerp(c.tl, c.bl, t);
        const b = lerp(c.tr, c.br, t);
        ops.push(`${n(a[0])} ${n(a[1])} m ${n(b[0])} ${n(b[1])} l S`);
      }
    }
  } else if (type === 'Text') {
    // A note's icon, 20 points square from its top-left corner, as Acrobat's Comment icon is.
    const x = rect ? rect[0] : 0;
    const top = rect ? rect[3] : 20;
    rect = [x, top - 20, x + 20, top];
    const [x0, y0] = [x, top - 20];
    ops.push('0.25 0.25 0.25 RG 0.8 w', `${rgb(colour)} rg`);
    ops.push(`${n(x0 + 1.5)} ${n(y0 + 6)} m ${n(x0 + 1.5)} ${n(y0 + 18.5)} l ${n(x0 + 18.5)} ${n(y0 + 18.5)} l ${n(x0 + 18.5)} ${n(y0 + 6)} l ${n(x0 + 9)} ${n(y0 + 6)} l ${n(x0 + 5)} ${n(y0 + 1.5)} l ${n(x0 + 5.5)} ${n(y0 + 6)} l h B`);
    for (const ly of [15, 12, 9]) ops.push(`${n(x0 + 4.5)} ${n(y0 + ly)} m ${n(x0 + 15.5)} ${n(y0 + ly)} l S`);
  } else if (type === 'FreeText') {
    const size = num(spec.fontSize, 12);
    if (!rect) rect = [0, 0, 200, size * 2];
    const [x0, y0, x1, y1] = rect;
    if (spec.interior) ops.push(`${rgb(spec.interior)} rg ${n(x0)} ${n(y0)} ${n(x1 - x0)} ${n(y1 - y0)} re f`);
    if (width > 0 && spec.border !== false && spec.borderColor) ops.push(`${rgb(spec.borderColor)} RG ${n(width)} w ${n(x0 + width / 2)} ${n(y0 + width / 2)} ${n(x1 - x0 - width)} ${n(y1 - y0 - width)} re S`);
    resources.set('Font', new Dict([['Helv', extras.fontRef || helvetica()]]));
    const pad = 3;
    const lines = wrapLines(spec.contents || '', size, Math.max(10, x1 - x0 - pad * 2));
    ops.push(`${n(x0)} ${n(y0)} ${n(x1 - x0)} ${n(y1 - y0)} re W n`, 'BT', `${rgb(colour)} rg`, `/Helv ${n(size)} Tf`, `${n(size * 1.2)} TL`);
    ops.push(`${n(x0 + pad)} ${n(y1 - pad - size * 0.9)} Td`);
    lines.forEach((line, i) => { ops.push(`${i ? 'T* ' : ''}${literal(encode(line))} Tj`); });
    ops.push('ET');
  } else if (type === 'Ink') {
    const strokes = (spec.ink || []).filter((s) => s.length >= 2);
    rect = boundsOf(strokes.flat(), width + 1);
    if (opacity < 1) gs([['CA', opacity], ['ca', opacity]]);
    ops.push(`${rgb(colour)} RG ${n(width)} w 1 J 1 j`);
    for (const s of strokes) {
      if (s.length === 2) { ops.push(`${n(s[0])} ${n(s[1])} m ${n(s[0] + 0.01)} ${n(s[1])} l S`); continue; }
      // Smoothed through the midpoints of its segments, as a pen draws.
      let path = `${n(s[0])} ${n(s[1])} m`;
      for (let i = 2; i + 3 < s.length; i += 2) {
        const mx = (s[i] + s[i + 2]) / 2; const my = (s[i + 1] + s[i + 3]) / 2;
        path += ` ${n(s[i])} ${n(s[i + 1])} ${n(s[i])} ${n(s[i + 1])} ${n(mx)} ${n(my)} c`;
      }
      path += ` ${n(s[s.length - 2])} ${n(s[s.length - 1])} l S`;
      ops.push(path);
    }
  } else if (type === 'Square' || type === 'Circle') {
    if (!rect) rect = [0, 0, 100, 100];
    const h = width / 2;
    const [x0, y0, x1, y1] = [rect[0] + h, rect[1] + h, rect[2] - h, rect[3] - h];
    if (opacity < 1) gs([['CA', opacity], ['ca', opacity]]);
    ops.push(`${rgb(colour)} RG ${n(width)} w`);
    if (spec.interior) ops.push(`${rgb(spec.interior)} rg`);
    const paint = spec.interior ? 'B' : 'S';
    if (type === 'Square') ops.push(`${n(x0)} ${n(y0)} ${n(x1 - x0)} ${n(y1 - y0)} re ${paint}`);
    else ops.push(`${ellipsePath(x0, y0, x1, y1)} ${paint}`);
  } else if (type === 'Line') {
    const [ax, ay, bx, by] = spec.line;
    const head = Math.max(6, width * 4);
    rect = boundsOf([ax, ay, bx, by], head + width);
    if (opacity < 1) gs([['CA', opacity], ['ca', opacity]]);
    ops.push(`${rgb(colour)} RG ${n(width)} w 1 J 1 j`);
    ops.push(`${n(ax)} ${n(ay)} m ${n(bx)} ${n(by)} l S`);
    if (spec.arrow) {
      const ang = Math.atan2(by - ay, bx - ax);
      const wing = (a) => [bx - head * Math.cos(a), by - head * Math.sin(a)];
      const p1 = wing(ang + Math.PI / 6);
      const p2 = wing(ang - Math.PI / 6);
      ops.push(`${n(p1[0])} ${n(p1[1])} m ${n(bx)} ${n(by)} l ${n(p2[0])} ${n(p2[1])} l S`);
    }
  } else if (type === 'Stamp') {
    if (!rect) rect = [0, 0, 120, 40];
    const [x0, y0, x1, y1] = rect;
    if (extras.imageRef) {
      resources.set('XObject', new Dict([['Im0', extras.imageRef]]));
      ops.push(`q ${n(x1 - x0)} 0 0 ${n(y1 - y0)} ${n(x0)} ${n(y0)} cm /Im0 Do Q`);
    }
    if (spec.paths && spec.paths.length) {
      ops.push(`${rgb(colour)} RG ${n(width)} w 1 J 1 j`);
      for (const s of spec.paths) {
        if (s.length < 4) continue;
        let path = `${n(s[0])} ${n(s[1])} m`;
        for (let i = 2; i + 1 < s.length; i += 2) path += ` ${n(s[i])} ${n(s[i + 1])} l`;
        ops.push(`${path} S`);
      }
    }
  } else if (type === 'Redact') {
    const quads = spec.quads || null;
    rect = quads ? boundsOf(quads.flat(), 0) : rect;
    ops.push('0.86 0.1 0.1 RG 1 w');
    if (quads) {
      for (const q of quads) {
        const c = corners(q);
        ops.push(`${n(c.tl[0])} ${n(c.tl[1])} m ${n(c.tr[0])} ${n(c.tr[1])} l ${n(c.br[0])} ${n(c.br[1])} l ${n(c.bl[0])} ${n(c.bl[1])} l h S`);
      }
    } else if (rect) ops.push(`${n(rect[0] + 0.5)} ${n(rect[1] + 0.5)} ${n(rect[2] - rect[0] - 1)} ${n(rect[3] - rect[1] - 1)} re S`);
  }
  return { rect: rect || [0, 0, 0, 0], content: ops.join('\n'), resources };
}

// ── writing ───────────────────────────────────────────────────────────────

const colourArray = (c) => (c ? c.slice(0, 3).map((x) => Math.round(Math.max(0, Math.min(1, x)) * 1000) / 1000) : null);
const uid = () => `rutba-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;

/** A picture for a stamp: PNG bytes, or { width, height, data (RGB), alpha }. */
function imageXObject(up, image) {
  let pic = image;
  if (Buffer.isBuffer(image) || image instanceof Uint8Array) {
    const png = decodePng(Buffer.from(image));
    pic = { width: png.width, height: png.height, data: png.data, alpha: png.alpha, colorSpace: png.colorSpace };
  }
  const data = Buffer.from(pic.data);
  const dict = new Dict([
    ['Type', new Name('XObject')], ['Subtype', new Name('Image')], ['Width', pic.width], ['Height', pic.height],
    ['ColorSpace', new Name(pic.colorSpace || (data.length >= pic.width * pic.height * 3 ? 'DeviceRGB' : 'DeviceGray'))], ['BitsPerComponent', 8],
  ]);
  if (pic.alpha) {
    const mask = new Dict([
      ['Type', new Name('XObject')], ['Subtype', new Name('Image')], ['Width', pic.width], ['Height', pic.height],
      ['ColorSpace', new Name('DeviceGray')], ['BitsPerComponent', 8], ['Filter', new Name('FlateDecode')],
    ]);
    dict.set('SMask', up.add(new Stream(mask, zlib.deflateSync(Buffer.from(pic.alpha)))));
  }
  dict.set('Filter', new Name('FlateDecode'));
  return up.add(new Stream(dict, zlib.deflateSync(data)));
}

/**
 * Change the comments of a PDF, as one incremental update.
 *
 * @param {Buffer|Uint8Array} bytes the whole file
 * @param {{
 *   add?: object[],                       // new annotations: { page, type, rect, quads, ink, line, arrow, color, interior, width, opacity, contents, fontSize, paths, image, subject, icon, overlay }
 *   reply?: { to: string, contents: string }[],
 *   resolve?: { id: string, resolved: boolean }[],
 *   edit?: { id: string, contents: string }[],
 *   remove?: string[],                    // ids, as readAnnotations names them
 * }} changes
 * @param {{ author?: string, date?: Date }} [options]
 * @returns {{ bytes: Buffer, added: string[] }} the file with the update after it, and the new annotations' ids
 */
function annotate(bytes, changes = {}, { author = '', date = new Date() } = {}) {
  const up = new PdfUpdate(bytes, { what: 'A comment' });
  const { file } = up;
  const pages = file.pages();
  const stamp = pdfDate(date);
  const who = textString(author || '');

  // Each page's list as it will be, and where the list lives.
  const lists = new Map();
  const listOf = (p) => {
    if (!lists.has(p)) {
      const page = pages[p];
      if (!page) throw new UpdateError(`There is no page ${p + 1} in this PDF.`, 'page');
      if (!page.ref) throw new UpdateError(`Page ${p + 1} cannot be changed.`, 'page');
      const raw = page.dict.get('Annots');
      const arr = file.resolve(raw);
      lists.set(p, { raw, items: Array.isArray(arr) ? [...arr] : [], changed: false });
    }
    return lists.get(p);
  };

  // Where each existing annotation is.
  const existing = readAnnotations(file);
  const byId = new Map(existing.map((a) => [a.id, a]));
  const dictOf = (a) => {
    if (a.ref) return file.getObject(a.ref);
    const m = /^d(\d+)\.(\d+)$/.exec(a.id);
    return m ? file.resolve(file.resolve(pages[+m[1]].dict.get('Annots'))[+m[2]]) : null;
  };
  const itemOf = (a) => {
    if (a.ref) return listOf(a.page).items.find((x) => x instanceof Ref && x.num === a.ref);
    return dictOf(a);
  };

  const added = [];
  const write = (spec) => {
    const p = Number(spec.page);
    const page = pages[p];
    if (!page) throw new UpdateError(`There is no page ${p + 1} in this PDF.`, 'page');
    const type = spec.type === 'Arrow' ? 'Line' : spec.type;
    if (!COMMENT_TYPES.has(type)) throw new UpdateError(`${spec.type} is not a kind of comment.`, 'page');
    const s = { ...spec, type, arrow: spec.arrow || spec.type === 'Arrow' };
    const extras = {};
    if (type === 'Stamp' && s.image) extras.imageRef = imageXObject(up, s.image);
    const look = appearance(s, extras);
    const dict = new Dict([
      ['Type', new Name('Annot')], ['Subtype', new Name(type)], ['Rect', look.rect.map((v) => Math.round(v * 1000) / 1000)],
      ['P', page.ref], ['NM', textString(uid())], ['M', textString(stamp)], ['CreationDate', textString(stamp)],
      ['F', FLAG_PRINT], ['T', who],
      ['Subj', textString(s.subject || SUBJECTS[s.arrow ? 'Arrow' : type] || type)],
    ]);
    if (s.contents) dict.set('Contents', textString(s.contents));
    const colour = colourArray(s.color || DEFAULT_COLOURS[type]);
    if (type !== 'FreeText' && colour) dict.set('C', colour);
    if (s.opacity !== undefined && s.opacity < 1) dict.set('CA', Math.round(s.opacity * 1000) / 1000);
    if (s.quads) dict.set('QuadPoints', s.quads.flat().map((v) => Math.round(v * 1000) / 1000));
    if (['Ink', 'Square', 'Circle', 'Line', 'Stamp', 'FreeText'].includes(type)) {
      dict.set('BS', new Dict([['Type', new Name('Border')], ['W', num(s.width, type === 'Ink' ? 2 : type === 'FreeText' ? 0 : 1.5)], ['S', new Name('S')]]));
    }
    if (type === 'Ink') dict.set('InkList', (s.ink || []).map((st) => st.map((v) => Math.round(v * 1000) / 1000)));
    if (type === 'Line') {
      dict.set('L', s.line.map((v) => Math.round(v * 1000) / 1000));
      dict.set('LE', [new Name('None'), new Name(s.arrow ? 'OpenArrow' : 'None')]);
    }
    if ((type === 'Square' || type === 'Circle') && s.interior) dict.set('IC', colourArray(s.interior));
    if (type === 'Text') { dict.set('Name', new Name(s.icon || 'Comment')); dict.set('Open', false); }
    if (type === 'FreeText') {
      const size = num(s.fontSize, 12);
      const c = colourArray(s.color || [0, 0, 0]);
      dict.set('DA', textString(`${c.map(n).join(' ')} rg /Helv ${n(size)} Tf`));
      dict.set('Q', 0);
      if (s.borderColor) dict.set('C', colourArray(s.borderColor));
    }
    if (type === 'Stamp') dict.set('Name', new Name(s.icon || 'Signature'));
    if (type === 'Redact') {
      dict.set('IC', [0, 0, 0]);
      dict.set('OC', [0.86, 0.1, 0.1]);
      if (s.overlay) dict.set('OverlayText', textString(s.overlay));
    }
    const form = flateStream([
      ['Type', new Name('XObject')], ['Subtype', new Name('Form')], ['BBox', look.rect.map((v) => Math.round(v * 1000) / 1000)], ['Resources', look.resources],
    ], look.content);
    dict.set('AP', new Dict([['N', up.add(form)]]));
    const ref = up.add(dict);
    const list = listOf(p);
    list.items.push(ref);
    list.changed = true;
    added.push(String(ref.num));
    return ref;
  };

  for (const spec of changes.add || []) write(spec);

  // A reply, or a review state: a Text annotation in reply to the comment.
  const answer = (target, { contents = '', state = null }) => {
    const a = byId.get(String(target));
    if (!a || !a.ref) throw new UpdateError('That comment is no longer in the file.', 'missing');
    const page = pages[a.page];
    const dict = new Dict([
      ['Type', new Name('Annot')], ['Subtype', new Name('Text')], ['Rect', a.rect || [0, 0, 20, 20]], ['P', page.ref],
      ['IRT', new Ref(a.ref, file.entries.get(a.ref)?.gen || 0)], ['NM', textString(uid())], ['M', textString(stamp)], ['CreationDate', textString(stamp)],
      ['T', who], ['Name', new Name('Comment')], ['Open', false],
    ]);
    if (state) {
      dict.set('F', 30);
      dict.set('StateModel', textString('Review'));
      dict.set('State', textString(state));
      dict.set('Contents', textString(state === 'Completed' ? `Completed set by ${author || 'a reviewer'}` : `None set by ${author || 'a reviewer'}`));
    } else {
      dict.set('F', 28);
      dict.set('Contents', textString(contents));
      dict.set('RT', new Name('R'));
    }
    const ref = up.add(dict);
    const list = listOf(a.page);
    list.items.push(ref);
    list.changed = true;
    added.push(String(ref.num));
  };
  for (const r of changes.reply || []) {
    if (String(r.contents || '').trim()) answer(r.to, { contents: String(r.contents) });
  }

  // Taking annotations out: the comment, its replies and their pop-ups.
  const removing = new Set();
  const take = (id) => {
    if (removing.has(id)) return;
    const a = byId.get(id);
    if (!a) return;
    removing.add(id);
    for (const r of existing) if (r.replyTo === id) take(r.id);
  };
  for (const id of changes.remove || []) take(String(id));
  // Resolving replaces the comment's review states with one, or with none.
  for (const r of changes.resolve || []) {
    for (const s of existing) if (s.replyTo === String(r.id) && s.state && (!s.stateModel || s.stateModel === 'Review')) take(s.id);
    if (r.resolved) answer(r.id, { state: 'Completed' });
  }
  for (const id of removing) {
    const a = byId.get(id);
    const list = listOf(a.page);
    const item = itemOf(a);
    const dict = dictOf(a);
    const at = list.items.indexOf(item);
    if (at >= 0) { list.items.splice(at, 1); list.changed = true; }
    if (dict instanceof Dict) {
      const popup = dict.get('Popup');
      if (popup instanceof Ref) {
        const pi = list.items.findIndex((x) => x instanceof Ref && x.num === popup.num);
        if (pi >= 0) list.items.splice(pi, 1);
        up.free(popup);
      }
    }
    if (a.ref) up.free(new Ref(a.ref, file.entries.get(a.ref)?.gen || 0));
  }

  // Edits: new words, and for a text box the appearance drawn again.
  for (const e of changes.edit || []) {
    const a = byId.get(String(e.id));
    if (!a || !a.ref || removing.has(a.id)) continue;
    const dict = copyDict(file.getObject(a.ref));
    dict.set('Contents', textString(e.contents || ''));
    dict.set('M', textString(stamp));
    if (a.type === 'FreeText') {
      const look = appearance({ type: 'FreeText', rect: a.rect, contents: e.contents, fontSize: a.fontSize || 12, color: colourFromDA(textOf(file, dict.get('DA'))), borderColor: a.color });
      const form = flateStream([
        ['Type', new Name('XObject')], ['Subtype', new Name('Form')], ['BBox', look.rect], ['Resources', look.resources],
      ], look.content);
      dict.set('AP', new Dict([['N', up.add(form)]]));
    }
    up.set(a.ref, dict);
  }

  // Each page's list written back where it lives.
  for (const [p, list] of lists) {
    if (!list.changed) continue;
    const page = pages[p];
    if (list.raw instanceof Ref && Array.isArray(file.resolve(list.raw))) {
      up.set(list.raw, list.items);
    } else {
      const dict = copyDict(page.dict);
      if (list.items.length) dict.set('Annots', list.items);
      else dict.delete('Annots');
      up.set(page.ref, dict);
    }
  }
  return { bytes: up.toBuffer(), added };
}

function colourFromDA(da) {
  const m = /([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+rg/.exec(da || '');
  if (m) return [Number(m[1]), Number(m[2]), Number(m[3])];
  const g = /([\d.]+)\s+g\b/.exec(da || '');
  return g ? [Number(g[1]), Number(g[1]), Number(g[1])] : [0, 0, 0];
}

module.exports = {
  readAnnotations, annotate, appearance, parseDate, pdfDate, wrapLines, COMMENT_TYPES, SUBJECTS, DEFAULT_COLOURS,
};
