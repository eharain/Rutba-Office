// Presentation: what the accessibility checker reads, the fixes it writes,
// and the text the spelling pass walks — against a Deck (presentation).
//
// The deck's scenes say what is on each slide and where; the slide's own XML
// says what the scene leaves out — each object's description and whether it
// is decorative, and whether a table's first row is a header. Fixes write
// the slide part; the document service takes the deck's undo snapshot
// around every op, so each is one step.

import { findElements, readAltProps, writeAltProps } from './alt-text.js';
import { normaliseColour } from './colour.js';

const shortText = (s, n = 40) => {
  const t = String(s ?? '').replace(/\s+/g, ' ').trim();
  return t.length > n ? `${t.slice(0, n - 1)}…` : t;
};

const textOf = (shape) => (shape?.text?.paragraphs || []).map((p) => (p.runs || []).map((r) => r.text || '').join('')).join('\n');

/** The `p:cNvPr` of a shape by id, in a slide's XML: `{ start, end, xml }`. */
function cNvPrOf(xml, id) {
  return findElements(xml, 'p:cNvPr').find((e) => new RegExp(`\\sid="${String(id).replace(/[^\w-]/g, '')}"`).test(e.open)) || null;
}

/** The element (p:sp, p:pic, p:graphicFrame, p:grpSp …) whose own cNvPr has this id. */
function shapeElementOf(xml, id) {
  const c = cNvPrOf(xml, id);
  if (!c) return null;
  for (const tag of ['p:graphicFrame', 'p:pic', 'p:sp', 'p:grpSp', 'p:cxnSp']) {
    const open = Math.max(xml.lastIndexOf(`<${tag}>`, c.start), xml.lastIndexOf(`<${tag} `, c.start));
    if (open < 0) continue;
    const close = xml.indexOf(`</${tag}>`, c.start);
    if (close < 0) continue;
    // The nearest opening before the cNvPr must not be closed before it.
    if (xml.indexOf(`</${tag}>`, open) < c.start) continue;
    return { tag, start: open, end: close + tag.length + 3, xml: xml.slice(open, close + tag.length + 3) };
  }
  return null;
}

const isTitle = (s) => s.placeholder && (s.placeholder.type === 'title' || s.placeholder.type === 'ctrTitle');

/** The deck's language: the most common `lang` its slides' runs carry. */
export function deckLanguage(deck) {
  const counts = new Map();
  for (let i = 0; i < deck.slideCount; i++) {
    const part = deck.slideParts[i]?.part;
    if (!part) continue;
    for (const m of deck.pkg.text(part).matchAll(/<a:(?:rPr|endParaRPr|defRPr)\b[^>]*\blang="([^"]+)"/g)) counts.set(m[1], (counts.get(m[1]) || 0) + 1);
  }
  return [...counts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || null;
}

/** The deck, described for the rules in accessibility.js. */
export function describeDeck(deck) {
  const model = { app: 'slides', objects: [], tables: [], texts: [], links: [], slides: [] };
  for (let i = 0; i < deck.slideCount; i++) {
    const scene = deck.slide(i);
    const xml = deck.pkg.text(scene.part);
    const top = scene.shapes.filter((s) => s.groupId == null);
    const title = top.find(isTitle) || scene.shapes.find(isTitle) || null;
    const slideBg = scene.background?.type === 'solid' ? normaliseColour(scene.background.color) : null;
    model.slides.push({
      index: i,
      where: { slide: i },
      title: title ? textOf(title) : '',
      titleShape: title ? title.id : null,
      shapes: top.filter((s) => s.geometry && s.kind !== 'connector').map((s) => ({ id: s.id, name: s.name, hidden: s.hidden, x: s.geometry.x, y: s.geometry.y, w: s.geometry.w, h: s.geometry.h })),
    });

    for (const s of top) {
      if (s.hidden) continue;
      const hasWords = Boolean(textOf(s).trim());
      const math = (s.text?.paragraphs || []).some((p) => (p.runs || []).some((r) => r.math));
      const needsAlt = s.kind === 'picture' || s.kind === 'chart' || s.kind === 'group' || (s.kind === 'shape' && !hasWords && !s.placeholder && !math && s.preset !== 'line') || s.kind === 'unsupported';
      if (needsAlt) {
        const c = cNvPrOf(xml, s.id);
        const props = c ? readAltProps(c.xml) : { descr: '', decorative: false };
        model.objects.push({
          key: `${i}:${s.id}`, kind: s.kind,
          label: `${s.name || s.kind} (slide ${i + 1})`,
          alt: props.descr, decorative: props.decorative,
          where: { slide: i, shape: s.id }, target: { slide: i, shape: s.id },
        });
      }
      if (s.kind === 'table') {
        const frame = shapeElementOf(xml, s.id);
        const tblPr = /<a:tblPr\b[^>]*>/.exec(frame?.xml || '')?.[0] || '';
        model.tables.push({
          key: `${i}:${s.id}`, label: `${s.name || 'Table'} (slide ${i + 1})`,
          hasHeader: /\sfirstRow="(1|true)"/.test(tblPr),
          merged: /<a:tc\b[^>]*\s(gridSpan|rowSpan|hMerge|vMerge)="/.test(frame?.xml || ''),
          where: { slide: i, shape: s.id }, target: { slide: i, shape: s.id },
          headerFix: 'Make the first row a header row',
        });
      }
    }

    // Text against what is behind it: the shape's own fill, else the slide's background.
    for (const s of scene.shapes) {
      if (s.hidden || !s.text?.paragraphs) continue;
      const fill = s.fill?.type === 'solid' ? normaliseColour(s.fill.color) : null;
      const bg = fill || (s.fill && s.fill.type !== 'none' ? null : slideBg || '#FFFFFF');
      if (!bg) continue;
      for (const p of s.text.paragraphs) {
        const lv = s.textStyle?.[Math.min(8, p.level || 0)] || {};
        for (const r of p.runs || []) {
          if (!String(r.text || '').trim() || r.math) continue;
          const fg = normaliseColour(r.color || (r.link ? lv.linkColor : null) || lv.color || s.styleText?.color || '#1A1A1A');
          model.texts.push({ key: `${i}:${s.id}`, label: `${s.name || 'Text'} (slide ${i + 1}): ${shortText(textOf(s), 30)}`, where: { slide: i, shape: s.id }, fg, bg, sizePt: r.size || lv.size || 18, bold: Boolean(r.bold ?? lv.bold), target: { slide: i, shape: s.id } });
        }
      }
      for (const p of s.text.paragraphs) {
        let current = null;
        const flush = () => {
          if (current) model.links.push({ key: `${i}:${s.id}:${model.links.length}`, text: current.text, url: current.url, label: `"${shortText(current.text)}" (slide ${i + 1})`, where: { slide: i, shape: s.id } });
          current = null;
        };
        for (const r of p.runs || []) {
          const url = r.link ? (typeof r.link === 'object' ? r.link.url || r.link.id : String(r.link)) : null;
          if (url && current && current.url === url) current.text += r.text || '';
          else {
            flush();
            if (url) current = { url, text: r.text || '' };
          }
        }
        flush();
      }
    }
  }
  return model;
}

/* ── the fixes ─────────────────────────────────────────────────────────── */

function writeSlide(deck, part, xml) {
  deck.pkg.write_(part, Buffer.from(xml, 'utf8'));
  deck.dirty = true;
  deck._scenes?.delete?.(part);
}

/** Alt text on a shape's own `p:cNvPr`. */
export function setDeckAltText(deck, { slide, shape, descr = '', decorative = false }) {
  const part = deck.slideParts[Number(slide)]?.part;
  if (!part) throw new RangeError(`no slide at index ${slide}`);
  const xml = deck.pkg.text(part);
  const c = cNvPrOf(xml, shape);
  if (!c) throw new Error('that object is no longer on the slide');
  writeSlide(deck, part, xml.slice(0, c.start) + writeAltProps(c.xml, { descr, decorative }) + xml.slice(c.end));
  return true;
}

/** A table's first row as its header row (`a:tblPr firstRow="1"`). */
export function setDeckTableHeader(deck, { slide, shape }) {
  const part = deck.slideParts[Number(slide)]?.part;
  if (!part) throw new RangeError(`no slide at index ${slide}`);
  const xml = deck.pkg.text(part);
  const frame = shapeElementOf(xml, shape);
  if (!frame) throw new Error('that table is no longer on the slide');
  let next = frame.xml;
  if (/<a:tblPr\b[^>]*\sfirstRow="[^"]*"/.test(next)) next = next.replace(/(<a:tblPr\b[^>]*\s)firstRow="[^"]*"/, '$1firstRow="1"');
  else if (/<a:tblPr\b/.test(next)) next = next.replace(/<a:tblPr\b/, '<a:tblPr firstRow="1"');
  else next = next.replace(/<a:tbl>/, '<a:tbl><a:tblPr firstRow="1"/>');
  writeSlide(deck, part, xml.slice(0, frame.start) + next + xml.slice(frame.end));
  return true;
}

/** A shape's words all in one colour — the contrast fix. */
export function setDeckTextColour(deck, { slide, shape, colour }) {
  const scene = deck.slide(Number(slide));
  const s = scene.shapes.find((x) => String(x.id) === String(shape));
  if (!s?.text?.paragraphs) throw new Error('that text is no longer on the slide');
  const paragraphs = s.text.paragraphs.map(({ plain, runs, ...props }) => ({ ...props, runs: (runs || []).map((r) => (r.break || r.field ? r : { ...r, color: colour })) }));
  return deck.setText(Number(slide), s.id, paragraphs);
}

/** The slide's title, typed into its title placeholder. */
export function setDeckSlideTitle(deck, { slide, shape, title }) {
  const scene = deck.slide(Number(slide));
  const s = scene.shapes.find((x) => String(x.id) === String(shape));
  if (!s) throw new Error('that slide has no title placeholder');
  const first = s.text?.paragraphs?.[0];
  const look = first?.runs?.find((r) => r.text) || {};
  const { plain, runs, ...props } = first || {};
  const { text: _t, link, field, ...runLook } = look;
  return deck.setText(Number(slide), s.id, [{ ...props, runs: [{ ...runLook, text: String(title ?? '').trim() }] }]);
}

/* ── the words, for the spelling pass ──────────────────────────────────── */

/**
 * Every run of words on every slide, a table's cells among them, then the
 * slide's notes — slide by slide, in the order the shapes are drawn.
 */
export function deckSegments(deck) {
  const out = [];
  for (let i = 0; i < deck.slideCount; i++) {
    const scene = deck.slide(i);
    for (const s of scene.shapes) {
      const walk = (paragraphs, row, col) => paragraphs.forEach((p, pi) => (p.runs || []).forEach((r, ri) => {
        if (!r.text || r.field || r.break || r.math) return;
        out.push({ key: `r:${i}:${s.id}:${row ?? ''}:${col ?? ''}:${pi}:${ri}`, text: r.text, where: { slide: i, shape: s.id, row, col } });
      }));
      if (s.text?.paragraphs) walk(s.text.paragraphs, null, null);
      if (s.table) s.table.rows?.forEach?.((row, ri) => row.cells.forEach((cell, ci) => { if (cell.text?.paragraphs) walk(cell.text.paragraphs, ri, ci); }));
    }
    if (String(scene.notes || '').trim()) out.push({ key: `t:${i}`, text: scene.notes, where: { slide: i, notes: true } });
  }
  return out;
}

export function deckStart(deck, { slide = 0 } = {}) {
  return { slide: Number(slide) || 0 };
}

/** Replace spans of runs or notes. Within a run, the rightmost first. */
export function replaceDeckText(deck, edits) {
  const list = [...edits].sort((a, b) => (a.key < b.key ? -1 : a.key > b.key ? 1 : b.from - a.from));
  let changed = 0;
  const notes = new Map();
  for (const e of list) {
    const parts = e.key.split(':');
    if (parts[0] === 't') {
      const slide = Number(parts[1]);
      const text = notes.has(slide) ? notes.get(slide) : deck.slide(slide).notes || '';
      notes.set(slide, text.slice(0, e.from) + e.text + text.slice(e.to));
      continue;
    }
    const [, slide, shape, row, col, paragraph, run] = parts;
    const hit = { slide: Number(slide), shape, row: row === '' ? null : Number(row), col: col === '' ? null : Number(col), paragraph: Number(paragraph), run: Number(run), offset: e.from, length: e.to - e.from };
    if (deck.replace(hit, e.text)) changed += 1;
  }
  for (const [slide, text] of notes) {
    deck.setNotes(slide, text);
    changed += 1;
  }
  return changed;
}
