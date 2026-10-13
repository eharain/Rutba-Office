// What the engine read, as the suite hands it on: the text, its lines and its
// words, each with a box in the picture's own pixels and a confidence from
// 0 to 100.
//
//   { width, height, languages, angle,
//     text,                                    lines joined, a blank line between blocks
//     lines: [{ text, box, confidence, direction, block, words: [first, end) }],
//                                              block: the column the line is in
//     words: [{ text, box, confidence, line }] }
//
// A box is { x, y, width, height }. Words and lines are in reading order,
// found from where they are on the page (readingOrder); a line's text is
// its words joined by a space, except between two characters of Chinese or
// Japanese, which are written without one.

import { columns } from './layout.js';

const CJK = /[⺀-ヿ㐀-䶿一-鿿豈-﫿＀-￯]/u;
const RTL = /[֐-ࣿיִ-﷿ﹰ-ﻼ]/gu;
const LTR = /[A-Za-zÀ-ɏऀ-ॿ぀-ヿ一-鿿]/gu;

/** Words joined as their script writes them. */
export function joinWords(words) {
  let out = '';
  for (const w of words) {
    if (out && !(CJK.test(out.at(-1)) && CJK.test(w[0]))) out += ' ';
    out += w;
  }
  return out;
}

/** Which way a line of text reads, by the letters it holds. */
export function directionOf(text) {
  const r = (text.match(RTL) || []).length;
  const l = (text.match(LTR) || []).length;
  return r > l ? 'rtl' : 'ltr';
}

const ARABIC_PUNCTUATION = { ',': '،', ';': '؛', '?': '؟' };

const boxOf = (b) => ({ x: b.x0, y: b.y0, width: Math.max(0, b.x1 - b.x0), height: Math.max(0, b.y1 - b.y0) });

/**
 * The engine's lines in reading order: the columns of the page (right to
 * left when most of its lines read that way), in each the lines top to
 * bottom, and lines the engine split along one row (the end of a line set
 * apart as a block of its own) joined again, each line's words in its own
 * direction. The engine's own block order can put such a piece, or a
 * column, after the rest of the page; geometry decides here (layout.js).
 * @param {{ box, words: object[] }[]} found
 * @returns {{ column, box, words, direction }[]}
 */
export function readingOrder(found) {
  if (!found.length) return [];
  const lineDir = (l) => directionOf(joinWords(l.words.map((w) => w.text)));
  const rtlPage = found.filter((l) => lineDir(l) === 'rtl').length > found.length / 2;
  const cols = columns(found.map((l) => ({ text: '', box: l.box })));
  if (rtlPage) cols.reverse();
  const inCol = cols.map(() => []);
  for (const l of found) {
    const cx = l.box.x + l.box.width / 2;
    let best = 0, bestGap = Infinity;
    cols.forEach((c, i) => {
      const gap = cx < c.x0 ? c.x0 - cx : cx > c.x1 ? cx - c.x1 : 0;
      if (gap < bestGap) { bestGap = gap; best = i; }
    });
    inCol[best].push(l);
  }
  const out = [];
  inCol.forEach((ls, column) => {
    ls.sort((p, q) => p.box.y - q.box.y);
    const rows = [];
    for (const l of ls) {
      const last = rows[rows.length - 1];
      const overlap = last ? (Math.min(last.box.y + last.box.height, l.box.y + l.box.height) - Math.max(last.box.y, l.box.y)) / Math.max(1, Math.min(last.box.height, l.box.height)) : 0;
      const apart = last && (l.box.x >= last.box.x + last.box.width - 2 || l.box.x + l.box.width <= last.box.x + 2);
      if (last && overlap > 0.5 && apart) {
        last.parts.push(l);
        const x0 = Math.min(last.box.x, l.box.x), y0 = Math.min(last.box.y, l.box.y);
        const x1 = Math.max(last.box.x + last.box.width, l.box.x + l.box.width), y1 = Math.max(last.box.y + last.box.height, l.box.y + l.box.height);
        last.box = { x: x0, y: y0, width: x1 - x0, height: y1 - y0 };
      } else rows.push({ box: { ...l.box }, parts: [l] });
    }
    for (const row of rows) {
      const rtl = directionOf(joinWords(row.parts.flatMap((p) => p.words.map((w) => w.text)))) === 'rtl';
      // Pieces of one row in its direction; each piece keeps the engine's own order of its words.
      const parts = [...row.parts].sort((p, q) => (rtl ? q.box.x - p.box.x : p.box.x - q.box.x));
      out.push({ column, box: row.box, words: parts.flatMap((p) => p.words), direction: rtl ? 'rtl' : 'ltr' });
    }
  });
  return out;
}

/**
 * @param {object} json the engine's GetJSONText answer
 * @param {{ width: number, height: number, languages: string[], angle?: number,
 *   map?: (box) => box }} page the picture's size, and how a box in what the
 *   engine read maps back onto the picture (after a deskew or a scale)
 */
export function fromEngine(json, { width, height, languages, angle = 0, map = (b) => b }) {
  const clamp = (b) => {
    const x = Math.max(0, Math.min(width, Math.round(b.x)));
    const y = Math.max(0, Math.min(height, Math.round(b.y)));
    return { x, y, width: Math.max(0, Math.min(width, Math.round(b.x + b.width)) - x), height: Math.max(0, Math.min(height, Math.round(b.y + b.height)) - y) };
  };
  const arabicScript = languages.some((l) => l === 'urd' || l === 'ara');
  const found = [];
  for (const block of json?.blocks || []) {
    for (const para of block.paragraphs || []) {
      for (const line of para.lines || []) {
        const own = [];
        for (const word of line.words || []) {
          const text = String(word.text || '').trim();
          if (text) own.push({ text, box: clamp(map(boxOf(word.bbox))), confidence: Math.round(word.confidence ?? 0) });
        }
        // The models of Arabic script often read its comma, semicolon and question mark as the Latin ones.
        if (arabicScript && own.length && directionOf(joinWords(own.map((w) => w.text))) === 'rtl') {
          for (const w of own) w.text = w.text.replace(/([؀-ۿ])([,;?])$/u, (_, a, p) => a + ARABIC_PUNCTUATION[p]);
        }
        const lineBox = clamp(map(boxOf(line.bbox)));
        // A word the engine boxed as a sliver (seen on Arabic words: the baseline only) takes its line's height.
        for (const w of own) {
          if (w.box.height < lineBox.height * 0.4) w.box = { x: w.box.x, y: lineBox.y, width: w.box.width, height: lineBox.height };
        }
        const kept = own.filter((w) => w.box.width > 0 && w.box.height > 0);
        if (kept.length) found.push({ box: lineBox, words: kept });
      }
    }
  }
  const words = [];
  const lines = [];
  const columnText = [];
  for (const line of readingOrder(found)) {
    const first = words.length;
    for (const w of line.words) words.push({ ...w, line: lines.length });
    const text = joinWords(line.words.map((w) => w.text));
    lines.push({
      text,
      box: line.box,
      confidence: Math.round(line.words.reduce((n, w) => n + w.confidence, 0) / line.words.length),
      direction: line.direction,
      block: line.column,
      words: [first, words.length],
    });
    (columnText[line.column] ||= []).push(text);
  }
  return { width, height, languages, angle, text: columnText.filter(Boolean).map((c) => c.join('\n')).join('\n\n'), lines, words };
}
