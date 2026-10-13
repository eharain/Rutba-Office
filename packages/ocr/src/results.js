// What the engine read, as the suite hands it on: the text, its lines and its
// words, each with a box in the picture's own pixels and a confidence from
// 0 to 100.
//
//   { width, height, languages, angle,
//     text,                                    lines joined, a blank line between blocks
//     lines: [{ text, box, confidence, direction, block, words: [first, end) }],
//     words: [{ text, box, confidence, line }] }
//
// A box is { x, y, width, height }. Words keep the engine's reading order;
// a line's text is its words joined by a space, except between two
// characters of Chinese or Japanese, which are written without one.

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

const boxOf = (b) => ({ x: b.x0, y: b.y0, width: Math.max(0, b.x1 - b.x0), height: Math.max(0, b.y1 - b.y0) });

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
  const words = [];
  const lines = [];
  const blocks = [];
  (json?.blocks || []).forEach((block, b) => {
    const inBlock = [];
    for (const para of block.paragraphs || []) {
      for (const line of para.lines || []) {
        const first = words.length;
        for (const word of line.words || []) {
          const text = String(word.text || '').trim();
          if (!text) continue;
          words.push({ text, box: clamp(map(boxOf(word.bbox))), confidence: Math.round(word.confidence ?? 0), line: lines.length });
        }
        if (words.length === first) continue;
        const own = words.slice(first);
        const text = joinWords(own.map((w) => w.text));
        lines.push({
          text,
          box: clamp(map(boxOf(line.bbox))),
          confidence: Math.round(own.reduce((n, w) => n + w.confidence, 0) / own.length),
          direction: directionOf(text),
          block: b,
          words: [first, words.length],
        });
        inBlock.push(text);
      }
    }
    if (inBlock.length) blocks.push(inBlock.join('\n'));
  });
  return { width, height, languages, angle, text: blocks.join('\n\n'), lines, words };
}
