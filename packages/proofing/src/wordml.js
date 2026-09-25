// A WordprocessingML paragraph's words as one string, and a way back from a
// span of that string to the `w:t` elements that hold it.
//
// Word splits a paragraph's words across runs wherever the formatting, the
// revision marks or the proofing marks change — sometimes in the middle of
// a word. Replacing one misspelt word in a header, a footnote or a text box
// means finding the characters wherever they fell, rewriting the first
// piece and taking the rest out of the pieces after it, so every run keeps
// its own look.

const decode = (s) => String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n))).replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16))).replace(/&amp;/g, '&');
const encode = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** Every `<w:p>` in a fragment, with its position. Paragraphs inside a text box inside this one are left for their own box. */
export function paragraphsIn(xml) {
  const out = [];
  const re = /<w:p\b[^>]*?(\/?)>|<\/w:p>/g;
  let depth = 0;
  let start = -1;
  let m;
  while ((m = re.exec(xml))) {
    if (m[0] === '</w:p>') {
      depth -= 1;
      if (depth === 0 && start >= 0) {
        const end = m.index + m[0].length;
        out.push({ start, end, xml: xml.slice(start, end) });
        start = -1;
      }
      depth = Math.max(0, depth);
    } else if (m[1] === '/') {
      if (depth === 0) out.push({ start: m.index, end: m.index + m[0].length, xml: m[0] });
    } else {
      if (depth === 0) start = m.index;
      depth += 1;
    }
  }
  return out;
}

/** The paragraphs of each text box in a fragment (the `mc:Choice` only — the VML fallback repeats them). */
export function textBoxesIn(xml) {
  const out = [];
  const src = String(xml);
  const fallbacks = [...src.matchAll(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g)].map((m) => [m.index, m.index + m[0].length]);
  for (const m of src.matchAll(/<w:txbxContent\b[^>]*>([\s\S]*?)<\/w:txbxContent>/g)) {
    if (fallbacks.some(([a, b]) => m.index >= a && m.index < b)) continue;
    const innerStart = m.index + m[0].indexOf('>') + 1;
    out.push({ start: innerStart, xml: m[1], paragraphs: paragraphsIn(m[1]) });
  }
  return out;
}

/**
 * The paragraph's text and the pieces it came from. Deleted text (a tracked
 * deletion's `w:delText`) and field codes are not text; a tab is a tab and a
 * break a line break, so no two words run together.
 */
export function paragraphText(pXml) {
  const xml = String(pXml).replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, (m) => ' '.repeat(m.length));
  const pieces = [];
  let text = '';
  const re = /<w:t\b[^>]*\/>|<w:t\b([^>]*)>([\s\S]*?)<\/w:t>|<w:tab\b[^>]*\/>|<w:br\b[^>]*\/>|<w:cr\b[^>]*\/>|<w:txbxContent\b[\s\S]*?<\/w:txbxContent>/g;
  let m;
  while ((m = re.exec(xml))) {
    const tag = m[0];
    if (tag.startsWith('<w:txbxContent')) continue;
    if (tag.startsWith('<w:tab')) { text += '\t'; continue; }
    if (tag.startsWith('<w:br') || tag.startsWith('<w:cr')) { text += '\n'; continue; }
    if (m[2] === undefined) continue;
    const decoded = decode(m[2]);
    const contentStart = m.index + tag.indexOf('>') + 1;
    pieces.push({ from: text.length, to: text.length + decoded.length, decoded, openStart: m.index, contentStart, contentEnd: contentStart + m[2].length });
    text += decoded;
  }
  return { text, pieces };
}

/** The paragraph with the characters `[from, to)` replaced by `replacement`. */
export function replaceInParagraph(pXml, from, to, replacement) {
  const { pieces } = paragraphText(pXml);
  let xml = String(pXml);
  let placed = false;
  const touched = pieces.filter((p) => p.to > from && p.from < to);
  // Rewritten from the last piece back, so earlier offsets hold.
  const edits = touched.map((p, i) => {
    const a = Math.max(0, from - p.from);
    const b = Math.min(p.decoded.length, to - p.from);
    const insert = i === 0 ? String(replacement) : '';
    if (i === 0) placed = true;
    return { p, next: p.decoded.slice(0, a) + insert + p.decoded.slice(b) };
  });
  if (!placed) {
    // A span in no piece — an empty paragraph asked to take a word: nothing to do.
    return xml;
  }
  for (const { p, next } of edits.reverse()) {
    const open = xml.slice(p.openStart, p.contentStart);
    const needsSpace = /^\s|\s$/.test(next) && !/xml:space="preserve"/.test(open);
    const newOpen = needsSpace ? open.replace(/^<w:t\b/, '<w:t xml:space="preserve"') : open;
    xml = xml.slice(0, p.openStart) + newOpen + encode(next) + xml.slice(p.contentEnd);
  }
  return xml;
}
