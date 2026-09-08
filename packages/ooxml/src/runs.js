/**
 * Runs — the unit of character formatting, and the text inside one.
 *
 * Split out of document.js so the table reader can use them without the two
 * files importing each other. Nothing here knows about paragraphs, tables or
 * packages: it is the lowest layer of the WordprocessingML reader.
 */
import { esc } from './package.js';
import { unesc } from './workbook.js';

/**
 * Text of a run/paragraph fragment, honouring w:tab and w:br like Word does.
 *
 * Walked in document order: `<w:t>` gives its words, a bare `<w:tab/>` a tab,
 * `<w:br/>`/`<w:cr/>` a line break. The first version substituted the tab and
 * break characters into the XML and then kept only what sat inside `<w:t>` —
 * which a run's `<w:tab/>` never does — so no tab ever reached the text, and
 * every form line in the corpus ran its label into its value. Only a BARE
 * `<w:tab/>` is a tab: the `<w:tab w:val w:pos/>` in a paragraph's `<w:tabs>`
 * is a tab STOP, and counting it put phantom tabs at the head of every
 * paragraph that defined one.
 */
export function textOf(xmlFragment) {
  let out = '';
  const re = /<w:t\b[^>]*?(?:\/>|>([\s\S]*?)<\/w:t>)|<w:tab\s*\/>|<w:(?:br|cr)\b[^>]*\/>/g;
  for (const m of String(xmlFragment).matchAll(re)) {
    const tag = m[0];
    if (tag.startsWith('<w:tab')) out += '\t';
    else if (tag.startsWith('<w:t')) out += unesc(m[1] ?? '');
    else out += '\n';
  }
  return out;
}

/** Render a run that carries the given properties verbatim. */
function renderRun(rPrXml, text) {
  const props = rPrXml ? rPrXml : '';
  // A tab and a line break are elements in the file, not characters: a
  // literal tab inside <w:t> is something Word tolerates, not something it
  // writes. xml:space="preserve" or Word eats leading and trailing spaces.
  const body = String(text).split(/([\t\n])/).map((piece) => {
    if (piece === '\t') return '<w:tab/>';
    if (piece === '\n') return '<w:br/>';
    return piece ? '<w:t xml:space="preserve">' + esc(piece) + '</w:t>' : '';
  }).join('');
  return '<w:r>' + props + body + '</w:r>';
}

/** First run's <w:rPr>…</w:rPr>, so an edit inherits the original formatting. */
function firstRunProps(fragment) {
  const run = /<w:r\b[^>]*>([\s\S]*?)<\/w:r>/.exec(fragment);
  if (!run) return null;
  const rPr = /<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>|<w:rPr\b[^>]*\/>/.exec(run[1]);
  return rPr ? rPr[0] : null;
}

const RPR_RE = /<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>|<w:rPr\b[^>]*\/>/;

/**
 * Split a paragraph into RUNS.
 *
 * A run is the unit of character formatting: `<w:r><w:rPr><w:b/></w:rPr><w:t>Hello</w:t></w:r>`.
 * A sentence with one bold word is three runs, and typing in the middle of it
 * must splice into the right one and leave the others exactly as they were —
 * which is why `rPr` is carried verbatim rather than parsed and regenerated. We
 * understand `w:b`, `w:i`, `w:u` and `w:strike` well enough to toggle them;
 * everything else in `rPr` (fonts, colours, spacing, language, effects) rides
 * along untouched.
 *
 * @returns {Array<{rPr: string|null, text: string, bold: boolean, italic: boolean, underline: boolean, strike: boolean}>}
 */
/** One `<w:r>` inner content -> a run record, or null when it has no text. */
function runFromInner(inner, link = null) {
  // A run with no <w:t> carries something else — a break, a comment anchor, a
  // drawing. It has no text to edit, so it is not offered as one. The two
  // exceptions are a footnote or endnote REFERENCE (the little number in the
  // body) and the matching mark at the head of the note: they are offered as
  // empty runs carrying `noteRef`/`noteMark`, so the page can draw the number
  // in place. Text-free, so no caret offset moves; the paragraph is
  // structural, so no rebuild ever drops them.
  // A run holding only a tab or a line break IS text — Word writes the tab
  // between a label and its value as a run of its own, and dropping it ran
  // every form line's label into its value and left `text` and `runs`
  // disagreeing about where the caret was.
  if (!/<w:t\b|<w:tab\s*\/>|<w:(?:br|cr)\b/.test(inner)) {
    const ref = /<w:(footnote|endnote)Reference\b[^>]*\bw:id="([^"]+)"/.exec(inner);
    const mark = /<w:(footnote|endnote)Ref\b/.exec(inner);
    if (!ref && !mark) return null;
    const rPrMatch = RPR_RE.exec(inner);
    return {
      rPr: rPrMatch ? rPrMatch[0] : null,
      text: '',
      bold: false, italic: false, underline: false, strike: false,
      ...(ref ? { noteRef: { kind: ref[1], id: ref[2] } } : { noteMark: mark[1] }),
      ...(link !== null ? { link } : {}),
    };
  }
  const rPrMatch = RPR_RE.exec(inner);
  const rPr = rPrMatch ? rPrMatch[0] : null;
  return {
    rPr,
    text: textOf(inner),
    bold: hasToggle(rPr, 'b'),
    italic: hasToggle(rPr, 'i'),
    underline: hasToggle(rPr, 'u'),
    // `w:strike` never matches `w:dstrike` — the regex wants the exact tag.
    strike: hasToggle(rPr, 'strike'),
    ...(link !== null ? { link } : {}),
  };
}

/** Every text run in a fragment, flat — the pre-hyperlink behaviour. */
function flatRuns(fragment, out, link = null) {
  for (const m of String(fragment).matchAll(/<w:r\b(?![a-zA-Z])[^>]*>([\s\S]*?)<\/w:r>/g)) {
    const run = runFromInner(m[1], link);
    if (run) out.push(run);
  }
}

export function parseRuns(paragraphXml) {
  const xml = String(paragraphXml);
  const runs = [];

  // A `<w:hyperlink>` is a GROUP of runs wearing a target: its runs are as
  // editable as any others, and `link` — the wrapper's attributes, verbatim —
  // is what lets a rebuild put the wrapper back. Everything else at the top
  // level (an sdt's body, a smart tag) contributes its text runs flat, exactly
  // as this function always has; nesting inside those is display-only because
  // paragraphs carrying them are structural and never rebuilt.
  const re = /<w:hyperlink\b([^>]*)>([\s\S]*?)<\/w:hyperlink>/g;
  let cursor = 0;
  let m;
  while ((m = re.exec(xml))) {
    flatRuns(xml.slice(cursor, m.index), runs);
    flatRuns(m[2], runs, m[1]);
    cursor = m.index + m[0].length;
  }
  flatRuns(xml.slice(cursor), runs);
  return runs;
}

/** `<w:b/>` and `<w:b w:val="1"/>` are on; `<w:b w:val="0"/>` is off. */
export function hasToggle(rPr, tag) {
  if (!rPr) return false;
  const m = new RegExp('<w:' + tag + '\\b([^>]*)/?>').exec(rPr);
  if (!m) return false;
  const val = /w:val="([^"]*)"/.exec(m[1]);
  if (!val) return true;
  return !['0', 'false', 'off'].includes(val[1].toLowerCase());
}

/** Add or remove a toggle inside an rPr, preserving everything else in it. */
export function withToggle(rPr, tag, on) {
  const existing = new RegExp('<w:' + tag + '\\b[^>]*/?>', 'g');
  if (!on) {
    if (!rPr) return null;
    const stripped = rPr.replace(existing, '');
    return /<w:rPr\b[^>]*>\s*<\/w:rPr>/.test(stripped) ? null : stripped;
  }
  if (hasToggle(rPr, tag)) return rPr;
  const element = tag === 'u' ? '<w:u w:val="single"/>' : '<w:' + tag + '/>';
  if (!rPr) return '<w:rPr>' + element + '</w:rPr>';
  if (/<w:rPr\b[^>]*\/>/.test(rPr)) return '<w:rPr>' + element + '</w:rPr>';
  // Order matters to Word: toggles belong near the front of rPr.
  return rPr.replace(/^(<w:rPr\b[^>]*>)/, '$1' + element);
}

/**
 * Render a list of runs back to XML, dropping any that ended up empty.
 * Consecutive runs sharing a `link` come back wrapped in one `<w:hyperlink>`
 * carrying exactly the attributes the parse captured — which is what lets a
 * keystroke inside a link leave the link standing.
 */
export function renderRuns(runs) {
  const out = [];
  let openLink = null;
  for (const r of runs.filter((run) => run.text !== '')) {
    const link = r.link ?? null;
    if (link !== openLink) {
      if (openLink !== null) out.push('</w:hyperlink>');
      if (link !== null) out.push('<w:hyperlink' + link + '>');
      openLink = link;
    }
    out.push(renderRun(r.rPr, r.text));
  }
  if (openLink !== null) out.push('</w:hyperlink>');
  return out.join('');
}

export { renderRun, firstRunProps, RPR_RE };
