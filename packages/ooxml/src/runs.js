/**
 * Runs — the unit of character formatting, and the text inside one.
 *
 * Split out of document.js so the table reader can use them without the two
 * files importing each other. Nothing here knows about paragraphs, tables or
 * packages: it is the lowest layer of the WordprocessingML reader.
 */
import { esc } from './package.js';
import { unesc } from './workbook.js';

/** Text of a run/paragraph fragment, honouring w:tab and w:br like Word does. */
export function textOf(xmlFragment) {
  return String(xmlFragment)
    .replace(/<w:tab\b[^>]*\/>/g, '\t')
    .replace(/<w:br\b[^>]*\/>/g, '\n')
    .replace(/<w:cr\b[^>]*\/>/g, '\n')
    .split(/<w:t\b[^>]*>/)
    .slice(1)
    .map((chunk) => unesc(chunk.split('</w:t>')[0]))
    .join('');
}

/** Render a run that carries the given properties verbatim. */
function renderRun(rPrXml, text) {
  const props = rPrXml ? rPrXml : '';
  // xml:space="preserve" or Word eats leading and trailing spaces.
  return '<w:r>' + props + '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r>';
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
  // drawing. It has no text to edit, so it is not offered as one.
  if (!/<w:t\b/.test(inner)) return null;
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
