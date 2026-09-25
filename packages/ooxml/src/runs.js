/**
 * Runs — the unit of character formatting, and the text inside one.
 *
 * Split out of document.js so the table reader can use them without the two
 * files importing each other. Nothing here knows about paragraphs, tables or
 * packages: it is the lowest layer of the WordprocessingML reader.
 */
import { esc, attrs } from './package.js';
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
  const re = /<w:t\b[^>]*?(?:\/>|>([\s\S]*?)<\/w:t>)|<w:tab\s*\/>|<w:(?:br|cr)\b[^>]*\/>|<w:(?:footnote|endnote)Reference\b[^>]*\/>/g;
  for (const m of String(xmlFragment).matchAll(re)) {
    const tag = m[0];
    if (tag.startsWith('<w:tab')) out += '\t';
    else if (tag.startsWith('<w:t')) out += unesc(m[1] ?? '');
    // A reference is told apart by its OWN tag, not by whether the words it
    // sits beside happen to contain "Reference" — a run reading "See the
    // Reference guide" (or Word's own "Error! Reference source not found.")
    // is not a footnote, and `tag` here is the WHOLE match, words and all,
    // so a substring check on it read that word as the element.
    else if (/^<w:(?:footnote|endnote)Reference\b/.test(tag)) out += NOTE_MARK;
    else out += '\n';
  }
  return out;
}

/**
 * A footnote or endnote REFERENCE is one character of the text, as Word
 * counts it: U+FFFC, the object-replacement character. So a caret steps over
 * it, a selection can delete it, and `text` and `runs` agree to the letter —
 * while the file never sees the character, because `renderRun` writes the
 * reference element back from the run that carries it.
 */
export const NOTE_MARK = '￼';

/** Render a run that carries the given properties verbatim. */
function renderRun(rPrXml, text, run = null) {
  const props = rPrXml ? rPrXml : '';
  // A note reference or mark is an element with no words: written back from
  // what the run carries, never from its text.
  if (run?.noteRef) return '<w:r>' + props + '<w:' + run.noteRef.kind + 'Reference w:id="' + esc(String(run.noteRef.id)) + '"/></w:r>';
  if (run?.noteMark) return '<w:r>' + props + '<w:' + run.noteMark + 'Ref/></w:r>';
  // A tab and a line break are elements in the file, not characters: a
  // literal tab inside <w:t> is something Word tolerates, not something it
  // writes. xml:space="preserve" or Word eats leading and trailing spaces.
  // The reference character never reaches the file either.
  const body = String(text).replace(/￼/g, '').split(/([\t\n])/).map((piece) => {
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
      // The reference is a character; the mark at the head of a note is not
      // (a note's paragraph is never edited through the body's text).
      text: ref ? NOTE_MARK : '',
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
    // An explicit size, in points, when the run sets its own rather than
    // riding its style's — a drop cap's letter always does. Undefined
    // (never null) when the run sets none, so a spread onto a run record
    // never plants the key.
    ...(sizeOf(rPr) !== undefined ? { fontSize: sizeOf(rPr) } : {}),
    ...(link !== null ? { link } : {}),
  };
}

/** A run's own explicit size, in points, off `<w:sz>` — undefined when it sets none. */
function sizeOf(rPr) {
  if (!rPr) return undefined;
  const m = /<w:sz\b[^>]*\bw:val="(\d+)"/.exec(rPr);
  return m ? Number(m[1]) / 2 : undefined;
}

/** Every text run in a fragment, flat — the pre-hyperlink behaviour. */
function flatRuns(fragment, out, link = null) {
  for (const m of String(fragment).matchAll(/<w:r\b(?![a-zA-Z])[^>]*>([\s\S]*?)<\/w:r>/g)) {
    const run = runFromInner(m[1], link);
    if (run) out.push(run);
  }
}

/**
 * `w:ins`/`w:del`'s own attributes, read once for every run the wrapper
 * carries: who made the change and when, and the element's `w:id` (Word's
 * own small per-change counter, not to be confused with a bookmark's).
 */
function parseTrackAttrs(attrsText) {
  const a = attrs(attrsText);
  return { id: a['w:id'] ?? '0', author: a['w:author'] ? unesc(a['w:author']) : '', date: a['w:date'] ?? null };
}

/**
 * Text of a run's content read as a DELETION would write it: `<w:delText>`
 * in place of `<w:t>`, everything else (a tab, a break) the same as `textOf`.
 * Word writes a deleted run's words to `<w:delText>` rather than `<w:t>` —
 * the one difference between a surviving run and a struck-through one.
 */
function textOfDel(xmlFragment) {
  let out = '';
  const re = /<w:delText\b[^>]*?(?:\/>|>([\s\S]*?)<\/w:delText>)|<w:tab\s*\/>|<w:(?:br|cr)\b[^>]*\/>/g;
  for (const m of String(xmlFragment).matchAll(re)) {
    const tag = m[0];
    if (tag.startsWith('<w:tab')) out += '\t';
    else if (tag.startsWith('<w:delText')) out += unesc(m[1] ?? '');
    else out += '\n';
  }
  return out;
}

/**
 * Every run inside a `w:ins`, tagged with the insertion it belongs to — an
 * ordinary run in every other respect, since inserted words are already
 * part of the document's current text (Word counts them, the caret walks
 * through them) and only need marking for the display and for Accept/Reject.
 */
function flatInsRuns(fragment, out, insMeta) {
  const before = out.length;
  flatRuns(fragment, out);
  for (let i = before; i < out.length; i++) out[i] = { ...out[i], ins: insMeta };
}

/**
 * Every run inside a `w:del`, read from its `w:delText`. Unlike an insertion,
 * a deletion is NOT part of the document's current text or its caret's
 * address space — `text` here is always empty, the way `textOf` (which never
 * matches `<w:delText>`) already leaves it out of a paragraph's own `text`.
 * The struck-through words a reviewing pane or All Markup needs to SHOW ride
 * on `del.text` instead, a side channel the offset math never counts.
 */
function flatDelRuns(fragment, out, delMeta) {
  for (const m of String(fragment).matchAll(/<w:r\b(?![a-zA-Z])[^>]*>([\s\S]*?)<\/w:r>/g)) {
    const inner = m[1];
    if (!/<w:delText\b|<w:tab\s*\/>|<w:(?:br|cr)\b/.test(inner)) continue;
    const rPrMatch = RPR_RE.exec(inner);
    out.push({ rPr: rPrMatch ? rPrMatch[0] : null, text: '', del: { ...delMeta, text: textOfDel(inner) } });
  }
}

/**
 * A field's instruction, split the way Word's own fields are: the first
 * word says what kind of field it is; a REF's second word is the bookmark
 * it names, a SEQ's the sequence (label) it counts — `Figure`, `Table`,
 * `Equation`, or a document's own. `PAGE`, `DATE` and the rest carry no
 * name — the editor only ever follows a REF or groups a SEQ by it, so those
 * are the only two kinds worth a second word.
 */
function parseFieldInstr(instr) {
  const words = String(instr).trim().split(/\s+/);
  const kind = (words[0] || '').toLowerCase();
  return { kind, name: (kind === 'ref' || kind === 'seq') ? (words[1] ?? null) : null };
}

/**
 * A COMPLETE complex field — `w:fldChar` begin/separate/end around an
 * `<w:instrText>` — matched as three pieces: the run carrying `begin`, the
 * runs between it and the run carrying `end` (the instruction and, after
 * `separate`, the cached result), and the `end` run itself. `[\s\S]*?`
 * non-greedy between `<w:r>` and its own `</w:r>` keeps one run's fldChar
 * from swallowing its neighbour's; a run with no text between its tags
 * still matches, which is exactly the shape `begin` and `end` runs are.
 * Incomplete (no `separate`, the way a hand-built fixture sometimes is, or
 * a field mid-edit Word never actually saves) is left to the caller to
 * notice and skip — this regex only finds the begin…end SPAN.
 */
const COMPLEX_FIELD_RE = () => /(<w:r\b[^>]*>(?:(?!<\/w:r>)[\s\S])*?<w:fldChar\b[^>]*\bw:fldCharType="begin"[^>]*\/>(?:(?!<\/w:r>)[\s\S])*?<\/w:r>)([\s\S]*?)(<w:r\b[^>]*>(?:(?!<\/w:r>)[\s\S])*?<w:fldChar\b[^>]*\bw:fldCharType="end"[^>]*\/>(?:(?!<\/w:r>)[\s\S])*?<\/w:r>)/g;

/**
 * Visit every complete complex field in a fragment. `visit` gets the field
 * code, the run holding its current cached result (or null for an empty
 * one), and the raw pieces a caller needs to rewrite just the result — and
 * returns either `undefined` (leave this field exactly as it was) or a
 * replacement for the WHOLE matched span. Shared by `foldComplexFields`
 * (turns one into the one-run shape `fieldRunFromFldSimple` already reads)
 * and `mapComplexFieldResults` (rewrites a SEQ's cached number on Update
 * Fields) — one reading of the begin/separate/end shape, two uses of it.
 */
function forEachComplexField(xml, visit) {
  return String(xml).replace(COMPLEX_FIELD_RE(), (whole, begin, middle, end) => {
    const sep = /<w:fldChar\b[^>]*\bw:fldCharType="separate"[^>]*\/>/.exec(middle);
    if (!sep) return whole; // incomplete — no cached result to read or fold
    const instrPart = middle.slice(0, sep.index);
    const instr = [...instrPart.matchAll(/<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g)]
      .map((m) => unesc(m[1])).join('');
    if (!instr) return whole;
    const resultPart = middle.slice(sep.index + sep[0].length);
    const runMatch = /<w:r\b[^>]*>([\s\S]*?)<\/w:r>/.exec(resultPart);
    const replacement = visit({ begin, middle, end, instr, runMatch, resultPart, sepIndex: sep.index, sepTag: sep[0] });
    return replacement === undefined ? whole : replacement;
  });
}

/**
 * A complex field folded to look like a `<w:fldSimple>` — same one-run
 * shape, same `field` record — so a SEQ caption shades and reads exactly
 * as a REF does, with nothing downstream needing to learn a second field
 * shape. Called before anything else in this module sees the paragraph's
 * XML, the way `<w:hyperlink>` and `<w:fldSimple>` already are.
 */
function foldComplexFields(xml) {
  return forEachComplexField(xml, ({ instr, runMatch }) =>
    '<w:fldSimple w:instr="' + esc(instr) + '">' + (runMatch ? runMatch[0] : '<w:r><w:t xml:space="preserve"></w:t></w:r>') + '</w:fldSimple>');
}

/**
 * Update Fields for a SEQ caption: every complex field in a fragment gets
 * its cached result run offered to `mapper(instr, currentText)`; a string
 * back replaces just that run's words (its `w:rPr`, if it had one, rides
 * along), `undefined` or the same text leaves the field untouched. Only the
 * result changes — the instruction, the begin/separate/end markers, and
 * everything else in the paragraph are byte-identical to what was there.
 */
export function mapComplexFieldResults(xml, mapper) {
  return forEachComplexField(xml, ({ begin, middle, end, instr, runMatch, resultPart, sepIndex, sepTag }) => {
    const currentText = runMatch ? textOf(runMatch[1]) : '';
    const next = mapper(instr, currentText);
    if (next === undefined || next === currentText) return undefined;
    let newResultPart;
    if (runMatch) {
      const rPrMatch = RPR_RE.exec(runMatch[1]);
      const newRun = '<w:r>' + (rPrMatch ? rPrMatch[0] : '') + '<w:t xml:space="preserve">' + esc(next) + '</w:t></w:r>';
      newResultPart = resultPart.slice(0, runMatch.index) + newRun + resultPart.slice(runMatch.index + runMatch[0].length);
    } else {
      newResultPart = '<w:r><w:t xml:space="preserve">' + esc(next) + '</w:t></w:r>' + resultPart;
    }
    return begin + middle.slice(0, sepIndex) + sepTag + newResultPart + end;
  });
}

/**
 * A `<w:fldSimple>` — Word's cached-result field — as ONE run. Its `w:instr`
 * attribute is the field code (` REF Summary \h `); its content is the last
 * result Word computed, a run like any other, which is why the run inside is
 * read with `runFromInner` rather than reinvented. A field with no readable
 * run inside (an empty result) still becomes a run, with no text, so an
 * empty REF is not silently dropped from the paragraph.
 */
function fieldRunFromFldSimple(attrsText, inner) {
  const instrMatch = /\bw:instr="([^"]*)"/.exec(attrsText);
  const instr = instrMatch ? unesc(instrMatch[1]) : '';
  const runMatch = /<w:r\b(?![a-zA-Z])[^>]*>([\s\S]*?)<\/w:r>/.exec(inner);
  const inside = runMatch ? runFromInner(runMatch[1]) : null;
  return {
    rPr: inside?.rPr ?? null,
    text: inside?.text ?? '',
    field: { instr, ...parseFieldInstr(instr) },
  };
}

export function parseRuns(paragraphXml) {
  // A complex field is folded to a `<w:fldSimple>`-shaped span BEFORE
  // anything else here reads the paragraph, so a SEQ caption's number is
  // one run carrying `field`, same as a REF's — see `foldComplexFields`.
  const xml = foldComplexFields(String(paragraphXml));
  const runs = [];

  // A `<w:hyperlink>` is a GROUP of runs wearing a target: its runs are as
  // editable as any others, and `link` — the wrapper's attributes, verbatim —
  // is what lets a rebuild put the wrapper back. A `<w:fldSimple>` is the
  // opposite shape — Word's cached-result field collapses to ONE run, its
  // instruction and result carried on `field` — so it is read here rather
  // than flattened into the run(s) it wraps. Everything else at the top
  // level (an sdt's body, a smart tag) contributes its text runs flat, exactly
  // as this function always has; nesting inside those is display-only because
  // paragraphs carrying them are structural and never rebuilt.
  // `w:ins`/`w:del` are read at this same top level, one more group wrapper
  // beside hyperlink and fldSimple — not inside either of those (a tracked
  // change inside a hyperlink or a field is rare enough, and costly enough
  // to get wrong, that it is left to ride through as plain text there, the
  // same stance already taken on nesting hyperlink inside fldSimple or back).
  const re = /<w:hyperlink\b([^>]*)>([\s\S]*?)<\/w:hyperlink>|<w:fldSimple\b([^>]*)>([\s\S]*?)<\/w:fldSimple>|<w:ins\b([^>]*)>([\s\S]*?)<\/w:ins>|<w:del\b([^>]*)>([\s\S]*?)<\/w:del>/g;
  let cursor = 0;
  let m;
  while ((m = re.exec(xml))) {
    flatRuns(xml.slice(cursor, m.index), runs);
    if (m[1] !== undefined) flatRuns(m[2], runs, m[1]);
    else if (m[3] !== undefined) runs.push(fieldRunFromFldSimple(m[3], m[4]));
    else if (m[5] !== undefined) flatInsRuns(m[6], runs, parseTrackAttrs(m[5]));
    else flatDelRuns(m[8], runs, parseTrackAttrs(m[7]));
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

/** `<w:ins>`/`<w:del>` wrapping some inner XML, with an id, an author and a date its own attributes. */
function renderTrackWrap(tag, meta, inner) {
  const id = meta?.id ?? '0';
  const author = ' w:author="' + esc(String(meta?.author ?? '')) + '"';
  const date = meta?.date ? ' w:date="' + esc(String(meta.date)) + '"' : '';
  return '<w:' + tag + ' w:id="' + esc(String(id)) + '"' + author + date + '>' + inner + '</w:' + tag + '>';
}

/**
 * Render a list of runs back to XML, dropping any that ended up empty.
 * Consecutive runs sharing a `link` come back wrapped in one `<w:hyperlink>`
 * carrying exactly the attributes the parse captured — which is what lets a
 * keystroke inside a link leave the link standing. An `ins`/`del` run wraps
 * in `<w:ins>`/`<w:del>` the same way — see `renderTrackWrap`.
 */
export function renderRuns(runs) {
  const out = [];
  let openLink = null;
  // A deletion carries no text of its own (see `flatDelRuns`) — `del.text`
  // is what must survive to `<w:delText>`, so it alone earns an empty run a
  // place in the file.
  for (const r of runs.filter((run) => run.text !== '' || run.noteRef || run.noteMark || run.field || run.del)) {
    const link = r.link ?? null;
    if (link !== openLink) {
      if (openLink !== null) out.push('</w:hyperlink>');
      if (link !== null) out.push('<w:hyperlink' + link + '>');
      openLink = link;
    }
    // A field run wraps its rendered run in the `<w:fldSimple>` its `field`
    // remembers — the cached result Word shows until Update Fields is next
    // pressed, right where parseRuns found it among the paragraph's runs.
    if (r.field) out.push('<w:fldSimple w:instr="' + esc(r.field.instr) + '">' + renderRun(r.rPr, r.text) + '</w:fldSimple>');
    else if (r.del) out.push(renderTrackWrap('del', r.del, '<w:r>' + (r.rPr || '') + '<w:delText xml:space="preserve">' + esc(r.del.text ?? '') + '</w:delText></w:r>'));
    else if (r.ins) out.push(renderTrackWrap('ins', r.ins, renderRun(r.rPr, r.text, r)));
    else out.push(renderRun(r.rPr, r.text, r));
  }
  if (openLink !== null) out.push('</w:hyperlink>');
  return out.join('');
}

export { renderRun, firstRunProps, RPR_RE };
