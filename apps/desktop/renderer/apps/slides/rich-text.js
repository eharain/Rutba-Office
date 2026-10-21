// The words of a text box as runs, edited a range at a time (owner,
// 2026-10-21: "the entire text box has single formatting and does not allow
// formatting selected text separately; any change applies to entire text
// box"). A box's text is paragraphs of runs, each run its words and their
// look; these functions change some of the words and leave the rest as they
// were, so a bold word stays bold when the line around it is retyped, and
// Bold on three selected words makes those three bold.
//
// A place in the text is an offset into its plain form: each paragraph's
// runs one after another, a line break inside a paragraph one character
// (\v), and the paragraphs joined by \n. That is the text the box editor
// reports its caret in, and the text Insert Symbol and the thesaurus splice.

/** The plain form of one run: its words, a line break as \v, a field or equation as what it shows. */
export const runPlain = (r) => (r?.break ? '\v' : r?.math ? (r.text || '□') : String(r?.text ?? ''));

/** The plain text of paragraphs, as the editor counts it. */
export const plainOf = (paragraphs) => (paragraphs || []).map((p) => (p.runs || []).map(runPlain).join('')).join('\n');

/** A run that can be cut or joined: plain words, not a field, a break, an equation or a picture. */
const plainRun = (r) => !r.break && !r.math && !r.field && !r.picture;

/** A run's look without its words. */
export function lookOf(r) {
  const { text, field, break: brk, math, picture, ...look } = r || {};
  return look;
}

const sameLook = (a, b) => JSON.stringify(lookOf(a)) === JSON.stringify(lookOf(b));

/** Runs next to each other with the same look made one; empty plain runs dropped, one kept so the paragraph keeps a look. */
export function tidy(runs) {
  const out = [];
  for (const r of runs) {
    if (plainRun(r) && r.text === '') continue;
    const last = out[out.length - 1];
    if (last && plainRun(last) && plainRun(r) && sameLook(last, r) && !last.link === !r.link && JSON.stringify(last.link) === JSON.stringify(r.link)) last.text += r.text;
    else out.push({ ...r });
  }
  if (!out.length) {
    const any = runs.find(plainRun);
    out.push({ ...(any ? lookOf(any) : {}), text: '' });
  }
  return out;
}

/** Where `offset` falls: the paragraph, and the offset inside it. */
function locate(paragraphs, offset) {
  let at = Math.max(0, offset);
  for (let i = 0; i < paragraphs.length; i++) {
    const len = (paragraphs[i].runs || []).map(runPlain).join('').length;
    if (at <= len) return { p: i, at };
    at -= len + 1;
  }
  const last = paragraphs.length - 1;
  return { p: last, at: (paragraphs[last]?.runs || []).map(runPlain).join('').length };
}

/** A paragraph's runs cut at `at` (a plain run is split there), with the index of the first run at or after it. */
function cutAt(runs, at) {
  const out = [];
  let pos = 0;
  let index = null;
  for (const r of runs) {
    const len = runPlain(r).length;
    if (index === null && at <= pos) index = out.length;
    if (index === null && at > pos && at < pos + len && plainRun(r)) {
      out.push({ ...r, text: r.text.slice(0, at - pos) });
      index = out.length;
      out.push({ ...r, text: r.text.slice(at - pos) });
    } else out.push({ ...r });
    pos += len;
  }
  if (index === null) index = out.length;
  return { runs: out, index };
}

/**
 * Every paragraph and run, each with the part of it between `start` and
 * `end` given to `fn` (a run cut at the ends): fn(run, paragraphIndex) -> run.
 */
function mapRange(paragraphs, start, end, fn) {
  const a = locate(paragraphs, Math.min(start, end));
  const b = locate(paragraphs, Math.max(start, end));
  return paragraphs.map((p, i) => {
    if (i < a.p || i > b.p) return p;
    const from = i === a.p ? a.at : 0;
    const to = i === b.p ? b.at : Infinity;
    let runs = p.runs || [];
    const cutEnd = cutAt(runs, to === Infinity ? 1e12 : to);
    runs = cutEnd.runs;
    const cutStart = cutAt(runs, from);
    runs = cutStart.runs;
    const endIndex = to === Infinity ? runs.length : cutAt(runs, to).index;
    return { ...p, runs: tidy(runs.map((r, ri) => (ri >= cutStart.index && ri < endIndex ? fn(r, i) : r))) };
  });
}

/** The runs (with words) between `start` and `end`; the one at the caret when they are equal. */
export function runsIn(paragraphs, start, end) {
  const found = [];
  if (start === end) {
    const { p, at } = locate(paragraphs, start);
    const runs = paragraphs[p]?.runs || [];
    let pos = 0;
    for (const r of runs) {
      const len = runPlain(r).length;
      // The run the caret is in, or the one before it at a boundary, as a word processor reads it.
      if (at > pos && at <= pos + len) return [r];
      pos += len;
    }
    return runs.length ? [runs[0]] : [];
  }
  mapRange(paragraphs, start, end, (r) => { if (runPlain(r).length) found.push(r); return r; });
  return found;
}

/** The look a format delta gives one run. Toggles are resolved by the caller. */
export function formatRun(r, d) {
  const out = { ...r };
  for (const k of ['bold', 'italic', 'underline', 'strike']) if (k in d) out[k] = d[k] || undefined;
  if ('size' in d && d.size != null) out.size = d.size;
  if ('color' in d && d.color) out.color = d.color;
  if ('font' in d) out.font = d.font || undefined;
  if ('highlight' in d) out.highlight = d.highlight || undefined;
  if ('spacing' in d) out.spacing = d.spacing || undefined;
  if ('lang' in d && d.lang) out.lang = d.lang;
  if ('noProof' in d) out.noProof = d.noProof || undefined;
  if (d.case && plainRun(r)) out.text = recase(r.text, d.case);
  for (const k of Object.keys(out)) if (out[k] === undefined) delete out[k];
  return out;
}

/** Change case as PowerPoint's does. */
export function recase(s, mode, first = true) {
  return mode === 'upper' ? s.toUpperCase()
    : mode === 'lower' ? s.toLowerCase()
      : mode === 'title' ? s.replace(/(^|\s)(\S)/g, (m, sp, ch) => sp + ch.toUpperCase())
        : mode === 'sentence' ? (first ? s.charAt(0).toUpperCase() + s.slice(1).toLowerCase() : s.toLowerCase())
          : s;
}

/** A toggle ('toggle') resolved once for the whole range: on unless every run in it already is. */
export function resolveToggles(delta, runs) {
  const out = { ...delta };
  for (const k of ['bold', 'italic', 'underline', 'strike']) {
    if (out[k] === 'toggle') out[k] = !(runs.length && runs.every((r) => r[k]));
  }
  return out;
}

/** Run keys a format delta may carry; the rest belong to paragraphs. */
export const RUN_KEYS = ['bold', 'italic', 'underline', 'strike', 'size', 'color', 'font', 'highlight', 'spacing', 'lang', 'noProof', 'case'];
export const PARAGRAPH_KEYS = ['align', 'bullet', 'level', 'lineHeight', 'rtl'];

/** The look `delta` gives the words between `start` and `end`, and nothing else. */
export function formatRange(paragraphs, start, end, delta) {
  const d = resolveToggles(delta, runsIn(paragraphs, start, end));
  return mapRange(paragraphs, start, end, (r) => formatRun(r, d));
}

/** One paragraph with a paragraph delta applied: bullets, level, spacing, direction, alignment. */
export function formatParagraph(p, arg) {
  const level = 'level' in arg ? Math.max(0, Math.min(8, (p.level || 0) + Number(arg.level || 0))) : p.level;
  const bullet = arg.bullet === 'char' ? { type: 'char', char: '•' }
    : arg.bullet === 'number' ? { type: 'number', scheme: 'arabicPeriod', start: 1 }
      : arg.bullet === 'none' ? { type: 'none' }
        : p.bullet;
  const out = {
    ...p,
    level: level || undefined,
    bullet,
    lineHeight: 'lineHeight' in arg ? arg.lineHeight : p.lineHeight,
    rtl: 'rtl' in arg ? Boolean(arg.rtl) : p.rtl,
    align: arg.align ?? ('rtl' in arg && Boolean(arg.rtl) !== Boolean(p.rtl)
      ? (arg.rtl ? (!p.align || p.align === 'left' ? 'right' : p.align) : (p.align === 'right' ? 'left' : p.align))
      : p.align),
  };
  for (const k of Object.keys(out)) if (out[k] === undefined) delete out[k];
  return out;
}

/** The paragraphs the range touches, given the paragraph part of a delta. */
export function formatParagraphsIn(paragraphs, start, end, arg) {
  const a = locate(paragraphs, Math.min(start, end)).p;
  const b = locate(paragraphs, Math.max(start, end)).p;
  return paragraphs.map((p, i) => (i >= a && i <= b ? formatParagraph(p, arg) : p));
}

/**
 * `text` put in place of the words between `start` and `end`, with the look
 * of the run it is typed into (the run before the caret, as typing reads
 * it, or `look` when given). A \n in it starts a new paragraph, which keeps
 * the paragraph's own settings; a \v is a line break.
 */
export function replaceRange(paragraphs, start, end, text, look = null) {
  const lo = Math.min(start, end);
  const hi = Math.max(start, end);
  const a = locate(paragraphs, lo);
  const b = locate(paragraphs, hi);
  const pa = paragraphs[a.p];
  const pb = paragraphs[b.p];
  const head = cutAt(pa.runs || [], a.at);
  const before = head.runs.slice(0, head.index);
  const tail = cutAt(pb.runs || [], b.at);
  const after = tail.runs.slice(tail.index);
  const typingLook = look ?? lookOf([...before].reverse().find(plainRun) || after.find(plainRun) || {});
  const { text: _t, plain: _p, runs: _r, ...paraProps } = pa;
  const lines = String(text).split('\n');
  const made = lines.map((line, i) => {
    const pieces = line.split('\v');
    const runs = [];
    pieces.forEach((piece, j) => {
      if (j) runs.push({ break: true, text: '\n' });
      if (piece) runs.push({ ...typingLook, text: piece });
    });
    return runs;
  });
  const out = [];
  if (made.length === 1) {
    out.push({ ...paraProps, runs: tidy([...before, ...made[0], ...after]) });
  } else {
    out.push({ ...paraProps, runs: tidy([...before, ...made[0]]) });
    for (let i = 1; i < made.length - 1; i++) out.push({ ...paraProps, runs: tidy(made[i].length ? made[i] : [{ ...typingLook, text: '' }]) });
    const { runs: _rb, plain: _pb, ...lastProps } = pb;
    out.push({ ...(a.p === b.p ? paraProps : lastProps), runs: tidy([...made[made.length - 1], ...after].length ? [...made[made.length - 1], ...after] : [{ ...typingLook, text: '' }]) });
  }
  return [...paragraphs.slice(0, a.p).map(strip), ...out, ...paragraphs.slice(b.p + 1).map(strip)];
}

const strip = (p) => { const { plain, ...rest } = p; return rest; };

/**
 * The paragraphs once their plain text is `next`: the change found as the
 * words between what is the same at the start and at the end, so every run
 * outside it keeps its look. What Insert Symbol and the thesaurus commit.
 */
export function withPlainText(paragraphs, next) {
  const before = plainOf(paragraphs);
  const after = String(next);
  if (before === after) return paragraphs.map(strip);
  let s = 0;
  while (s < before.length && s < after.length && before[s] === after[s]) s++;
  let e = 0;
  while (e < before.length - s && e < after.length - s && before[before.length - 1 - e] === after[after.length - 1 - e]) e++;
  return replaceRange(paragraphs, s, before.length - e, after.slice(s, after.length - e));
}

/** What the ribbon shows for the words between `start` and `end`: on only when every run is. */
export function formatState(paragraphs, start, end, defaults = {}) {
  const runs = runsIn(paragraphs, start, end).filter((r) => !r.break);
  const r = runs[0] || {};
  const all = (k) => runs.length > 0 && runs.every((x) => x[k]);
  const one = (k) => (runs.every((x) => x[k] === r[k]) ? r[k] : null);
  const p = paragraphs[locate(paragraphs, Math.min(start, end)).p] || {};
  return {
    bold: all('bold'), italic: all('italic'), underline: all('underline'), strike: all('strike'),
    size: one('size') || defaults.size || null, color: one('color') || null, font: one('font') || '', highlight: one('highlight') || null, spacing: one('spacing') || 0,
    align: p.align || defaults.align || 'left', bullet: p.bullet?.type || null, level: p.level || 0, lineHeight: p.lineHeight || null, rtl: Boolean(p.rtl ?? defaults.rtl),
  };
}
