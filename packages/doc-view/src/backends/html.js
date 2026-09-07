/**
 * HTML backend — an email body.
 *
 * The same editor that types into a `.docx` types into this, which is the point
 * of the port. Mail composes with `@rutba/doc-view` + this backend and never
 * loads the OOXML layer.
 *
 * Email HTML is not web HTML, and the differences are the reason this is a
 * backend rather than "just innerHTML":
 *
 *   - **Inline styles only.** Gmail strips `<style>` blocks in several contexts
 *     and Outlook ignores most of a stylesheet, so formatting has to ride on the
 *     element. `rPr` is therefore a small canonical token set rather than a class.
 *   - **A conservative tag set.** `<b>`, `<i>`, `<u>`, `<p>`, `<span>` render
 *     everywhere. Semantic `<strong>`/`<em>` are fine too, but the presentational
 *     tags are what every client agrees on.
 *   - **Nothing is trusted on the way in.** Parsing an email body means parsing
 *     something an attacker may have written, so `<script>`, event handlers and
 *     `javascript:` never survive the round trip.
 *
 * `rPr` is a sorted token string — `'b|i'` — so the editor's coalesce, which
 * compares `rPr` by equality, merges runs correctly with no format knowledge.
 */

import { DocView } from '../view.js';

const TOKENS = { b: 'b', i: 'i', u: 'u', s: 's' };

const parseTokens = (rPr) => new Set(String(rPr ?? '').split('|').filter(Boolean));
const formatTokens = (set) => [...set].sort().join('') === '' ? null : [...set].sort().join('|');

const escapeHtml = (s) =>
  String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
const unescapeHtml = (s) =>
  String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'").replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&');

/** Strip anything that could execute. An email body is untrusted input. */
function sanitise(html) {
  return String(html)
    .replace(/<script\b[\s\S]*?<\/script>/gi, '')
    .replace(/<style\b[\s\S]*?<\/style>/gi, '')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/\son\w+\s*=\s*"[^"]*"/gi, '')
    .replace(/\son\w+\s*=\s*'[^']*'/gi, '')
    .replace(/javascript:/gi, '');
}

export class HtmlBackend {
  /**
   * @param {string} html  an email body; a bare string becomes one paragraph
   */
  constructor(html = '') {
    this.blocks = HtmlBackend.parse(html);
    if (!this.blocks.length) this.blocks = [{ style: null, runs: [] }];
    this.dirty = false;
  }

  static open(html) { return new HtmlBackend(html); }

  /** Parse a body into blocks of runs. Unknown tags contribute their text only. */
  static parse(html) {
    const clean = sanitise(html ?? '');
    const blocks = [];
    const blockRe = /<(p|div|h1|h2|h3|li)\b[^>]*>([\s\S]*?)<\/\1>/gi;
    let matched = false;
    let m;
    while ((m = blockRe.exec(clean))) {
      matched = true;
      blocks.push({
        style: /^h[123]$/i.test(m[1]) ? 'Heading' + m[1].slice(1) : (m[1].toLowerCase() === 'li' ? 'ListParagraph' : null),
        runs: HtmlBackend.parseRuns(m[2]),
      });
    }
    if (!matched && clean.trim()) blocks.push({ style: null, runs: HtmlBackend.parseRuns(clean) });
    return blocks;
  }

  /** Inline markup -> runs. Nesting accumulates, so `<b><i>x` is both. */
  static parseRuns(html) {
    const runs = [];
    const active = new Set();
    let link = null;
    const re = /<\/?([a-z0-9]+)\b[^>]*>|([^<]+)/gi;
    let m;
    while ((m = re.exec(html))) {
      if (m[2] !== undefined) {
        const text = unescapeHtml(m[2]).replace(/\s*\n\s*/g, ' ');
        if (text) runs.push({ rPr: formatTokens(active), text, ...(link ? { link } : {}) });
        continue;
      }
      const tag = m[1].toLowerCase();
      const closing = m[0].startsWith('</');
      if (tag === 'a') {
        // The href IS the link token here. Only web-safe schemes survive —
        // an email body is untrusted input.
        if (closing) { link = null; continue; }
        const href = /\bhref\s*=\s*"([^"]*)"|\bhref\s*=\s*'([^']*)'/i.exec(m[0]);
        const target = unescapeHtml(href?.[1] ?? href?.[2] ?? '');
        link = /^(https?:\/\/|mailto:)/i.test(target) ? target : null;
        continue;
      }
      const token = { b: 'b', strong: 'b', i: 'i', em: 'i', u: 'u', s: 's', strike: 's', del: 's' }[tag];
      if (!token) continue;
      if (closing) active.delete(token);
      else active.add(token);
    }
    return runs;
  }

  // ---- the port ----------------------------------------------------------

  paragraphCount() { return this.blocks.length; }

  paragraph(index) {
    const b = this.blocks[index];
    if (!b) return null;
    const runs = b.runs.map((r) => {
      const tokens = parseTokens(r.rPr);
      return {
        rPr: r.rPr ?? null,
        text: r.text,
        bold: tokens.has('b'),
        italic: tokens.has('i'),
        underline: tokens.has('u'),
        strike: tokens.has('s'),
        ...(r.link ? { link: r.link } : {}),
      };
    });
    return {
      index,
      text: runs.map((r) => r.text).join(''),
      style: b.style,
      runs,
      // An email body has no field codes or content controls to protect, so
      // every paragraph is freely editable.
      structural: false,
      structuralTags: [],
    };
  }

  setParagraphRuns(index, runs) {
    if (!this.blocks[index]) throw new Error('no paragraph at index ' + index);
    this.blocks[index].runs = runs
      .filter((r) => r.text !== '')
      .map((r) => ({ rPr: r.rPr ?? null, text: r.text, ...(r.link ? { link: r.link } : {}) }));
    this.dirty = true;
    return this;
  }

  splitParagraph(index, runIndex, offset) {
    const b = this.blocks[index];
    if (!b) throw new Error('no paragraph at index ' + index);
    const before = [];
    const after = [];
    b.runs.forEach((run, i) => {
      if (i < runIndex) before.push(run);
      else if (i > runIndex) after.push(run);
      else {
        before.push({ ...run, text: run.text.slice(0, offset) });
        after.push({ ...run, text: run.text.slice(offset) });
      }
    });
    b.runs = before.filter((r) => r.text !== '');
    this.blocks.splice(index + 1, 0, { style: b.style, runs: after.filter((r) => r.text !== '') });
    this.dirty = true;
    return this;
  }

  mergeWithNext(index) {
    const b = this.blocks[index];
    const next = this.blocks[index + 1];
    if (!b || !next) throw new Error('nothing to merge at index ' + index);
    b.runs = [...b.runs, ...next.runs];
    this.blocks.splice(index + 1, 1);
    this.dirty = true;
    return this;
  }

  removeParagraph(index) {
    if (this.blocks.length <= 1) throw new Error('a body must keep at least one paragraph');
    this.blocks.splice(index, 1);
    this.dirty = true;
    return this;
  }

  insertParagraphAfter(index, runs = [], { inheritStyle = true } = {}) {
    const b = this.blocks[index];
    this.blocks.splice(index + 1, 0, {
      style: inheritStyle ? (b?.style ?? null) : null,
      runs: runs.filter((r) => r.text !== ''),
    });
    this.dirty = true;
    return this;
  }

  toggleRunFormat(rPr, tag, on) {
    const token = TOKENS[tag];
    if (!token) throw new Error('unknown format: ' + tag);
    const set = parseTokens(rPr);
    if (on) set.add(token);
    else set.delete(token);
    return formatTokens(set);
  }

  /** Mint (or resolve) a link token — for email HTML the href IS the token. */
  makeLink(url) {
    const target = String(url);
    if (!/^(https?:\/\/|mailto:)/i.test(target)) {
      throw new Error('a link must start http://, https:// or mailto: — got ' + target);
    }
    return target;
  }

  linkTarget(token) { return token; }

  /** Email HTML: inline, conservative tags, escaped text. */
  save() {
    return this.blocks.map((b) => {
      const pieces = [];
      let openLink = null;
      for (const run of b.runs) {
        let html = escapeHtml(run.text);
        const tokens = parseTokens(run.rPr);
        if (tokens.has('s')) html = '<s>' + html + '</s>';
        if (tokens.has('u')) html = '<u>' + html + '</u>';
        if (tokens.has('i')) html = '<i>' + html + '</i>';
        if (tokens.has('b')) html = '<b>' + html + '</b>';
        const link = run.link ?? null;
        if (link !== openLink) {
          if (openLink !== null) pieces.push('</a>');
          if (link !== null) pieces.push('<a href="' + escapeHtml(link) + '">');
          openLink = link;
        }
        pieces.push(html);
      }
      if (openLink !== null) pieces.push('</a>');
      const inner = pieces.join('');
      if (b.style === 'ListParagraph') return '<li>' + inner + '</li>';
      if (b.style && /^Heading[123]$/.test(b.style)) {
        const level = b.style.slice(-1);
        return '<h' + level + '>' + inner + '</h' + level + '>';
      }
      return '<p>' + (inner || '<br>') + '</p>';
    }).join('');
  }

  /** Plain-text alternative — every email should carry one. */
  toPlainText() {
    return this.blocks.map((b) => b.runs.map((r) => r.text).join('')).join('\n');
  }

  modifiedParts() { return this.dirty ? ['body'] : []; }
}

export { sanitise, escapeHtml };

/** Open an email body for editing. The Mail entry point. */
export const openHtml = (body) => new DocView(HtmlBackend.open(body));
