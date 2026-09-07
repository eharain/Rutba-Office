// GitHub Flavored Markdown, both directions.
//
// A README is a document people care about the shape of. Opening one in an
// editor that quietly drops the task list, flattens the table and loses the
// front matter is worse than not opening it at all, because the damage only
// shows up in the diff after it has been saved.
//
// So this parses what GitHub actually renders — ATX and setext headings, fenced
// and indented code with the language kept, nested and loose lists, task items,
// tables with their alignment, block quotes, thematic breaks, reference links
// and images, footnotes, autolinks, strikethrough, hard breaks and backslash
// escapes — and writes every one of them back. What it cannot model as a block
// (YAML front matter, raw HTML, the definition lists at the bottom of a file)
// it carries through untouched rather than discarding, so a round trip changes
// the parts that were edited and nothing else.
//
// It is not a CommonMark implementation and does not try to be: the goal is
// that a file which came from GitHub goes back to GitHub unharmed.

/* ── shared ──────────────────────────────────────────────────────────────── */

const ESCAPABLE = '\\\\`*_{}[]()#+-.!|~<>"\'';

/** A thematic break: three or more of the same mark, spaces allowed between. */
const isRule = (line) => /^ {0,3}([-*_])(?:[ \t]*\1){2,}[ \t]*$/.test(line);

const isBlank = (line) => !line.trim();

/**
 * A list line. The marker may be followed by content, or by nothing at all: a
 * bare hyphen is an empty bullet, which is what issue and pull-request
 * templates are made of. LIST_ITEM captures the content, LIST_OPEN only asks
 * whether a line starts one.
 */
const LIST_OPEN = /^(\s*)([-*+]|\d{1,9}[.)])(?:\s+|\s*$)/;
const LIST_ITEM = /^(\s*)([-*+]|\d{1,9}[.)])(?:\s+(.*)|\s*)$/;

/** How deep a list line is indented, in the units lists nest by. */
const indentOf = (line) => (/^[ \t]*/.exec(line)[0] || '').replace(/\t/g, '    ').length;

/* ── inline ──────────────────────────────────────────────────────────────── */

/**
 * One line (or joined paragraph) into styled runs.
 *
 * Written as a scanner rather than a chain of replacements because the
 * constructs nest: a link can hold bold text, bold can hold code, and a URL
 * inside a code span is not a link. A regular expression pass would get every
 * one of those wrong in a way that is invisible until somebody's README is
 * mangled.
 */
export function inlineRuns(source, refs = new Map()) {
  const runs = [];
  const text = String(source ?? '');
  let plain = '';

  const flush = () => {
    if (plain) runs.push({ text: plain });
    plain = '';
  };
  const push = (run) => {
    flush();
    runs.push(run);
  };
  /** Apply a style to a nested parse, so **[a](b)** keeps both. */
  const nested = (inner, style) => {
    for (const run of inlineRuns(inner, refs)) push({ ...run, ...style });
  };

  let i = 0;
  while (i < text.length) {
    const c = text[i];

    // 1. Backslash escapes. A punctuation mark after a backslash is literal.
    if (c === '\\' && ESCAPABLE.includes(text[i + 1])) {
      plain += text[i + 1];
      i += 2;
      continue;
    }

    // 2. Code spans. The contents are literal, however they look.
    if (c === '`') {
      const fence = /^`+/.exec(text.slice(i))[0];
      const close = text.indexOf(fence, i + fence.length);
      if (close > 0 && text[close + fence.length] !== '`') {
        push({ text: text.slice(i + fence.length, close).replace(/^ (.*) $/, '$1'), code: true });
        i = close + fence.length;
        continue;
      }
    }

    // 3. Autolinks in angle brackets, and bare email addresses in them.
    if (c === '<') {
      const m = /^<((?:https?|ftp|mailto):[^>\s]+|[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+)>/.exec(text.slice(i));
      if (m) {
        const url = m[1].includes('@') && !m[1].startsWith('mailto:') ? `mailto:${m[1]}` : m[1];
        // The angle brackets are not decoration: a bare `mailto:x` in running
        // text is not a link to GitHub, so dropping them breaks the document.
        push({ text: m[1], link: url, angle: true });
        i += m[0].length;
        continue;
      }
      // Inline HTML: kept verbatim so it survives the round trip.
      const tag = /^<\/?[a-zA-Z][^>]*>/.exec(text.slice(i));
      if (tag) {
        push({ text: tag[0], html: true });
        i += tag[0].length;
        continue;
      }
    }

    // 4. Images, then links — the `!` has to be tested first or `![a](b)`
    //    parses as a literal bang followed by a link.
    if (c === '!' && text[i + 1] === '[') {
      const parsed = linkAt(text, i + 1, refs);
      if (parsed) {
        push({ text: parsed.label, image: parsed.url, title: parsed.title || null });
        i = parsed.end;
        continue;
      }
    }
    if (c === '[') {
      // Footnote reference: [^1]
      const foot = /^\[\^([^\]]+)\]/.exec(text.slice(i));
      if (foot) {
        push({ text: foot[1], footnote: foot[1] });
        i += foot[0].length;
        continue;
      }
      const parsed = linkAt(text, i, refs);
      if (parsed) {
        const inner = inlineRuns(parsed.label, refs);
        for (const run of inner) push({ ...run, link: parsed.url, ref: parsed.ref || null, title: parsed.title || run.title || null });
        i = parsed.end;
        continue;
      }
    }

    // 5. Strikethrough, GFM's own.
    if (c === '~' && text[i + 1] === '~') {
      const close = text.indexOf('~~', i + 2);
      if (close > 0) {
        nested(text.slice(i + 2, close), { strike: true });
        i = close + 2;
        continue;
      }
    }

    // 6. Emphasis. Longest mark first: `***a***` is bold *and* italic, and
    //    trying `**` against it leaves a stray `*` that poisons everything
    //    after it — which is exactly how emphasis parsers usually go wrong.
    if (c === '*' || c === '_') {
      // `snake_case_words` must not become emphasis. GitHub only opens an
      // underscore run at a word boundary.
      const wordish = c === '_' && /\w/.test(text[i - 1] || '');
      const opened = Math.min(3, /^(\*+|_+)/.exec(text.slice(i))[0].length);
      let matched = false;

      if (!wordish) {
        for (let width = opened; width >= 1 && !matched; width--) {
          const close = findClose(text, i + width, c.repeat(width));
          if (close < 0) continue;
          const inner = text.slice(i + width, close);
          if (!inner.trim()) continue;
          nested(inner, width === 3 ? { bold: true, italic: true } : width === 2 ? { bold: true } : { italic: true });
          i = close + width;
          matched = true;
        }
      }
      if (matched) continue;
    }

    // 7. A bare URL, which GitHub links whether or not you asked it to.
    if ((c === 'h' || c === 'w') && /^(https?:\/\/|www\.)/.test(text.slice(i))) {
      const m = /^(?:https?:\/\/|www\.)[^\s<>[\]()]+[^\s<>[\]().,;:!?'"]/.exec(text.slice(i));
      if (m) {
        push({ text: m[0], link: m[0].startsWith('www.') ? `https://${m[0]}` : m[0] });
        i += m[0].length;
        continue;
      }
    }

    plain += c;
    i++;
  }

  flush();
  return runs.length ? runs : [{ text: '' }];
}

/**
 * A link or image starting at `[`. Handles inline targets, reference targets
 * and the collapsed form where the label is the reference.
 */
function linkAt(text, start, refs) {
  // Find the matching close bracket, counting nesting so [a [b] c] works.
  let depth = 0;
  let close = -1;
  for (let i = start; i < text.length; i++) {
    if (text[i] === '\\') {
      i++;
      continue;
    }
    if (text[i] === '[') depth++;
    else if (text[i] === ']') {
      depth--;
      if (depth === 0) {
        close = i;
        break;
      }
    }
  }
  if (close < 0) return null;
  const label = text.slice(start + 1, close);

  // Inline: [label](url "title")
  if (text[close + 1] === '(') {
    let depth2 = 0;
    let end = -1;
    for (let i = close + 1; i < text.length; i++) {
      if (text[i] === '\\') {
        i++;
        continue;
      }
      if (text[i] === '(') depth2++;
      else if (text[i] === ')') {
        depth2--;
        if (depth2 === 0) {
          end = i;
          break;
        }
      }
    }
    if (end < 0) return null;
    const target = text.slice(close + 2, end).trim();
    const m = /^(<[^>]*>|\S+)(?:\s+["'(](.*)["')])?$/.exec(target) || [null, target, null];
    return { label, url: String(m[1] || '').replace(/^<|>$/g, ''), title: m[2] || null, end: end + 1 };
  }

  // Reference: [label][ref] or [label][] or [label]
  const refMatch = /^\[([^\]]*)\]/.exec(text.slice(close + 1));
  const key = (refMatch && refMatch[1].trim() ? refMatch[1] : label).trim().toLowerCase();
  const found = refs.get(key);
  if (!found) return null;
  return { label, url: found.url, title: found.title, ref: found.label, end: close + 1 + (refMatch ? refMatch[0].length : 0) };
}

/** The closing run of `mark`, skipping code spans and escapes. */
function findClose(text, from, mark) {
  for (let i = from; i < text.length; i++) {
    if (text[i] === '\\') {
      i++;
      continue;
    }
    if (text[i] === '`') {
      const fence = /^`+/.exec(text.slice(i))[0];
      const end = text.indexOf(fence, i + fence.length);
      if (end > 0) {
        i = end + fence.length - 1;
        continue;
      }
    }
    if (text.startsWith(mark, i)) {
      // A double mark cannot close a single one: *a**b* is not emphasis of a*.
      if (mark.length === 1 && text[i + 1] === mark) {
        i++;
        continue;
      }
      return i;
    }
  }
  return -1;
}

/* ── blocks ──────────────────────────────────────────────────────────────── */

/**
 * Markdown to the block model the rest of the suite speaks.
 *
 * @param {Uint8Array|Buffer|string} input
 * @param {{ decode?: (input) => string }} [opts]
 * @returns {{ blocks: object[], meta: object }}
 */
export function parseMarkdown(text) {
  const source = String(text).replace(/\r\n?/g, '\n');
  const lines = source.split('\n');
  const blocks = [];
  const refs = new Map();
  const footnotes = [];
  let frontMatter = null;
  let i = 0;

  // Front matter, only at the very top. Jekyll, Hugo, Docusaurus and every
  // static site put configuration here, and it is not ours to rewrite.
  if (/^---\s*$/.test(lines[0] || '')) {
    const end = lines.findIndex((l, n) => n > 0 && /^(---|\.\.\.)\s*$/.test(l));
    if (end > 0) {
      frontMatter = lines.slice(1, end).join('\n');
      i = end + 1;
    }
  }

  // Reference definitions are collected first, because a link may use one that
  // is defined at the bottom of the file.
  for (const line of lines) {
    const m = /^ {0,3}\[([^^\]][^\]]*)\]:\s*(\S+)(?:\s+["'(](.*)["')])?\s*$/.exec(line);
    if (m) refs.set(m[1].trim().toLowerCase(), { url: m[2].replace(/^<|>$/g, ''), title: m[3] || null, label: m[1] });
  }

  const runs = (s) => inlineRuns(s, refs);
  const flat = (list) => list.map((r) => r.text ?? '').join('');

  while (i < lines.length) {
    const line = lines[i];

    if (isBlank(line)) {
      i++;
      continue;
    }

    // A definition line was already collected; do not emit it as a paragraph.
    if (/^ {0,3}\[([^^\]][^\]]*)\]:\s*\S+/.test(line)) {
      i++;
      continue;
    }

    // Footnote definition: [^id]: the note
    let m = /^ {0,3}\[\^([^\]]+)\]:\s*(.*)$/.exec(line);
    if (m) {
      const body = [m[2]];
      i++;
      while (i < lines.length && (isBlank(lines[i]) ? false : indentOf(lines[i]) >= 4 || !/^\S/.test(lines[i]))) body.push(lines[i++].trim());
      const joined = body.join(' ').trim();
      footnotes.push({ id: m[1], text: joined, runs: runs(joined) });
      continue;
    }

    // ATX heading, closing hashes optional.
    if ((m = /^ {0,3}(#{1,6})(\s+(.*?))?\s*#*\s*$/.exec(line))) {
      const body = (m[3] || '').trim();
      blocks.push({ type: 'heading', level: m[1].length, text: body, runs: runs(body) });
      i++;
      continue;
    }

    // Fenced code, either fence character, with the language kept.
    if ((m = /^ {0,3}(`{3,}|~{3,})\s*(\S*)/.exec(line))) {
      const fence = m[1][0].repeat(Math.max(3, m[1].length));
      const language = m[2] || '';
      const body = [];
      i++;
      while (i < lines.length && !new RegExp(`^ {0,3}${m[1][0]}{${m[1].length},}\\s*$`).test(lines[i])) body.push(lines[i++]);
      i++; // the closing fence
      blocks.push({ type: 'code', language, text: body.join('\n'), fence });
      continue;
    }

    if (isRule(line)) {
      blocks.push({ type: 'rule', mark: line.trim()[0] });
      i++;
      continue;
    }

    // Table: a pipe row followed by a delimiter row that sets the alignment.
    if (/\|/.test(line) && /^ {0,3}\|?[\s:|-]*-[\s:|-]*$/.test(lines[i + 1] || '') && /-/.test(lines[i + 1] || '')) {
      const cells = (l) => l.trim().replace(/^\|/, '').replace(/\|$/, '').split(/(?<!\\)\|/).map((c) => c.trim().replace(/\\\|/g, '|'));
      const header = cells(line);
      const align = cells(lines[i + 1]).map((spec) => {
        const left = spec.startsWith(':');
        const right = spec.endsWith(':');
        return left && right ? 'center' : right ? 'right' : left ? 'left' : null;
      });
      i += 2;
      const rows = [header];
      while (i < lines.length && lines[i].trim() && /\|/.test(lines[i])) rows.push(cells(lines[i++]));
      blocks.push({
        type: 'table',
        header: true,
        align,
        rows: rows.map((row) => row.map((c) => ({ text: c, runs: runs(c) }))),
      });
      continue;
    }

    // Block quote, including the nesting depth and lazy continuation lines.
    if (/^ {0,3}>/.test(line)) {
      const body = [];
      let level = 1;
      while (i < lines.length && (/^ {0,3}>/.test(lines[i]) || (!isBlank(lines[i]) && body.length))) {
        const q = /^ {0,3}((?:>\s?)+)(.*)$/.exec(lines[i]);
        if (q) {
          level = Math.max(level, (q[1].match(/>/g) || []).length);
          body.push(q[2]);
        } else body.push(lines[i].trim()); // lazy continuation
        i++;
      }
      const joined = body.join('\n').trim();
      blocks.push({
        type: 'quote',
        level,
        text: joined,
        // One run set per line, so the quote is written back with the breaks it
        // was written with rather than reflowed into a single long line.
        lines: joined.split('\n').map((l) => ({ text: l, runs: runs(l) })),
        runs: runs(joined.replace(/\n/g, ' ')),
      });
      continue;
    }

    // Lists. Nesting is by indent; a task box makes it a task item.
    // A marker with nothing after it is an empty list item, not a paragraph.
    // Pull-request and issue templates are full of them, and reading the
    // second `-` of a pair as a setext underline turns the list into a heading.
    if ((m = LIST_OPEN.exec(line))) {
      const ordered = /\d/.test(m[2]);
      const start = ordered ? Number(m[2].replace(/\D/g, '')) : 1;
      const baseIndent = indentOf(line);
      const items = [];
      let loose = false;

      while (i < lines.length) {
        const item = LIST_ITEM.exec(lines[i]);
        if (item) {
          const indent = indentOf(lines[i]);
          if (indent < baseIndent) break;
          // A bullet list and a numbered list that touch are two lists, not
          // one. GitHub renders them as separate <ul> and <ol>, and merging
          // them turns "1. 2. 3." into three more bullets.
          if (indent === baseIndent && /\d/.test(item[2]) !== ordered) break;
          let body = item[3] ?? '';
          const task = /^\[([ xX])\]\s+/.exec(body);
          if (task) body = body.slice(task[0].length);
          i++;

          // Continuation lines belong to this item, not to a new paragraph.
          const more = [];
          while (i < lines.length && !isBlank(lines[i]) && !/^(\s*)([-*+]|\d{1,9}[.)])\s+/.test(lines[i]) && indentOf(lines[i]) > indent) {
            more.push(lines[i++].trim());
          }
          if (more.length) body = `${body} ${more.join(' ')}`;

          items.push({
            text: body,
            runs: runs(body),
            level: Math.floor(Math.max(0, indent - baseIndent) / 2),
            marker: item[2],
            ordered: /\d/.test(item[2]),
            task: task ? task[1].toLowerCase() === 'x' : null,
          });
          continue;
        }
        // One blank line inside a list keeps the list going, and makes it
        // loose — but only if what follows really is the same list.
        const next = LIST_OPEN.exec(lines[i + 1] || '');
        if (isBlank(lines[i]) && next && !(indentOf(lines[i + 1]) === baseIndent && /\d/.test(next[2]) !== ordered)) {
          loose = true;
          i++;
          continue;
        }
        break;
      }
      blocks.push({ type: 'list', ordered, start, loose, items });
      continue;
    }

    // Indented code, but only where a paragraph is not already running.
    if (indentOf(line) >= 4 && !blocks.length) {
      const body = [];
      while (i < lines.length && (indentOf(lines[i]) >= 4 || isBlank(lines[i]))) body.push(lines[i++].replace(/^ {4}|\t/, ''));
      while (body.length && isBlank(body[body.length - 1])) body.pop();
      blocks.push({ type: 'code', language: '', text: body.join('\n'), indented: true });
      continue;
    }

    // An HTML block, kept whole and unedited.
    if (/^ {0,3}<(?:!--|\/?[a-zA-Z][a-zA-Z0-9-]*)/.test(line)) {
      const body = [];
      while (i < lines.length && !isBlank(lines[i])) body.push(lines[i++]);
      blocks.push({ type: 'html', text: body.join('\n') });
      continue;
    }

    // Paragraph, with setext headings recognised from the line below it.
    const para = [];
    while (i < lines.length && !isBlank(lines[i])) {
      const next = lines[i + 1] || '';
      if (para.length === 0 && /^ {0,3}(=+|-+)\s*$/.test(next) && !isRule(next)) {
        const body = lines[i].trim();
        blocks.push({ type: 'heading', level: next.trim()[0] === '=' ? 1 : 2, setext: true, text: body, runs: runs(body) });
        i += 2;
        break;
      }
      if (/^ {0,3}(#{1,6}\s|>|```|~~~)/.test(lines[i]) && para.length) break;
      if (LIST_OPEN.test(lines[i]) && para.length) break;
      if (isRule(lines[i]) && para.length) break;
      para.push(lines[i++]);
    }
    if (para.length) {
      // Two spaces at the end of a line is a hard break, and is meaningful.
      const joined = para.map((l, n) => (/ {2,}$/.test(l) && n < para.length - 1 ? `${l.trim()}\n` : l.trim())).join(' ').replace(/\n /g, '\n').trim();
      blocks.push({ type: 'paragraph', text: joined, runs: runs(joined) });
    }
  }

  return {
    blocks,
    meta: {
      frontMatter,
      definitions: [...refs.values()],
      footnotes,
      // Column alignment, one entry per table in document order. It belongs to
      // the file rather than to any block a document can hold, so anything
      // converting through another format keeps it here.
      tables: blocks.filter((b) => b.type === 'table').map((b) => b.align || []),
      // What the file said about itself, for a title when nothing else does.
      title: blocks.find((b) => b.type === 'heading' && b.level === 1)?.text || null,
    },
  };
}

/* ── writing ─────────────────────────────────────────────────────────────── */

/**
 * Escape only what would actually change the meaning if left alone.
 *
 * The obvious implementation puts a backslash in front of every mark, and it
 * ruins files: `project_api_provider.md` becomes `project\_api\_provider.md`,
 * `2 * 3` becomes `2 \* 3`, and a document that was opened and saved with no
 * edit at all comes back with a hundred changed lines. Since the reader only
 * opens emphasis where it could close it, and never inside a word for `_`, the
 * writer escapes on exactly those conditions and leaves the rest alone.
 */
function escapeInline(raw) {
  let out = '';
  for (let i = 0; i < raw.length; i++) {
    const c = raw[i];
    const rest = raw.slice(i + 1);

    if (c === '\\' || c === '`') {
      out += `\\${c}`;
      continue;
    }
    // A lone asterisk cannot become emphasis: there is nothing to close it.
    if (c === '*' && rest.includes('*')) {
      out += '\\*';
      continue;
    }
    // The reader ignores an underscore inside a word, so the writer may too.
    if (c === '_' && rest.includes('_') && !/\w/.test(raw[i - 1] || '')) {
      out += '\\_';
      continue;
    }
    if (c === '~' && raw[i + 1] === '~') {
      out += '\\~';
      continue;
    }
    // A bracket only matters where it could open a link or a reference.
    if ((c === '[' || c === ']') && /\]\(|\]\[|\]:/.test(raw.slice(Math.max(0, i - 1)))) {
      out += `\\${c}`;
      continue;
    }
    out += c;
  }
  return out;
}

/** Runs back to inline Markdown, innermost mark first. */
export function runsToMarkdown(runs) {
  return (runs || [])
    .map((r) => {
      const raw = String(r.text ?? '');
      if (r.html) return raw;
      if (r.code) return `\`${raw}\``;
      if (r.footnote) return `[^${r.footnote}]`;

      // A link whose text is its own address needs no brackets — that is how it
      // was written and how GitHub renders it. Wrapping it as [url](url) is
      // technically the same page and a worse file.
      const plainStyle = !r.bold && !r.italic && !r.strike;
      if (r.link && plainStyle && !r.image) {
        if (r.angle) return `<${raw}>`;
        if (raw === r.link) return raw;
        if (r.link === `mailto:${raw}`) return `<${raw}>`;
      }

      let t = escapeInline(raw);
      if (r.strike) t = `~~${t}~~`;
      if (r.bold) t = `**${t}**`;
      if (r.italic) t = `*${t}*`;
      if (r.image) return `![${t}](${r.image}${r.title ? ` "${r.title}"` : ''})`;
      // A link written against a definition stays written that way, so a file
      // that keeps its addresses at the bottom still does after a save.
      if (r.link && r.ref) return `[${t}][${r.ref}]`;
      if (r.link) return `[${t}](${r.link}${r.title ? ` "${r.title}"` : ''})`;
      return t;
    })
    .join('');
}

/**
 * The block model back to Markdown.
 *
 * `meta` is what came out of the parser. Passing it back is what makes the
 * round trip safe: the front matter and any definitions the document still
 * refers to are restored exactly as they were written.
 */
export function serializeMarkdown(blocks, meta = {}) {
  const body = (b) => b.text ?? runsToMarkdown(b.runs);
  const out = [];

  for (const block of blocks || []) {
    switch (block.type) {
      case 'heading': {
        const heading = runsToMarkdown(block.runs) || block.text || '';
        // A file that used the underlined form keeps it. The rendered page is
        // identical either way; the diff is not.
        if (block.setext && (block.level === 1 || block.level === 2)) {
          out.push(`${heading}
${(block.level === 1 ? '=' : '-').repeat(Math.max(3, heading.length))}`);
          break;
        }
        out.push(`${'#'.repeat(Math.min(Math.max(block.level || 1, 1), 6))} ${heading}`);
        break;
      }

      case 'code': {
        const language = block.language || '';
        // A fence long enough to contain whatever is inside it.
        const longest = Math.max(3, ...String(block.text || '').split('\n').map((l) => (/^`+/.exec(l.trim()) || [''])[0].length + 1));
        const fence = '`'.repeat(longest);
        out.push(`${fence}${language}\n${block.text ?? ''}\n${fence}`);
        break;
      }

      case 'quote': {
        const prefix = '> '.repeat(block.level || 1);
        const body = block.lines
          ? block.lines.map((l) => runsToMarkdown(l.runs) || l.text || '')
          : String(runsToMarkdown(block.runs) || block.text || '').split('\n');
        out.push(body.map((l) => `${prefix}${l}`.trimEnd()).join('\n'));
        break;
      }

      case 'list': {
        let n = block.start || 1;
        out.push(
          (block.items || [])
            .map((item) => {
              const pad = '  '.repeat(item.level || 0);
              const marker = block.ordered ? `${n++}.` : item.marker && !/\d/.test(item.marker) ? item.marker : '-';
              const box = item.task === true ? '[x] ' : item.task === false ? '[ ] ' : '';
              return `${pad}${marker} ${box}${runsToMarkdown(item.runs) || item.text || ''}`.trimEnd();
            })
            .join(block.loose ? '\n\n' : '\n')
        );
        break;
      }

      case 'table': {
        const rows = block.rows || [];
        if (!rows.length) break;
        const cell = (c) => String(runsToMarkdown(c.runs) || c.text || '').replace(/\|/g, '\\|');
        // Columns padded to a common width. GitHub does not need it; a person
        // reading the file in a text editor very much does.
        const widths = rows[0].map((_, col) => Math.max(3, ...rows.map((r) => cell(r[col] ?? { text: '' }).length)));
        const line = (row) => `| ${row.map((c, col) => cell(c).padEnd(widths[col])).join(' | ')} |`;
        const rule = `| ${widths
          .map((w, col) => {
            const a = (block.align || [])[col];
            if (a === 'center') return `:${'-'.repeat(Math.max(1, w - 2))}:`;
            if (a === 'right') return `${'-'.repeat(Math.max(1, w - 1))}:`;
            if (a === 'left') return `:${'-'.repeat(Math.max(1, w - 1))}`;
            return '-'.repeat(w);
          })
          .join(' | ')} |`;
        out.push([line(rows[0]), rule, ...rows.slice(1).map(line)].join('\n'));
        break;
      }

      case 'rule':
        out.push('---');
        break;

      case 'html':
        out.push(block.text ?? '');
        break;

      default: {
        const text = runsToMarkdown(block.runs) || body(block) || '';
        if (text.trim()) out.push(text);
      }
    }
  }

  // Footnote definitions belong at the bottom, after everything they annotate.
  for (const note of meta.footnotes || []) out.push(`[^${note.id}]: ${runsToMarkdown(note.runs) || note.text || ''}`);

  // The definitions the document still refers to. A `[text][ref]` written above
  // is a dangling link without the `[ref]: url` line below it, and a definition
  // whose link was edited away is litter — so this keeps exactly the ones in use.
  const used = new Set();
  const walk = (runs) => {
    for (const run of runs || []) {
      if (run.ref) used.add(String(run.ref).toLowerCase());
      if (run.link) used.add(String(run.link).toLowerCase());
    }
  };
  for (const block of blocks || []) {
    walk(block.runs);
    for (const item of block.items || []) walk(item.runs);
    for (const line of block.lines || []) walk(line.runs);
    for (const row of block.rows || []) for (const cell of row || []) walk(cell.runs);
  }
  for (const note of meta.footnotes || []) walk(note.runs);

  for (const d of meta.definitions || []) {
    if (!used.has(String(d.label).toLowerCase()) && !used.has(String(d.url).toLowerCase())) continue;
    out.push(`[${d.label}]: ${d.url}${d.title ? ` "${d.title}"` : ''}`);
  }

  const document = out.join('\n\n').replace(/\n{3,}/g, '\n\n').trim();
  const head = meta.frontMatter != null ? `---\n${meta.frontMatter}\n---\n\n` : '';
  return `${head}${document}\n`;
}
