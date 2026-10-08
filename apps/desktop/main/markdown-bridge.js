// Markdown in and out of the word processor.
//
// The suite's document engine speaks OOXML, so a .md file has to become a
// document to be edited and a file again to be saved. The whole difficulty is
// the second direction: a bridge that only carries text turns a README into a
// wall of paragraphs, and the damage is invisible until the pull request.
//
// So the crossing is made in terms a document can hold and Markdown can be
// rebuilt from:
//
//   heading      Heading1…6, the styles Word already has
//   list item    ListBullet / ListNumber, the marker and any task box written
//                into the text, indentation as leading spaces
//   code         one paragraph per line in the Code style, fences included, so
//                the block is visible and reconstructs exactly
//   quote        the Quote style
//   table        a real table, not tab-separated text
//   rule         a paragraph holding the three hyphens it came from
//   link         a real hyperlink relationship, which the engine models as a
//                run's `link` and writes back
//
// Everything that has no place in a document at all — YAML front matter, link
// definitions, footnote bodies — is kept beside the session and restored on the
// way out. That is the difference between an editor you can point at a repo and
// one you cannot.

import { parseMarkdown, serializeMarkdown, runsToMarkdown } from '@rutba/office-formats/markdown';

const HEADING = /^Heading([1-6])$/i;

/** A task box a person can see and tick, rather than a hidden attribute. */
const BOX = { true: '☑ ', false: '☐ ' };

/**
 * Markdown blocks → the paragraph list `buildDocx` takes.
 * @param {object[]} blocks from parseMarkdown
 */
export function markdownToParagraphs(blocks) {
  const out = [];

  for (const block of blocks || []) {
    switch (block.type) {
      case 'heading':
        out.push({ style: `Heading${Math.min(Math.max(block.level || 1, 1), 6)}`, runs: runsFor(block.runs, block.text) });
        break;

      case 'code': {
        const fence = '```';
        const lines = String(block.text ?? '').split('\n');
        // The fences are part of what is on screen. A person editing a README
        // needs to see where the code block starts and stops, and keeping them
        // means the block reconstructs by reading the lines back as Markdown.
        for (const line of [`${fence}${block.language || ''}`, ...lines, fence]) {
          out.push({ style: 'Code', runs: [{ text: line, font: 'Consolas' }] });
        }
        break;
      }

      case 'quote':
        for (const line of block.lines || [{ runs: block.runs, text: block.text }]) {
          out.push({ style: 'Quote', runs: runsFor(line.runs, line.text) });
        }
        break;

      case 'list':
        for (const [n, item] of (block.items || []).entries()) {
          const marker = block.ordered ? `${(block.start || 1) + n}. ` : '- ';
          const lead = '  '.repeat(item.level || 0);
          const box = item.task === null || item.task === undefined ? '' : BOX[String(item.task)];
          out.push({
            style: block.ordered ? 'ListNumber' : 'ListBullet',
            runs: [{ text: `${lead}${marker}${box}` }, ...runsFor(item.runs, item.text)],
          });
        }
        break;

      case 'table':
        out.push({
          table: {
            header: true,
            rows: (block.rows || []).map((row) => row.map((cell) => cell.text ?? runsToMarkdown(cell.runs))),
          },
        });
        break;

      case 'rule':
        out.push({ style: 'Rule', text: '---' });
        break;

      case 'html':
        for (const line of String(block.text ?? '').split('\n')) out.push({ style: 'Code', runs: [{ text: line, font: 'Consolas' }] });
        break;

      default:
        out.push({ runs: runsFor(block.runs, block.text) });
    }
  }

  return out.length ? out : [{ text: '' }];
}

/** Markdown runs → document runs. The styling names differ; the shape does not. */
function runsFor(runs, fallback = '') {
  const list = (runs || []).filter((r) => r.text !== '' || r.image);
  if (!list.length) return [{ text: fallback ?? '' }];
  return list.map((r) => ({
    // An image in Markdown is a reference, and until the picture itself is
    // fetched the honest thing to show is the reference.
    text: r.image ? `![${r.text}](${r.image})` : r.text ?? '',
    bold: r.bold || undefined,
    italic: r.italic || undefined,
    strike: r.strike || undefined,
    font: r.code ? 'Consolas' : undefined,
    link: r.link || undefined,
  }));
}

/**
 * The document back to Markdown.
 *
 * Reconstruction is by paragraph style, which is why the crossing above is
 * careful to use one. Anything unrecognised is a paragraph, which is both the
 * right answer and the safe one.
 *
 * A list the document numbers itself (made in the editor, or in Word) is
 * known by its labels — `render().listLabels` — and written as a Markdown
 * list: a bullet as "-", a number as the count it has reached, each level
 * in by its indent.
 *
 * @param {object[]} blocks `render().blocks` from the document engine
 * @param {object} [meta] what parseMarkdown returned on the way in
 * @param {object} [labels] `render().listLabels`: each list item's label by its block's index
 */
export function paragraphsToMarkdown(blocks, meta = {}, labels = null) {
  const pieces = [];
  let run = null; // a group of adjacent lines that belong together
  // A run of numbered items: the indents open above this one, and each level's count.
  let indents = [];
  let counts = [];

  const closeRun = () => {
    if (run) pieces.push(run.lines.join('\n'));
    run = null;
    indents = [];
    counts = [];
  };
  const inRun = (kind, line) => {
    if (!run || run.kind !== kind) {
      closeRun();
      run = { kind, lines: [] };
    }
    run.lines.push(line);
  };

  // A table reaches here as one paragraph per cell, each carrying the address
  // of the cell it lives in — `t1024:r2:c0`, the table's offset then the row
  // and cell ordinals. Rebuilding the grid means collecting those addresses
  // until the table ends; a cell's text is simply the paragraph's text.
  let table = null;
  let tableCount = 0;
  const closeTable = () => {
    if (!table) return;
    const rows = table.rows.map((row) => (row || []).map((c) => c ?? ''));
    if (rows.length) {
      const columns = Math.max(...rows.map((r) => r.length));
      const widths = Array.from({ length: columns }, (_, c) => Math.max(3, ...rows.map((r) => String(r[c] ?? '').length)));
      const line = (row) => `| ${Array.from({ length: columns }, (_, i) => String(row[i] ?? '').padEnd(widths[i])).join(' | ')} |`;

      // Column alignment is a property of the Markdown table and has nowhere to
      // live in the document, so it is remembered from the way in and applied
      // to the same table on the way out. Tables added while editing simply
      // have none, which is the correct default.
      const align = (meta.tables || [])[tableCount] || [];
      const rule = `| ${widths
        .map((w, c) => {
          if (align[c] === 'center') return `:${'-'.repeat(Math.max(1, w - 2))}:`;
          if (align[c] === 'right') return `${'-'.repeat(Math.max(1, w - 1))}:`;
          if (align[c] === 'left') return `:${'-'.repeat(Math.max(1, w - 1))}`;
          return '-'.repeat(w);
        })
        .join(' | ')} |`;

      pieces.push([line(rows[0]), rule, ...rows.slice(1).map(line)].join('\n'));
    }
    tableCount++;
    table = null;
  };

  for (const block of blocks || []) {
    const at = /^(t\d+(?::r\d+:c\d+:t\d+)*):r(\d+):c(\d+)$/.exec(String(block.container || ''));
    if (at) {
      closeRun();
      if (!table || table.id !== at[1]) {
        closeTable();
        table = { id: at[1], rows: [] };
      }
      const row = Number(at[2]);
      const cell = Number(at[3]);
      table.rows[row] = table.rows[row] || [];
      // A cell may hold several paragraphs. Markdown cells hold one line, so
      // they are joined rather than one of them being chosen.
      //
      // The header row's bold comes from the table style, not from the author —
      // GitHub renders a header row bold anyway — so it is not written back as
      // emphasis. A `**word**` a person typed into a header would be lost; a
      // header row that grew asterisks on every save is worse.
      const text = row === 0
        ? (block.runs || []).map((r) => r.text ?? '').join('') || block.text || ''
        : runsToMarkdown(toMarkdownRuns(block.runs)) || block.text || '';
      table.rows[row][cell] = [table.rows[row][cell], text].filter(Boolean).join(' ');
      continue;
    }
    closeTable();

    // A table handed over whole, which is what the reader produces before the
    // document engine has seen it.
    if (block.table || block.type === 'table') {
      closeRun();
      const rows = (block.table?.rows || block.rows || []).map((row) =>
        row.map((cell) => (typeof cell === 'string' ? cell : cell.text ?? runsToMarkdown(cell.runs)))
      );
      if (rows.length) {
        const widths = rows[0].map((_, c) => Math.max(3, ...rows.map((r) => String(r[c] ?? '').length)));
        const line = (row) => `| ${row.map((c, i) => String(c ?? '').padEnd(widths[i])).join(' | ')} |`;
        pieces.push([line(rows[0]), `| ${widths.map((w) => '-'.repeat(w)).join(' | ')} |`, ...rows.slice(1).map(line)].join('\n'));
      }
      continue;
    }

    const style = String(block.style || '');
    const text = runsToMarkdown(toMarkdownRuns(block.runs)) || block.text || '';

    const heading = HEADING.exec(style);
    if (heading) {
      closeRun();
      pieces.push(`${'#'.repeat(Number(heading[1]))} ${text}`);
      continue;
    }

    if (/^Code$/i.test(style)) {
      // Code lines are already literal — do not let the inline writer escape
      // the asterisks and underscores inside somebody's source.
      inRun('code', (block.runs || []).map((r) => r.text ?? '').join('') || block.text || '');
      continue;
    }

    if (/^Quote$/i.test(style)) {
      inRun('quote', `> ${text}`);
      continue;
    }

    if (/^List(Bullet|Number|Paragraph)$/i.test(style)) {
      // The marker, the indent and the task box are already in the text: they
      // were written there so a person could see and edit them. The box goes
      // back to being the two characters Markdown uses — a tick in the file
      // would render as a tick on GitHub, not as a checked box.
      inRun('list', text.replace(/☑\s*/, '[x] ').replace(/☐\s*/, '[ ] '));
      continue;
    }

    const label = labels?.[block.index];
    if (label?.label) {
      // Deeper than the item above it is a level in; shallower closes levels back to its own.
      const indent = Number(label.indentPx) || 0;
      if (run?.kind !== 'numbered') { closeRun(); run = { kind: 'numbered', lines: [] }; }
      while (indents.length && indent < indents[indents.length - 1] - 0.5) { indents.pop(); counts.pop(); }
      if (!indents.length || indent > indents[indents.length - 1] + 0.5) { indents.push(indent); counts.push(0); }
      const level = indents.length - 1;
      // This level's own count is the last number its label shows ("3.1." is the first under three).
      const shown = /(\d+)\D*$/.exec(label.label);
      counts[level] = shown ? Number(shown[1]) : counts[level] + 1;
      const marker = label.bullet ? '-' : `${counts[level]}.`;
      run.lines.push(`${'    '.repeat(level)}${marker} ${text}`);
      continue;
    }

    if (/^Rule$/i.test(style) || /^-{3,}$/.test(text.trim())) {
      closeRun();
      pieces.push('---');
      continue;
    }

    closeRun();
    if (text.trim()) pieces.push(text);
  }
  closeRun();
  closeTable();

  // Footnote bodies and the definitions still in use ride out on the tail, and
  // the front matter goes back on the front, exactly as it arrived.
  for (const note of meta.footnotes || []) pieces.push(`[^${note.id}]: ${runsToMarkdown(note.runs) || note.text || ''}`);

  const body = pieces.join('\n\n').replace(/\n{3,}/g, '\n\n').trim();
  const head = meta.frontMatter != null ? `---\n${meta.frontMatter}\n---\n\n` : '';
  const document = `${head}${body}\n`;

  // One pass through the reader and writer, so what lands on disk is normalised
  // the same way as a file the parser produced — and so the definitions that
  // are still referenced come back.
  const reparsed = parseMarkdown(document);
  return serializeMarkdown(reparsed.blocks, { ...reparsed.meta, definitions: meta.definitions || reparsed.meta.definitions });
}

/** Document runs → the shape the Markdown writer expects. */
function toMarkdownRuns(runs) {
  return (runs || []).map((r) => ({
    text: r.text ?? '',
    bold: r.bold || undefined,
    italic: r.italic || undefined,
    strike: r.strike || undefined,
    code: /consolas|courier|mono/i.test(String(r.font || r.rPr?.font || '')) || undefined,
    link: r.link || undefined,
  }));
}
