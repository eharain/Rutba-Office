// A document written as HTML, plain text or Rich Text from the document
// model odt.js and docx-read.js read into — so a Word document saved as one
// of them keeps its headings, its lists with their numbers, its tables with
// their spans and its pictures, where each was written from the editor's
// plain lines and came out as a paragraph.
//
//   - listLabels: each list item's label as the list's levels count it —
//     a bullet's character, or a number in its format (1, a, A, i, I) with
//     its prefix and suffix and the levels above it where the level shows
//     them, counting on from the level's start and starting again under a
//     new item above;
//   - writeHtmlDocument: headings, paragraphs in their alignment and
//     indents, runs in their looks, links, lists as nested <ul>/<ol> in
//     their style, tables with their widths, spans, shading and borders,
//     pictures embedded as data, the page's margins;
//   - writePlainDocument: the words, each list item after its label, a
//     table's rows with tabs between their cells;
//   - toRtfBlocks: the blocks writeRtf takes, labels, indents, links,
//     pictures, spans and widths included.

import { writeRtf } from './rtf.js';

const escHtml = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const MEDIA = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp' };

function roman(n) {
  const table = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
  let out = '';
  let v = Math.max(1, Math.min(3999, n));
  for (const [k, s] of table) while (v >= k) { out += s; v -= k; }
  return out;
}
function letters(n) {
  let out = '';
  let v = Math.max(1, n);
  while (v > 0) { v -= 1; out = String.fromCharCode(97 + (v % 26)) + out; v = Math.floor(v / 26); }
  return out;
}
/** A count in a level's number format. */
function formatNumber(n, format) {
  switch (format) {
    case 'a': return letters(n);
    case 'A': return letters(n).toUpperCase();
    case 'i': return roman(n);
    case 'I': return roman(n).toUpperCase();
    default: return String(n);
  }
}

/**
 * Each list item's label, as a Map from block to { label, level, kind, indent, hanging }.
 * Blocks inside table cells are counted too, in the order they come.
 */
export function listLabels(doc) {
  const out = new Map();
  const counters = new Map(); // list id → counts by level
  const visit = (blocks) => {
    for (const b of blocks || []) {
      if (b.type === 'table') { for (const row of b.rows || []) for (const cell of row) if (!cell?.covered) visit(cell.blocks); continue; }
      if (!b.list) continue;
      const levels = doc.lists?.get?.(b.list.style) || [];
      const level = Math.max(0, Math.min(8, b.list.level || 0));
      const lv = levels[level] || { kind: 'bullet', char: '•', indent: 48 + 24 * level, hanging: 24 };
      if (!counters.has(b.list.id)) counters.set(b.list.id, []);
      const counts = counters.get(b.list.id);
      for (let k = 0; k < level; k++) if (counts[k] == null) counts[k] = levels[k]?.start ?? 1;
      counts[level] = counts[level] == null ? (lv.start ?? 1) : counts[level] + 1;
      counts.length = level + 1; // a deeper level starts again under this item
      let label;
      if (lv.kind === 'number') {
        const shown = Math.max(1, Math.min(lv.display || 1, level + 1));
        const parts = [];
        for (let k = level - shown + 1; k <= level; k++) parts.push(formatNumber(counts[k] ?? 1, k === level ? lv.format : levels[k]?.format || '1'));
        label = `${lv.prefix || ''}${parts.join('.')}${lv.suffix ?? '.'}`;
      } else label = lv.char || '•';
      out.set(b, { label, level, count: counts[level], kind: lv.kind, format: lv.format, start: lv.start, indent: lv.indent ?? 48 + 24 * level, hanging: lv.hanging ?? 24 });
    }
  };
  visit(doc.blocks);
  return out;
}

const textOfRuns = (runs) => (runs || []).map((r) => (r.tab ? '\t' : r.br ? '\n' : r.image ? '' : r.text || '')).join('');

/* ── HTML ──────────────────────────────────────────────────────────────── */

const CSS_LIST = { 1: 'decimal', a: 'lower-alpha', A: 'upper-alpha', i: 'lower-roman', I: 'upper-roman' };

/** The document as one HTML page, its pictures in it as data. */
export function writeHtmlDocument(doc = {}) {
  const runStyle = (r) => {
    const css = [];
    if (r.color) css.push(`color:${r.color}`);
    if (r.size) css.push(`font-size:${r.size}pt`);
    if (r.font) css.push(`font-family:'${String(r.font).replace(/'/g, '')}'`);
    if (r.highlight) css.push(`background-color:${r.highlight}`);
    return css.join(';');
  };
  const runsHtml = (runs) => {
    let out = '';
    for (const r of runs || []) {
      let piece;
      if (r.image) {
        const data = doc.images?.get?.(r.image.href);
        if (!data) continue;
        const ext = (/\.([a-z0-9]+)$/i.exec(r.image.href)?.[1] || 'png').toLowerCase();
        piece = `<img src="data:${MEDIA[ext] || 'application/octet-stream'};base64,${Buffer.from(data).toString('base64')}" width="${Math.round(r.image.width)}" height="${Math.round(r.image.height)}" alt="${escHtml(r.image.name || '')}">`;
      } else if (r.tab) piece = '&emsp;';
      else if (r.br) piece = '<br>';
      else {
        piece = escHtml(r.text);
        if (r.vertical === 'superscript') piece = `<sup>${piece}</sup>`;
        if (r.vertical === 'subscript') piece = `<sub>${piece}</sub>`;
        if (r.strike) piece = `<s>${piece}</s>`;
        if (r.underline) piece = `<u>${piece}</u>`;
        if (r.italic) piece = `<em>${piece}</em>`;
        if (r.bold) piece = `<strong>${piece}</strong>`;
        const css = runStyle(r);
        if (css) piece = `<span style="${escHtml(css)}">${piece}</span>`;
      }
      if (r.link) piece = `<a href="${escHtml(r.link)}">${piece}</a>`;
      out += piece;
    }
    return out;
  };
  const paraStyle = (b) => {
    const css = [];
    const align = { center: 'center', right: 'right', both: 'justify' }[b.align];
    if (align) css.push(`text-align:${align}`);
    if (!b.list && b.indentLeft) css.push(`margin-left:${Math.round(b.indentLeft)}px`);
    if (!b.list && b.indentFirst) css.push(`text-indent:${Math.round(b.indentFirst)}px`);
    if (b.indentRight) css.push(`margin-right:${Math.round(b.indentRight)}px`);
    if (b.pageBreakBefore) css.push('break-before:page');
    return css.length ? ` style="${escHtml(css.join(';'))}"` : '';
  };
  const paragraphHtml = (b) => {
    const tag = b.heading ? `h${Math.min(6, b.heading)}` : 'p';
    return `<${tag}${paraStyle(b)}${b.rtl ? ' dir="rtl"' : ''}>${runsHtml(b.runs)}</${tag}>`;
  };
  const tableHtml = (t) => {
    const widths = t.columns || [];
    const cols = widths.some(Boolean) ? `<colgroup>${widths.map((w) => (w ? `<col style="width:${Math.round(w)}px">` : '<col>')).join('')}</colgroup>` : '';
    const rows = (t.rows || []).map((row) => `<tr>${row.map((c) => {
      if (!c || c.covered) return '';
      const css = [c.fill ? `background-color:${c.fill}` : '', c.border ? 'border:1px solid #000' : ''].filter(Boolean).join(';');
      return `<td${c.colspan > 1 ? ` colspan="${c.colspan}"` : ''}${c.rowspan > 1 ? ` rowspan="${c.rowspan}"` : ''}${css ? ` style="${css}"` : ''}>${blocksHtml(c.blocks || [])}</td>`;
    }).join('')}</tr>`).join('\n');
    return `<table style="border-collapse:collapse">${cols}\n${rows}\n</table>`;
  };
  const labels = listLabels(doc);
  /** One list's run of items as nested lists, each level in its own list type. */
  const listHtml = (items) => {
    let out = '';
    const open = []; // tag per depth
    let itemOpen = false;
    const openList = (info) => {
      const tag = info.kind === 'number' ? 'ol' : 'ul';
      const type = info.kind === 'number' ? CSS_LIST[info.format] || 'decimal' : 'disc';
      // A list that carries on after something else starts at the number it had reached.
      const start = info.kind === 'number' && info.count && info.count !== 1 ? ` start="${info.count}"` : '';
      out += `<${tag} style="list-style-type:${type}"${start}>`;
      open.push(tag);
      itemOpen = false;
    };
    for (const b of items) {
      const info = labels.get(b);
      const level = info.level;
      while (open.length > level + 1) { out += `${itemOpen ? '</li>' : ''}</${open.pop()}>`; itemOpen = true; }
      while (open.length < level + 1) {
        if (open.length && !itemOpen) out += '<li style="list-style:none">';
        if (open.length) itemOpen = false;
        openList(open.length === level ? info : { kind: 'bullet' });
      }
      if (itemOpen) out += '</li>';
      const inner = b.heading ? paragraphHtml(b) : `${runsHtml(b.runs)}`;
      out += `<li${b.heading ? '' : paraStyle(b)}>${inner}`;
      itemOpen = true;
    }
    if (itemOpen) out += '</li>';
    while (open.length) { out += `</${open.pop()}>`; if (open.length) out += '</li>'; }
    return out;
  };
  function blocksHtml(blocks) {
    const out = [];
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.type === 'table') { out.push(tableHtml(b)); continue; }
      if (b.list && labels.has(b)) {
        const items = [];
        while (i < blocks.length && blocks[i].type !== 'table' && blocks[i].list?.id === b.list.id) items.push(blocks[i++]);
        i -= 1;
        out.push(listHtml(items));
        continue;
      }
      out.push(paragraphHtml(b));
    }
    return out.join('\n');
  }
  const page = doc.page;
  const pageCss = page ? `@page { size: ${(page.width / 96).toFixed(2)}in ${(page.height / 96).toFixed(2)}in; margin: ${(page.top / 96).toFixed(2)}in ${(page.right / 96).toFixed(2)}in ${(page.bottom / 96).toFixed(2)}in ${(page.left / 96).toFixed(2)}in; }\n` : '';
  const base = doc.defaults?.run || {};
  const bodyCss = [base.font ? `font-family:'${String(base.font).replace(/'/g, '')}'` : '', base.size ? `font-size:${base.size}pt` : ''].filter(Boolean).join(';');
  return `<!doctype html>\n<html>\n<head>\n<meta charset="utf-8">\n<title>${escHtml(doc.title || '')}</title>\n<style>\n${pageCss}${bodyCss ? `body { ${bodyCss} }\n` : ''}td { vertical-align: top; padding: 2px 6px; }\n</style>\n</head>\n<body>\n${blocksHtml(doc.blocks || [])}\n</body>\n</html>\n`;
}

/* ── plain text ────────────────────────────────────────────────────────── */

/** The document's words: a blank line between paragraphs, a list item after its label, a table row's cells between tabs. */
export function writePlainDocument(doc = {}) {
  const labels = listLabels(doc);
  const lines = [];
  const visit = (blocks) => {
    for (const b of blocks || []) {
      // A list ends with a blank line before whatever follows it.
      if (!labels.has(b) && lines.length && lines[lines.length - 1] !== '') lines.push('');
      if (b.type === 'table') {
        for (const row of b.rows || []) lines.push(row.map((c) => (c && !c.covered ? (c.blocks || []).map((x) => (x.type === 'table' ? '' : textOfRuns(x.runs))).filter(Boolean).join(' ') : '')).join('\t'));
        lines.push('');
        continue;
      }
      const info = labels.get(b);
      lines.push(`${info ? `${'    '.repeat(info.level)}${info.label} ` : ''}${textOfRuns(b.runs)}`);
      if (!info) lines.push('');
    }
  };
  visit(doc.blocks);
  return `${lines.join('\n').replace(/\n{3,}/g, '\n\n').trim()}\n`;
}

/* ── Rich Text ─────────────────────────────────────────────────────────── */

/** The blocks writeRtf takes, from the model: labels and indents, links, pictures, widths and spans. */
export function toRtfBlocks(doc = {}) {
  const labels = listLabels(doc);
  const runs = (rs) => (rs || []).flatMap((r) => {
    if (r.image) {
      const data = doc.images?.get?.(r.image.href);
      const ext = (/\.([a-z0-9]+)$/i.exec(r.image.href)?.[1] || '').toLowerCase();
      return data && (ext === 'png' || ext === 'jpg' || ext === 'jpeg') ? [{ picture: { data: Buffer.from(data), kind: ext === 'png' ? 'png' : 'jpeg', width: r.image.width, height: r.image.height } }] : [];
    }
    if (r.tab) return [{ ...r, tab: undefined, text: '\t' }];
    if (r.br) return [{ ...r, br: undefined, text: '\n' }];
    return [r];
  });
  const block = (b) => {
    if (b.type === 'table') {
      return {
        type: 'table',
        bordered: (b.rows || []).some((row) => row.some((c) => c && !c.covered && c.border)),
        columns: (b.columns || []).map((w) => Math.round((w || 0) * 15)),
        rows: (b.rows || []).map((row, ri) => row.map((c, ci) => {
          if (!c) return { runs: [] };
          if (c.covered) {
            // Under a span from above, a continuation; beside one, nothing of its own.
            const above = (() => { for (let r = ri - 1; r >= 0; r--) { const x = b.rows[r][ci]; if (x && !x.covered) return x.rowspan > ri - r ? x : null; } return null; })();
            return above ? { continues: true, colspan: above.colspan || 1, runs: [] } : { covered: true };
          }
          const paras = (c.blocks || []).filter((x) => x.type !== 'table');
          const cellRuns = paras.flatMap((x, k) => [...(k ? [{ text: '\n' }] : []), ...(labels.get(x) ? [{ text: `${labels.get(x).label}\t` }] : []), ...runs(x.runs)]);
          return { runs: cellRuns, colspan: c.colspan || 1, rowspan: c.rowspan || 1, fill: c.fill || null, align: paras[0]?.align || null };
        })),
      };
    }
    const info = labels.get(b);
    return {
      type: b.heading ? 'heading' : 'paragraph',
      ...(b.heading ? { level: b.heading } : {}),
      align: b.align || null,
      ...(info ? { label: info.label, indentTwips: Math.round(info.indent * 15), hangTwips: Math.round(info.hanging * 15) } : b.indentLeft || b.indentFirst ? { indentTwips: Math.round((b.indentLeft || 0) * 15), firstTwips: Math.round((b.indentFirst || 0) * 15) } : {}),
      ...(b.pageBreakBefore ? { pageBreakBefore: true } : {}),
      ...(b.rtl ? { rtl: true } : {}),
      runs: runs(b.runs),
    };
  };
  return (doc.blocks || []).map(block);
}

/** The document as Rich Text. */
export function writeRtfDocument(doc = {}) {
  return writeRtf({ blocks: toRtfBlocks(doc), title: doc.title || '' });
}
