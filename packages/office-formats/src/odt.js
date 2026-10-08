// An OpenDocument text document, read whole for the Word engine.
//
// odf.js reads an .odt into plain blocks (a heading's words, a list's items,
// a table's cells' words) for what only wants the text. This reads what a
// document needs to look as it was made:
//
//   - each paragraph's style — a heading's level, its alignment, its
//     indents, a page break before it — and each run's look: bold, italic,
//     underline, strike, colour, size and face, from the automatic and named
//     styles along their parent chains, a span's over its paragraph's;
//   - lists as lists: the list style each one names (bullets and their
//     characters, numbers and their format, suffix and start, level by
//     level), each item at its level, a nested list a level in;
//   - tables with their columns' widths, their cells' spans (a covered cell
//     where a span reaches), each cell's paragraphs and its background;
//   - pictures where they stand in the words, at their size; a text box's
//     paragraphs after its anchor; links; tabs and line breaks;
//   - the page: its size and margins, from the page layout.

import { readZip } from '@rutba/ooxml/zip';
import { parse, all, kids, first, textOf } from './xml.js';
import { lengthPx } from './odf.js';

/* ── styles ────────────────────────────────────────────────────────────── */

/** A text-properties element as run looks; only what it says. */
function textLook(props) {
  if (!props) return {};
  const a = props.attrs;
  const out = {};
  if (a['fo:font-weight']) out.bold = a['fo:font-weight'] === 'bold' || Number(a['fo:font-weight']) >= 600;
  if (a['fo:font-style']) out.italic = a['fo:font-style'] === 'italic' || a['fo:font-style'] === 'oblique';
  if (a['style:text-underline-style']) out.underline = a['style:text-underline-style'] !== 'none';
  if (a['style:text-line-through-style']) out.strike = a['style:text-line-through-style'] !== 'none';
  if (a['fo:color'] && /^#[0-9a-f]{6}$/i.test(a['fo:color'])) out.color = a['fo:color'].toUpperCase();
  if (a['fo:font-size'] && /pt$/.test(a['fo:font-size'])) out.size = Number.parseFloat(a['fo:font-size']);
  const face = a['style:font-name'] || a['fo:font-family'];
  if (face) out.font = String(face).replace(/^'|'$/g, '');
  if (a['fo:background-color'] && /^#[0-9a-f]{6}$/i.test(a['fo:background-color'])) out.highlight = a['fo:background-color'].toUpperCase();
  if (a['style:text-position']) {
    const pos = String(a['style:text-position']);
    if (/^super/.test(pos) || /^[1-9]/.test(pos)) out.vertical = 'superscript';
    else if (/^sub/.test(pos) || /^-/.test(pos)) out.vertical = 'subscript';
  }
  return out;
}

/** A paragraph-properties element as paragraph looks; only what it says. */
function paragraphLook(props) {
  if (!props) return {};
  const a = props.attrs;
  const out = {};
  const align = { start: 'left', left: 'left', center: 'center', end: 'right', right: 'right', justify: 'both' }[a['fo:text-align']];
  if (align) out.align = align;
  if (a['fo:margin-left'] != null) out.indentLeft = lengthPx(a['fo:margin-left']) ?? 0;
  if (a['fo:margin-right'] != null) out.indentRight = lengthPx(a['fo:margin-right']) ?? 0;
  if (a['fo:text-indent'] != null) out.indentFirst = lengthPx(a['fo:text-indent']) ?? 0;
  if (a['fo:margin-top'] != null) out.spaceBefore = lengthPx(a['fo:margin-top']) ?? 0;
  if (a['fo:margin-bottom'] != null) out.spaceAfter = lengthPx(a['fo:margin-bottom']) ?? 0;
  if (a['fo:break-before'] === 'page') out.pageBreakBefore = true;
  if (a['style:writing-mode'] === 'rl-tb') out.rtl = true;
  return out;
}

/**
 * Every style a document names, automatic and named, each resolved along
 * its parents: { text, paragraph, cell, column, named } — `named` whether
 * the style is one of the document's own named styles (not automatic).
 */
function readStyles(roots) {
  const raw = new Map();
  for (const { root, named } of roots) {
    if (!root) continue;
    for (const st of all(root, 'style:style')) {
      const name = st.attrs['style:name'];
      raw.set(`${st.attrs['style:family']}:${name}`, {
        parent: st.attrs['style:parent-style-name'] || null,
        family: st.attrs['style:family'],
        named,
        text: textLook(first(st, 'style:text-properties')),
        paragraph: paragraphLook(first(st, 'style:paragraph-properties')),
        cell: (() => {
          const c = first(st, 'style:table-cell-properties');
          if (!c) return {};
          const out = {};
          const bg = c.attrs['fo:background-color'];
          if (bg && /^#[0-9a-f]{6}$/i.test(bg)) out.fill = bg.toUpperCase();
          const border = c.attrs['fo:border'] || c.attrs['fo:border-top'] || c.attrs['fo:border-bottom'] || c.attrs['fo:border-left'] || c.attrs['fo:border-right'];
          if (border && border !== 'none') out.border = true;
          return out;
        })(),
        column: lengthPx(first(st, 'style:table-column-properties')?.attrs['style:column-width']),
        outline: st.attrs['style:default-outline-level'] ? Number(st.attrs['style:default-outline-level']) : null,
        display: st.attrs['style:display-name'] || name,
      });
    }
  }
  const resolve = (family, name, depth = 0) => {
    const own = raw.get(`${family}:${name}`);
    if (!own) return { text: {}, paragraph: {}, cell: {}, column: null, outline: null, chain: [] };
    const up = own.parent && depth < 12 ? resolve(family, own.parent, depth + 1) : { text: {}, paragraph: {}, cell: {}, column: null, outline: null, chain: [] };
    return {
      text: { ...up.text, ...own.text },
      paragraph: { ...up.paragraph, ...own.paragraph },
      cell: { ...up.cell, ...own.cell },
      column: own.column ?? up.column,
      outline: own.outline ?? up.outline,
      chain: [own.display, ...up.chain],
    };
  };
  return resolve;
}

/** Each list style: its levels' kinds, bullets, number formats, suffixes, starts and indents. */
function readListStyles(roots) {
  const out = new Map();
  for (const root of roots) {
    if (!root) continue;
    for (const ls of all(root, 'text:list-style')) {
      const levels = [];
      for (const lv of kids(ls)) {
        if (typeof lv === 'string') continue;
        const level = Number(lv.attrs['text:level'] || 1) - 1;
        if (level < 0 || level > 8) continue;
        const props = first(lv, 'style:list-level-properties');
        const align = props ? first(props, 'style:list-level-label-alignment') : null;
        const indent = align ? lengthPx(align.attrs['fo:margin-left']) : (lengthPx(props?.attrs['text:space-before']) ?? 0) + (lengthPx(props?.attrs['text:min-label-width']) ?? 0);
        const hanging = align ? -(lengthPx(align.attrs['fo:text-indent']) ?? 0) : lengthPx(props?.attrs['text:min-label-width']);
        if (lv.name === 'text:list-level-style-number') {
          levels[level] = {
            kind: 'number', format: lv.attrs['style:num-format'] ?? '1', prefix: lv.attrs['style:num-prefix'] || '', suffix: lv.attrs['style:num-suffix'] ?? '.',
            start: Number(lv.attrs['text:start-value'] || 1), display: Number(lv.attrs['text:display-levels'] || 1), indent, hanging,
          };
        } else if (lv.name === 'text:list-level-style-bullet') {
          levels[level] = { kind: 'bullet', char: lv.attrs['text:bullet-char'] || '•', indent, hanging };
        } else if (lv.name === 'text:list-level-style-image') {
          levels[level] = { kind: 'bullet', char: '•', indent, hanging };
        }
      }
      out.set(ls.attrs['style:name'], levels);
    }
  }
  return out;
}

/** The page: its size and margins, from the first page layout. */
function readPage(stylesRoot) {
  const props = stylesRoot ? first(stylesRoot, 'style:page-layout-properties') : null;
  if (!props) return null;
  const a = props.attrs;
  return {
    width: lengthPx(a['fo:page-width']), height: lengthPx(a['fo:page-height']),
    top: lengthPx(a['fo:margin-top']), bottom: lengthPx(a['fo:margin-bottom']), left: lengthPx(a['fo:margin-left']), right: lengthPx(a['fo:margin-right']),
    landscape: a['style:print-orientation'] === 'landscape',
  };
}

/* ── the body ──────────────────────────────────────────────────────────── */

/**
 * Read an .odt's body: { blocks, page } — blocks of
 * { type: 'paragraph', heading, align, indentLeft, indentRight, indentFirst,
 *   spaceBefore, spaceAfter, pageBreakBefore, rtl, list: { id, style, level },
 *   runs: [{ text, bold, … } | { tab } | { br } | { image: { href, width, height } }] }
 * and { type: 'table', columns, rows: [[{ blocks, colspan, rowspan, covered, fill, border }]] };
 * `lists` maps each list style's name to its levels.
 */
export function readOdtDocument(contentXml, stylesXml) {
  const content = parse(contentXml);
  const stylesRoot = stylesXml ? parse(stylesXml) : null;
  const automatic = first(content, 'office:automatic-styles');
  const named = stylesRoot ? first(stylesRoot, 'office:styles') : null;
  const stylesAuto = stylesRoot ? first(stylesRoot, 'office:automatic-styles') : null;
  const style = readStyles([{ root: named, named: true }, { root: stylesAuto, named: false }, { root: automatic, named: false }]);
  const lists = readListStyles([named, stylesAuto, automatic]);
  const body = first(first(content, 'office:body') || content, 'office:text') || content;
  let listCount = 0;
  const lastList = new Map(); // a list style → the list last written in it

  /** A paragraph's runs, its spans' looks over its own, its frames and breaks among them. */
  const runsOf = (node, base, extra) => {
    const runs = [];
    const push = (text, look) => {
      if (!text) return;
      const last = runs[runs.length - 1];
      if (last && last.text != null && JSON.stringify({ ...last, text: '' }) === JSON.stringify({ ...look, text: '' })) last.text += text;
      else runs.push({ ...look, text });
    };
    const walk = (n, look) => {
      for (const c of n.children || []) {
        if (typeof c === 'string') { push(c.replace(/[\r\n\t ]+/g, ' '), look); continue; }
        switch (c.name) {
          case 'text:span': walk(c, { ...look, ...style('text', c.attrs['text:style-name']).text }); break;
          case 'text:a': walk(c, { ...look, link: c.attrs['xlink:href'] || null }); break;
          case 'text:s': push(' '.repeat(Math.max(1, Number(c.attrs['text:c'] || 1))), look); break;
          case 'text:tab': runs.push({ ...look, tab: true }); break;
          case 'text:line-break': runs.push({ ...look, br: true }); break;
          case 'text:soft-page-break': break;
          case 'text:note': break; // the note's body has no place in the line
          case 'text:bookmark': case 'text:bookmark-start': case 'text:bookmark-end': break;
          case 'draw:frame': {
            const img = first(c, 'draw:image');
            const box = first(c, 'draw:text-box');
            if (img) runs.push({ image: { href: img.attrs['xlink:href'] || '', width: lengthPx(c.attrs['svg:width']) ?? 96, height: lengthPx(c.attrs['svg:height']) ?? 96, name: c.attrs['draw:name'] || null } });
            else if (box) extra.push(...blocksOf(box, null));
            break;
          }
          case 'draw:a': walk(c, look); break;
          default: walk(c, look);
        }
      }
    };
    walk(node, base);
    // Spaces a paragraph's markup left at its ends are not words.
    if (runs[0]?.text != null) runs[0].text = runs[0].text.replace(/^ +/, '');
    const last = runs[runs.length - 1];
    if (last?.text != null) last.text = last.text.replace(/ +$/, '');
    return runs.filter((r) => r.text !== '' || r.tab || r.br || r.image);
  };

  const paragraphOf = (node, list) => {
    const st = style('paragraph', node.attrs['text:style-name']);
    const heading = node.name === 'text:h' ? Math.min(9, Math.max(1, Number(node.attrs['text:outline-level'] || st.outline || 1))) : null;
    // A heading's look is its Heading style's; only what an automatic style adds on top comes with it.
    const base = heading ? style('paragraph', node.attrs['text:style-name']).text : st.text;
    const extra = [];
    const runs = runsOf(node, heading ? {} : base, extra);
    return [{ type: 'paragraph', heading, ...st.paragraph, ...(list ? { list } : {}), runs }, ...extra];
  };

  const tableOf = (node) => {
    // Its own columns and rows — directly, or in their header and group elements — not a nested table's.
    const own = (n, leaf, groups) => kids(n).flatMap((c) => (typeof c === 'string' ? [] : c.name === leaf ? [c] : groups.includes(c.name) ? own(c, leaf, groups) : []));
    const columns = [];
    for (const col of own(node, 'table:table-column', ['table:table-columns', 'table:table-header-columns', 'table:table-column-group'])) {
      const n = Math.min(64, Number(col.attrs['table:number-columns-repeated'] || 1));
      const w = style('table-column', col.attrs['table:style-name']).column;
      for (let i = 0; i < n; i++) columns.push(w || null);
    }
    const rows = [];
    for (const row of own(node, 'table:table-row', ['table:table-rows', 'table:table-header-rows', 'table:table-row-group'])) {
      const cells = [];
      for (const cell of kids(row)) {
        if (typeof cell === 'string') continue;
        if (cell.name === 'table:covered-table-cell') {
          const n = Number(cell.attrs['table:number-columns-repeated'] || 1);
          for (let i = 0; i < n; i++) cells.push({ covered: true });
        } else if (cell.name === 'table:table-cell') {
          const n = Math.min(64, Number(cell.attrs['table:number-columns-repeated'] || 1));
          const look = style('table-cell', cell.attrs['table:style-name']).cell;
          for (let i = 0; i < n; i++) {
            cells.push({
              blocks: blocksOf(cell, null),
              colspan: Number(cell.attrs['table:number-columns-spanned'] || 1),
              rowspan: Number(cell.attrs['table:number-rows-spanned'] || 1),
              ...look,
            });
          }
        }
      }
      rows.push(cells);
    }
    const width = Math.max(columns.length, ...rows.map((r) => r.length));
    while (columns.length < width) columns.push(null);
    return { type: 'table', columns, rows };
  };

  /** A list's items, each paragraph at the list's level and a nested list a level in. */
  const listOf = (node, styleName, level, id) => {
    const out = [];
    const name = node.attrs['text:style-name'] || styleName;
    for (const item of kids(node)) {
      if (typeof item === 'string' || (item.name !== 'text:list-item' && item.name !== 'text:list-header')) continue;
      for (const c of kids(item)) {
        if (typeof c === 'string') continue;
        if (c.name === 'text:p' || c.name === 'text:h') out.push(...paragraphOf(c, item.name === 'text:list-header' ? null : { id, style: name, level: Math.min(8, level) }));
        else if (c.name === 'text:list') out.push(...listOf(c, name, level + 1, id));
        else if (c.name === 'table:table') out.push(tableOf(c));
      }
    }
    return out;
  };

  function blocksOf(node, inherited) {
    const out = [];
    for (const c of kids(node)) {
      if (typeof c === 'string') continue;
      if (c.name === 'text:p' || c.name === 'text:h') out.push(...paragraphOf(c, inherited));
      else if (c.name === 'text:list') {
        // A list that carries on numbering after something else is the list it carries on.
        const name = c.attrs['text:style-name'] || null;
        const carries = c.attrs['text:continue-numbering'] === 'true' || c.attrs['text:continue-list'];
        const id = carries && lastList.has(name) ? lastList.get(name) : ++listCount;
        lastList.set(name, id);
        out.push(...listOf(c, null, 0, id));
      }
      else if (c.name === 'table:table') out.push(tableOf(c));
      else if (c.name === 'text:section' || c.name === 'text:index-body' || c.name === 'text:table-of-content' || c.name === 'text:illustration-index' || c.name === 'text:alphabetical-index') out.push(...blocksOf(c, inherited));
      else if (c.name === 'draw:frame') {
        const extra = [];
        const runs = runsOf({ children: [c] }, {}, extra);
        if (runs.length) out.push({ type: 'paragraph', heading: null, runs });
        out.push(...extra);
      }
    }
    return out;
  }

  return { blocks: blocksOf(body, null), lists, page: readPage(stylesRoot) };
}

/** An .odt file read whole: its body, list styles and page, and its pictures by their names in the archive. */
export function readOdt(bytes) {
  const { entries } = readZip(Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes));
  const byName = new Map(entries.map((e) => [e.name, e]));
  const text = (name) => (byName.has(name) ? byName.get(name).data.toString('utf8') : '');
  const content = text('content.xml');
  if (!content) throw new Error('OpenDocument file has no content.xml');
  const images = new Map();
  for (const [name, e] of byName) if (/^(Pictures|media|ObjectReplacements)\//.test(name)) images.set(name, new Uint8Array(e.data));
  const meta = text('meta.xml');
  const title = meta ? textOf(first(parse(meta), 'dc:title')) || null : null;
  return { ...readOdtDocument(content, text('styles.xml')), images, title };
}
