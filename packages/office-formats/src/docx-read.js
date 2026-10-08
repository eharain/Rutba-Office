// A .docx read into the document model odt.js reads an .odt into, so that
// what the Word engine holds can be written as an OpenDocument text whole.
//
// odt.js gives { blocks, lists, page, images, title }: paragraphs with their
// heading level, alignment, indents, spacing, page break and direction, a
// list item's list and level, and runs in their looks (with tabs, breaks,
// links and pictures among them); tables by grid position, a span's covered
// places marked; each list's levels; the page's size and margins. This reads
// the same from WordprocessingML:
//
//   - a paragraph's style along its basedOn chain (a heading by its outline
//     level or its name), its own properties over the style's, a run's
//     character style and direct properties over the paragraph style's;
//   - numbering: each w:num's abstract levels (format, text, start, indent),
//     a w:lvlOverride's start, a list item by its numId and ilvl;
//   - tables: the grid's widths, a cell's gridSpan as places to its right
//     and a w:vMerge as places below, its shading and whether it has borders;
//   - pictures (w:drawing's blip) at their extent, hyperlinks by their
//     relationship, inserted text kept and deleted text left out;
//   - the last section's page size, orientation and margins.

import { readZip } from '@rutba/ooxml/zip';
import { parse, all, kids, first, textOf } from './xml.js';

const twipsPx = (v) => (v == null || v === '' ? null : Number(v) / 15);
const emuPx = (v) => Number(v || 0) / 9525;
/** A direct child: a property never taken from a nested paragraph, cell or tracked change's record. */
const own = (node, name) => (node ? kids(node, name)[0] || null : null);
const on = (el) => Boolean(el) && !/^(0|false|off|none)$/i.test(el.attrs['w:val'] ?? 'true');

/** Word's highlight names as the colours they paint. */
const HIGHLIGHT = {
  yellow: '#FFFF00', green: '#00FF00', cyan: '#00FFFF', magenta: '#FF00FF', blue: '#0000FF', red: '#FF0000',
  darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
  darkYellow: '#808000', darkGray: '#808080', lightGray: '#C0C0C0', black: '#000000', white: '#FFFFFF',
};

/** WordprocessingML number formats as ODF's num-format. */
const NUMFMT = { decimal: '1', decimalZero: '1', lowerLetter: 'a', upperLetter: 'A', lowerRoman: 'i', upperRoman: 'I', ordinal: '1', cardinalText: '1', ordinalText: '1' };

/** A w:rPr as run looks; only what it says. */
function runLook(rPr) {
  if (!rPr) return {};
  const out = {};
  const b = own(rPr, 'w:b');
  if (b) out.bold = on(b);
  const i = own(rPr, 'w:i');
  if (i) out.italic = on(i);
  const u = own(rPr, 'w:u');
  if (u) out.underline = (u.attrs['w:val'] || 'single') !== 'none';
  const strike = own(rPr, 'w:strike') || own(rPr, 'w:dstrike');
  if (strike) out.strike = on(strike);
  const colour = own(rPr, 'w:color')?.attrs['w:val'];
  if (colour && /^[0-9a-f]{6}$/i.test(colour)) out.color = `#${colour.toUpperCase()}`;
  const sz = own(rPr, 'w:sz')?.attrs['w:val'];
  if (sz && Number(sz) > 0) out.size = Number(sz) / 2;
  const fonts = own(rPr, 'w:rFonts');
  const face = fonts && (fonts.attrs['w:ascii'] || fonts.attrs['w:hAnsi'] || fonts.attrs['w:cs']);
  if (face) out.font = face;
  const hl = own(rPr, 'w:highlight')?.attrs['w:val'];
  if (hl && HIGHLIGHT[hl]) out.highlight = HIGHLIGHT[hl];
  else {
    const fill = own(rPr, 'w:shd')?.attrs['w:fill'];
    if (fill && /^[0-9a-f]{6}$/i.test(fill)) out.highlight = `#${fill.toUpperCase()}`;
  }
  const va = own(rPr, 'w:vertAlign')?.attrs['w:val'];
  if (va === 'superscript' || va === 'subscript') out.vertical = va;
  return out;
}

/** A w:pPr as paragraph looks; only what it says. */
function paragraphLook(pPr) {
  if (!pPr) return {};
  const out = {};
  const jc = own(pPr, 'w:jc')?.attrs['w:val'];
  const align = { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'both', distribute: 'both' }[jc];
  if (align) out.align = align;
  const ind = own(pPr, 'w:ind');
  if (ind) {
    const a = ind.attrs;
    const left = a['w:start'] ?? a['w:left'];
    const right = a['w:end'] ?? a['w:right'];
    if (left != null) out.indentLeft = twipsPx(left);
    if (right != null) out.indentRight = twipsPx(right);
    if (a['w:hanging'] != null) out.indentFirst = -twipsPx(a['w:hanging']);
    else if (a['w:firstLine'] != null) out.indentFirst = twipsPx(a['w:firstLine']);
  }
  const spacing = own(pPr, 'w:spacing');
  if (spacing?.attrs['w:before'] != null) out.spaceBefore = twipsPx(spacing.attrs['w:before']);
  if (spacing?.attrs['w:after'] != null) out.spaceAfter = twipsPx(spacing.attrs['w:after']);
  if (on(own(pPr, 'w:pageBreakBefore'))) out.pageBreakBefore = true;
  if (on(own(pPr, 'w:bidi'))) out.rtl = true;
  return out;
}

/** The document's styles, each resolved along basedOn: paragraph and run looks, and a heading's level. */
function readStyles(root) {
  const raw = new Map();
  let defaults = { run: {}, paragraph: {} };
  if (root) {
    const dd = first(root, 'w:docDefaults');
    if (dd) defaults = { run: runLook(first(first(dd, 'w:rPrDefault') || dd, 'w:rPr')), paragraph: paragraphLook(first(first(dd, 'w:pPrDefault') || dd, 'w:pPr')) };
    for (const st of all(root, 'w:style')) {
      const pPr = own(st, 'w:pPr');
      const name = own(st, 'w:name')?.attrs['w:val'] || '';
      const outline = own(pPr, 'w:outlineLvl')?.attrs['w:val'];
      const byName = /^heading\s*([1-9])$/i.exec(name);
      raw.set(st.attrs['w:styleId'], {
        type: st.attrs['w:type'],
        basedOn: own(st, 'w:basedOn')?.attrs['w:val'] || null,
        run: runLook(own(st, 'w:rPr')),
        paragraph: paragraphLook(pPr),
        heading: outline != null && Number(outline) < 9 ? Number(outline) + 1 : byName ? Number(byName[1]) : null,
        numPr: own(pPr, 'w:numPr'),
      });
    }
  }
  const cache = new Map();
  const resolve = (id, depth = 0) => {
    if (!id || !raw.has(id) || depth > 12) return { run: {}, paragraph: {}, heading: null, numPr: null };
    if (cache.has(id)) return cache.get(id);
    const self = raw.get(id);
    const up = resolve(self.basedOn, depth + 1);
    const out = { run: { ...up.run, ...self.run }, paragraph: { ...up.paragraph, ...self.paragraph }, heading: self.heading ?? up.heading, numPr: self.numPr ?? up.numPr };
    cache.set(id, out);
    return out;
  };
  return { resolve, defaults };
}

/** Each w:num as an ODF list style's levels, by numId. */
function readNumbering(root) {
  const lists = new Map();
  if (!root) return lists;
  const abstract = new Map();
  for (const an of all(root, 'w:abstractNum')) abstract.set(an.attrs['w:abstractNumId'], an);
  const levelOf = (lvl) => {
    const fmt = first(lvl, 'w:numFmt')?.attrs['w:val'] || 'decimal';
    const text = first(lvl, 'w:lvlText')?.attrs['w:val'] ?? '';
    const ind = first(first(lvl, 'w:pPr') || lvl, 'w:ind');
    const indent = ind ? twipsPx(ind.attrs['w:start'] ?? ind.attrs['w:left']) ?? 0 : 0;
    const hanging = ind && ind.attrs['w:hanging'] != null ? twipsPx(ind.attrs['w:hanging']) : 24;
    if (fmt === 'bullet' || fmt === 'none') return { kind: 'bullet', char: fmt === 'none' ? ' ' : (text || '•'), indent, hanging };
    // "%1." is a prefix, this level's number and a suffix; "%1.%2." shows the levels above too.
    const self = Number(lvl.attrs['w:ilvl'] || 0) + 1;
    const at = text.lastIndexOf(`%${self}`);
    const firstMark = text.indexOf('%');
    const shown = (text.match(/%[1-9]/g) || []).length;
    return {
      kind: 'number', format: NUMFMT[fmt] || '1',
      prefix: firstMark > 0 ? text.slice(0, firstMark) : '', suffix: at >= 0 ? text.slice(at + 2) : '.',
      start: Number(first(lvl, 'w:start')?.attrs['w:val'] || 1), display: Math.max(1, shown), indent, hanging,
    };
  };
  for (const num of all(root, 'w:num')) {
    const an = abstract.get(first(num, 'w:abstractNumId')?.attrs['w:val']);
    if (!an) continue;
    const levels = [];
    for (const lvl of kids(an, 'w:lvl')) {
      const i = Number(lvl.attrs['w:ilvl'] || 0);
      if (i >= 0 && i < 9) levels[i] = levelOf(lvl);
    }
    for (const o of kids(num, 'w:lvlOverride')) {
      const i = Number(o.attrs['w:ilvl'] || 0);
      const start = first(o, 'w:startOverride')?.attrs['w:val'];
      const lvl = first(o, 'w:lvl');
      if (lvl) levels[i] = levelOf(lvl);
      if (start != null && levels[i]?.kind === 'number') levels[i] = { ...levels[i], start: Number(start) };
    }
    lists.set(`WWNum${num.attrs['w:numId']}`, levels);
  }
  return lists;
}

/**
 * Read a .docx whole: { blocks, lists, page, images, title } in odt.js's
 * shapes, its pictures named as an OpenDocument keeps them (Pictures/…),
 * and `defaults.run`, the look every run has unless it says otherwise.
 */
export function readDocxDocument(bytes) {
  const { entries } = readZip(Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes));
  const byName = new Map(entries.map((e) => [e.name, e]));
  const text = (name) => (byName.has(name) ? byName.get(name).data.toString('utf8') : '');
  const docXml = text('word/document.xml');
  if (!docXml) throw new Error('Word document has no word/document.xml');
  const doc = parse(docXml);
  const { resolve, defaults } = readStyles(text('word/styles.xml') ? parse(text('word/styles.xml')) : null);
  // The document's own default look — its defaults and its Normal style's — is the written file's default style, not each run's.
  const plain = { ...defaults.run, ...resolve('Normal').run };
  const beyond = (look) => Object.fromEntries(Object.entries(look).filter(([k, v]) => plain[k] !== v));
  const lists = readNumbering(text('word/numbering.xml') ? parse(text('word/numbering.xml')) : null);
  const rels = new Map();
  const relsXml = text('word/_rels/document.xml.rels');
  if (relsXml) for (const r of all(parse(relsXml), 'Relationship')) rels.set(r.attrs.Id, { target: r.attrs.Target, external: r.attrs.TargetMode === 'External' });
  const images = new Map();
  const picture = (rid) => {
    const rel = rels.get(rid);
    if (!rel || rel.external) return null;
    const part = rel.target.startsWith('/') ? rel.target.slice(1) : `word/${rel.target}`.replace(/\/[^/]+\/\.\.\//g, '/');
    const e = byName.get(part);
    if (!e) return null;
    const name = `Pictures/${part.split('/').pop()}`;
    images.set(name, new Uint8Array(e.data));
    return name;
  };

  let pendingBreak = false;

  /** A paragraph's runs: text, tabs, breaks, pictures and links, each run in its look over the paragraph style's. */
  const runsOf = (p, base) => {
    const runs = [];
    const push = (t, look) => {
      if (!t) return;
      const last = runs[runs.length - 1];
      if (last && last.text != null && JSON.stringify({ ...last, text: '' }) === JSON.stringify({ ...look, text: '' })) last.text += t;
      else runs.push({ ...look, text: t });
    };
    const run = (r, extra) => {
      const rPr = own(r, 'w:rPr');
      const look = { ...base, ...resolve(own(rPr, 'w:rStyle')?.attrs['w:val']).run, ...runLook(rPr), ...extra };
      // A link in Word's own link look is a link: the written file's link style gives it that look.
      if (look.link && look.color === '#0563C1' && look.underline) { delete look.color; delete look.underline; }
      for (const c of kids(r)) {
        if (typeof c === 'string') continue;
        switch (c.name) {
          case 'w:t': push(textOf(c), look); break;
          case 'w:tab': runs.push({ ...look, tab: true }); break;
          case 'w:br': case 'w:cr':
            if (c.attrs['w:type'] === 'page') pendingBreak = true;
            else runs.push({ ...look, br: true });
            break;
          case 'w:noBreakHyphen': push('-', look); break;
          case 'w:sym': push(String.fromCharCode(Number.parseInt(c.attrs['w:char'] || '20', 16) & 0xffff), look); break;
          case 'w:drawing': {
            const blip = all(c, 'a:blip')[0];
            const extent = first(c, 'wp:extent');
            const name = blip ? picture(blip.attrs['r:embed']) : null;
            if (name) runs.push({ image: { href: name, width: emuPx(extent?.attrs.cx) || 96, height: emuPx(extent?.attrs.cy) || 96, name: first(c, 'wp:docPr')?.attrs.name || null } });
            break;
          }
          default: break;
        }
      }
    };
    const walk = (node, extra) => {
      for (const c of kids(node)) {
        if (typeof c === 'string') continue;
        if (c.name === 'w:r') run(c, extra);
        else if (c.name === 'w:hyperlink') {
          const rel = rels.get(c.attrs['r:id']);
          const link = rel ? rel.target + (c.attrs['w:anchor'] ? `#${c.attrs['w:anchor']}` : '') : c.attrs['w:anchor'] ? `#${c.attrs['w:anchor']}` : null;
          walk(c, link ? { ...extra, link } : extra);
        } else if (c.name === 'w:del' || c.name === 'w:moveFrom' || c.name === 'w:pPr') continue;
        else if (c.name === 'w:ins' || c.name === 'w:moveTo' || c.name === 'w:smartTag' || c.name === 'w:customXml' || c.name === 'w:fldSimple' || c.name === 'w:sdtContent') walk(c, extra);
        else if (c.name === 'w:sdt') walk(own(c, 'w:sdtContent') || c, extra);
      }
    };
    walk(p, {});
    return runs.filter((r) => r.text !== '' || r.tab || r.br || r.image);
  };

  const paragraphOf = (p) => {
    const pPr = own(p, 'w:pPr');
    const st = resolve(own(pPr, 'w:pStyle')?.attrs['w:val'] || 'Normal');
    const direct = paragraphLook(pPr);
    const heading = (() => {
      const outline = own(pPr, 'w:outlineLvl')?.attrs['w:val'];
      return outline != null && Number(outline) < 9 ? Number(outline) + 1 : st.heading;
    })();
    // A heading's look is its Heading style's in the file written; a body paragraph's style look rides its runs.
    const base = heading ? {} : beyond(st.run);
    const numPr = own(pPr, 'w:numPr') || st.numPr;
    const numId = own(numPr, 'w:numId')?.attrs['w:val'] ?? null;
    const level = Number(own(numPr, 'w:ilvl')?.attrs['w:val'] || 0);
    const list = numId && numId !== '0' && lists.has(`WWNum${numId}`) ? { id: `n${numId}`, style: `WWNum${numId}`, level: Math.min(8, level) } : null;
    const look = { ...st.paragraph, ...direct };
    // A list item's indents are its list level's, as the file written gives them.
    if (list) { delete look.indentLeft; delete look.indentFirst; }
    const runs = runsOf(p, base);
    const block = { type: 'paragraph', heading: heading || null, ...look, ...(list ? { list } : {}), runs };
    if (pendingBreak) { block.pageBreakBefore = true; pendingBreak = false; }
    // A break at the end of this paragraph starts the next one on a fresh page.
    return block;
  };

  const tableOf = (tbl) => {
    const columns = kids(own(tbl, 'w:tblGrid'), 'w:gridCol').map((g) => twipsPx(g.attrs['w:w']) || null);
    const rows = [];
    const anchors = new Map(); // grid column → the cell a vertical merge started in
    for (const tr of kids(tbl, 'w:tr')) {
      const cells = [];
      let col = 0;
      const before = own(own(tr, 'w:trPr'), 'w:gridBefore')?.attrs['w:val'];
      for (let k = 0; k < Number(before || 0); k++) { cells.push({ blocks: [], colspan: 1, rowspan: 1 }); col += 1; }
      const tcs = kids(tr).flatMap((c) => (typeof c === 'string' ? [] : c.name === 'w:tc' ? [c] : c.name === 'w:sdt' ? kids(own(c, 'w:sdtContent') || c, 'w:tc') : []));
      for (const tc of tcs) {
        const tcPr = own(tc, 'w:tcPr');
        const span = Math.max(1, Number(own(tcPr, 'w:gridSpan')?.attrs['w:val'] || 1));
        const vm = own(tcPr, 'w:vMerge');
        if (vm && (vm.attrs['w:val'] || 'continue') === 'continue' && anchors.has(col)) {
          anchors.get(col).rowspan += 1;
          for (let k = 0; k < span; k++) cells.push({ covered: true });
        } else {
          const fill = own(tcPr, 'w:shd')?.attrs['w:fill'];
          const cell = {
            blocks: blocksOf(tc),
            colspan: span, rowspan: 1,
            ...(fill && /^[0-9a-f]{6}$/i.test(fill) ? { fill: `#${fill.toUpperCase()}` } : {}),
            ...(own(tcPr, 'w:tcBorders') ? { border: true } : {}),
          };
          cells.push(cell);
          for (let k = 1; k < span; k++) cells.push({ covered: true });
          if (vm) anchors.set(col, cell);
          else for (let k = 0; k < span; k++) anchors.delete(col + k);
        }
        col += span;
      }
      rows.push(cells);
    }
    const width = Math.max(columns.length, ...rows.map((r) => r.length));
    while (columns.length < width) columns.push(null);
    // A table that says it has borders gives every cell one.
    const borders = own(own(tbl, 'w:tblPr'), 'w:tblBorders');
    if (borders && kids(borders).some((b) => typeof b !== 'string' && b.attrs['w:val'] && b.attrs['w:val'] !== 'nil' && b.attrs['w:val'] !== 'none')) {
      for (const row of rows) for (const cell of row) if (!cell.covered) cell.border = true;
    }
    return { type: 'table', columns, rows };
  };

  function blocksOf(node) {
    const out = [];
    for (const c of kids(node)) {
      if (typeof c === 'string') continue;
      if (c.name === 'w:p') out.push(paragraphOf(c));
      else if (c.name === 'w:tbl') out.push(tableOf(c));
      else if (c.name === 'w:sdt') out.push(...blocksOf(own(c, 'w:sdtContent') || c));
      else if (c.name === 'w:customXml' || c.name === 'w:ins' || c.name === 'w:moveTo') out.push(...blocksOf(c));
    }
    return out;
  }

  const body = first(doc, 'w:body') || doc;
  const blocks = blocksOf(body);
  const sect = own(body, 'w:sectPr') || all(body, 'w:sectPr').pop();
  let page = null;
  if (sect) {
    const sz = own(sect, 'w:pgSz');
    const mar = own(sect, 'w:pgMar');
    page = {
      width: twipsPx(sz?.attrs['w:w']) ?? 816, height: twipsPx(sz?.attrs['w:h']) ?? 1056,
      top: twipsPx(mar?.attrs['w:top']) ?? 96, bottom: twipsPx(mar?.attrs['w:bottom']) ?? 96,
      left: twipsPx(mar?.attrs['w:left']) ?? 96, right: twipsPx(mar?.attrs['w:right']) ?? 96,
      landscape: sz?.attrs['w:orient'] === 'landscape',
    };
  }
  const core = text('docProps/core.xml');
  const title = core ? textOf(first(parse(core), 'dc:title')) || null : null;
  return { blocks, lists, page, images, title, defaults: { run: plain } };
}
