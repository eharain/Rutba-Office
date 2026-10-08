// Rich Text read into the document model odt.js reads an .odt into, so that
// an .rtf opens as it was made — Word's, WordPad's or this suite's own.
//
// rtf.js runs the stream for its words and their looks: paragraphs of runs,
// a table as rows of cells. This reads the file whole. The stream is first
// gathered into its groups (each \uN with its fallback skipped, each \'xx
// decoded in the file's code page), then run with a state per group:
//
//   - the font, colour, style, list and list-override tables, read before
//     the body is, so a run's \fN and \cfN and a paragraph's \sN and \lsN
//     mean what the file says;
//   - paragraphs in their alignment, indents, spacing, page break and
//     direction, a heading by its outline level or its style's name, a list
//     item by its list override and level (Word's \listtext label left out,
//     since the list draws its own), WordPad's \pntext lists as lists too;
//   - runs in their looks — bold, italic, underline, strike, colour, size,
//     face, highlight, raised and lowered — tabs and line breaks among them;
//   - HYPERLINK fields as links, their results as the linked words;
//   - PNG and JPEG pictures at the size they are shown;
//   - tables from their rows' cell edges: the columns every row's edges make,
//     a cell's span across them, a merge down (\clvmgf/\clvmrg) and across
//     (\clmgf/\clmrg), shading and borders, each cell's own paragraphs;
//   - the page's size, orientation and margins, and the title.

import { decoderFor } from './codepage.js';

const CODEPAGES = { ansi: 'windows-1252', mac: 'macintosh', pc: 'ibm437', pca: 'ibm850', 1250: 'windows-1250', 1251: 'windows-1251', 1252: 'windows-1252', 1253: 'windows-1253', 1254: 'windows-1254', 1255: 'windows-1255', 1256: 'windows-1256', 1257: 'windows-1257', 1258: 'windows-1258', 874: 'windows-874', 932: 'shift_jis', 936: 'gbk', 949: 'euc-kr', 950: 'big5', 65001: 'utf-8' };
/** Groups whose words are not the document's. */
const SKIP = new Set(['fonttbl', 'colortbl', 'stylesheet', 'listtable', 'listoverridetable', 'info', 'header', 'headerl', 'headerr', 'headerf', 'footer', 'footerl', 'footerr', 'footerf', 'footnote', 'annotation', 'atnid', 'atnauthor', 'xmlns', 'themedata', 'colorschememapping', 'latentstyles', 'datastore', 'generator', 'rsidtbl', 'filetbl', 'revtbl', 'userprops', 'docvar', 'pn', 'pntext', 'listtext', 'nonshppict', 'object', 'nonesttables', 'mmathPr', 'bkmkstart', 'bkmkend', 'xe', 'tc', 'pgdsctbl', 'wgrffmtfilter', 'fchars', 'lchars', 'aftnsep', 'ftnsep', 'ftnsepc', 'aftnsepc', 'ftncn', 'aftncn', 'listpicture', 'protusertbl', 'mmath', 'shp', 'shpinst']);
/** \levelnfc numbers as ODF's number formats; 23 is a bullet, 255 none. */
const NFC = { 0: '1', 1: 'I', 2: 'i', 3: 'A', 4: 'a', 22: '1' };
const SYMBOL_BULLETS = { 0xf0b7: '•', 0xf0a7: '▪', 0xf06f: 'o', 0xf0d8: '➢', 0xf0fc: '✓', 0xf076: '❖', 0xb7: '•' };

/** The stream as nested groups of { word, num } control words and decoded text. */
function groupsOf(text, decoder) {
  const root = { items: [] };
  const stack = [root];
  let cur = root;
  let bytes = [];
  let uc = 1;
  let skip = 0;
  const flush = () => {
    if (bytes.length) { cur.items.push({ text: decoder.decode(Uint8Array.from(bytes)) }); bytes = []; }
  };
  let i = 0;
  while (i < text.length) {
    const ch = text[i];
    if (ch === '{') { flush(); const g = { items: [], uc }; cur.items.push(g); stack.push(g); cur = g; skip = 0; i += 1; continue; }
    if (ch === '}') { flush(); stack.pop(); cur = stack[stack.length - 1] || root; uc = cur.uc ?? 1; skip = 0; i += 1; continue; }
    if (ch === '\r' || ch === '\n') { i += 1; continue; }
    if (ch !== '\\') {
      if (skip > 0) skip -= 1; else bytes.push(text.charCodeAt(i) & 0xff);
      i += 1;
      continue;
    }
    const next = text[i + 1];
    if (next === '\\' || next === '{' || next === '}') { if (skip > 0) skip -= 1; else bytes.push(next.charCodeAt(0)); i += 2; continue; }
    if (next === "'") {
      const byte = parseInt(text.slice(i + 2, i + 4), 16);
      if (skip > 0) skip -= 1; else if (Number.isFinite(byte)) bytes.push(byte);
      i += 4;
      continue;
    }
    if (next === '\r' || next === '\n') { flush(); cur.items.push({ word: 'par', num: null }); i += 2; continue; }
    if (next === '~') { flush(); cur.items.push({ text: ' ' }); i += 2; continue; }
    if (next === '_') { flush(); cur.items.push({ text: '‑' }); i += 2; continue; }
    if (next === '-') { i += 2; continue; }
    if (next === '*') { flush(); cur.items.push({ word: '*', num: null }); i += 2; continue; }
    let j = i + 1;
    while (j < text.length && /[a-zA-Z]/.test(text[j])) j += 1;
    const word = text.slice(i + 1, j);
    let num = '';
    if (text[j] === '-') { num = '-'; j += 1; }
    while (j < text.length && /[0-9]/.test(text[j])) { num += text[j]; j += 1; }
    if (text[j] === ' ') j += 1;
    i = j;
    if (!word) { i += 1; continue; }
    const n = num === '' || num === '-' ? null : Number.parseInt(num, 10);
    if (word === 'bin' && n > 0) { i += n; continue; }
    if (word === 'u' && n != null) { flush(); if (skip <= 0) cur.items.push({ text: String.fromCodePoint(n < 0 ? n + 65536 : n) }); skip = uc; continue; }
    if (word === 'uc') { uc = n ?? 1; cur.uc = uc; continue; }
    if (skip > 0) { skip -= 1; continue; }
    flush();
    cur.items.push({ word, num: n });
  }
  flush();
  return root;
}

const firstWord = (g) => {
  for (const it of g.items) {
    if (it.word === '*') continue;
    if (it.word) return it.word;
    if (it.items || (it.text && it.text.trim())) return null;
  }
  return null;
};
const starred = (g) => g.items[0]?.word === '*';
const textIn = (g) => g.items.map((it) => (it.text != null ? it.text : it.items ? textIn(it) : '')).join('');
const find = (g, name, out = []) => {
  for (const it of g.items || []) {
    if (!it.items) continue;
    if (firstWord(it) === name) out.push(it);
    else find(it, name, out);
  }
  return out;
};
const wordIn = (g, name) => g.items.find((it) => it.word === name)?.num;

/** The font table: each font's name by its number. */
function fontsOf(root) {
  const fonts = new Map();
  for (const table of find(root, 'fonttbl')) {
    const entries = table.items.filter((it) => it.items);
    // A table written without a group per font is one run of words and names.
    if (!entries.length) entries.push(table);
    for (const f of entries) {
      let at = null;
      let name = '';
      for (const it of f.items) {
        if (it.word === 'f' && at == null) at = it.num;
        else if (it.text != null) name += it.text;
        else if (it.items && starred(it)) continue;
      }
      if (at != null) fonts.set(at, name.replace(/;.*$/s, '').trim());
    }
  }
  return fonts;
}

/** The colour table: index 0 is the reader's own colour, the rest in order. */
function coloursOf(root) {
  const colours = [];
  for (const table of find(root, 'colortbl')) {
    let rgb = null;
    for (const it of table.items) {
      if (it.word === 'red' || it.word === 'green' || it.word === 'blue') { rgb = rgb || { red: 0, green: 0, blue: 0 }; rgb[it.word] = it.num || 0; }
      else if (it.text != null) {
        for (const c of it.text) if (c === ';') { colours.push(rgb ? `#${[rgb.red, rgb.green, rgb.blue].map((v) => Math.max(0, Math.min(255, v)).toString(16).padStart(2, '0')).join('').toUpperCase()}` : null); rgb = null; }
      }
    }
  }
  return colours;
}

/** The style sheet: a paragraph style's heading level, by its number. */
function headingStylesOf(root) {
  const out = new Map();
  for (const sheet of find(root, 'stylesheet')) {
    for (const st of sheet.items.filter((it) => it.items)) {
      const s = wordIn(st, 's');
      if (s == null) continue;
      const outline = wordIn(st, 'outlinelevel');
      const name = textIn(st).replace(/;.*$/s, '').trim();
      const m = /^heading\s*([1-9])$/i.exec(name);
      if (outline != null && outline < 9) out.set(s, outline + 1);
      else if (m) out.set(s, Number(m[1]));
    }
  }
  return out;
}

/** The list tables: each override's levels, as ODF list levels, by its \ls number. */
function listsOf(root) {
  const byId = new Map();
  for (const table of find(root, 'listtable')) {
    for (const list of table.items.filter((it) => it.items && firstWord(it) === 'list')) {
      const levels = [];
      for (const lv of list.items.filter((it) => it.items && firstWord(it) === 'listlevel')) {
        const nfc = wordIn(lv, 'levelnfcn') ?? wordIn(lv, 'levelnfc') ?? 0;
        const start = wordIn(lv, 'levelstartat') ?? 1;
        const li = wordIn(lv, 'li') ?? wordIn(lv, 'lin') ?? 720 * (levels.length + 1);
        const fi = wordIn(lv, 'fi') ?? -360;
        const lt = lv.items.find((it) => it.items && firstWord(it) === 'leveltext');
        const raw = lt ? textIn(lt).replace(/;$/, '') : '';
        // The first character counts the rest; a character below ten stands for a level's number.
        const body = raw.length ? raw.slice(1, 1 + raw.charCodeAt(0)) : '';
        const level = levels.length;
        if (nfc === 23 || nfc === 255) {
          const c = body.codePointAt(0);
          levels.push({ kind: 'bullet', char: nfc === 255 ? ' ' : SYMBOL_BULLETS[c] || (c && c >= 32 ? String.fromCodePoint(c) : '•'), indent: li / 15, hanging: -fi / 15 });
        } else {
          const marks = [...body].map((ch, k) => ({ ch, k })).filter(({ ch }) => ch.charCodeAt(0) < 10);
          const own = marks.findLast?.(({ ch }) => ch.charCodeAt(0) === level) ?? [...marks].reverse().find(({ ch }) => ch.charCodeAt(0) === level);
          levels.push({
            kind: 'number', format: NFC[nfc] || '1',
            prefix: marks.length ? body.slice(0, marks[0].k) : '', suffix: own ? body.slice(own.k + 1) : '.',
            start, display: Math.max(1, marks.length), indent: li / 15, hanging: -fi / 15,
          });
        }
      }
      byId.set(wordIn(list, 'listid'), levels);
    }
  }
  const byLs = new Map();
  for (const table of find(root, 'listoverridetable')) {
    for (const o of table.items.filter((it) => it.items && firstWord(it) === 'listoverride')) {
      const levels = byId.get(wordIn(o, 'listid'));
      if (!levels) continue;
      // An override may start its levels elsewhere: its \lfolevel groups, a level each in order.
      const own = levels.map((l) => ({ ...l }));
      o.items.filter((it) => it.items && firstWord(it) === 'lfolevel').forEach((lfo, k) => {
        const start = wordIn(lfo, 'levelstartat');
        if (lfo.items.some((it) => it.word === 'listoverridestartat') && start != null && own[k]?.kind === 'number') own[k].start = start;
      });
      byLs.set(wordIn(o, 'ls'), own);
    }
  }
  return byLs;
}

/**
 * Read an .rtf whole: { blocks, lists, page, images, title, defaults } in
 * odt.js's shapes.
 */
export function readRtfDocument(input) {
  const text = typeof input === 'string' ? input : new TextDecoder('latin1').decode(input);
  const cp = /\\(ansicpg(\d+)|mac|pc|pca|ansi)\b/.exec(text.slice(0, 4096));
  const decoder = decoderFor(cp ? CODEPAGES[cp[2] ? Number(cp[2]) : cp[1]] || 'windows-1252' : 'windows-1252');
  const root = groupsOf(text, decoder);
  const fonts = fontsOf(root);
  const colours = coloursOf(root);
  const headingStyles = headingStylesOf(root);
  const rtfLists = listsOf(root);
  const info = find(root, 'info')[0];
  const titleGroup = info ? find(info, 'title')[0] : null;
  const title = titleGroup ? textIn(titleGroup).trim() || null : null;
  const top = root.items.find((it) => it.items) || root;
  const deff = wordIn(top, 'deff') ?? 0;

  const lists = new Map();
  const images = new Map();
  const blocks = [];
  let imageCount = 0;
  let page = null;
  const pageWords = {};

  // The paragraph and run being written, and the table being gathered.
  let runs = [];
  let cellBlocks = [];
  let rowCells = [];
  let rowDefs = [];
  let cellDef = {};
  let tableRows = [];
  let pnLabel = null;

  const fresh = () => ({ bold: false, italic: false, underline: false, strike: false, color: 0, size: 0, font: null, highlight: 0, vertical: null, link: null });
  const freshPara = () => ({ align: null, li: 0, ri: 0, fi: 0, sb: null, sa: null, pagebb: false, rtl: false, outline: null, style: null, ls: null, ilvl: 0, intbl: false });

  const lookOf = (st) => {
    const look = {};
    if (st.bold) look.bold = true;
    if (st.italic) look.italic = true;
    if (st.underline) look.underline = true;
    if (st.strike) look.strike = true;
    if (st.color && colours[st.color]) look.color = colours[st.color];
    if (st.size) look.size = st.size;
    if (st.font != null && st.font !== deff && fonts.get(st.font)) look.font = fonts.get(st.font);
    if (st.highlight && colours[st.highlight]) look.highlight = colours[st.highlight];
    if (st.vertical) look.vertical = st.vertical;
    if (st.link) look.link = st.link;
    return look;
  };
  const push = (t, st) => {
    if (!t) return;
    const look = lookOf(st);
    const last = runs[runs.length - 1];
    if (last && last.text != null && JSON.stringify({ ...last, text: '' }) === JSON.stringify({ ...look, text: '' })) last.text += t;
    else runs.push({ ...look, text: t });
  };

  const closeTable = () => {
    if (!tableRows.length) return;
    // The columns every row's cell edges make, left to right.
    const edges = [...new Set(tableRows.flatMap((r) => r.defs.map((d) => d.right)))].filter((x) => Number.isFinite(x)).sort((a, b) => a - b);
    const left0 = Math.min(0, ...tableRows.map((r) => r.left ?? 0));
    const columns = edges.map((e, k) => Math.max(1, (e - (k ? edges[k - 1] : left0)) / 15));
    const grid = tableRows.map(() => []);
    const anchors = new Map(); // grid column → the cell a merge down started in
    tableRows.forEach((row, ri) => {
      let col = 0;
      let prev = null;
      row.defs.forEach((def, k) => {
        const end = edges.indexOf(def.right);
        const span = Math.max(1, end - col + 1);
        const content = row.cells[k] || { blocks: [] };
        if (def.hmerge === 'cont' && prev) {
          prev.colspan += span;
          for (let s = 0; s < span; s++) grid[ri][col + s] = { covered: true };
        } else if (def.vmerge === 'cont' && anchors.has(col)) {
          anchors.get(col).rowspan += 1;
          for (let s = 0; s < span; s++) grid[ri][col + s] = { covered: true };
          prev = null;
        } else {
          const cell = { blocks: content.blocks, colspan: span, rowspan: 1, ...(def.fill && colours[def.fill] ? { fill: colours[def.fill] } : {}), ...(def.border ? { border: true } : {}) };
          grid[ri][col] = cell;
          for (let s = 1; s < span; s++) grid[ri][col + s] = { covered: true };
          if (def.vmerge === 'first') anchors.set(col, cell); else anchors.delete(col);
          prev = cell;
        }
        col = Math.max(col + span, end + 1);
      });
      for (let c = 0; c < edges.length; c++) if (!grid[ri][c]) grid[ri][c] = { blocks: [], colspan: 1, rowspan: 1 };
    });
    blocks.push({ type: 'table', columns, rows: grid });
    tableRows = [];
  };

  const endParagraph = (para) => {
    // A \pntext label made this a list item, as WordPad writes one.
    let list = null;
    if (para.ls != null && rtfLists.has(para.ls)) {
      const style = `RTFList${para.ls}`;
      if (!lists.has(style)) lists.set(style, rtfLists.get(para.ls));
      list = { id: `r${para.ls}`, style, level: Math.max(0, Math.min(8, para.ilvl || 0)) };
    } else if (pnLabel != null) {
      const numbered = /\d|^[a-z]\W*$|^[ivxlc]+\W*$/i.test(pnLabel.trim());
      const style = numbered ? 'RTFpnNumber' : 'RTFpnBullet';
      if (!lists.has(style)) lists.set(style, [numbered ? { kind: 'number', format: /[a-z]/.test(pnLabel) ? 'a' : /[A-Z]/.test(pnLabel) ? 'A' : '1', prefix: '', suffix: /\)/.test(pnLabel) ? ')' : '.', start: 1, display: 1, indent: 48, hanging: 24 } : { kind: 'bullet', char: '•', indent: 48, hanging: 24 }]);
      list = { id: style, style, level: 0 };
    }
    pnLabel = null;
    // Spaces a writer left at a paragraph's end are not words.
    const last = runs[runs.length - 1];
    if (last?.text != null) last.text = last.text.replace(/ +$/, '');
    const out = runs.filter((r) => r.text !== '' || r.tab || r.br || r.image);
    runs = [];
    const heading = para.outline != null && para.outline < 9 ? para.outline + 1 : para.style != null ? headingStyles.get(para.style) ?? null : null;
    const block = {
      type: 'paragraph', heading,
      ...(para.align ? { align: para.align } : {}),
      ...(!list && para.li ? { indentLeft: para.li / 15 } : {}),
      ...(!list && para.fi ? { indentFirst: para.fi / 15 } : {}),
      ...(para.ri ? { indentRight: para.ri / 15 } : {}),
      ...(para.sb != null ? { spaceBefore: para.sb / 15 } : {}),
      ...(para.sa != null ? { spaceAfter: para.sa / 15 } : {}),
      ...(para.pagebb ? { pageBreakBefore: true } : {}),
      ...(para.rtl ? { rtl: true } : {}),
      ...(list ? { list } : {}),
      runs: out,
    };
    if (para.intbl) cellBlocks.push(block);
    else { closeTable(); blocks.push(block); }
  };

  /** Run a group's words with its own copy of the state. */
  const run = (group, st, para) => {
    const w0 = firstWord(group);
    if (w0 && (SKIP.has(w0) || (starred(group) && !['shppict', 'fldinst', 'pgdsc'].includes(w0)))) {
      if (w0 === 'pntext') pnLabel = textIn(group);
      return para;
    }
    if (w0 === 'pict') { picture(group, st); return para; }
    if (w0 === 'field') { field(group, st, para); return para; }
    st = { ...st };
    for (const it of group.items) {
      if (it.items) { para = run(it, st, para); continue; }
      if (it.text != null) { push(it.text, st); continue; }
      const { word, num } = it;
      switch (word) {
        case 'par': endParagraph(para); para = { ...para }; break;
        case 'sect': case 'page': endParagraph(para); para = { ...para, pagebb: false }; if (word === 'page') para.pagebb = true; break;
        case 'line': runs.push({ ...lookOf(st), br: true }); break;
        case 'tab': runs.push({ ...lookOf(st), tab: true }); break;
        case 'emdash': push('—', st); break;
        case 'endash': push('–', st); break;
        case 'bullet': push('•', st); break;
        case 'lquote': push('‘', st); break;
        case 'rquote': push('’', st); break;
        case 'ldblquote': push('“', st); break;
        case 'rdblquote': push('”', st); break;
        case 'pard': para = { ...freshPara() }; break;
        case 'plain': Object.assign(st, fresh(), { link: st.link }); break;
        case 'b': st.bold = num !== 0; break;
        case 'i': st.italic = num !== 0; break;
        case 'ul': case 'uld': case 'uldb': case 'uldash': case 'ulw': case 'ulwave': case 'ulth': st.underline = num !== 0; break;
        case 'ulnone': st.underline = false; break;
        case 'strike': case 'striked': st.strike = num !== 0; break;
        case 'cf': st.color = num || 0; break;
        case 'fs': st.size = (num ?? 24) / 2; break;
        case 'f': st.font = num; break;
        case 'highlight': case 'cb': case 'chcbpat': st.highlight = num || 0; break;
        case 'super': st.vertical = 'superscript'; break;
        case 'sub': st.vertical = 'subscript'; break;
        case 'nosupersub': st.vertical = null; break;
        case 'up': st.vertical = num ? 'superscript' : null; break;
        case 'dn': st.vertical = num ? 'subscript' : null; break;
        case 'ql': para.align = 'left'; break;
        case 'qc': para.align = 'center'; break;
        case 'qr': para.align = 'right'; break;
        case 'qj': case 'qd': para.align = 'both'; break;
        case 'li': case 'lin': para.li = num || 0; break;
        case 'ri': case 'rin': para.ri = num || 0; break;
        case 'fi': para.fi = num || 0; break;
        case 'sb': para.sb = num || 0; break;
        case 'sa': para.sa = num || 0; break;
        case 'pagebb': para.pagebb = num !== 0; break;
        case 'rtlpar': para.rtl = true; break;
        case 'ltrpar': para.rtl = false; break;
        case 'outlinelevel': para.outline = num; break;
        case 's': para.style = num; break;
        case 'ls': para.ls = num; break;
        case 'ilvl': para.ilvl = num || 0; break;
        case 'intbl': para.intbl = true; break;
        case 'itap': para.intbl = (num ?? 1) > 0; break;
        case 'trowd': rowDefs = []; cellDef = {}; break;
        case 'trleft': rowDefs.left = num || 0; break;
        case 'clvmgf': cellDef.vmerge = 'first'; break;
        case 'clvmrg': cellDef.vmerge = 'cont'; break;
        case 'clmgf': cellDef.hmerge = 'first'; break;
        case 'clmrg': cellDef.hmerge = 'cont'; break;
        case 'clcbpat': cellDef.fill = num; break;
        case 'clbrdrt': case 'clbrdrb': case 'clbrdrl': case 'clbrdrr': cellDef.borderSide = true; break;
        case 'brdrs': case 'brdrth': case 'brdrdb': case 'brdrdot': case 'brdrdash': case 'brdrsh': if (cellDef.borderSide) cellDef.border = true; break;
        case 'brdrnone': case 'brdrnil': cellDef.borderSide = false; break;
        case 'cellx': rowDefs.push({ ...cellDef, right: num || 0 }); cellDef = {}; break;
        case 'cell': case 'nestcell':
          if (runs.length || !cellBlocks.length) endParagraph({ ...para, intbl: true });
          rowCells.push({ blocks: cellBlocks });
          cellBlocks = [];
          break;
        case 'row': case 'nestrow':
          tableRows.push({ defs: rowDefs.length ? rowDefs : rowCells.map((_, k) => ({ right: 1440 * (k + 1) })), cells: rowCells, left: rowDefs.left });
          rowCells = [];
          break;
        case 'paperw': case 'paperh': case 'margl': case 'margr': case 'margt': case 'margb': pageWords[word] = num; break;
        case 'landscape': pageWords.landscape = true; break;
        default: break;
      }
    }
    return para;
  };

  /** A picture: PNG or JPEG bytes, at the size it is shown. */
  const picture = (group, st) => {
    let kind = null;
    let hex = '';
    const w = {};
    for (const it of group.items) {
      if (it.word === 'pngblip') kind = 'png';
      else if (it.word === 'jpegblip') kind = 'jpg';
      else if (it.word) w[it.word] = it.num;
      else if (it.text != null) hex += it.text;
    }
    if (!kind) return;
    const data = Buffer.from(hex.replace(/[^0-9a-f]/gi, ''), 'hex');
    if (!data.length) return;
    const name = `Pictures/image${++imageCount}.${kind}`;
    images.set(name, new Uint8Array(data));
    const width = ((w.picwgoal || 0) / 15 || w.picw || 96) * (w.picscalex || 100) / 100;
    const height = ((w.pichgoal || 0) / 15 || w.pich || 96) * (w.picscaley || 100) / 100;
    runs.push({ ...(st.link ? { link: st.link } : {}), image: { href: name, width, height, name: null } });
  };

  /** A field: a HYPERLINK's result as linked words; any other field's result as its words. */
  const field = (group, st, para) => {
    const inst = group.items.find((it) => it.items && firstWord(it) === 'fldinst');
    const result = group.items.find((it) => it.items && firstWord(it) === 'fldrslt');
    const code = inst ? textIn(inst) : '';
    const m = /HYPERLINK\s+(?:\\l\s+)?"([^"]*)"/i.exec(code) || /HYPERLINK\s+(\S+)/i.exec(code);
    const anchor = /\\l\s+"([^"]*)"/.exec(code);
    const link = m ? (/\\l\s+"/.test(code.slice(0, code.indexOf(m[1]))) ? `#${m[1]}` : m[1] + (anchor && anchor[1] !== m[1] ? `#${anchor[1]}` : '')) : null;
    if (result) run({ items: result.items.filter((it) => it.word !== 'fldrslt') }, { ...st, ...(link ? { link } : {}) }, para);
  };

  const para = run({ items: top.items }, fresh(), freshPara());
  if (runs.length) endParagraph(para);
  closeTable();
  // The empty paragraph a closing \par leaves behind is not the document's.
  while (blocks.length && blocks[blocks.length - 1].type === 'paragraph' && !blocks[blocks.length - 1].runs.length && !blocks[blocks.length - 1].pageBreakBefore) blocks.pop();
  if (pageWords.paperw || pageWords.paperh) {
    page = {
      width: (pageWords.paperw || 12240) / 15, height: (pageWords.paperh || 15840) / 15,
      left: (pageWords.margl ?? 1800) / 15, right: (pageWords.margr ?? 1800) / 15, top: (pageWords.margt ?? 1440) / 15, bottom: (pageWords.margb ?? 1440) / 15,
      landscape: Boolean(pageWords.landscape),
    };
  }
  const defaultFont = fonts.get(deff);
  return { blocks, lists, page, images, title, defaults: { run: defaultFont ? { font: defaultFont } : {} } };
}
