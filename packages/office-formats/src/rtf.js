// Rich Text Format.
//
// RTF is a stream of groups and control words, not a tree of content, so the
// only way to read it is to run it: keep a formatting state, push and pop it
// with the braces, and emit text as it arrives. That is what this does.
//
// The awkward parts, all of which appear in real files:
//   - `\uN` gives a Unicode code point *and* a fallback the reader must skip,
//     with `\ucN` saying how many characters the fallback occupies.
//   - `\'xx` is a byte in the current code page, which is announced once at the
//     top of the file and never again.
//   - `\*\destination` marks a group whose contents are for a reader that
//     understands it, and are to be ignored by one that does not.

import { decoderFor } from './codepage.js';

const DESTINATIONS = new Set([
  'fonttbl', 'colortbl', 'stylesheet', 'info', 'pict', 'object', 'header', 'footer',
  'headerl', 'headerr', 'headerf', 'footerl', 'footerr', 'footerf', 'footnote',
  'xmlns', 'themedata', 'colorschememapping', 'latentstyles', 'datastore',
  'generator', 'listtable', 'listoverridetable', 'rsidtbl', 'mmathPr', 'fldinst',
  'filetbl', 'revtbl', 'userprops', 'panose', 'falt', 'atnid', 'atnauthor',
]);

const CODEPAGES = {
  ansi: 'windows-1252',
  1250: 'windows-1250',
  1251: 'windows-1251',
  1252: 'windows-1252',
  1253: 'windows-1253',
  1254: 'windows-1254',
  1255: 'windows-1255',
  1256: 'windows-1256',
  1257: 'windows-1257',
  1258: 'windows-1258',
  874: 'windows-874',
  932: 'shift_jis',
  936: 'gbk',
  949: 'euc-kr',
  950: 'big5',
  65001: 'utf-8',
};

/**
 * Parse RTF into paragraphs of formatted runs.
 * @param {string|Uint8Array} input
 * @returns {{ blocks: object[], meta: object, tables: number }}
 */
export function readRtf(input) {
  const text = typeof input === 'string' ? input : new TextDecoder('latin1').decode(input);

  let codepage = 'windows-1252';
  const cpMatch = /\\(ansicpg(\d+)|mac|pc|pca|ansi)\b/.exec(text.slice(0, 4096));
  if (cpMatch) codepage = CODEPAGES[cpMatch[2] ? Number(cpMatch[2]) : cpMatch[1]] || 'windows-1252';
  let decoder = decoderFor(codepage);

  const initial = { bold: false, italic: false, underline: false, strike: false, size: 0, color: 0, align: 'left', level: 0 };
  let state = { ...initial };
  const stack = [];

  const blocks = [];
  let runs = [];
  let buffer = '';
  let bufferBytes = [];
  let pendingSkip = 0; // characters of \uN fallback still to swallow
  let ucSkip = 1;
  let inTable = false;
  let rowCells = [];
  let tableRows = [];
  let tableCount = 0;
  const meta = {};
  let metaTarget = null;

  const flushBytes = () => {
    if (bufferBytes.length) {
      buffer += decoder.decode(new Uint8Array(bufferBytes));
      bufferBytes = [];
    }
  };

  const flushRun = () => {
    flushBytes();
    if (buffer) {
      if (metaTarget) meta[metaTarget] = (meta[metaTarget] || '') + buffer;
      else runs.push({ text: buffer, bold: state.bold, italic: state.italic, underline: state.underline, strike: state.strike, size: state.size, color: state.color });
      buffer = '';
    }
  };

  const endParagraph = () => {
    flushRun();
    if (runs.length || blocks.length === 0 || true) {
      blocks.push({ type: 'paragraph', align: state.align, level: state.level, runs });
    }
    runs = [];
  };

  const endCell = () => {
    flushRun();
    rowCells.push({ runs });
    runs = [];
  };

  const endRow = () => {
    if (rowCells.length) tableRows.push(rowCells);
    rowCells = [];
    inTable = true;
  };

  const closeTable = () => {
    if (tableRows.length) {
      blocks.push({ type: 'table', rows: tableRows.map((r) => r.map((c) => ({ runs: c.runs }))) });
      tableCount++;
    }
    tableRows = [];
    inTable = false;
  };

  let i = 0;
  let depth = 0;
  let skipDepth = -1;

  while (i < text.length) {
    const ch = text[i];

    if (ch === '{') {
      stack.push({ ...state, metaTarget });
      depth++;
      i++;
      continue;
    }
    if (ch === '}') {
      flushRun();
      const popped = stack.pop();
      if (popped) {
        metaTarget = popped.metaTarget;
        delete popped.metaTarget;
        state = popped;
      }
      depth--;
      if (skipDepth >= 0 && depth < skipDepth) skipDepth = -1;
      i++;
      continue;
    }
    if (ch === '\\') {
      const next = text[i + 1];
      // Escaped literal characters.
      if (next === '\\' || next === '{' || next === '}') {
        if (skipDepth < 0) bufferBytes.push(next.charCodeAt(0));
        i += 2;
        continue;
      }
      if (next === "'") {
        const hex = text.slice(i + 2, i + 4);
        const byte = parseInt(hex, 16);
        if (pendingSkip > 0) pendingSkip--;
        else if (skipDepth < 0 && Number.isFinite(byte)) bufferBytes.push(byte);
        i += 4;
        continue;
      }
      if (next === '*') {
        // \*\dest — a destination this reader may skip wholesale.
        skipDepth = skipDepth < 0 ? depth : skipDepth;
        i += 2;
        continue;
      }
      if (next === '\n' || next === '\r') {
        if (skipDepth < 0) endParagraph();
        i += 2;
        continue;
      }
      // Control word: letters, optional signed number, optional single space.
      let j = i + 1;
      while (j < text.length && /[a-zA-Z]/.test(text[j])) j++;
      const word = text.slice(i + 1, j);
      let numStr = '';
      if (text[j] === '-') {
        numStr = '-';
        j++;
      }
      while (j < text.length && /[0-9]/.test(text[j])) {
        numStr += text[j];
        j++;
      }
      if (text[j] === ' ') j++;
      const num = numStr === '' || numStr === '-' ? null : parseInt(numStr, 10);
      i = j;

      if (word === 'u') {
        if (skipDepth < 0 && num != null) {
          flushBytes();
          buffer += String.fromCodePoint(num < 0 ? num + 65536 : num);
        }
        pendingSkip = ucSkip;
        continue;
      }
      if (pendingSkip > 0 && !DESTINATIONS.has(word)) {
        // A control word inside the fallback run counts as one skipped char.
        pendingSkip--;
        continue;
      }

      if (DESTINATIONS.has(word)) {
        if (word === 'title' || word === 'author' || word === 'subject' || word === 'company') metaTarget = word;
        else skipDepth = skipDepth < 0 ? depth : skipDepth;
        continue;
      }
      if (word === 'title' || word === 'author' || word === 'subject' || word === 'company' || word === 'keywords') {
        flushRun();
        metaTarget = word;
        continue;
      }
      if (skipDepth >= 0) continue;

      switch (word) {
        case 'uc':
          ucSkip = num ?? 1;
          break;
        case 'par':
        case 'sect':
          endParagraph();
          break;
        case 'line':
          flushRun();
          runs.push({ text: '\n', ...state });
          break;
        case 'tab':
          flushBytes();
          buffer += '\t';
          break;
        case 'cell':
          endCell();
          break;
        case 'row':
          endRow();
          break;
        case 'intbl':
          inTable = true;
          break;
        case 'pard':
          if (inTable && !rowCells.length && tableRows.length) closeTable();
          state = { ...state, ...initial };
          break;
        case 'plain':
          Object.assign(state, { bold: false, italic: false, underline: false, strike: false, size: 0 });
          break;
        case 'b':
          flushRun();
          state = { ...state, bold: num !== 0 };
          break;
        case 'i':
          flushRun();
          state = { ...state, italic: num !== 0 };
          break;
        case 'ul':
          flushRun();
          state = { ...state, underline: num !== 0 };
          break;
        case 'ulnone':
          flushRun();
          state = { ...state, underline: false };
          break;
        case 'strike':
          flushRun();
          state = { ...state, strike: num !== 0 };
          break;
        case 'fs':
          flushRun();
          state = { ...state, size: (num ?? 24) / 2 };
          break;
        case 'cf':
          flushRun();
          state = { ...state, color: num ?? 0 };
          break;
        case 'ql':
        case 'qc':
        case 'qr':
        case 'qj':
          state = { ...state, align: { ql: 'left', qc: 'center', qr: 'right', qj: 'justify' }[word] };
          break;
        case 'ansicpg':
          decoder = decoderFor(CODEPAGES[num] || 'windows-1252');
          break;
        case 'outlinelevel':
          state = { ...state, level: (num ?? 0) + 1 };
          break;
        default:
          break; // unknown control words carry no text
      }
      continue;
    }

    if (ch === '\r' || ch === '\n') {
      i++;
      continue;
    }
    if (pendingSkip > 0) {
      pendingSkip--;
      i++;
      continue;
    }
    if (skipDepth < 0) bufferBytes.push(text.charCodeAt(i) & 0xff);
    i++;
  }

  flushRun();
  if (runs.length) blocks.push({ type: 'paragraph', align: state.align, level: state.level, runs });
  if (tableRows.length) closeTable();

  // Trim the empty paragraph a trailing \par leaves behind.
  while (blocks.length && blocks[blocks.length - 1].type === 'paragraph' && !blocks[blocks.length - 1].runs?.length) {
    blocks.pop();
  }

  for (const k of Object.keys(meta)) meta[k] = meta[k].replace(/ /g, '').trim();

  return { blocks, meta, tables: tableCount };
}

/** Plain text of an RTF file — what search and previews want. */
export function rtfToText(input) {
  const { blocks } = readRtf(input);
  const line = (b) =>
    b.type === 'table'
      ? b.rows.map((r) => r.map((c) => c.runs.map((x) => x.text).join('')).join('\t')).join('\n')
      : (b.runs || []).map((r) => r.text).join('');
  return blocks.map(line).join('\n');
}

/* ── writing ────────────────────────────────────────────────────────────── */

/**
 * Blocks back out as Rich Text Format.
 *
 * The installer tells Windows this suite is the EDITOR of `.rtf`, and until
 * now opening one and pressing Ctrl+S refused: a trap the operating system
 * set on our behalf. This is the other half.
 *
 * Writing RTF is reading it backwards. A reader runs the stream, so a writer
 * declares the fonts and the colours up front — a run can only name them by
 * index — and then emits each paragraph as a set of state changes followed by
 * its text, turning each change off again at the end so the next paragraph
 * starts clean. What goes out is the subset this suite's own reader carries:
 * the words, their font, size, weight, slant, underline, strike and colour,
 * the paragraph's alignment and list level, and a table as rows of cells.
 */

/** Text as RTF text: braces and backslashes escape, and anything above ASCII becomes \uN. */
function rtfText(text) {
  let out = '';
  for (const ch of String(text ?? '')) {
    const code = ch.codePointAt(0);
    if (ch === '\\' || ch === '{' || ch === '}') out += '\\' + ch;
    else if (ch === '\t') out += '\\tab ';
    else if (ch === '\n') out += '\\line ';
    else if (code < 128) out += ch;
    else if (code <= 0xffff) out += '\\u' + (code > 32767 ? code - 65536 : code) + '?';
    else {
      // Above the basic plane. RTF's \u takes a signed 16-bit word, so a
      // code point that needs a surrogate pair is written as the two words
      // a reader will put back together.
      const v = code - 0x10000;
      out += '\\u' + (0xd800 + (v >> 10) - 65536) + '?';
      out += '\\u' + (0xdc00 + (v & 0x3ff) - 65536) + '?';
    }
  }
  return out;
}

const ALIGN_WORD = { left: '\\ql', center: '\\qc', centre: '\\qc', right: '\\qr', justify: '\\qj', both: '\\qj' };

/** #rrggbb as the colour table entry RTF wants. */
function colourEntry(colour) {
  const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(String(colour || ''));
  return m ? '\\red' + parseInt(m[1], 16) + '\\green' + parseInt(m[2], 16) + '\\blue' + parseInt(m[3], 16) + ';' : null;
}

/** RTF's \levelnfc for ODF's number formats. */
const LEVEL_NFC = { 1: 0, I: 1, i: 2, A: 3, a: 4 };

/**
 * Word's list tables: a definition per list style, its nine levels each a
 * bullet or a number (format, start, the text round its number and the
 * levels above it shown), placed at its indent and hang; an override per
 * list, so each list counts on its own.
 */
function listTablesRtf(lists) {
  if (!lists?.definitions?.length) return '';
  const hex = (n) => "\\'" + n.toString(16).padStart(2, '0');
  const levelRtf = (lv, k) => {
    const indent = Math.round(((lv?.indent ?? 48 + 24 * k) || 48) * 15);
    const hang = Math.round((lv?.hanging ?? 24) * 15);
    const place = '\\fi-' + hang + '\\li' + indent + '\\lin' + indent + ' ';
    if (!lv || lv.kind !== 'number') {
      const ch = (lv?.char || '\u2022').codePointAt(0);
      return '{\\listlevel\\levelnfc23\\levelnfcn23\\leveljc0\\leveljcn0\\levelfollow0\\levelstartat1\\levelspace0\\levelindent0{\\leveltext' + hex(1) + '\\u' + (ch > 32767 ? ch - 65536 : ch) + ' ?;}{\\levelnumbers;}' + place + '}';
    }
    // The text round the number: the prefix, the levels shown (a byte below ten for each), the suffix.
    const shown = Math.max(1, Math.min(lv.display || 1, k + 1));
    const parts = [];
    const numbers = [];
    let text = rtfText(lv.prefix || '');
    let length = [...String(lv.prefix || '')].length;
    for (let s = k - shown + 1; s <= k; s++) {
      if (s > k - shown + 1) { text += '.'; length += 1; }
      text += hex(s);
      length += 1;
      numbers.push(length);
      parts.push(s);
    }
    text += rtfText(lv.suffix ?? '.');
    length += [...String(lv.suffix ?? '.')].length;
    const nfc = LEVEL_NFC[lv.format] ?? 0;
    return '{\\listlevel\\levelnfc' + nfc + '\\levelnfcn' + nfc + '\\leveljc0\\leveljcn0\\levelfollow0\\levelstartat' + (lv.start || 1) + '\\levelspace0\\levelindent0{\\leveltext' + hex(length) + text + ';}{\\levelnumbers' + numbers.map(hex).join('') + ';}' + place + '}';
  };
  const defs = lists.definitions.map((d) => '{\\list\\listtemplateid' + d.listid + '\\listhybrid' + Array.from({ length: 9 }, (_, k) => levelRtf(d.levels[k] || d.levels[d.levels.length - 1] && { ...d.levels[d.levels.length - 1], indent: 48 + 24 * k }, k)).join('') + '{\\listname ;}\\listid' + d.listid + '}').join('');
  const overrides = lists.overrides.map((o) => '{\\listoverride\\listid' + o.listid + '\\listoverridecount0\\ls' + o.ls + '}').join('');
  return '{\\*\\listtable' + defs + '}\n{\\*\\listoverridetable' + overrides + '}\n';
}

export function writeRtf({ blocks = [], title = '', lists = null } = {}) {
  // Both tables are built while the body is written and emitted before it,
  // because a run refers to a font and a colour by index and neither is known
  // until the run that wants it turns up.
  const fonts = ['Calibri'];
  const colours = [null]; // index 0 is "whatever the reader defaults to"
  const fontIndex = (name) => {
    const font = name || 'Calibri';
    const at = fonts.indexOf(font);
    return at < 0 ? fonts.push(font) - 1 : at;
  };
  const colourIndex = (colour) => {
    const entry = colourEntry(colour);
    if (!entry) return 0;
    const at = colours.indexOf(entry);
    return at < 0 ? colours.push(entry) - 1 : at;
  };

  /**
   * A run with any formatting of its own is written as a group.
   *
   * Turning the formatting off again afterwards — `\b Bold\b0 and normal` —
   * is the obvious way and it is wrong: a control word is delimited by one
   * space, and that space is eaten rather than printed, so "Bold and normal"
   * came back "Boldand normal". A group ends the formatting at its closing
   * brace and leaves the next run's leading space alone. It is also what Word
   * writes.
   */
  const runOut = (run) => {
    // A picture: PNG or JPEG bytes as hex, at the size it is shown (twips), as Word writes one.
    if (run.picture) {
      const p = run.picture;
      const px = (v) => Math.max(1, Math.round(Number(v) || 96));
      return '{\\pict' + (p.kind === 'png' ? '\\pngblip' : '\\jpegblip') + '\\picw' + px(p.width) + '\\pich' + px(p.height) +
        '\\picwgoal' + px(p.width) * 15 + '\\pichgoal' + px(p.height) * 15 + ' ' + Buffer.from(p.data).toString('hex') + '}';
    }
    const on = [];
    const font = fontIndex(run.font || run.fontName);
    if (font) on.push('\\f' + font);
    const colour = colourIndex(run.colour || run.color);
    if (colour) on.push('\\cf' + colour);
    // RTF measures type in half-points, as Word's own format does.
    if (run.size) on.push('\\fs' + Math.round(Number(run.size) * 2));
    if (run.bold) on.push('\\b');
    if (run.italic) on.push('\\i');
    if (run.underline) on.push('\\ul');
    if (run.strike) on.push('\\strike');
    const text = rtfText(run.text);
    return on.length ? '{' + on.join('') + ' ' + text + '}' : text;
  };

  const runsOf = (holder) => ((holder.runs || []).length ? holder.runs : [{ text: holder.text ?? '' }]);

  /** Runs, a link round the runs that share one: a HYPERLINK field, its result the runs. */
  const runsOut = (runs) => {
    let out = '';
    for (let i = 0; i < runs.length;) {
      const link = runs[i].link;
      if (!link) { out += runOut(runs[i]); i += 1; continue; }
      let inner = '';
      while (i < runs.length && runs[i].link === link) { inner += runOut(runs[i]); i += 1; }
      out += '{\\field{\\*\\fldinst{HYPERLINK "' + rtfText(String(link).replace(/"/g, '%22')) + '"}}{\\fldrslt{' + inner + '}}}';
    }
    return out;
  };

  const paragraph = (block) => {
    const align = ALIGN_WORD[block.align] || '';
    const heading = block.type === 'heading';
    // A list item hangs its label at its level's indent; a heading's level is its outline level, not an indent.
    const indent = block.indentTwips != null
      ? '\\li' + block.indentTwips + (block.hangTwips ? '\\fi-' + block.hangTwips : block.firstTwips ? '\\fi' + block.firstTwips : '')
      : block.level && !heading ? '\\li' + block.level * 360 : '';
    const size = heading ? Math.max(20, 36 - (block.level || 1) * 4) : 0;
    const flags = (heading ? '\\outlinelevel' + Math.max(0, Math.min(8, (block.level || 1) - 1)) : '') + (block.pageBreakBefore ? '\\pagebb' : '') + (block.rtl ? '\\rtlpar' : '') +
      // A list item names its list and level; its label rides in \listtext for a reader without lists.
      (block.list ? '\\ls' + block.list.ls + '\\ilvl' + block.list.level : '');
    return '{\\pard' + align + indent + flags + '\\sa120 ' +
      (heading ? '\\b\\fs' + size + ' ' : '') +
      (block.label ? (block.list ? '{\\listtext ' + rtfText(block.label) + '\\tab}' : rtfText(block.label) + '\\tab ') : '') +
      runsOut(runsOf(block)) +
      (heading ? '\\b0' : '') +
      '\\par}';
  };

  /**
   * A table row is not a group.
   *
   * The row definition, then each cell as its own `\pard\intbl … \cell`, then
   * `\row`; a `\pard` after the last row is what tells a reader the table has
   * ended. Wrapping a row in braces instead — which looks tidier — puts the
   * cells inside a group whose state is popped at the closing brace, and both
   * this suite's reader and Word's read the row back as one empty cell.
   */
  const table = (block) => {
    const rows = block.rows || [];
    const out = [];
    // The grid: the table's own column widths when it gives them, else the row shared out evenly.
    const grid = (block.columns || []).length && block.columns.every((w) => w > 0) ? block.columns : null;
    for (const row of rows) {
      const even = Math.floor(9000 / Math.max(1, row.length));
      const edge = (col, span) => (grid ? grid.slice(0, col + span).reduce((a, b) => a + b, 0) : even * (col + span));
      const defs = [];
      const cells = [];
      let col = 0;
      for (const cell of row) {
        // A place a span from the left covers has no cell of its own: the span's cell reaches over it.
        if (cell.covered) continue;
        const span = Math.max(1, cell.colspan || 1);
        const merge = cell.continues ? '\\clvmrg' : (cell.rowspan || 1) > 1 ? '\\clvmgf' : '';
        const fill = cell.fill ? colourIndex(cell.fill) : 0;
        // Ruled unless the table says it has no borders (a writer that does not say gets them, as before).
        const rules = block.bordered === false ? '' : '\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrb\\brdrs\\clbrdrr\\brdrs';
        defs.push(merge + (fill ? '\\clcbpat' + fill : '') + rules + '\\cellx' + edge(col, span));
        cells.push(cell);
        col += span;
      }
      out.push('\\trowd\\trgaph108' + defs.join(''));
      // `\intbl` alone, not `\pard\intbl`. A reader ends a table when a
      // paragraph resets outside a row, and this one — ours, and it is not
      // alone in this — reads the `\pard` that opens the second row's first
      // cell as that reset: a table of four rows came back as four tables of
      // one. The row's properties come from `\trowd` either way.
      for (const cell of cells) out.push('\\intbl' + (ALIGN_WORD[cell.align] || '') + ' ' + (cell.continues ? '' : runsOut(runsOf(cell))) + '\\cell');
      out.push('\\row');
    }
    return out.join('\n') + '\n\\pard';
  };

  const body = (blocks.length ? blocks : [{ type: 'paragraph', runs: [] }])
    .map((block) => (block.type === 'table' ? table(block) : paragraph(block)))
    .join('\n');

  const fontTable = fonts.map((name, i) => '{\\f' + i + '\\fnil\\fcharset0 ' + rtfText(name) + ';}').join('');
  const colourTable = '{\\colortbl;' + colours.slice(1).join('') + '}';
  const info = title ? '{\\info{\\title ' + rtfText(title) + '}}' : '';
  const listTables = listTablesRtf(lists);
  return '{\\rtf1\\ansi\\ansicpg1252\\deff0\\uc1{\\fonttbl' + fontTable + '}' + colourTable + listTables + info +
    '\n\\viewkind4\\fs22\n' + body + '\n}';
}

export default { readRtf, rtfToText, writeRtf };
