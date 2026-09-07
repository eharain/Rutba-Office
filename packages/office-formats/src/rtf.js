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

function decoderFor(label) {
  try {
    return new TextDecoder(label);
  } catch {
    return new TextDecoder('windows-1252');
  }
}

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

export default { readRtf, rtfToText };
