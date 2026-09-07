/**
 * List labels — the "3." in front of the third item.
 *
 * The backend says WHICH list a paragraph belongs to and at what level; it
 * cannot say what the label is, because "3." depends on the two items before it,
 * not on the paragraph's own XML. Counting is order-dependent view work, exactly
 * like choosing a header band for a page — so it lives here, format-free: the
 * definitions arrive as plain data and an HTML backend could supply the same
 * shape for `<ol>` lists.
 *
 * The counter rules are Word's, and the non-obvious one is RESTART: entering a
 * deeper level resets that level's counter... but returning to a shallower one
 * does not reset it — 1, 1.1, 1.2, 2 continues at 2, and the next 2.x starts at
 * 1 again. Getting this wrong is invisible on flat lists and instantly wrong on
 * a contract's clause numbering.
 */

const toRoman = (n) => {
  const table = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'],
    [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
  let out = '';
  for (const [value, glyph] of table) while (n >= value) { out += glyph; n -= value; }
  return out;
};
const toLetter = (n) => {
  // 1 -> a, 26 -> z, 27 -> aa: spreadsheet-column style, which is Word's too.
  let out = '';
  while (n > 0) { n -= 1; out = String.fromCharCode(97 + (n % 26)) + out; n = Math.floor(n / 26); }
  return out;
};

export function formatCounter(format, n) {
  switch (format) {
    case 'bullet': return '•';
    case 'none': return '';
    case 'lowerLetter': return toLetter(n);
    case 'upperLetter': return toLetter(n).toUpperCase();
    case 'lowerRoman': return toRoman(n);
    case 'upperRoman': return toRoman(n).toUpperCase();
    default: return String(n);
  }
}

/**
 * Walk the flow in document order and label every list paragraph.
 *
 * @param {Array} flow      document order — the same array the paginator walks
 * @param {Array} blocks    paragraphs carrying `numbering: {numId, level}|null`
 * @param {Record<string, Array>} defs  numId -> levels, from the backend
 * @returns {Map<number, {label: string, indentPx: number, bullet: boolean}>}
 */
export function computeListLabels(flow, blocks, defs) {
  const out = new Map();
  if (!defs || !Object.keys(defs).length) return out;
  const byIndex = new Map(blocks.map((b) => [b.index, b]));

  /** numId -> per-level counters. Separate lists never share a count. */
  const counters = new Map();

  const count = (blockIndex) => {
    const block = byIndex.get(blockIndex);
    const numbering = block?.numbering;
    if (!numbering) return;
    const levels = defs[numbering.numId];
    const def = levels?.[numbering.level];
    if (!def) return;

    if (!counters.has(numbering.numId)) counters.set(numbering.numId, []);
    const c = counters.get(numbering.numId);

    // A bullet never counts; a counter level increments and RESETS everything
    // deeper, so the next sub-list starts again at its own `start`.
    if (def.format !== 'bullet' && def.format !== 'none') {
      c[numbering.level] = (c[numbering.level] ?? (levels[numbering.level].start - 1)) + 1;
      c.length = numbering.level + 1;
    }

    // lvlText like "%1.%2." pulls in the counters of the SHALLOWER levels too —
    // that is how "2.3." knows about the 2.
    const label = def.format === 'bullet' ? '•' : def.lvlText.replace(/%(\d)/g, (_, d) => {
      const level = Number(d) - 1;
      const value = c[level] ?? levels[level]?.start ?? 1;
      return formatCounter(levels[level]?.format ?? 'decimal', value);
    });

    out.set(block.index, {
      label,
      indentPx: def.indentPx,
      bullet: def.format === 'bullet',
    });
  };

  // Tables descend: a list item inside a cell is still item N of its list,
  // counted in document order with everything around it. (Cell list labels
  // went unpainted for a long time precisely because this walk skipped
  // table entries.)
  const walkTable = (table) => {
    for (const row of table.rows ?? []) {
      for (const cell of row.cells ?? []) {
        for (const b of cell.blocks ?? []) {
          if (b.kind === 'table') walkTable(b.table);
          else if (Number.isInteger(b.blockIndex)) count(b.blockIndex);
        }
      }
    }
  };

  for (const entry of flow ?? []) {
    if (entry.kind === 'paragraph') count(entry.paragraphIndex);
    else if (entry.kind === 'table') walkTable(entry.table);
  }
  return out;
}
