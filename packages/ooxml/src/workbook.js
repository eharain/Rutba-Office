/**
 * A workbook VIEW over a package — the smallest thing that can edit a cell
 * without disturbing anything else in the file.
 *
 * The discipline, in descending order of how much fidelity each buys:
 *
 * 1. Parts we do not edit are never rewritten (the package layer guarantees it).
 * 2. Inside a sheet part, everything outside `<sheetData>` is opaque text held
 *    verbatim — `cols`, `mergeCells`, `conditionalFormatting`, `dataValidations`,
 *    `autoFilter`, `pageSetup`, `sheetProtection`, `extLst`. These are exactly
 *    the elements a naive rewriter silently drops, and losing `sheetProtection`
 *    or `dataValidations` from a file that goes to a bank is not a cosmetic bug.
 * 3. Inside `<sheetData>`, rows we do not touch keep their original XML.
 * 4. Inside a row we do touch, cells we do not touch keep their original XML.
 * 5. In a cell we DO touch, the style index `s=` is carried over unchanged, so
 *    editing a value never changes how it is formatted.
 *
 * String writes use `t="inlineStr"` rather than appending to `sharedStrings.xml`.
 * Both are core OOXML; inline strings mean a value edit touches ONE part instead
 * of two, and `sharedStrings.xml` — which the rest of the workbook indexes into
 * by position — is never renumbered. Trading a slightly less idiomatic file for
 * not rewriting a shared index is the right way round.
 */
import { OoxmlPackage, attrs, esc } from './package.js';

const SHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

export function colToIndex(letters) {
  let n = 0;
  for (const ch of String(letters).toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
export function indexToCol(n) {
  let s = '';
  let v = n + 1;
  while (v > 0) {
    const r = (v - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    v = Math.floor((v - 1) / 26);
  }
  return s;
}
export function parseRef(ref) {
  const m = /^([A-Za-z]+)(\d+)$/.exec(String(ref).trim());
  if (!m) throw new Error('bad cell reference: ' + ref);
  return { col: colToIndex(m[1]), row: Number(m[2]) - 1 };
}
export const makeRef = (row, col) => indexToCol(col) + (row + 1);

/** One sheet part, split so the bits we do not understand stay verbatim. */
class SheetPart {
  constructor(xml) {
    this.original = xml;
    const open = /<sheetData\s*\/>|<sheetData(\s[^>]*)?>/.exec(xml);
    if (!open) throw new Error('sheet part has no <sheetData>');
    if (open[0].endsWith('/>')) {
      this.prefix = xml.slice(0, open.index);
      this.openTag = '<sheetData>';
      this.body = '';
      this.suffix = xml.slice(open.index + open[0].length);
    } else {
      const close = xml.indexOf('</sheetData>', open.index);
      if (close < 0) throw new Error('unterminated <sheetData>');
      this.prefix = xml.slice(0, open.index);
      this.openTag = open[0];
      this.body = xml.slice(open.index + open[0].length, close);
      this.suffix = xml.slice(close + '</sheetData>'.length);
    }
    this.rows = this._parseRows();
    this.dirty = false;
  }

  /** Ordered row records; `xml` is kept verbatim until the row is touched. */
  _parseRows() {
    const rows = [];
    const re = /<row\b([^>]*?)(\/>|>([\s\S]*?)<\/row>)/g;
    let m;
    while ((m = re.exec(this.body))) {
      const a = attrs(m[1]);
      rows.push({
        index: a.r ? Number(a.r) - 1 : rows.length,
        attrsStr: m[1],
        inner: m[2] === '/>' ? '' : m[3],
        xml: m[0],
        dirty: false,
      });
    }
    return rows;
  }

  _rowAt(rowIndex) {
    return this.rows.find((r) => r.index === rowIndex) ?? null;
  }

  static _parseCells(inner) {
    const cells = [];
    const re = /<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g;
    let m;
    while ((m = re.exec(inner))) {
      const a = attrs(m[1]);
      cells.push({
        ref: a.r ?? null,
        col: a.r ? parseRef(a.r).col : cells.length,
        attrsStr: m[1],
        style: a.s ?? null,
        type: a.t ?? null,
        inner: m[2] === '/>' ? '' : m[3],
        xml: m[0],
      });
    }
    return cells;
  }

  getCell(rowIndex, colIndex) {
    const row = this._rowAt(rowIndex);
    if (!row) return null;
    const cell = SheetPart._parseCells(row.inner).find((c) => c.col === colIndex);
    return cell ?? null;
  }

  /**
   * Write a value, carrying the existing style index across.
   * @param {(index: string) => string} [resolveSharedString]
   */
  setCell(rowIndex, colIndex, value, resolveSharedString) {
    const ref = makeRef(rowIndex, colIndex);
    let row = this._rowAt(rowIndex);
    if (!row) {
      row = { index: rowIndex, attrsStr: ' r="' + (rowIndex + 1) + '"', inner: '', xml: '', dirty: true };
      const at = this.rows.findIndex((r) => r.index > rowIndex);
      if (at < 0) this.rows.push(row);
      else this.rows.splice(at, 0, row);
    }
    const cells = SheetPart._parseCells(row.inner);
    const existing = cells.find((c) => c.col === colIndex);
    const style = existing?.style ?? null;
    const cellXml = renderCell(ref, value, style);

    if (existing) {
      row.inner = row.inner.replace(existing.xml, cellXml);
    } else {
      const after = cells.filter((c) => c.col < colIndex).pop();
      if (after) row.inner = row.inner.replace(after.xml, after.xml + cellXml);
      else row.inner = cellXml + row.inner;
    }
    row.dirty = true;
    this.dirty = true;
    void resolveSharedString;
    return this;
  }

  /**
   * Point a cell at a different style index, leaving its value alone.
   *
   * The counterpart of `setCell`, which carries the style across and changes
   * the value. Formatting an empty cell is normal — a blank column can be
   * shaded or given a number format before anything is typed into it — so this
   * creates the `<c>` when there is not one.
   */
  setCellStyle(rowIndex, colIndex, styleIndex) {
    const ref = makeRef(rowIndex, colIndex);
    let row = this._rowAt(rowIndex);
    if (!row) {
      row = { index: rowIndex, attrsStr: ' r="' + (rowIndex + 1) + '"', inner: '', xml: '', dirty: true };
      const at = this.rows.findIndex((r) => r.index > rowIndex);
      if (at < 0) this.rows.push(row);
      else this.rows.splice(at, 0, row);
    }
    const cells = SheetPart._parseCells(row.inner);
    const existing = cells.find((c) => c.col === colIndex);

    if (existing) {
      // Rewrite only the attribute. The cell's contents — a value, a formula,
      // a cached result, an inline string — are none of this method's business.
      const head = existing.xml.slice(0, existing.xml.indexOf('>') + 1);
      const withStyle = /\bs="\d+"/.test(head)
        ? head.replace(/\bs="\d+"/, styleIndex === null ? '' : 's="' + styleIndex + '"')
        : (styleIndex === null ? head : head.replace(/(<c\b[^>]*?)(\/?>)$/, '$1 s="' + styleIndex + '"$2'));
      row.inner = row.inner.replace(existing.xml, withStyle + existing.xml.slice(head.length));
    } else {
      const cellXml = '<c r="' + ref + '"' + (styleIndex === null ? '' : ' s="' + styleIndex + '"') + '/>';
      const before = cells.filter((c) => c.col < colIndex).pop();
      if (before) row.inner = row.inner.replace(before.xml, before.xml + cellXml);
      else row.inner = cellXml + row.inner;
    }
    row.dirty = true;
    this.dirty = true;
    return this;
  }

  // ── column widths and row heights ─────────────────────────────────────────

  /**
   * The `<col>` records this sheet declares, with `min`/`max` pulled out and
   * every other attribute kept verbatim so hidden/style/bestFit survive a
   * width edit untouched. `<cols>` lives in the PREFIX — the schema puts it
   * before `<sheetData>` — which is why this reads prefix text, not suffix.
   */
  colEntries() {
    const block = /<cols\b[^>]*>([\s\S]*?)<\/cols>/.exec(this.prefix);
    if (!block) return [];
    const out = [];
    for (const m of block[1].matchAll(/<col\b([^>]*?)\/?>(?:<\/col>)?/g)) {
      const a = attrs(m[1]);
      const min = Number(a.min);
      const max = Number(a.max);
      if (!min || !max) continue;
      // Everything except min/max is carried as raw text so nothing is lost.
      const rest = m[1]
        .replace(/\s*\bmin="[^"]*"/, '')
        .replace(/\s*\bmax="[^"]*"/, '');
      out.push({ min, max, rest });
    }
    return out;
  }

  /** Replace the whole `<cols>` block, or drop it when no entries remain. */
  setColEntries(entries) {
    const sorted = [...entries].sort((a, b) => a.min - b.min);
    const block = sorted.length
      ? '<cols>' + sorted.map((e) => '<col min="' + e.min + '" max="' + e.max + '"' + e.rest + '/>').join('') + '</cols>'
      : '';
    const paired = /<cols\b[^>]*>[\s\S]*?<\/cols>/;
    const empty = /<cols\b[^>]*\/>/;
    if (paired.test(this.prefix)) this.prefix = this.prefix.replace(paired, block);
    else if (empty.test(this.prefix)) this.prefix = this.prefix.replace(empty, block);
    // The prefix ends exactly where <sheetData> begins, and the schema puts
    // cols immediately before sheetData, so appending is the correct position.
    else if (block) this.prefix += block;
    this.dirty = true;
    return this;
  }

  /**
   * Set ONE column's width in character units, carving it out of any range
   * record that covers it so the neighbours keep their own settings.
   */
  setColWidth(colIndex, widthChars) {
    const target = colIndex + 1; // <col> is 1-based
    const withWidth = (rest) => {
      let r = /\bwidth="[^"]*"/.test(rest)
        ? rest.replace(/\bwidth="[^"]*"/, 'width="' + widthChars + '"')
        : rest + ' width="' + widthChars + '"';
      if (/\bcustomWidth="[^"]*"/.test(r)) r = r.replace(/\bcustomWidth="[^"]*"/, 'customWidth="1"');
      else r += ' customWidth="1"';
      return r;
    };
    const out = [];
    let placed = false;
    for (const e of this.colEntries()) {
      if (e.max < target || e.min > target) { out.push(e); continue; }
      if (e.min <= target - 1) out.push({ min: e.min, max: target - 1, rest: e.rest });
      out.push({ min: target, max: target, rest: withWidth(e.rest) });
      if (e.max >= target + 1) out.push({ min: target + 1, max: e.max, rest: e.rest });
      placed = true;
    }
    if (!placed) out.push({ min: target, max: target, rest: ' width="' + widthChars + '" customWidth="1"' });
    return this.setColEntries(out);
  }

  /**
   * Set one row's height in points. `customHeight="1"` marks it as chosen by a
   * person rather than fitted to the font, which is what makes Excel keep it.
   */
  setRowHeight(rowIndex, points) {
    let row = this._rowAt(rowIndex);
    if (!row) {
      row = { index: rowIndex, attrsStr: ' r="' + (rowIndex + 1) + '"', inner: '', xml: '', dirty: true };
      const at = this.rows.findIndex((r) => r.index > rowIndex);
      if (at < 0) this.rows.push(row);
      else this.rows.splice(at, 0, row);
    }
    if (/\bht="[^"]*"/.test(row.attrsStr)) row.attrsStr = row.attrsStr.replace(/\bht="[^"]*"/, 'ht="' + points + '"');
    else row.attrsStr += ' ht="' + points + '"';
    if (/\bcustomHeight="[^"]*"/.test(row.attrsStr)) row.attrsStr = row.attrsStr.replace(/\bcustomHeight="[^"]*"/, 'customHeight="1"');
    else row.attrsStr += ' customHeight="1"';
    row.dirty = true;
    this.dirty = true;
    return this;
  }

  // ── structural shifts (rows and columns) ──────────────────────────────────

  /** Rewrite a row record to live at a new index: its `r` and every cell ref. */
  _renumberRow(row, newIndex) {
    const r = newIndex + 1;
    if (/\br="\d+"/.test(row.attrsStr)) row.attrsStr = row.attrsStr.replace(/\br="\d+"/, 'r="' + r + '"');
    else row.attrsStr += ' r="' + r + '"';
    row.inner = row.inner.replace(/(<c\b[^>]*?\br=")([A-Za-z]+)\d+(?=")/g, '$1$2' + r);
    row.index = newIndex;
    row.dirty = true;
  }

  /** Shift every row at or below `at` down by `count`. Sparse: no XML is made
   *  for the inserted rows — an empty row needs no record. */
  insertRowsShift(at, count) {
    for (const row of this.rows) {
      if (row.index >= at) this._renumberRow(row, row.index + count);
    }
    this.dirty = true;
    return this;
  }

  /** Remove rows `[at, at+count)` and close the gap beneath them. */
  deleteRowsRange(at, count) {
    this.rows = this.rows.filter((row) => row.index < at || row.index >= at + count);
    for (const row of this.rows) {
      if (row.index >= at + count) this._renumberRow(row, row.index - count);
    }
    this.dirty = true;
    return this;
  }

  /**
   * Shift cells sideways in every row: `op` is 'insert' or 'delete' at 0-based
   * column `at`. Cells inside a deleted slice vanish; the rest are re-lettered.
   * The optional `spans` hint is recomputed where present so it never lies.
   */
  shiftCells(at, count, op) {
    for (const row of this.rows) {
      if (!row.inner) continue;
      const cells = SheetPart._parseCells(row.inner);
      let touched = false;
      for (const cell of cells) {
        if (op === 'delete' && cell.col >= at && cell.col < at + count) {
          row.inner = row.inner.replace(cell.xml, '');
          touched = true;
        } else if (cell.col >= (op === 'insert' ? at : at + count)) {
          const next = op === 'insert' ? cell.col + count : cell.col - count;
          const moved = cell.xml.replace(/(<c\b[^>]*?\br=")[A-Za-z]+(\d+")/, '$1' + indexToCol(next) + '$2');
          row.inner = row.inner.replace(cell.xml, moved);
          touched = true;
        }
      }
      if (!touched) continue;
      if (/\bspans="[^"]*"/.test(row.attrsStr)) {
        const left = SheetPart._parseCells(row.inner);
        row.attrsStr = left.length
          ? row.attrsStr.replace(/\bspans="[^"]*"/,
            'spans="' + (Math.min(...left.map((c) => c.col)) + 1) + ':' + (Math.max(...left.map((c) => c.col)) + 1) + '"')
          : row.attrsStr.replace(/\s*\bspans="[^"]*"/, '');
      }
      row.dirty = true;
    }
    this.dirty = true;
    return this;
  }

  /** Shift the `<col>` width records for a column insert/delete. */
  shiftColEntries(at, count, op) {
    const entries = this.colEntries();
    if (!entries.length) return this;
    const out = [];
    let changed = false;
    for (const e of entries) {
      const lo = e.min - 1;
      const hi = e.max - 1;
      if (op === 'insert') {
        if (lo >= at) { out.push({ min: e.min + count, max: e.max + count, rest: e.rest }); changed = true; }
        else if (hi >= at) { out.push({ min: e.min, max: e.max + count, rest: e.rest }); changed = true; }
        else out.push(e);
      } else {
        const end = at + count;
        const nlo = lo >= end ? lo - count : lo >= at ? at : lo;
        const nhi = hi >= end ? hi - count : hi >= at ? at - 1 : hi;
        if (nlo > nhi) { changed = true; continue; }
        if (nlo !== lo || nhi !== hi) changed = true;
        out.push({ min: nlo + 1, max: nhi + 1, rest: e.rest });
      }
    }
    if (changed) this.setColEntries(out);
    return this;
  }

  /** Keep the `<dimension>` hint honest after a structural edit. */
  shiftDimension(axis, op, at, count) {
    this.prefix = this.prefix.replace(/(<dimension\b[^>]*?\bref=")([^"]+)(")/, (m, open, ref, close) => {
      const parts = ref.split(':');
      let from;
      let to;
      try {
        from = parseRef(parts[0]);
        to = parts[1] ? parseRef(parts[1]) : from;
      } catch {
        return m;
      }
      const lo = axis === 'row' ? from.row : from.col;
      const hi = axis === 'row' ? to.row : to.col;
      const shifted = shiftSpan(lo, hi, at, count, op);
      const [nlo, nhi] = shifted ?? [0, 0];
      const a = axis === 'row' ? makeRef(nlo, from.col) : makeRef(from.row, nlo);
      const b = axis === 'row' ? makeRef(nhi, to.col) : makeRef(to.row, nhi);
      return open + (parts[1] ? a + ':' + b : a) + close;
    });
    return this;
  }

  /**
   * Run an adjuster over every stored formula: the `<f>` text and, where a
   * shared-formula master carries a `ref` range, that attribute too. The
   * original bytes survive wherever the adjuster changes nothing.
   */
  adjustFormulas(fn) {
    let touched = false;
    for (const row of this.rows) {
      if (!/<f[\s>]/.test(row.inner)) continue;
      const next = row.inner.replace(/(<f\b[^>]*>)([\s\S]*?)(<\/f>)/g, (m, open, body, close) => {
        const source = unesc(body);
        const adjusted = fn(source);
        const openAdj = open.replace(/\bref="([^"]+)"/, (mm, refText) => 'ref="' + fn(refText) + '"');
        if (adjusted === source && openAdj === open) return m;
        return openAdj + (adjusted === source ? body : esc(adjusted)) + close;
      });
      if (next !== row.inner) {
        row.inner = next;
        row.dirty = true;
        touched = true;
      }
    }
    if (touched) this.dirty = true;
    return this;
  }

  /**
   * The merge refs this sheet declares, e.g. `['A1:C1', 'B4:B6']`.
   *
   * `<mergeCells>` lives OUTSIDE `<sheetData>`, in the opaque suffix this class
   * otherwise never touches, so it is read straight out of that text.
   */
  mergeRefs() {
    const block = /<mergeCells\b[^>]*>([\s\S]*?)<\/mergeCells>/.exec(this.suffix);
    if (!block) return [];
    return [...block[1].matchAll(/<mergeCell\b[^>]*\bref="([^"]+)"/g)].map((m) => m[1]);
  }

  /**
   * Replace the whole `<mergeCells>` block with these refs, or drop it entirely
   * when none remain — an empty `<mergeCells count="0"/>` is not a valid element.
   *
   * Placement follows the schema: `mergeCells` comes after `</sheetData>`, which
   * is exactly where the suffix begins, so a fresh block is prepended to it. An
   * existing block is edited in place so its original position is preserved.
   */
  setMerges(refs) {
    const unique = [...new Set(refs)];
    const block = unique.length
      ? '<mergeCells count="' + unique.length + '">'
        + unique.map((r) => '<mergeCell ref="' + esc(r) + '"/>').join('')
        + '</mergeCells>'
      : '';
    const paired = /<mergeCells\b[^>]*>[\s\S]*?<\/mergeCells>/;
    const empty = /<mergeCells\b[^>]*\/>/;
    if (paired.test(this.suffix)) {
      this.suffix = this.suffix.replace(paired, block);
    } else if (empty.test(this.suffix)) {
      this.suffix = this.suffix.replace(empty, block);
    } else if (block) {
      this.suffix = block + this.suffix;
    }
    this.dirty = true;
    return this;
  }

  /** The frozen pane the sheet's view declares: {rows, cols}, or null. */
  frozenPane() {
    const m = /<pane\b([^>]*)\/?>/.exec(this.prefix);
    if (!m) return null;
    const a = attrs(m[1]);
    if (a.state !== 'frozen' && a.state !== 'frozenSplit') return null;
    const rows = Number(a.ySplit ?? 0) || 0;
    const cols = Number(a.xSplit ?? 0) || 0;
    return rows || cols ? { rows, cols } : null;
  }

  /**
   * Freeze `rows` rows and `cols` columns at the top-left — 0 and 0 thaws.
   * The pane element is the FIRST child of `<sheetView>`; a sheet with no
   * view settings at all gains a minimal `<sheetViews>` where the schema
   * puts it, after `<dimension>` when there is one.
   */
  setFrozenPane(rows, cols) {
    this.prefix = this.prefix.replace(/<pane\b[^>]*\/>|<pane\b[^>]*>[\s\S]*?<\/pane>/, '');
    if (rows > 0 || cols > 0) {
      const activePane = rows && cols ? 'bottomRight' : rows ? 'bottomLeft' : 'topRight';
      const pane = '<pane'
        + (cols ? ' xSplit="' + cols + '"' : '')
        + (rows ? ' ySplit="' + rows + '"' : '')
        + ' topLeftCell="' + makeRef(rows, cols) + '" activePane="' + activePane + '" state="frozen"/>';
      if (/<sheetView\b[^>]*\/>/.test(this.prefix)) {
        this.prefix = this.prefix.replace(/<sheetView\b([^>]*)\/>/, '<sheetView$1>' + pane + '</sheetView>');
      } else if (/<sheetView\b[^>]*>/.test(this.prefix)) {
        this.prefix = this.prefix.replace(/(<sheetView\b[^>]*>)/, '$1' + pane);
      } else {
        const views = '<sheetViews><sheetView workbookViewId="0">' + pane + '</sheetView></sheetViews>';
        if (/<dimension\b[^>]*\/>/.test(this.prefix)) {
          this.prefix = this.prefix.replace(/(<dimension\b[^>]*\/>)/, '$1' + views);
        } else {
          this.prefix = this.prefix.replace(/(<worksheet\b[^>]*>)/, '$1' + views);
        }
      }
    }
    this.dirty = true;
    return this;
  }

  /** The `<sheetProtection>` element's attributes, or null. In the suffix. */
  sheetProtection() {
    const m = /<sheetProtection\b([^>]*?)\/?>/.exec(this.suffix);
    return m ? attrs(m[1]) : null;
  }

  /**
   * Turn protection on (with the attribute set given) or off (null).
   * Placement per schema: after `<sheetCalcPr>` when one exists, else at the
   * head of the tail — protection precedes scenarios and everything after.
   */
  setSheetProtection(attrsText) {
    this.suffix = this.suffix.replace(/<sheetProtection\b[^>]*\/?>(<\/sheetProtection>)?/, '');
    if (attrsText !== null) {
      const el = '<sheetProtection' + attrsText + '/>';
      if (/<sheetCalcPr\b[^>]*\/?>/.test(this.suffix)) {
        this.suffix = this.suffix.replace(/(<sheetCalcPr\b[^>]*\/?>)/, '$1' + el);
      } else {
        this.suffix = el + this.suffix;
      }
    }
    this.dirty = true;
    return this;
  }

  /**
   * Reference a drawing part from this sheet: `<drawing r:id>`, placed per
   * schema — after everything else in the tail except tableParts and extLst.
   * A sheet holds at most one; a second call is a no-op.
   */
  setDrawingRef(relId) {
    if (/<drawing\b/.test(this.suffix)) return this;
    const el = '<drawing r:id="' + esc(relId)
      + '" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>';
    // The suffix INCLUDES the closing </worksheet> — a blind append would put
    // the element outside the root, which is exactly what the Office gate
    // caught on this method's first day. tableParts and extLst are the only
    // schema elements that may follow drawing.
    const before = /<tableParts\b|<extLst\b|<\/worksheet>/.exec(this.suffix);
    this.suffix = this.suffix.slice(0, before.index) + el + this.suffix.slice(before.index);
    this.dirty = true;
    return this;
  }

  /**
   * Append one `<conditionalFormatting>` block: after the last existing one,
   * else in schema position — before dataValidations and everything that
   * follows it in the tail.
   */
  addConditionalFormatting(block) {
    const close = '</conditionalFormatting>';
    const last = this.suffix.lastIndexOf(close);
    if (last >= 0) {
      const at = last + close.length;
      this.suffix = this.suffix.slice(0, at) + block + this.suffix.slice(at);
    } else {
      const anchor = /<dataValidations\b|<hyperlinks\b|<printOptions\b|<pageMargins\b|<pageSetup\b|<drawing\b|<tableParts\b|<extLst\b|<\/worksheet>/.exec(this.suffix);
      this.suffix = this.suffix.slice(0, anchor.index) + block + this.suffix.slice(anchor.index);
    }
    this.dirty = true;
    return this;
  }

  /** Remove the blocks whose sqref the predicate claims. Returns how many. */
  removeConditionalFormattings(pred) {
    let removed = 0;
    this.suffix = this.suffix.replace(
      /<conditionalFormatting\b([^>]*)>[\s\S]*?<\/conditionalFormatting>/g,
      (m, attrsText) => {
        const sqref = /sqref="([^"]*)"/.exec(attrsText)?.[1] ?? '';
        if (pred(sqref)) { removed += 1; return ''; }
        return m;
      },
    );
    if (removed) this.dirty = true;
    return removed;
  }

  /**
   * Append one `<dataValidation>` entry. Unlike conditional formatting's
   * many blocks, validations live in ONE `<dataValidations count>` block —
   * created in schema position (after conditionalFormatting, before
   * hyperlinks and the rest of the tail) when the sheet has none.
   */
  addDataValidation(entryXml) {
    const block = /<dataValidations\b([^>]*)>([\s\S]*?)<\/dataValidations>/.exec(this.suffix);
    if (block) {
      const count = (block[2].match(/<dataValidation\b/g) ?? []).length + 1;
      const attrs = block[1].replace(/\s*count="[^"]*"/, '') + ' count="' + count + '"';
      this.suffix = this.suffix.replace(block[0],
        '<dataValidations' + attrs.replace(/^ +/, ' ') + '>' + block[2] + entryXml + '</dataValidations>');
    } else {
      const el = '<dataValidations count="1">' + entryXml + '</dataValidations>';
      const anchor = /<hyperlinks\b|<printOptions\b|<pageMargins\b|<pageSetup\b|<drawing\b|<tableParts\b|<extLst\b|<\/worksheet>/.exec(this.suffix);
      this.suffix = this.suffix.slice(0, anchor.index) + el + this.suffix.slice(anchor.index);
    }
    this.dirty = true;
    return this;
  }

  /** Remove the entries the predicate claims (by sqref); an emptied block goes. */
  removeDataValidations(pred) {
    const block = /<dataValidations\b([^>]*)>([\s\S]*?)<\/dataValidations>/.exec(this.suffix);
    if (!block) return 0;
    let removed = 0;
    const kept = block[2].replace(
      /<dataValidation\b([^>]*?)(?:\/>|>[\s\S]*?<\/dataValidation>)/g,
      (m, attrsText) => {
        const sqref = /sqref="([^"]*)"/.exec(attrsText)?.[1] ?? '';
        if (pred(sqref)) { removed += 1; return ''; }
        return m;
      },
    );
    if (!removed) return 0;
    const count = (kept.match(/<dataValidation\b/g) ?? []).length;
    this.suffix = count
      ? this.suffix.replace(block[0],
        '<dataValidations' + block[1].replace(/\s*count="[^"]*"/, '') + ' count="' + count + '">' + kept + '</dataValidations>')
      : this.suffix.replace(block[0], '');
    this.dirty = true;
    return removed;
  }

  /** The sheet-level `<autoFilter>` block, verbatim, or null. In the suffix. */
  autoFilterXml() {
    return (/<autoFilter\b[^>]*(?:\/>|>[\s\S]*?<\/autoFilter>)/.exec(this.suffix) ?? [null])[0];
  }

  /**
   * Replace or remove the sheet-level autoFilter. A fresh block lands where
   * the schema wants it: after sheetCalcPr / sheetProtection /
   * protectedRanges / scenarios, before sortState, mergeCells and the rest.
   */
  setAutoFilter(block) {
    const existing = /<autoFilter\b[^>]*(?:\/>|>[\s\S]*?<\/autoFilter>)/;
    if (existing.test(this.suffix)) {
      this.suffix = this.suffix.replace(existing, block ?? '');
    } else if (block) {
      const before = /(<sheetCalcPr\b[^>]*\/?>|<sheetProtection\b[^>]*\/?>|<protectedRanges\b[^>]*>[\s\S]*?<\/protectedRanges>|<scenarios\b[^>]*>[\s\S]*?<\/scenarios>)/g;
      let at = 0;
      let m;
      while ((m = before.exec(this.suffix))) at = m.index + m[0].length;
      this.suffix = this.suffix.slice(0, at) + block + this.suffix.slice(at);
    }
    this.dirty = true;
    return this;
  }

  /** The `<scenarios>` block, verbatim, or null. Lives in the suffix. */
  scenariosXml() {
    return (/<scenarios\b[^>]*>[\s\S]*?<\/scenarios>/.exec(this.suffix) ?? [null])[0];
  }

  /**
   * Replace the whole `<scenarios>` block, or drop it when the text is empty.
   * A fresh block lands where the schema wants it: after sheetCalcPr /
   * sheetProtection / protectedRanges, before everything else in the tail.
   */
  setScenarios(block) {
    const paired = /<scenarios\b[^>]*>[\s\S]*?<\/scenarios>/;
    const empty = /<scenarios\b[^>]*\/>/;
    if (paired.test(this.suffix)) {
      this.suffix = this.suffix.replace(paired, block);
    } else if (empty.test(this.suffix)) {
      this.suffix = this.suffix.replace(empty, block);
    } else if (block) {
      const before = /(<sheetCalcPr\b[^>]*\/?>|<sheetProtection\b[^>]*\/?>|<protectedRanges\b[^>]*>[\s\S]*?<\/protectedRanges>)/g;
      let at = 0;
      let m;
      while ((m = before.exec(this.suffix))) at = m.index + m[0].length;
      this.suffix = this.suffix.slice(0, at) + block + this.suffix.slice(at);
    }
    this.dirty = true;
    return this;
  }

  render() {
    const body = this.rows
      .map((r) => (r.dirty ? '<row' + r.attrsStr + '>' + r.inner + '</row>' : r.xml))
      .join('');
    return this.prefix + this.openTag + body + '</sheetData>' + this.suffix;
  }
}

function renderCell(ref, value, style) {
  const s = style === null || style === undefined ? '' : ' s="' + esc(style) + '"';
  if (value === null || value === undefined || value === '') {
    return '<c r="' + ref + '"' + s + '/>';
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    return '<c r="' + ref + '"' + s + '><v>' + value + '</v></c>';
  }
  if (typeof value === 'boolean') {
    return '<c r="' + ref + '"' + s + ' t="b"><v>' + (value ? 1 : 0) + '</v></c>';
  }
  const text = String(value);
  if (text.startsWith('=')) {
    // Drop any cached <v>: a stale cached result next to a new formula is worse
    // than none, and the consumer recalculates on open.
    return '<c r="' + ref + '"' + s + '><f>' + esc(text.slice(1)) + '</f></c>';
  }
  return '<c r="' + ref + '"' + s + ' t="inlineStr"><is><t xml:space="preserve">' + esc(text) + '</t></is></c>';
}

/**
 * Canonicalise a range ref to `topLeft:bottomRight`, so `'C1:A1'` and `'A1:C1'`
 * compare equal and a single cell like `'B2'` becomes the degenerate `'B2:B2'`.
 * This is what lets addMerge de-duplicate and removeMerge match by identity.
 */
export function normalizeRange(ref) {
  const [a, b] = String(ref).split(':');
  const from = parseRef(a);
  const to = b ? parseRef(b) : from;
  const top = Math.min(from.row, to.row);
  const bottom = Math.max(from.row, to.row);
  const left = Math.min(from.col, to.col);
  const right = Math.max(from.col, to.col);
  return makeRef(top, left) + ':' + makeRef(bottom, right);
}

// ── structural reference adjustment ──────────────────────────────────────────

/** One index through an insert/delete at `at`. Null means it was deleted. */
function shiftIndex(v, at, count, op) {
  if (op === 'insert') return v >= at ? v + count : v;
  if (v >= at + count) return v - count;
  if (v >= at) return null;
  return v;
}

/**
 * A span `[lo, hi]` through an insert/delete. Excel semantics: an insertion
 * before the span moves it, inside it grows it; a deletion shrinks it by the
 * rows/columns removed and returns null once nothing of it remains.
 */
function shiftSpan(lo, hi, at, count, op) {
  if (op === 'insert') {
    return [lo >= at ? lo + count : lo, hi >= at ? hi + count : hi];
  }
  const end = at + count;
  const nlo = lo >= end ? lo - count : lo >= at ? at : lo;
  const nhi = hi >= end ? hi - count : hi >= at ? at - 1 : hi;
  return nlo > nhi ? null : [nlo, nhi];
}

const sameSheetName = (a, b) =>
  a !== null && b !== null && String(a).toLowerCase() === String(b).toLowerCase();

/** `{ colAbs, col, rowAbs, row, end }` for an A1 token at `pos`, or null. */
function matchCellAt(src, pos) {
  const re = /(\$?)([A-Za-z]{1,3})(\$?)(\d+)/y;
  re.lastIndex = pos;
  const m = re.exec(src);
  if (!m) return null;
  return {
    colAbs: m[1] === '$', col: colToIndex(m[2]),
    rowAbs: m[3] === '$', row: Number(m[4]) - 1,
    end: pos + m[0].length,
  };
}

/**
 * The reference starting at `pos`: a cell, a cell range, a whole-column range
 * (`A:C`) or a whole-row range (`1:5`). Null when the text there is not one.
 */
function matchRefAt(src, pos) {
  const a = matchCellAt(src, pos);
  if (a) {
    if (src[a.end] === ':') {
      const b = matchCellAt(src, a.end + 1);
      if (b) return { kind: 'range', a, b, end: b.end };
    }
    return { kind: 'cell', a, end: a.end };
  }
  const colRe = /(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/y;
  colRe.lastIndex = pos;
  const cm = colRe.exec(src);
  if (cm && !/[A-Za-z0-9_.]/.test(src[pos + cm[0].length] ?? '')) {
    return {
      kind: 'cols',
      a: { abs: cm[1] === '$', index: colToIndex(cm[2]) },
      b: { abs: cm[3] === '$', index: colToIndex(cm[4]) },
      end: pos + cm[0].length,
    };
  }
  const rowRe = /(\$?)(\d+):(\$?)(\d+)/y;
  rowRe.lastIndex = pos;
  const rm = rowRe.exec(src);
  if (rm && !/[0-9]/.test(src[pos + rm[0].length] ?? '')) {
    return {
      kind: 'rows',
      a: { abs: rm[1] === '$', index: Number(rm[2]) - 1 },
      b: { abs: rm[3] === '$', index: Number(rm[4]) - 1 },
      end: pos + rm[0].length,
    };
  }
  return null;
}

const cellText = (p) => (p.colAbs ? '$' : '') + indexToCol(p.col) + (p.rowAbs ? '$' : '') + (p.row + 1);

/** The adjusted text for one matched reference, `#REF!` when it is gone. */
function adjustedRefText(src, start, ref, ctx) {
  const { axis, op, at, count } = ctx;
  const original = src.slice(start, ref.end);
  if (ref.kind === 'cell') {
    const v = axis === 'row' ? ref.a.row : ref.a.col;
    const shifted = shiftIndex(v, at, count, op);
    if (shifted === null) return '#REF!';
    if (shifted === v) return original;
    const next = { ...ref.a };
    if (axis === 'row') next.row = shifted;
    else next.col = shifted;
    return cellText(next);
  }
  if (ref.kind === 'range') {
    const lo = axis === 'row' ? ref.a.row : ref.a.col;
    const hi = axis === 'row' ? ref.b.row : ref.b.col;
    const span = shiftSpan(lo, hi, at, count, op);
    if (span === null) return '#REF!';
    if (span[0] === lo && span[1] === hi) return original;
    const a = { ...ref.a };
    const b = { ...ref.b };
    if (axis === 'row') { a.row = span[0]; b.row = span[1]; }
    else { a.col = span[0]; b.col = span[1]; }
    return cellText(a) + ':' + cellText(b);
  }
  // Whole-column and whole-row ranges only move on their own axis.
  if ((ref.kind === 'cols') !== (axis === 'col')) return original;
  const span = shiftSpan(ref.a.index, ref.b.index, at, count, op);
  if (span === null) return '#REF!';
  if (span[0] === ref.a.index && span[1] === ref.b.index) return original;
  const letter = (p, i) => (p.abs ? '$' : '') + (ref.kind === 'cols' ? indexToCol(i) : String(i + 1));
  return letter(ref.a, span[0]) + ':' + letter(ref.b, span[1]);
}

/**
 * Move every A1-style reference in a formula through a structural edit.
 *
 * Text level, deliberately: the stored formula IS text, and the calc layer
 * re-parses on reload, so adjusting the text is both necessary and sufficient.
 * String literals are skipped; sheet qualifiers (quoted and bare) are honoured
 * so only references aimed at the edited sheet move; a reference into a
 * deleted slice becomes `#REF!` with its qualifier kept, which is how Excel
 * writes it.
 *
 * @param {string} text the formula WITHOUT its leading `=` (as stored in `<f>`)
 * @param {object} ctx { editedSheet, currentSheet, axis: 'row'|'col',
 *                       op: 'insert'|'delete', at, count } — `currentSheet` is
 *                       the sheet the formula lives on (null for a defined
 *                       name, whose unqualified refs bind to nothing).
 */
export function adjustFormula(text, ctx) {
  const src = String(text);
  const n = src.length;
  let out = '';
  let i = 0;
  const idChar = (ch) => /[A-Za-z0-9_.]/.test(ch ?? '');

  const refBoundaryOk = (end) => {
    const next = src[end] ?? '';
    return next !== '(' && next !== '!' && !idChar(next) && next !== '$';
  };

  const emit = (start, refStart, qualifier, ref) => {
    const target = qualifier ?? ctx.currentSheet;
    if (!sameSheetName(target, ctx.editedSheet)) { out += src.slice(start, ref.end); return; }
    out += src.slice(start, refStart) + adjustedRefText(src, refStart, ref, ctx);
  };

  while (i < n) {
    const ch = src[i];
    if (ch === '"') {
      // A string literal, with "" as the escape. Nothing inside is a reference.
      let j = i + 1;
      while (j < n) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') j += 2;
          else { j += 1; break; }
        } else j += 1;
      }
      out += src.slice(i, j);
      i = j;
      continue;
    }
    if (ch === "'") {
      // A quoted sheet name, with '' as the escape, then `!` and a reference.
      let j = i + 1;
      let name = '';
      while (j < n) {
        if (src[j] === "'") {
          if (src[j + 1] === "'") { name += "'"; j += 2; }
          else { j += 1; break; }
        } else { name += src[j]; j += 1; }
      }
      if (src[j] === '!') {
        const ref = matchRefAt(src, j + 1);
        if (ref && refBoundaryOk(ref.end)) { emit(i, j + 1, name, ref); i = ref.end; continue; }
        out += src.slice(i, j + 1);
        i = j + 1;
        continue;
      }
      out += src.slice(i, j);
      i = j;
      continue;
    }
    const prev = i > 0 ? src[i - 1] : '';
    const starts = /[A-Za-z_$0-9]/.test(ch) && !idChar(prev) && prev !== '$';
    if (starts) {
      // A bare sheet qualifier first: `Data!B2`. A name that parses as a cell
      // reference is never a sheet name here — Excel quotes those.
      const qualRe = /([A-Za-z_][A-Za-z0-9_.]*)!/y;
      qualRe.lastIndex = i;
      const qm = qualRe.exec(src);
      if (qm) {
        const ref = matchRefAt(src, i + qm[0].length);
        if (ref && refBoundaryOk(ref.end)) { emit(i, i + qm[0].length, qm[1], ref); i = ref.end; continue; }
        out += qm[0];
        i += qm[0].length;
        continue;
      }
      const ref = matchRefAt(src, i);
      if (ref && refBoundaryOk(ref.end)) { emit(i, i, null, ref); i = ref.end; continue; }
      // Not a reference: consume the whole identifier so `LOG10(` is never
      // re-scanned from the middle and mistaken for the cell LOG10.
      let j = i + 1;
      while (j < n && (idChar(src[j]) || src[j] === '$')) j += 1;
      out += src.slice(i, j);
      i = j;
      continue;
    }
    out += ch;
    i += 1;
  }
  return out;
}

export class Workbook {
  constructor(pkg) {
    this.pkg = pkg;
    this.mainPart = pkg.mainDocument();
    if (!this.mainPart.startsWith('xl/')) throw new Error('not a spreadsheet package: ' + this.mainPart);
    this._sheets = null;
    this._loaded = new Map(); // part name -> SheetPart
    this._sharedStrings = null;
  }

  static open(buf) { return new Workbook(OoxmlPackage.read(buf)); }

  /** @returns {Array<{name: string, sheetId: string, rId: string, part: string}>} */
  sheets() {
    if (this._sheets) return this._sheets;
    const xml = this.pkg.text(this.mainPart);
    const rels = new Map(this.pkg.rels(this.mainPart).map((r) => [r.Id, r.Target]));
    const out = [];
    for (const m of xml.matchAll(/<sheet\b([^>]*)\/?>/g)) {
      const a = attrs(m[1]);
      const rId = a['r:id'] ?? a['relationships:id'] ?? a.id;
      const target = rels.get(rId);
      if (!target) continue;
      out.push({
        name: a.name,
        sheetId: a.sheetId,
        rId,
        part: OoxmlPackage.resolveTarget(this.mainPart, target),
      });
    }
    this._sheets = out;
    return out;
  }

  sheetNames() { return this.sheets().map((s) => s.name); }

  _sheetPart(sheetName) {
    const sheet = this.sheets().find((s) => s.name === sheetName);
    if (!sheet) throw new Error('no such sheet: ' + sheetName);
    if (!this._loaded.has(sheet.part)) {
      this._loaded.set(sheet.part, new SheetPart(this.pkg.text(sheet.part)));
    }
    return { sheet, part: this._loaded.get(sheet.part) };
  }

  sharedStrings() {
    if (this._sharedStrings) return this._sharedStrings;
    const name = 'xl/sharedStrings.xml';
    if (!this.pkg.has(name)) return (this._sharedStrings = []);
    const xml = this.pkg.text(name);
    const out = [];
    for (const m of xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)) {
      out.push([...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => unesc(t[1])).join(''));
    }
    return (this._sharedStrings = out);
  }

  /** Resolved display value, following shared-string indices. */
  getCell(sheetName, ref) {
    const { row, col } = parseRef(ref);
    const { part } = this._sheetPart(sheetName);
    const cell = part.getCell(row, col);
    if (!cell) return null;
    // A dataTable formula has no expression of its own — the element is the
    // MARKER of a what-if table, and the cell's truth is its cached value.
    // Returning '=' plus its empty body would hand every consumer a garbage
    // formula where Excel shows a number. (A SELF-CLOSING marker never
    // matched this regex and always fell through to the value — the paired
    // form now does the same.)
    const f = /<f([^>]*)>([\s\S]*?)<\/f>/.exec(cell.inner);
    if (f && !/\bt="dataTable"/.test(f[1])) return '=' + unesc(f[2]);
    const v = /<v>([\s\S]*?)<\/v>/.exec(cell.inner);
    if (cell.type === 's' && v) return this.sharedStrings()[Number(v[1])] ?? null;
    if (cell.type === 'inlineStr' || cell.type === 'str') {
      const t = /<t[^>]*>([\s\S]*?)<\/t>/.exec(cell.inner);
      if (t) return unesc(t[1]);
    }
    if (cell.type === 'b' && v) return v[1] === '1';
    if (cell.type === 'e' && v) return unesc(v[1]);
    if (!v) return null;
    const raw = unesc(v[1]);
    // A cell with no `t` attribute, or t="n", holds a NUMBER. Returning the
    // string would make it text to everything downstream — and a spreadsheet
    // skips text inside a range, so SUM would quietly ignore the whole column.
    if (!cell.type || cell.type === 'n') {
      const n = Number(raw);
      if (raw.trim() !== '' && Number.isFinite(n)) return n;
    }
    return raw;
  }

  setCell(sheetName, ref, value) {
    const { row, col } = parseRef(ref);
    const { part } = this._sheetPart(sheetName);
    part.setCell(row, col, value);
    return this;
  }

  /**
   * Splice a bodiless formula marker — `<f t="dataTable" .../>` — into a
   * cell that already holds its cached value. Any formula the cell carried
   * goes: a marker and an expression cannot both claim the same cell.
   */
  setCellFormulaMarker(sheetName, cellRef, attrsText) {
    const { row, col } = parseRef(cellRef);
    const { part } = this._sheetPart(sheetName);
    const rowRec = part._rowAt(row);
    const cell = rowRec ? SheetPart._parseCells(rowRec.inner).find((c) => c.col === col) : null;
    if (!cell) throw new Error('no cell at ' + cellRef + ' to mark');
    const inner = cell.inner.replace(/<f\b[^>]*(?:\/>|>[\s\S]*?<\/f>)/, '');
    const open = /^<c\b[^>]*>/.exec(cell.xml)?.[0]
      ?? cell.xml.replace(/\/>$/, '>');
    const rebuilt = open + '<f' + attrsText + '/>' + inner + '</c>';
    rowRec.inner = rowRec.inner.replace(cell.xml, rebuilt);
    rowRec.dirty = true;
    part.dirty = true;
    return this;
  }

  /** Point a cell at a style index, leaving its value alone. */
  setCellStyle(sheetName, ref, styleIndex) {
    const { row, col } = parseRef(ref);
    const { part } = this._sheetPart(sheetName);
    part.setCellStyle(row, col, styleIndex);
    return this;
  }

  /** Write a rectangular block anchored at a cell. */
  setRange(sheetName, anchor, values) {
    const { row, col } = parseRef(anchor);
    values.forEach((line, r) => {
      line.forEach((v, c) => {
        if (v === undefined) return;
        this.setCell(sheetName, makeRef(row + r, col + c), v);
      });
    });
    return this;
  }

  definedNames() {
    const xml = this.pkg.text(this.mainPart);
    const block = /<definedNames>([\s\S]*?)<\/definedNames>/.exec(xml);
    if (!block) return [];
    const out = [];
    for (const m of block[1].matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
      out.push({ name: attrs(m[1]).name, ref: unesc(m[2]), attrsStr: m[1] });
    }
    return out;
  }

  /** Add or replace a defined name, leaving the rest of workbook.xml alone. */
  setDefinedName(name, ref) {
    let xml = this.pkg.text(this.mainPart);
    const entry = '<definedName name="' + esc(name) + '">' + esc(ref) + '</definedName>';
    const existing = new RegExp('<definedName\\b[^>]*name="' + name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"[^>]*>[\\s\\S]*?</definedName>');
    if (existing.test(xml)) {
      xml = xml.replace(existing, entry);
    } else if (/<definedNames>/.test(xml)) {
      xml = xml.replace('</definedNames>', entry + '</definedNames>');
    } else {
      // Schema order matters: definedNames must follow sheets.
      xml = xml.replace('</sheets>', '</sheets><definedNames>' + entry + '</definedNames>');
    }
    this.pkg.write_(this.mainPart, xml);
    return this;
  }

  /**
   * Every table (ListObject) in the workbook, read from the table parts each
   * sheet's rels point at. `ref` covers header rows, data and totals rows;
   * `columns` are the header names structured references resolve against.
   */
  tables() {
    const out = [];
    for (const { name: sheetName, part } of this.sheets()) {
      for (const rel of this.pkg.rels(part)) {
        if (!String(rel.Type).endsWith('/table')) continue;
        const partName = OoxmlPackage.resolveTarget(part, rel.Target);
        if (!this.pkg.has(partName)) continue;
        const xml = this.pkg.text(partName);
        const a = attrs((/<table\b([^>]*?)>/.exec(xml) ?? [])[1] ?? '');
        if (!a.ref) continue;
        const styleInfo = attrs((/<tableStyleInfo\b([^>]*?)\/?>/.exec(xml) ?? [])[1] ?? '');
        out.push({
          sheet: sheetName,
          part: partName,
          name: a.displayName ?? a.name,
          ref: a.ref,
          headerRowCount: a.headerRowCount === undefined ? 1 : Number(a.headerRowCount),
          totalsRowCount: Number(a.totalsRowCount ?? 0),
          columns: [...xml.matchAll(/<tableColumn\b([^>]*?)\/?>/g)]
            .map((m) => attrs(m[1]).name)
            .filter((n) => n !== undefined)
            .map(unesc),
          // How the table asked to LOOK. The named Excel style is recorded but
          // painted in the renderer's own palette, the way charts are.
          styleName: styleInfo.name ?? null,
          showRowStripes: styleInfo.showRowStripes === '1',
        });
      }
    }
    return out;
  }

  /**
   * A table's value-filter state: colId -> array of allowed display texts.
   * A filter kind this editor does not model (customFilters, top10,
   * dynamicFilter) reads as null — present, opaque, and preserved.
   */
  tableFilters(tablePartName) {
    if (!this.pkg.has(tablePartName)) return new Map();
    const af = /<autoFilter\b[^>]*>([\s\S]*?)<\/autoFilter>/.exec(this.pkg.text(tablePartName));
    return af ? parseFilterColumns(af[1]) : new Map();
  }

  /**
   * Replace ONE column's value filter in a table part; `values` null or empty
   * clears it. Other columns' filterColumn blocks — including kinds we do not
   * model — ride through verbatim, re-assembled in colId order as Excel
   * writes them.
   */
  setTableFilter(tablePartName, colId, values) {
    let xml = this.pkg.text(tablePartName);
    if (!/<autoFilter\b/.test(xml)) {
      const ref = attrs((/<table\b([^>]*?)>/.exec(xml) ?? [])[1] ?? '').ref ?? '';
      xml = xml.replace(/(<table\b[^>]*>)/, '$1<autoFilter ref="' + esc(ref) + '"></autoFilter>');
    }
    const af = /<autoFilter\b[^>]*(?:\/>|>[\s\S]*?<\/autoFilter>)/.exec(xml);
    xml = xml.replace(af[0], withFilterColumn(af[0], colId, values));
    this.pkg.write_(tablePartName, xml);
    return this;
  }

  /**
   * The sheet-level (non-table) autofilter — Excel's Data → Filter over a
   * plain range. `{ ref, top, left, bottom, right, filters }` or null.
   */
  sheetAutoFilter(sheetName) {
    const block = this._sheetPart(sheetName).part.autoFilterXml();
    if (!block) return null;
    const ref = attrs((/<autoFilter\b([^>]*?)(?:\/>|>)/.exec(block) ?? [])[1] ?? '').ref;
    if (!ref) return null;
    const [a, b] = String(ref).split(':');
    let p1;
    let p2;
    try {
      p1 = parseRef(a.replace(/\$/g, ''));
      p2 = b ? parseRef(b.replace(/\$/g, '')) : p1;
    } catch {
      return null;
    }
    const inner = (/<autoFilter\b[^>]*>([\s\S]*?)<\/autoFilter>/.exec(block) ?? [null, ''])[1];
    return {
      ref,
      top: Math.min(p1.row, p2.row),
      left: Math.min(p1.col, p2.col),
      bottom: Math.max(p1.row, p2.row),
      right: Math.max(p1.col, p2.col),
      filters: parseFilterColumns(inner),
    };
  }

  /**
   * Turn the sheet autofilter on over `ref`, or off with null. Turning it on
   * fresh carries no filterColumn state; turning it off drops whatever state
   * it had, exactly as Excel's Data → Filter toggle behaves.
   */
  setSheetAutoFilter(sheetName, ref) {
    this._sheetPart(sheetName).part.setAutoFilter(
      ref === null || ref === undefined ? null : '<autoFilter ref="' + esc(String(ref)) + '"/>',
    );
    return this;
  }

  /** Append one dataValidation entry to a sheet's (one) validations block. */
  addDataValidation(sheetName, entryXml) {
    this._sheetPart(sheetName).part.addDataValidation(entryXml);
    return this;
  }

  /** Remove the sheet's validation entries the predicate claims (by sqref). */
  removeDataValidations(sheetName, pred) {
    return this._sheetPart(sheetName).part.removeDataValidations(pred);
  }

  /** Append one conditionalFormatting block to a sheet, in schema position. */
  addConditionalFormatting(sheetName, sqref, ruleXml) {
    this._sheetPart(sheetName).part.addConditionalFormatting(
      '<conditionalFormatting sqref="' + esc(sqref) + '">' + ruleXml + '</conditionalFormatting>',
    );
    return this;
  }

  /** Remove the sheet's CF blocks the predicate claims (by sqref). */
  removeConditionalFormattings(sheetName, pred) {
    return this._sheetPart(sheetName).part.removeConditionalFormattings(pred);
  }

  /** Replace ONE column's value filter in the sheet-level autofilter. */
  setSheetFilter(sheetName, colId, values) {
    const { part } = this._sheetPart(sheetName);
    const block = part.autoFilterXml();
    if (!block) throw new Error('sheet ' + sheetName + ' has no autofilter to set a filter on');
    part.setAutoFilter(withFilterColumn(block, colId, values));
    return this;
  }

  /**
   * Set exactly which rows of a band are hidden — what a filter means to the
   * sheet part. Rows in the band but not in `hiddenRows` are UNhidden, so
   * clearing a filter is the same call with an empty set. A hidden row that
   * does not exist yet is created empty, or there would be nothing to carry
   * the attribute.
   */
  setRowsHidden(sheetName, from, to, hiddenRows) {
    const { part } = this._sheetPart(sheetName);
    for (let r = from; r <= to; r++) {
      const wantHidden = hiddenRows.has(r);
      let row = part._rowAt(r);
      if (!row) {
        if (!wantHidden) continue;
        row = { index: r, attrsStr: ' r="' + (r + 1) + '"', inner: '', xml: '', dirty: true };
        part.rows.push(row);
        part.rows.sort((a, b) => a.index - b.index);
        part.dirty = true;
      }
      const has = /\shidden="[^"]*"/.test(row.attrsStr);
      if (wantHidden && !has) {
        row.attrsStr += ' hidden="1"';
        row.dirty = true;
        part.dirty = true;
      } else if (!wantHidden && has) {
        row.attrsStr = row.attrsStr.replace(/\s+hidden="[^"]*"/, '');
        row.dirty = true;
        part.dirty = true;
      }
    }
    return this;
  }

  /**
   * The sheet's drawing part, created when it does not exist yet: the part,
   * its content type, the sheet-rels entry and the `<drawing r:id>` element,
   * all wired in one call. Returns the part name either way, so an insert
   * needs no branch of its own.
   */
  ensureSheetDrawing(sheetName) {
    const sheetPartName = this.partNameFor(sheetName);
    const existing = this.pkg.rels(sheetPartName).find((r) => String(r.Type).endsWith('/drawing'));
    if (existing) return OoxmlPackage.resolveTarget(sheetPartName, existing.Target);

    const n = this.pkg.nextPartNumber('xl/drawings/', 'drawing');
    const partName = 'xl/drawings/drawing' + n + '.xml';
    this.pkg.addPart(partName,
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
      + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"></xdr:wsDr>',
      'application/vnd.openxmlformats-officedocument.drawing+xml');
    const relId = this.pkg.addRelationshipTo(sheetPartName,
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
      '../drawings/drawing' + n + '.xml');
    this._sheetPart(sheetName).part.setDrawingRef(relId);
    return partName;
  }

  /** Append one anchor to a drawing part. The next id, so callers need none. */
  appendDrawingAnchor(drawingPartName, anchorXmlOf) {
    const xml = this.pkg.text(drawingPartName);
    const ids = [...xml.matchAll(/<xdr:cNvPr\b[^>]*\bid="(\d+)"/g)].map((m) => Number(m[1]));
    const id = (ids.length ? Math.max(...ids) : 1) + 1;
    this.pkg.write_(drawingPartName, xml.replace('</xdr:wsDr>', anchorXmlOf(id) + '</xdr:wsDr>'));
    return id;
  }

  /**
   * Ask the next Excel to recalculate everything on open.
   *
   * Set when a save writes array markers: their ghost cells carry no cached
   * values, and this is the standing hint that makes every consumer compute
   * them. Placement matters — `calcPr` precedes `pivotCaches` in the schema,
   * and appending it at the end of a workbook that lists a pivot cache would
   * be exactly the kind of violation the Office gate exists to catch.
   */
  setFullCalcOnLoad() {
    let xml = this.pkg.text(this.mainPart);
    if (/fullCalcOnLoad="1"/.test(xml)) return this;
    if (/<calcPr\b/.test(xml)) {
      xml = xml.replace(/<calcPr\b([^>]*?)\/>/, (_, a) =>
        '<calcPr' + a.replace(/\s*fullCalcOnLoad="[^"]*"/, '') + ' fullCalcOnLoad="1"/>');
    } else if (/<pivotCaches\b/.test(xml)) {
      xml = xml.replace(/<pivotCaches\b/, '<calcPr calcId="191029" fullCalcOnLoad="1"/><pivotCaches');
    } else {
      xml = xml.replace('</workbook>', '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>');
    }
    this.pkg.write_(this.mainPart, xml);
    return this;
  }

  /** Remove a defined name. An emptied <definedNames> block goes with it. */
  deleteDefinedName(name) {
    let xml = this.pkg.text(this.mainPart);
    const existing = new RegExp('<definedName\\b[^>]*name="' + String(name).replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"[^>]*>[\\s\\S]*?</definedName>');
    if (!existing.test(xml)) return this;
    xml = xml.replace(existing, '').replace(/<definedNames>\s*<\/definedNames>/, '');
    this.pkg.write_(this.mainPart, xml);
    return this;
  }

  /**
   * A sheet's protection, digested for enforcement: is it on, does it carry
   * a password, and which of the two restrictions this editor honours are
   * relaxed. The attribute semantics are the spec's famous inversion —
   * `formatCells="1"` (the default) means formatting is FORBIDDEN.
   */
  sheetProtection(sheetName) {
    const a = this._sheetPart(sheetName).part.sheetProtection();
    if (!a || (a.sheet !== '1' && a.sheet !== 'true')) return { sheet: false, hasPassword: false };
    return {
      sheet: true,
      hasPassword: Boolean(a.password || a.hashValue || a.algorithmName),
      // Each flag here means "the author ALLOWED this" — true only when the
      // file says ="0", because the default of every restriction is on.
      formatCells: a.formatCells === '0',
      insertRows: a.insertRows === '0',
      deleteRows: a.deleteRows === '0',
    };
  }

  /** Protect a sheet (no password — see the view for why), or thaw it. */
  setSheetProtection(sheetName, on) {
    this._sheetPart(sheetName).part.setSheetProtection(
      on ? ' sheet="1" objects="1" scenarios="1"' : null,
    );
    return this;
  }

  /** The frozen pane on a sheet: {rows, cols} or null. */
  frozenPane(sheetName) {
    return this._sheetPart(sheetName).part.frozenPane();
  }

  /** Freeze rows/cols at a sheet's top-left; 0 and 0 thaws. */
  setFrozenPane(sheetName, rows, cols) {
    this._sheetPart(sheetName).part.setFrozenPane(rows, cols);
    return this;
  }

  // ── scenarios ─────────────────────────────────────────────────────────────

  /** A sheet's scenarios: name, comment, and the input cells with values. */
  scenarios(sheetName) {
    const { part } = this._sheetPart(sheetName);
    const block = part.scenariosXml();
    if (!block) return [];
    const out = [];
    for (const m of block.matchAll(/<scenario\b([^>]*)>([\s\S]*?)<\/scenario>/g)) {
      const a = attrs(m[1]);
      out.push({
        name: unesc(a.name ?? ''),
        comment: a.comment !== undefined ? unesc(a.comment) : null,
        cells: [...m[2].matchAll(/<inputCells\b([^>]*?)\/?>/g)].map((c) => {
          const ca = attrs(c[1]);
          return { ref: ca.r, value: ca.val !== undefined ? unesc(ca.val) : '' };
        }),
      });
    }
    return out;
  }

  /** Add or replace one scenario. Input values are text, as Excel stores them. */
  setScenario(sheetName, { name, comment, cells }) {
    const kept = this.scenarios(sheetName).filter((s) => s.name !== name);
    kept.push({ name, comment: comment ?? null, cells });
    this._writeScenarios(sheetName, kept);
    return this;
  }

  deleteScenario(sheetName, name) {
    this._writeScenarios(sheetName, this.scenarios(sheetName).filter((s) => s.name !== name));
    return this;
  }

  _writeScenarios(sheetName, list) {
    const { part } = this._sheetPart(sheetName);
    const block = list.length
      ? '<scenarios>' + list.map((s) =>
        '<scenario name="' + esc(s.name) + '" locked="1"'
          + (s.comment ? ' comment="' + esc(s.comment) + '"' : '')
          + ' count="' + s.cells.length + '">'
          + s.cells.map((c) => '<inputCells r="' + esc(c.ref) + '" val="' + esc(String(c.value)) + '"/>').join('')
          + '</scenario>').join('')
        + '</scenarios>'
      : '';
    part.setScenarios(block);
  }

  /** The merge ranges a sheet declares, as refs like `'A1:C1'`. */
  merges(sheetName) {
    const { part } = this._sheetPart(sheetName);
    return part.mergeRefs();
  }

  /**
   * Add a merged region to a sheet. Idempotent: an identical range is not
   * duplicated, so a re-merge cannot inflate the `count`.
   */
  addMerge(sheetName, ref) {
    const norm = normalizeRange(ref);
    const { part } = this._sheetPart(sheetName);
    const refs = part.mergeRefs();
    if (!refs.some((r) => normalizeRange(r) === norm)) {
      refs.push(norm);
      part.setMerges(refs);
    }
    return this;
  }

  /** Remove a merged region from a sheet, matching regardless of ref ordering. */
  removeMerge(sheetName, ref) {
    const norm = normalizeRange(ref);
    const { part } = this._sheetPart(sheetName);
    part.setMerges(part.mergeRefs().filter((r) => normalizeRange(r) !== norm));
    return this;
  }

  /** Replace every merged region on a sheet in one call. */
  setMerges(sheetName, refs) {
    const { part } = this._sheetPart(sheetName);
    part.setMerges((refs ?? []).map(normalizeRange));
    return this;
  }

  // ── column widths and row heights ─────────────────────────────────────────

  /** Set one column's width in OOXML character units. */
  setColWidthChars(sheetName, colIndex, widthChars) {
    const { part } = this._sheetPart(sheetName);
    part.setColWidth(colIndex, widthChars);
    return this;
  }

  /** Set one row's height in points. */
  setRowHeightPoints(sheetName, rowIndex, points) {
    const { part } = this._sheetPart(sheetName);
    part.setRowHeight(rowIndex, points);
    return this;
  }

  // ── structural edits: insert and delete rows/columns ──────────────────────

  /**
   * The merge a delete would cut in half, or null when the delete is safe.
   * A merge fully inside the deleted slice is fine (it goes with the slice);
   * one that pokes out on either side would have to be split, which is never
   * what the user meant, so the caller refuses before touching anything.
   */
  findMergeSplit(sheetName, axis, at, count) {
    const { part } = this._sheetPart(sheetName);
    const end = at + count;
    for (const ref of part.mergeRefs()) {
      const [a, b] = normalizeRange(ref).split(':');
      const from = parseRef(a);
      const to = parseRef(b);
      const lo = axis === 'row' ? from.row : from.col;
      const hi = axis === 'row' ? to.row : to.col;
      const intersects = !(hi < at || lo >= end);
      const inside = lo >= at && hi < end;
      if (intersects && !inside) return ref;
    }
    return null;
  }

  insertRows(sheetName, at, count = 1) { return this._structuralEdit(sheetName, 'row', 'insert', at, count); }
  deleteRows(sheetName, at, count = 1) { return this._structuralEdit(sheetName, 'row', 'delete', at, count); }
  insertCols(sheetName, at, count = 1) { return this._structuralEdit(sheetName, 'col', 'insert', at, count); }
  deleteCols(sheetName, at, count = 1) { return this._structuralEdit(sheetName, 'col', 'delete', at, count); }

  /**
   * One structural edit, Excel semantics throughout:
   *
   *   - rows below / columns right of the edit are renumbered, cell refs too;
   *   - merges move with their cells; an insertion inside a merge grows it; a
   *     deletion swallows merges fully inside the slice and REFUSES to split
   *     one that partially overlaps;
   *   - `<col>` width records and the `<dimension>` hint are shifted;
   *   - every formula in the WORKBOOK is adjusted at the text level — the
   *     edited sheet's own references and other sheets' qualified references
   *     alike — and so are the defined names in workbook.xml.
   *
   * Cached `<v>` values next to formulas are left alone here; the calc layer
   * re-caches every formula on save, which is the one place caches are owed.
   */
  _structuralEdit(sheetName, axis, op, at, count) {
    if (!Number.isInteger(at) || at < 0) throw new Error('bad index for a structural edit: ' + at);
    if (!Number.isInteger(count) || count < 1) throw new Error('bad count for a structural edit: ' + count);
    const { part } = this._sheetPart(sheetName);

    if (op === 'delete') {
      const split = this.findMergeSplit(sheetName, axis, at, count);
      if (split) {
        throw new Error('cannot delete ' + (axis === 'row' ? 'rows' : 'columns')
          + ': that would split a merged cell (' + split + ')');
      }
    }

    // Merges first, from the part as it still stands.
    const merges = part.mergeRefs();
    if (merges.length) {
      const next = [];
      for (const ref of merges) {
        const [a, b] = normalizeRange(ref).split(':');
        const from = parseRef(a);
        const to = parseRef(b);
        const lo = axis === 'row' ? from.row : from.col;
        const hi = axis === 'row' ? to.row : to.col;
        const span = shiftSpan(lo, hi, at, count, op);
        if (span === null) continue; // fully inside a deleted slice
        next.push(axis === 'row'
          ? makeRef(span[0], from.col) + ':' + makeRef(span[1], to.col)
          : makeRef(from.row, span[0]) + ':' + makeRef(to.row, span[1]));
      }
      const changed = next.length !== merges.length
        || next.some((r, i) => r !== normalizeRange(merges[i]));
      if (changed) part.setMerges(next);
    }

    if (axis === 'row') {
      if (op === 'insert') part.insertRowsShift(at, count);
      else part.deleteRowsRange(at, count);
    } else {
      part.shiftCells(at, count, op);
      part.shiftColEntries(at, count, op);
    }
    part.shiftDimension(axis, op, at, count);
    part.dirty = true;

    // Formulas everywhere: the edited sheet's own, and every other sheet's
    // qualified references into it.
    for (const { name } of this.sheets()) {
      const { part: p } = this._sheetPart(name);
      p.adjustFormulas((f) => adjustFormula(f, {
        editedSheet: sheetName, currentSheet: name, axis, op, at, count,
      }));
    }

    // Defined names live in workbook.xml and are always sheet-qualified, so
    // they go through the same adjuster with no current sheet.
    this._adjustDefinedNames((t) => adjustFormula(t, {
      editedSheet: sheetName, currentSheet: null, axis, op, at, count,
    }));
    return this;
  }

  /** Rewrite each defined name through `fn`, touching only the ones that move. */
  _adjustDefinedNames(fn) {
    let xml = this.pkg.text(this.mainPart);
    const block = /<definedNames>[\s\S]*?<\/definedNames>/.exec(xml);
    if (!block) return this;
    let changed = false;
    const next = block[0].replace(/(<definedName\b[^>]*>)([\s\S]*?)(<\/definedName>)/g, (m, open, body, close) => {
      const source = unesc(body);
      const adjusted = fn(source);
      if (adjusted === source) return m;
      changed = true;
      return open + esc(adjusted) + close;
    });
    if (changed) {
      xml = xml.replace(block[0], next);
      this.pkg.write_(this.mainPart, xml);
    }
    return this;
  }

  /**
   * The XML of specific parts, right now — the raw material of an undo step.
   *
   * Deliberately takes NAMES rather than snapshotting the whole package. A
   * snapshot is taken before every edit, and copying a 5MB workbook per keystroke
   * would make typing quadratic. An edit only ever touches the sheet it is on, so
   * the caller passes that part and pays for that part.
   *
   * @param {string[]} names  part names, e.g. ['xl/worksheets/sheet1.xml']
   * @returns {Record<string, string>}
   */
  snapshotParts(names) {
    const out = {};
    for (const name of names) {
      const loaded = this._loaded.get(name);
      // A loaded part may hold un-flushed edits, so ask IT rather than the
      // package — otherwise the snapshot is of a state already superseded.
      out[name] = loaded ? loaded.render() : this.pkg.text(name);
    }
    return out;
  }

  /**
   * Put those parts back. Anything parsed from them is dropped so it re-reads.
   *
   * The caller is responsible for rebuilding whatever it derived from the old
   * XML — the calculation engine, the geometry, the drawings. This layer only
   * guarantees the bytes.
   */
  restoreParts(parts) {
    for (const [name, xml] of Object.entries(parts ?? {})) {
      this.pkg.write_(name, xml);
      this._loaded.delete(name);
    }
    this._sheets = null;
    this._sharedStrings = null;
    return this;
  }

  /** The part backing a sheet, by name — what a caller snapshots before editing. */
  partNameFor(sheetName) {
    const sheet = this.sheets().find((s) => s.name === sheetName);
    if (!sheet) throw new Error('no such sheet: ' + sheetName);
    return sheet.part;
  }

  /** Flush edited sheet parts and return the package bytes. */
  save() {
    for (const [partName, part] of this._loaded) {
      if (part.dirty) this.pkg.write_(partName, part.render());
    }
    return this.pkg.write();
  }

  /** Parts this workbook has actually rewritten. */
  modifiedParts() {
    for (const [partName, part] of this._loaded) {
      if (part.dirty) this.pkg.write_(partName, part.render());
    }
    return this.pkg.modifiedParts();
  }
}

const unesc = (s) =>
  String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_, d) => String.fromCharCode(Number(d)))
    .replace(/&amp;/g, '&');

/**
 * The filterColumn children of an autoFilter, table-part and sheet-level
 * alike: colId -> allowed display texts, or null for a filter kind this
 * editor does not model (customFilters, top10, dynamicFilter) — present,
 * opaque, preserved.
 */
function parseFilterColumns(afInner) {
  const out = new Map();
  for (const m of String(afInner ?? '').matchAll(/<filterColumn\b([^>]*?)>([\s\S]*?)<\/filterColumn>/g)) {
    const colId = Number(attrs(m[1]).colId);
    if (!Number.isInteger(colId)) continue;
    out.set(colId, /<filters\b/.test(m[2])
      ? [...m[2].matchAll(/<filter\b[^>]*?\bval="([^"]*)"/g)].map((f) => unesc(f[1]))
      : null);
  }
  return out;
}

/**
 * One autoFilter block with ONE column's value filter replaced (or cleared,
 * with null/empty values). Every other filterColumn — including kinds we do
 * not model — rides through verbatim, re-assembled in colId order as Excel
 * writes them. Pure string in, string out: the same transform serves the
 * table part and the sheet part.
 */
function withFilterColumn(afBlock, colId, values) {
  let block = String(afBlock).replace(/<autoFilter\b([^>]*?)\/>/, '<autoFilter$1></autoFilter>');
  const af = /<autoFilter\b([^>]*)>([\s\S]*?)<\/autoFilter>/.exec(block);
  const kept = [];
  for (const m of af[2].matchAll(/<filterColumn\b([^>]*?)(?:\/>|>[\s\S]*?<\/filterColumn>)/g)) {
    const id = Number(attrs(m[1]).colId);
    if (id !== colId) kept.push({ colId: id, xml: m[0] });
  }
  if (values && values.length) {
    kept.push({
      colId,
      xml: '<filterColumn colId="' + colId + '"><filters>'
        + values.map((v) => '<filter val="' + esc(String(v)) + '"/>').join('')
        + '</filters></filterColumn>',
    });
  }
  kept.sort((a, b) => a.colId - b.colId);
  block = block.replace(af[0], '<autoFilter' + af[1] + '>' + kept.map((k) => k.xml).join('') + '</autoFilter>');
  // A block left with no children collapses back to the self-closing form a
  // fresh toggle writes, so on-then-off round-trips to the same bytes.
  return block.replace(/<autoFilter\b([^>]*)><\/autoFilter>/, '<autoFilter$1/>');
}

export { SHEET_NS, unesc };
