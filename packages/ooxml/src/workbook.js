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

const REL_HYPERLINK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const XMLNS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XMLNS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_COMMENTS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const REL_VML = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing';
const CT_COMMENTS = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
const CT_VML = 'application/vnd.openxmlformats-officedocument.vmlDrawing';
/** Where `<legacyDrawing>` goes in a worksheet: after the drawing, before what follows it. */
const AFTER_LEGACY_DRAWING = /<legacyDrawingHF\b|<picture\b|<oleObjects\b|<controls\b|<webPublishItems\b|<tableParts\b|<extLst\b|<\/worksheet>/;
const REL_TABLE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table';
const CT_TABLE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';
/** Where `<hyperlinks>` goes when a sheet has none: after the data validations, before the print options. */
const AFTER_HYPERLINKS = /<printOptions\b|<pageMargins\b|<pageSetup\b|<headerFooter\b|<rowBreaks\b|<colBreaks\b|<customProperties\b|<cellWatches\b|<ignoredErrors\b|<smartTags\b|<drawing\b|<legacyDrawing\b|<legacyDrawingHF\b|<picture\b|<oleObjects\b|<controls\b|<webPublishItems\b|<tableParts\b|<extLst\b|<\/worksheet>/;

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
/** The notes in a comments part: `[{ ref, author, text }]`, the author by name. */
export function parseComments(xml) {
  const authors = [...(/<authors>([\s\S]*?)<\/authors>/.exec(xml)?.[1] ?? '').matchAll(/<author>([\s\S]*?)<\/author>/g)].map((m) => unesc(m[1]));
  const out = [];
  for (const m of xml.matchAll(/<comment\b([^>]*)>([\s\S]*?)<\/comment>/g)) {
    const ref = /\bref="([A-Z]+\d+)"/.exec(m[1])?.[1];
    if (!ref) continue;
    const authorId = Number(/\bauthorId="(\d+)"/.exec(m[1])?.[1] ?? -1);
    const text = [...m[2].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((t) => unesc(t[1])).join('');
    out.push({ ref, author: authors[authorId] ?? '', text });
  }
  return out;
}

/** A comments part from its notes; the authors listed once each, in order of first use. */
export function commentsXml(list) {
  const authors = [];
  const items = list.map((c) => {
    let i = authors.indexOf(c.author || '');
    if (i < 0) {
      authors.push(c.author || '');
      i = authors.length - 1;
    }
    return '<comment ref="' + c.ref + '" authorId="' + i + '"><text><r><t xml:space="preserve">' + esc(c.text || '') + '</t></r></text></comment>';
  });
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<comments xmlns="' + XMLNS_MAIN + '"><authors>' + authors.map((a) => '<author>' + esc(a) + '</author>').join('') + '</authors>'
    + '<commentList>' + items.join('') + '</commentList></comments>';
}

/** The head of a VML drawing as Excel writes one for its notes: the layout and the note shape's type. */
const VML_HEAD = '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">'
  + '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>'
  + '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';

/** The box Excel draws a note in: hidden until the pointer rests on the cell, anchored beside it. */
function noteBox({ row, col }, id, z) {
  const left = col + 1;
  const top = Math.max(0, row - 1);
  return '<v:shape id="_x0000_s' + id + '" type="#_x0000_t202" style="position:absolute;margin-left:' + (left * 48) + 'pt;margin-top:' + (top * 11.25) + 'pt;width:108pt;height:59.25pt;z-index:' + z + ';visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">'
    + '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>'
    + '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>'
    + '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>'
    + '<x:Anchor>' + [left, 15, top, 2, left + 2, 15, top + 4, 16].join(', ') + '</x:Anchor>'
    + '<x:AutoFill>False</x:AutoFill><x:Row>' + row + '</x:Row><x:Column>' + col + '</x:Column></x:ClientData></v:shape>';
}
const VML_SHAPE = /<v:shape\b[\s\S]*?<\/v:shape>/g;
const isNoteBoxFor = (shape, { row, col }) =>
  /ObjectType="Note"/.test(shape) && new RegExp('<x:Row>' + row + '</x:Row>\\s*<x:Column>' + col + '</x:Column>').test(shape);
function withoutNoteBox(xml, at) {
  return xml.replace(VML_SHAPE, (shape) => (isNoteBoxFor(shape, at) ? '' : shape));
}
function withNoteBox(xml, at) {
  const without = withoutNoteBox(xml, at);
  const ids = [...without.matchAll(/id="_x0000_s(\d+)"/g)].map((m) => Number(m[1]));
  const id = ids.length ? Math.max(...ids) + 1 : 1025;
  const z = (without.match(/<v:shape\b/g) || []).length + 1;
  const end = without.lastIndexOf('</xml>');
  return without.slice(0, end) + noteBox(at, id, z) + without.slice(end);
}

/**
 * An attribute set, replaced or taken off in a run of attribute text:
 * `undefined` leaves it as it is, `null` removes it, a string sets it.
 */
export function withAttr(attrsText, name, value) {
  if (value === undefined) return attrsText;
  const re = new RegExp('\\s+' + name + '="[^"]*"');
  if (value === null) return attrsText.replace(re, '');
  const pair = ' ' + name + '="' + esc(value) + '"';
  return re.test(attrsText) ? attrsText.replace(re, pair) : attrsText + pair;
}

/* ── threaded comments ───────────────────────────────────────────────────── */

/**
 * Excel 365's comments — a conversation on a cell, with replies and a
 * resolved flag — live in `xl/threadedComments/threadedCommentN.xml` beside
 * the sheet, their authors in `xl/persons/person.xml` beside the workbook.
 * Each thread also keeps a shadow in the classic comments part (authored
 * `tc={id}`), with a VML box, which is what an Excel without threaded
 * comments shows.
 */
const REL_THREADED = 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment';
const CT_THREADED = 'application/vnd.ms-excel.threadedcomments+xml';
const REL_PERSON = 'http://schemas.microsoft.com/office/2017/10/relationships/person';
const CT_PERSON = 'application/vnd.ms-excel.person+xml';
const XMLNS_TC = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';

/** A GUID in braces, as Excel names threaded comments and people. */
export function newGuid() {
  const uuid = globalThis.crypto?.randomUUID
    ? globalThis.crypto.randomUUID()
    : 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
      const r = Math.floor(Math.random() * 16);
      return (c === 'x' ? r : (r % 4) + 8).toString(16);
    });
  return '{' + uuid.toUpperCase() + '}';
}

/** Excel's `dT`: UTC to the hundredth of a second, with no zone letter. */
export function threadDate(date = new Date()) {
  const d = date instanceof Date ? date : new Date(date);
  return d.toISOString().replace(/(\.\d\d)\dZ$/, '$1');
}

/**
 * The comments in a threaded-comments part, in the part's order:
 * `[{ id, ref, date, personId, parentId, done, text, extra }]` — `extra`
 * keeps what this editor does not model (mentions) for writing back.
 */
export function parseThreadedComments(xml) {
  const out = [];
  for (const m of String(xml || '').matchAll(/<threadedComment\b([^>]*?)(?:\/>|>([\s\S]*?)<\/threadedComment>)/g)) {
    const a = attrs(m[1]);
    const inner = m[2] ?? '';
    const text = /<text>([\s\S]*?)<\/text>/.exec(inner);
    out.push({
      id: a.id ?? '',
      ref: a.ref ?? '',
      date: a.dT ?? null,
      personId: a.personId ?? null,
      parentId: a.parentId ?? null,
      done: a.done === '1' || a.done === 'true',
      text: text ? unesc(text[1]) : '',
      extra: inner.replace(/<text>[\s\S]*?<\/text>|<text\/>/, ''),
    });
  }
  return out;
}

export function threadedCommentsXml(list) {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<ThreadedComments xmlns="' + XMLNS_TC + '" xmlns:x="' + XMLNS_MAIN + '">'
    + list.map((c) => '<threadedComment ref="' + esc(c.ref) + '"'
      + (c.date ? ' dT="' + esc(c.date) + '"' : '')
      + (c.personId ? ' personId="' + esc(c.personId) + '"' : '')
      + ' id="' + esc(c.id) + '"'
      + (c.parentId ? ' parentId="' + esc(c.parentId) + '"' : '')
      + (!c.parentId && c.done ? ' done="1"' : '')
      + '><text>' + esc(c.text ?? '') + '</text>' + (c.extra || '') + '</threadedComment>').join('')
    + '</ThreadedComments>';
}

/** The people who wrote threaded comments: `[{ id, displayName, userId, providerId }]`. */
export function parsePersons(xml) {
  return [...String(xml || '').matchAll(/<person\b([^>]*?)\/?>/g)].map((m) => {
    const a = attrs(m[1]);
    return { id: a.id ?? '', displayName: unesc(a.displayName ?? ''), userId: a.userId ? unesc(a.userId) : null, providerId: a.providerId ?? 'None' };
  });
}

function personsXml(list) {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<personList xmlns="' + XMLNS_TC + '" xmlns:x="' + XMLNS_MAIN + '">'
    + list.map((p) => '<person displayName="' + esc(p.displayName) + '" id="' + esc(p.id) + '"'
      + (p.userId ? ' userId="' + esc(p.userId) + '"' : '') + ' providerId="' + esc(p.providerId || 'None') + '"/>').join('')
    + '</personList>';
}

/**
 * What an Excel without threaded comments shows for a thread: the classic
 * note's words, the conversation laid out as Excel lays out its own.
 */
export function threadShadowText(thread) {
  const lines = ['[Threaded comment]', '',
    'This cell has a threaded comment. A version of Excel that reads threaded comments shows it as a conversation; a change made to this copy of it will not reach the conversation.',
    ''];
  thread.comments.forEach((c, i) => {
    lines.push(i === 0 ? 'Comment:' : 'Reply:');
    lines.push('    ' + String(c.text ?? '').replace(/\n/g, '\n    '));
  });
  return lines.join('\n');
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
    this._rowsChanged();
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

  /**
   * The row with this index.
   *
   * Indexed rather than scanned. A linear search here is invisible on a small
   * sheet and quadratic on a real one: loading a 20,000-row worksheet spent 2.8
   * of its 5.3 seconds inside this one `find`, because every row loaded looks
   * itself up. The index is rebuilt whenever the row list changes shape, which
   * `_rowsChanged` announces.
   */
  _rowAt(rowIndex) {
    if (!this._rowIndex) {
      this._rowIndex = new Map();
      for (const row of this.rows) this._rowIndex.set(row.index, row);
    }
    return this._rowIndex.get(rowIndex) ?? null;
  }

  /** A row was added, removed or renumbered, so the index no longer holds. */
  _rowsChanged() {
    this._rowIndex = null;
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
    return SheetPart.cellsByCol(row).get(colIndex) ?? null;
  }

  /**
   * A row's cells, parsed once per VERSION of its XML and indexed by column.
   *
   * `getCell` used to re-parse the whole row for every cell asked for, which
   * made loading a sheet quadratic in its width: a tender workbook of 2.4
   * million mostly-blank formatted cells took eleven minutes to open, all of
   * it here. The key is the row's XML string itself — every edit replaces
   * it, so a stale cache is not expressible.
   */
  static cellsByCol(row) {
    if (row._cellsFor !== row.inner) {
      row._cellsFor = row.inner;
      row._cellsByCol = new Map(SheetPart._parseCells(row.inner).map((c) => [c.col, c]));
    }
    return row._cellsByCol;
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
      this._rowsChanged();
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
      this._rowsChanged();
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
      this._rowsChanged();
    }
    if (/\bht="[^"]*"/.test(row.attrsStr)) row.attrsStr = row.attrsStr.replace(/\bht="[^"]*"/, 'ht="' + points + '"');
    else row.attrsStr += ' ht="' + points + '"';
    if (/\bcustomHeight="[^"]*"/.test(row.attrsStr)) row.attrsStr = row.attrsStr.replace(/\bcustomHeight="[^"]*"/, 'customHeight="1"');
    else row.attrsStr += ' customHeight="1"';
    row.dirty = true;
    this.dirty = true;
    return this;
  }

  // ── the outline: levels, folded groups, which side summaries sit ──────────

  /**
   * One row's outline attributes, as Excel writes them on `<row>`:
   * `outlineLevel` (absent at 0), `hidden` and `collapsed` (absent when
   * off). Each of `level`, `hidden` and `collapsed` is left alone when not
   * given. A row that does not exist is made only when something is set.
   */
  setRowOutline(rowIndex, { level, hidden, collapsed } = {}) {
    let row = this._rowAt(rowIndex);
    if (!row) {
      if (!(level > 0) && !hidden && !collapsed) return this;
      row = { index: rowIndex, attrsStr: ' r="' + (rowIndex + 1) + '"', inner: '', xml: '', dirty: true };
      const at = this.rows.findIndex((r) => r.index > rowIndex);
      if (at < 0) this.rows.push(row);
      else this.rows.splice(at, 0, row);
      this._rowsChanged();
    }
    const before = row.attrsStr;
    row.attrsStr = withAttr(row.attrsStr, 'outlineLevel', level === undefined ? undefined : level > 0 ? String(level) : null);
    row.attrsStr = withAttr(row.attrsStr, 'hidden', hidden === undefined ? undefined : hidden ? '1' : null);
    row.attrsStr = withAttr(row.attrsStr, 'collapsed', collapsed === undefined ? undefined : collapsed ? '1' : null);
    if (row.attrsStr !== before) {
      row.dirty = true;
      this.dirty = true;
    }
    return this;
  }

  /**
   * The mirror for one column: the `<col>` record that covers it is carved
   * out, as a width edit carves it, so its neighbours keep their own
   * settings. A column with nothing left to say drops its record.
   */
  setColOutline(colIndex, { level, hidden, collapsed } = {}) {
    const target = colIndex + 1;
    const edit = (rest) => {
      let r = withAttr(rest, 'outlineLevel', level === undefined ? undefined : level > 0 ? String(level) : null);
      r = withAttr(r, 'hidden', hidden === undefined ? undefined : hidden ? '1' : null);
      return withAttr(r, 'collapsed', collapsed === undefined ? undefined : collapsed ? '1' : null);
    };
    const out = [];
    let placed = false;
    for (const e of this.colEntries()) {
      if (e.max < target || e.min > target) { out.push(e); continue; }
      if (e.min <= target - 1) out.push({ min: e.min, max: target - 1, rest: e.rest });
      out.push({ min: target, max: target, rest: edit(e.rest) });
      if (e.max >= target + 1) out.push({ min: target + 1, max: e.max, rest: e.rest });
      placed = true;
    }
    if (!placed) out.push({ min: target, max: target, rest: edit('') });
    // A record with nothing left on it but its span says nothing.
    return this.setColEntries(out.filter((e) => e.rest.trim() !== ''));
  }

  /**
   * The sheet-wide outline properties: the deepest levels in
   * `sheetFormatPr` (`outlineLevelRow`, `outlineLevelCol`) and the sides
   * the summaries sit on in `sheetPr/outlinePr`, written only where they
   * differ from Excel's defaults (below, right). Each is left alone when
   * not given.
   */
  setOutlineProps({ rowLevels, colLevels, summaryBelow, summaryRight } = {}) {
    if (rowLevels !== undefined || colLevels !== undefined) {
      let fmt = /<sheetFormatPr\b[^>]*?\/?>/.exec(this.prefix)?.[0];
      if (!fmt) {
        if (!(rowLevels > 0) && !(colLevels > 0)) return this._outlinePr(summaryBelow, summaryRight);
        // Schema order: after sheetViews, before cols (which sits last).
        const fresh = '<sheetFormatPr defaultRowHeight="15"/>';
        const cols = /<cols\b/.exec(this.prefix);
        this.prefix = cols ? this.prefix.slice(0, cols.index) + fresh + this.prefix.slice(cols.index) : this.prefix + fresh;
        fmt = fresh;
      }
      let next = fmt.replace(/\s*\/?>$/, '');
      next = withAttr(next, 'outlineLevelRow', rowLevels === undefined ? undefined : rowLevels > 0 ? String(rowLevels) : null);
      next = withAttr(next, 'outlineLevelCol', colLevels === undefined ? undefined : colLevels > 0 ? String(colLevels) : null);
      next += fmt.endsWith('/>') ? '/>' : '>';
      if (next !== fmt) {
        this.prefix = this.prefix.replace(fmt, next);
        this.dirty = true;
      }
    }
    return this._outlinePr(summaryBelow, summaryRight);
  }

  _outlinePr(summaryBelow, summaryRight) {
    if (summaryBelow === undefined && summaryRight === undefined) return this;
    const current = /<outlinePr\b([^>]*?)\/?>/.exec(this.prefix);
    let attrsText = current ? current[1] : '';
    attrsText = withAttr(attrsText, 'summaryBelow', summaryBelow === undefined ? undefined : summaryBelow ? null : '0');
    attrsText = withAttr(attrsText, 'summaryRight', summaryRight === undefined ? undefined : summaryRight ? null : '0');
    const element = attrsText.trim() ? '<outlinePr' + attrsText + '/>' : '';
    if (current) {
      this.prefix = this.prefix.replace(current[0], element);
    } else if (element) {
      // tabColor comes first inside sheetPr, then outlinePr, then pageSetUpPr.
      if (/<sheetPr\b[^>]*\/>/.test(this.prefix)) {
        this.prefix = this.prefix.replace(/<sheetPr\b([^>]*)\/>/, '<sheetPr$1>' + element + '</sheetPr>');
      } else if (/<sheetPr\b/.test(this.prefix)) {
        const tab = /<tabColor\b[^>]*\/>/.exec(this.prefix);
        this.prefix = tab
          ? this.prefix.replace(tab[0], tab[0] + element)
          : this.prefix.replace(/<sheetPr\b([^>]*)>/, '<sheetPr$1>' + element);
      } else {
        this.prefix = this.prefix.replace(/(<worksheet\b[^>]*>)/, '$1<sheetPr>' + element + '</sheetPr>');
      }
    } else {
      return this;
    }
    this.dirty = true;
    return this;
  }

  /**
   * One attribute on `sheetPr` itself — `filterMode`, which Excel sets
   * while an advanced filter hides rows in place. Null takes it off.
   */
  setSheetPrAttr(name, value) {
    const open = /<sheetPr\b([^>]*?)(\/?)>/.exec(this.prefix);
    if (!open) {
      if (value === null || value === undefined) return this;
      this.prefix = this.prefix.replace(/(<worksheet\b[^>]*>)/, '$1<sheetPr ' + name + '="' + esc(String(value)) + '"/>');
      this.dirty = true;
      return this;
    }
    const next = '<sheetPr' + withAttr(open[1], name, value === null || value === undefined ? null : String(value)) + open[2] + '>';
    if (next !== open[0]) {
      this.prefix = this.prefix.replace(open[0], next);
      this.dirty = true;
    }
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
    this._rowsChanged();
    this.dirty = true;
    return this;
  }

  /** Remove rows `[at, at+count)` and close the gap beneath them. */
  deleteRowsRange(at, count) {
    this.rows = this.rows.filter((row) => row.index < at || row.index >= at + count);
    this._rowsChanged();
    for (const row of this.rows) {
      if (row.index >= at + count) this._renumberRow(row, row.index - count);
    }
    this._rowsChanged();
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

  /** The first `<sheetView>`'s attributes — where Excel keeps how a sheet is shown. */
  sheetViewAttrs() {
    const m = /<sheetView\b([^>]*?)\/?>/.exec(this.prefix);
    return m ? attrs(m[1]) : {};
  }

  /**
   * One attribute of the first `<sheetView>` set (a string) or taken off
   * (null), the element made — in `<sheetViews>`, where the schema puts it —
   * when the sheet has none.
   */
  setSheetViewAttr(name, value) {
    const re = /<sheetView\b([^>]*?)(\/?)>/;
    if (!re.test(this.prefix)) {
      if (value === null || value === undefined) return this;
      const views = '<sheetViews><sheetView workbookViewId="0"/></sheetViews>';
      if (/<dimension\b[^>]*\/>/.test(this.prefix)) this.prefix = this.prefix.replace(/(<dimension\b[^>]*\/>)/, '$1' + views);
      else this.prefix = this.prefix.replace(/(<worksheet\b[^>]*>)/, '$1' + views);
    }
    this.prefix = this.prefix.replace(re, (_, a, close) => '<sheetView' + withAttr(a, name, value === undefined ? null : value) + close + '>');
    this.dirty = true;
    return this;
  }

  /**
   * A split window — View → Split — as the sheet's view declares it: a
   * `<pane>` with no frozen state, xSplit and ySplit in twentieths of a
   * point from the window's edge, `topLeftCell` the first cell of the
   * bottom-right pane; the sheet view's own `topLeftCell` is the top-left
   * pane's. `{ xSplit, ySplit, topLeftCell, viewTopLeftCell }`, or null.
   */
  splitPane() {
    const m = /<pane\b([^>]*)\/?>/.exec(this.prefix);
    if (!m) return null;
    const a = attrs(m[1]);
    if (a.state === 'frozen' || a.state === 'frozenSplit') return null;
    const xSplit = Number(a.xSplit ?? 0) || 0;
    const ySplit = Number(a.ySplit ?? 0) || 0;
    if (!xSplit && !ySplit) return null;
    return { xSplit, ySplit, topLeftCell: a.topLeftCell ?? null, viewTopLeftCell: this.sheetViewAttrs().topLeftCell ?? null };
  }

  /**
   * Split the window (a spec) or take the split off (null). The pane goes
   * where a frozen one would — the first child of `<sheetView>` — and
   * replaces it: a sheet is split or frozen, never both.
   */
  setSplitPane(spec) {
    this.prefix = this.prefix.replace(/<pane\b[^>]*\/>|<pane\b[^>]*>[\s\S]*?<\/pane>/, '');
    if (!spec || (!spec.xSplit && !spec.ySplit)) {
      this.setSheetViewAttr('topLeftCell', null);
      this.dirty = true;
      return this;
    }
    const activePane = spec.xSplit && spec.ySplit ? 'bottomRight' : spec.ySplit ? 'bottomLeft' : 'topRight';
    const pane = '<pane'
      + (spec.xSplit ? ' xSplit="' + Math.round(spec.xSplit) + '"' : '')
      + (spec.ySplit ? ' ySplit="' + Math.round(spec.ySplit) + '"' : '')
      + (spec.topLeftCell ? ' topLeftCell="' + esc(spec.topLeftCell) + '"' : '')
      + ' activePane="' + activePane + '"/>';
    this.setSheetViewAttr('topLeftCell', spec.viewTopLeftCell || null);
    if (/<sheetView\b[^>]*\/>/.test(this.prefix)) {
      this.prefix = this.prefix.replace(/<sheetView\b([^>]*)\/>/, '<sheetView$1>' + pane + '</sheetView>');
    } else {
      this.prefix = this.prefix.replace(/(<sheetView\b[^>]*>)/, '$1' + pane);
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
  /** The sheet's hyperlinks as the file has them: ref, r:id, location, display, tooltip. */
  hyperlinks() {
    const block = /<hyperlinks\b[^>]*>([\s\S]*?)<\/hyperlinks>/.exec(this.suffix);
    if (!block) return [];
    const out = [];
    for (const m of block[1].matchAll(/<hyperlink\b([^>]*?)\/?>/g)) {
      const a = attrs(m[1]);
      if (!a.ref) continue;
      out.push({ ref: a.ref, rId: a['r:id'] ?? null, location: a.location ?? null, display: a.display ?? null, tooltip: a.tooltip ?? null });
    }
    return out;
  }

  /**
   * Put a hyperlink on a cell, in place of the one it had. The block sits
   * after the data validations and before the print options, and is made
   * where it belongs when the sheet has none; an r:id needs the prefix
   * declared on the root, which a sheet this suite built may not carry.
   */
  setHyperlink({ ref, rId = null, location = null, display = null, tooltip = null }) {
    const el = '<hyperlink ref="' + ref + '"'
      + (rId ? ' r:id="' + rId + '"' : '')
      + (location ? ' location="' + esc(location) + '"' : '')
      + (display ? ' display="' + esc(display) + '"' : '')
      + (tooltip ? ' tooltip="' + esc(tooltip) + '"' : '') + '/>';
    this.removeHyperlink(ref);
    const block = /<hyperlinks\b[^>]*>[\s\S]*?<\/hyperlinks>/.exec(this.suffix);
    if (block) {
      const at = block.index + block[0].length - '</hyperlinks>'.length;
      this.suffix = this.suffix.slice(0, at) + el + this.suffix.slice(at);
    } else {
      const anchor = AFTER_HYPERLINKS.exec(this.suffix);
      this.suffix = this.suffix.slice(0, anchor.index) + '<hyperlinks>' + el + '</hyperlinks>' + this.suffix.slice(anchor.index);
    }
    if (rId && !/\sxmlns:r=/.test(this.prefix)) {
      this.prefix = this.prefix.replace(/<worksheet\b/, '<worksheet xmlns:r="' + XMLNS_R + '"');
    }
    this.dirty = true;
    return this;
  }

  /** The r:id of the sheet's legacy (VML) drawing — where its notes' boxes live — or null. */
  legacyDrawingId() {
    return /<legacyDrawing\b[^>]*\br:id="([^"]+)"/.exec(this.suffix)?.[1] ?? null;
  }

  /** Point the sheet at a VML drawing part, where the notes' boxes are. */
  setLegacyDrawing(rId) {
    if (this.legacyDrawingId()) return this;
    const anchor = AFTER_LEGACY_DRAWING.exec(this.suffix);
    this.suffix = this.suffix.slice(0, anchor.index) + '<legacyDrawing r:id="' + rId + '"/>' + this.suffix.slice(anchor.index);
    if (!/\sxmlns:r=/.test(this.prefix)) {
      this.prefix = this.prefix.replace(/<worksheet\b/, '<worksheet xmlns:r="' + XMLNS_R + '"');
    }
    this.dirty = true;
    return this;
  }

  /** Take the hyperlink off a cell; a block emptied by it goes too. Returns whether there was one. */
  removeHyperlink(ref) {
    const block = /<hyperlinks\b[^>]*>([\s\S]*?)<\/hyperlinks>/.exec(this.suffix);
    if (!block) return false;
    const one = new RegExp('<hyperlink\\b[^>]*\\bref="' + ref.replace(/[$]/g, '\\$&') + '"[^>]*?/?>(?:</hyperlink>)?');
    if (!one.test(block[1])) return false;
    const inner = block[1].replace(one, '');
    const replacement = inner.trim() ? block[0].replace(block[1], () => inner) : '';
    this.suffix = this.suffix.slice(0, block.index) + replacement + this.suffix.slice(block.index + block[0].length);
    this.dirty = true;
    return true;
  }

  /**
   * Point this sheet at a table part: `<tableParts>` holds one
   * `<tablePart>` per table, in schema position — after everything in the
   * tail except extLst — its count kept right. The r prefix is declared on
   * the root if the sheet never needed it before.
   */
  addTablePart(rId) {
    const entry = '<tablePart r:id="' + esc(rId) + '"/>';
    const empty = /<tableParts\b[^>]*\/>/.exec(this.suffix);
    if (empty) {
      this.suffix = this.suffix.slice(0, empty.index) + '<tableParts count="1">' + entry + '</tableParts>' + this.suffix.slice(empty.index + empty[0].length);
    } else {
      const block = /<tableParts\b[^>]*>([\s\S]*?)<\/tableParts>/.exec(this.suffix);
      if (block) {
        const count = (block[1].match(/<tablePart\b/g) || []).length + 1;
        const replacement = '<tableParts count="' + count + '">' + block[1] + entry + '</tableParts>';
        this.suffix = this.suffix.slice(0, block.index) + replacement + this.suffix.slice(block.index + block[0].length);
      } else {
        const before = /<extLst\b|<\/worksheet>/.exec(this.suffix);
        this.suffix = this.suffix.slice(0, before.index) + '<tableParts count="1">' + entry + '</tableParts>' + this.suffix.slice(before.index);
      }
    }
    if (!/\sxmlns:r=/.test(this.prefix)) {
      this.prefix = this.prefix.replace(/<worksheet\b/, '<worksheet xmlns:r="' + XMLNS_R + '"');
    }
    this.dirty = true;
    return this;
  }

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
  /**
   * The order the schema gives the elements that follow sheetData.
   *
   * A worksheet's children are a sequence, not a set: an element in the wrong
   * place makes the file invalid, and Excel repairs it by throwing the part
   * away. Everything this engine writes into the tail of a sheet goes in
   * through `setTailElement`, which reads its position from here.
   */
  static TAIL_ORDER = [
    'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState',
    'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting',
    'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter',
    'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags',
    'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls',
    'webPublishItems', 'tableParts', 'extLst',
  ];

  /** Replace, insert or remove one element in the sheet's tail, in schema order. */
  setTailElement(name, xml) {
    const existing = new RegExp('<' + name + '\\b[^>]*(?:/>|>[\\s\\S]*?</' + name + '>)');
    if (existing.test(this.suffix)) {
      this.suffix = this.suffix.replace(existing, xml ?? '');
      this.dirty = true;
      return this;
    }
    if (!xml) return this;
    const after = SheetPart.TAIL_ORDER.slice(SheetPart.TAIL_ORDER.indexOf(name) + 1);
    const anchor = new RegExp('<(?:' + after.join('|') + ')\\b|</worksheet>').exec(this.suffix);
    const at = anchor ? anchor.index : this.suffix.length;
    this.suffix = this.suffix.slice(0, at) + xml + this.suffix.slice(at);
    this.dirty = true;
    return this;
  }

  /** One element of the tail, as it stands. */
  tailElement(name) {
    return (new RegExp('<' + name + '\\b[^>]*(?:/>|>[\\s\\S]*?</' + name + '>)').exec(this.suffix) ?? [null])[0];
  }

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

// ── sparklines (the x14 extension Excel writes them as) ─────────────────────

const SPARKLINE_EXT_URI = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}';
const XMLNS_X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const XMLNS_XM = 'http://schemas.microsoft.com/office/excel/2006/main';
/** Excel's own defaults for a fresh sparkline group — the colours nobody has chosen otherwise. */
const SPARKLINE_COLOURS = '<x14:colorSeries rgb="FF376092"/><x14:colorNegative rgb="FFD00000"/>'
  + '<x14:colorAxis rgb="FF000000"/><x14:colorMarkers rgb="FFD00000"/><x14:colorFirst rgb="FFD00000"/>'
  + '<x14:colorLast rgb="FFD00000"/><x14:colorHigh rgb="FFD00000"/><x14:colorLow rgb="FFD00000"/>';

/** A range ref's bounds, 0-based and ordered — `normalizeRange`'s numbers rather than its text. */
function rangeBounds(rangeRef) {
  const [a, b] = String(rangeRef).split(':');
  const from = parseRef(a);
  const to = b ? parseRef(b) : from;
  return {
    top: Math.min(from.row, to.row), bottom: Math.max(from.row, to.row),
    left: Math.min(from.col, to.col), right: Math.max(from.col, to.col),
  };
}
const boundsRef = (b) => {
  const a = makeRef(b.top, b.left);
  const c = makeRef(b.bottom, b.right);
  return a === c ? a : a + ':' + c;
};

/**
 * One sparkline per row of `at` (each taking the matching row of `data`) —
 * or, when `at` is itself a single row, one per COLUMN, matching columns
 * instead. A single-cell `at` takes the whole of `data` as its one series,
 * which is what a plain "select a row, sparkline in the next cell" makes.
 */
function planSparklines(dataRef, atRef) {
  const data = rangeBounds(dataRef);
  const at = rangeBounds(atRef);
  const atRows = at.bottom - at.top + 1;
  const atCols = at.right - at.left + 1;
  if (atRows === 1 && atCols === 1) {
    return [{ at: boundsRef(at), data: boundsRef(data) }];
  }
  if (atCols === 1 && atRows > 1) {
    if (data.bottom - data.top + 1 !== atRows) {
      throw new Error('the data and the sparklines need the same number of rows');
    }
    return Array.from({ length: atRows }, (_, i) => ({
      at: makeRef(at.top + i, at.left),
      data: boundsRef({ top: data.top + i, bottom: data.top + i, left: data.left, right: data.right }),
    }));
  }
  if (atRows === 1 && atCols > 1) {
    if (data.right - data.left + 1 !== atCols) {
      throw new Error('the data and the sparklines need the same number of columns');
    }
    return Array.from({ length: atCols }, (_, i) => ({
      at: makeRef(at.top, at.left + i),
      data: boundsRef({ top: data.top, bottom: data.bottom, left: data.left + i, right: data.left + i }),
    }));
  }
  throw new Error('the sparklines need a single row or a single column to sit in');
}

/** A sheet name, quoted the way a formula quotes one when it needs to be. */
const qualifySheetName = (name) =>
  (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : "'" + String(name).replace(/'/g, "''") + "'");

function sparklineGroupXml(type, sheetName, items) {
  const sparklines = items.map((it) =>
    '<x14:sparkline><xm:f>' + esc(qualifySheetName(sheetName) + '!' + it.data) + '</xm:f><xm:sqref>' + esc(it.at) + '</xm:sqref></x14:sparkline>').join('');
  return '<x14:sparklineGroup' + (type === 'column' ? ' type="column"' : '') + ' displayEmptyCellsAs="gap">'
    + SPARKLINE_COLOURS + '<x14:sparklines>' + sparklines + '</x14:sparklines></x14:sparklineGroup>';
}

const sparklineExtXml = (groupXml) => '<ext uri="' + SPARKLINE_EXT_URI + '" xmlns:x14="' + XMLNS_X14 + '">'
  + '<x14:sparklineGroups xmlns:xm="' + XMLNS_XM + '">' + groupXml + '</x14:sparklineGroups></ext>';

/** The sparkline groups an `extLst` block carries, decoded to `{type, colour, sparklines}`. */
function parseSparklineExt(extLstXml) {
  if (!extLstXml || !/<x14:sparklineGroups\b/.test(extLstXml)) return [];
  const groups = [];
  for (const gm of extLstXml.matchAll(/<x14:sparklineGroup\b([^>]*)>([\s\S]*?)<\/x14:sparklineGroup>/g)) {
    const gAttrs = attrs(gm[1]);
    const type = gAttrs.type === 'column' ? 'column' : 'line';
    const rgb = /<x14:colorSeries\b[^>]*\brgb="([^"]*)"/.exec(gm[2])?.[1] ?? 'FF376092';
    const colour = rgb.length === 8 ? rgb.slice(2) : rgb;
    const sparklines = [];
    for (const sm of gm[2].matchAll(/<x14:sparkline\b[^>]*>([\s\S]*?)<\/x14:sparkline>/g)) {
      const f = /<xm:f>([\s\S]*?)<\/xm:f>/.exec(sm[1]);
      const sq = /<xm:sqref>([\s\S]*?)<\/xm:sqref>/.exec(sm[1]);
      if (!f || !sq) continue;
      sparklines.push({ data: unesc(f[1]), at: unesc(sq[1]) });
    }
    groups.push({ type, colour, sparklines });
  }
  return groups;
}

/** Every sparkline whose cell falls in `at` (a single cell or a range) taken out; an emptied group, and an emptied ext, go with the last one. */
function removeSparklinesFromExt(extLstXml, at) {
  const bounds = rangeBounds(at);
  let removed = 0;
  let next = extLstXml.replace(/<x14:sparklineGroup\b([^>]*)>([\s\S]*?)<\/x14:sparklineGroup>/g, (whole, gAttrs, inner) => {
    const kept = inner.replace(/<x14:sparkline\b[^>]*>([\s\S]*?)<\/x14:sparkline>/g, (one, body) => {
      const sq = /<xm:sqref>([\s\S]*?)<\/xm:sqref>/.exec(body);
      if (!sq) return one;
      let cell;
      try { cell = parseRef(unesc(sq[1]).split(':')[0]); } catch { return one; }
      if (cell.row >= bounds.top && cell.row <= bounds.bottom && cell.col >= bounds.left && cell.col <= bounds.right) {
        removed += 1;
        return '';
      }
      return one;
    });
    if (!/<x14:sparkline\b/.test(kept)) return '';
    return '<x14:sparklineGroup' + gAttrs + '>' + kept + '</x14:sparklineGroup>';
  });
  if (!removed) return { removed: 0, xml: extLstXml };
  // An emptied sparklineGroups ext goes with it; whatever other ext elements
  // the sheet carries — the x14ac uid, a data bar's rule — ride through.
  next = next.replace(/<ext\b[^>]*\buri="\{05C60535-1F16-4fd2-B633-F4F36F0B64E0\}"[^>]*>\s*<x14:sparklineGroups\b[^>]*>\s*<\/x14:sparklineGroups>\s*<\/ext>/, '');
  return { removed, xml: next };
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

  /** A sheet name Excel would take: 1–31 characters, none of []:*?/\, not quoted, unique. */
  _checkSheetName(name, { except = null } = {}) {
    const trimmed = String(name ?? '').trim();
    if (!trimmed || trimmed.length > 31) throw new Error('a sheet name is 1 to 31 characters');
    if (/[\[\]:*?\/\\]/.test(trimmed)) throw new Error('a sheet name cannot contain [ ] : * ? / or \\');
    if (trimmed.startsWith("'") || trimmed.endsWith("'")) throw new Error('a sheet name cannot start or end with an apostrophe');
    const taken = this.sheetNames().some((n) => n !== except && n.toLowerCase() === trimmed.toLowerCase());
    if (taken) throw new Error('there is already a sheet called "' + trimmed + '"');
    return trimmed;
  }

  /**
   * A new, empty worksheet at the end of the tab order: the part, its
   * content type, the relationship from the workbook and the `<sheet>` entry
   * with the next sheetId — everything Excel needs to find it.
   *
   * @returns {{ name: string, part: string }}
   */
  addSheet(name) {
    const clean = this._checkSheetName(name);
    const n = this.pkg.nextPartNumber('xl/worksheets/', 'sheet');
    const part = 'xl/worksheets/sheet' + n + '.xml';
    this.pkg.addPart(
      part,
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/></worksheet>',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
    );
    const rId = this.pkg.addRelationshipTo(this.mainPart, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', 'worksheets/sheet' + n + '.xml');
    let xml = this.pkg.text(this.mainPart);
    const ids = [...xml.matchAll(/<sheet\b[^>]*\bsheetId="(\d+)"/g)].map((m) => Number(m[1]));
    const sheetId = (ids.length ? Math.max(...ids) : 0) + 1;
    if (!/<\/sheets>/.test(xml)) throw new Error('the workbook lists no sheets');
    xml = xml.replace('</sheets>', '<sheet name="' + esc(clean) + '" sheetId="' + sheetId + '" r:id="' + rId + '"/></sheets>');
    this.pkg.write_(this.mainPart, xml);
    this._sheets = null;
    return { name: clean, part };
  }

  /**
   * Take a sheet out of the workbook: its entry, its relationship and its
   * part go; names scoped to later sheets move down one, names scoped to
   * it go with it, and the active tab is the first sheet again. The last
   * sheet cannot go — a workbook with no sheet in it is not a workbook.
   */
  removeSheet(name) {
    const sheets = this.sheets();
    const index = sheets.findIndex((s) => s.name === name);
    if (index < 0) throw new Error('no such sheet: ' + name);
    if (sheets.length <= 1) throw new Error('a workbook keeps at least one sheet');
    const entry = sheets[index];
    let xml = this.pkg.text(this.mainPart);
    xml = xml.replace(new RegExp('<sheet\\b[^>]*\\br:id="' + entry.rId + '"[^>]*/>'), '');
    xml = xml.replace(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g, (m, attrsText, inner) => {
      const scoped = /\blocalSheetId="(\d+)"/.exec(attrsText);
      if (!scoped) return m;
      const k = Number(scoped[1]);
      if (k === index) return '';
      if (k > index) return '<definedName' + attrsText.replace(/\blocalSheetId="\d+"/, 'localSheetId="' + (k - 1) + '"') + '>' + inner + '</definedName>';
      return m;
    });
    xml = xml.replace(/<definedNames>\s*<\/definedNames>/, '');
    xml = xml.replace(/(<workbookView\b[^>]*?)\s*activeTab="\d+"/, '$1');
    this.pkg.write_(this.mainPart, xml);
    const relsPart = 'xl/_rels/workbook.xml.rels';
    if (this.pkg.has(relsPart)) {
      const rels = this.pkg.text(relsPart);
      this.pkg.write_(relsPart, rels.replace(new RegExp('<Relationship\\b[^>]*\\bId="' + entry.rId + '"[^>]*/>'), ''));
    }
    this.pkg.removePart(entry.part);
    this._loaded.delete(entry.part);
    this._sheets = null;
    return true;
  }

  /**
   * Rename a sheet, and everything that names it: the entry, the defined
   * names that point into it, and every formula in every sheet that reads
   * it — as Excel does when a tab is renamed.
   *
   * @returns {string[]} the parts rewritten besides the workbook part
   */
  renameSheet(from, to) {
    const sheets = this.sheets();
    const entry = sheets.find((s) => s.name === from);
    if (!entry) throw new Error('no such sheet: ' + from);
    const clean = this._checkSheetName(to, { except: from });
    if (clean === from) return [];
    const rx = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const quotedTo = /[^A-Za-z0-9_]/.test(clean) || /^\d/.test(clean) ? "'" + clean.replace(/'/g, "''") + "'" : clean;
    // A reference to the old name: quoted or bare, followed by the bang.
    const refRx = new RegExp("(^|[^A-Za-z0-9_'.])(?:'" + rx(from.replace(/'/g, "''")) + "'|" + rx(from) + ')!', 'g');
    const rewrite = (text) => text.replace(refRx, '$1' + quotedTo + '!');

    let xml = this.pkg.text(this.mainPart);
    xml = xml.replace(new RegExp('(<sheet\\b[^>]*\\bname=")' + rx(esc(from)) + '(")'), '$1' + esc(clean) + '$2');
    xml = xml.replace(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g, (m, attrsText, inner) => '<definedName' + attrsText + '>' + esc(rewrite(unesc(inner))) + '</definedName>');
    this.pkg.write_(this.mainPart, xml);

    const touched = [];
    for (const s of sheets) {
      const loaded = this._loaded.get(s.part);
      const source = loaded ? loaded.render() : this.pkg.text(s.part);
      if (!source.includes(from)) continue;
      const next = source.replace(/<f\b([^>]*)>([^<]*)<\/f>/g, (m, attrsText, f) => '<f' + attrsText + '>' + esc(rewrite(unesc(f))) + '</f>');
      if (next !== source) {
        this.pkg.write_(s.part, next);
        this._loaded.delete(s.part);
        touched.push(s.part);
      }
    }
    this._sheets = null;
    return touched;
  }

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
    return this._cellValue(cell);
  }

  /**
   * Every cell of a sheet that holds something, in row order, decoded —
   * `{ ref, row, col, value }`. One parse per row; a formatted blank (the
   * `<c r="F9" s="3"/>` a tender's answer grid is mostly made of) costs a
   * lookup and nothing else. What the calc model is built from.
   */
  *cells(sheetName) {
    const { part } = this._sheetPart(sheetName);
    // Transient, not through the row cache: a walk of every sheet must not
    // pin every parsed cell of every sheet in memory (900 MB on the tender
    // workbook). And a self-closing `<c …/>` — a formatted blank, which is
    // what 2.4 million of its 2.44 million cells are — is skipped before its
    // attributes are even looked at.
    const re = /<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g;
    for (const row of part.rows) {
      re.lastIndex = 0;
      let m;
      while ((m = re.exec(row.inner))) {
        if (m[2] === '/>') continue;
        const a = attrs(m[1]);
        if (!a.r) continue;
        const cell = { ref: a.r, col: parseRef(a.r).col, style: a.s ?? null, type: a.t ?? null, inner: m[3] };
        const value = this._cellValue(cell);
        if (value === null) continue;
        yield { ref: cell.ref, row: row.index, col: cell.col, value };
      }
    }
  }

  /** A parsed cell's value: formula text, shared or inline string, boolean, error, number. */
  _cellValue(cell) {
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

  /**
   * Add, replace or remove a defined name, leaving the rest of workbook.xml
   * alone.
   *
   * `attrs` carries the ones that change what a name means rather than what
   * it points at — `localSheetId` above all, which is what makes
   * `_xlnm.Print_Area` this sheet's print area rather than the workbook's.
   * A null `ref` removes the name, which is how a print area is cleared.
   */
  setDefinedName(name, ref, extra = {}) {
    let xml = this.pkg.text(this.mainPart);
    const more = Object.entries(extra)
      .filter(([, v]) => v !== null && v !== undefined && v !== '')
      .map(([k, v]) => ` ${k}="${esc(String(v))}"`)
      .join('');
    const entry = ref === null || ref === undefined
      ? ''
      : '<definedName name="' + esc(name) + '"' + more + '>' + esc(ref) + '</definedName>';
    // A sheet-scoped name is not unique by name: every sheet with a print
    // area has one called _xlnm.Print_Area, and matching on the name alone
    // would give the second sheet's the first sheet's range.
    const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const scope = extra.localSheetId === null || extra.localSheetId === undefined
      ? '(?![^>]*localSheetId)'
      : `(?=[^>]*localSheetId="${extra.localSheetId}")`;
    const existing = new RegExp('<definedName\\b(?=[^>]*name="' + escaped + '")' + scope + '[^>]*>[\\s\\S]*?</definedName>');
    if (existing.test(xml)) {
      xml = xml.replace(existing, entry);
    } else if (!entry) {
      return this; // nothing to remove
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
          id: Number(a.id) || null,
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

  /** The sparkline groups Excel wrote on this sheet (or this engine did): `{type, colour, sparklines: [{data, at}]}`. */
  sparklineGroups(sheetName) {
    return parseSparklineExt(this._sheetPart(sheetName).part.tailElement('extLst') || '');
  }

  /**
   * Insert → Sparklines: one group, written exactly as Excel writes one — an
   * `x14:sparklineGroups` extension at the end of the sheet — merged into any
   * `extLst` the sheet already carries, after whatever other extensions are
   * there. `type` is 'line' or 'column'; `data` and `at` are plain ranges on
   * this sheet, shaped as `planSparklines` above describes.
   */
  addSparklines(sheetName, { type, data, at }) {
    if (type !== 'line' && type !== 'column') throw new Error('a sparkline is "line" or "column"');
    const groupXml = sparklineGroupXml(type, sheetName, planSparklines(data, at));
    const { part } = this._sheetPart(sheetName);
    const extLst = part.tailElement('extLst');
    if (extLst && /<x14:sparklineGroups\b/.test(extLst)) {
      part.setTailElement('extLst', extLst.replace('</x14:sparklineGroups>', groupXml + '</x14:sparklineGroups>'));
    } else if (extLst) {
      part.setTailElement('extLst', extLst.replace('</extLst>', sparklineExtXml(groupXml) + '</extLst>'));
    } else {
      part.setTailElement('extLst', '<extLst>' + sparklineExtXml(groupXml) + '</extLst>');
    }
    return this;
  }

  /** Take the sparkline(s) off a cell or range. Returns how many went. */
  removeSparklines(sheetName, at) {
    const { part } = this._sheetPart(sheetName);
    const extLst = part.tailElement('extLst');
    if (!extLst) return 0;
    const { removed, xml } = removeSparklinesFromExt(extLst, at);
    if (removed) part.setTailElement('extLst', /<ext\b/.test(xml) ? xml : null);
    return removed;
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
    const xml = this.pkg.text(this.mainPart);
    if (/fullCalcOnLoad="1"/.test(xml)) return this;
    return this._setCalcPr({ fullCalcOnLoad: '1' });
  }

  /**
   * `<calcPr>` with some attributes set (a value) or taken off (null), the
   * element made where the schema puts it — after the defined names, before
   * `oleSize`, the custom views, the pivot caches and `extLst` — when the
   * workbook has none.
   */
  _setCalcPr(changes) {
    let xml = this.pkg.text(this.mainPart);
    const found = /<calcPr\b([^>]*?)(\/>|>)/.exec(xml);
    let attrsText = found ? found[1] : ' calcId="191029"';
    for (const [name, value] of Object.entries(changes)) {
      attrsText = attrsText.replace(new RegExp('\\s*\\b' + name + '="[^"]*"'), '');
      if (value !== null && value !== undefined) attrsText += ' ' + name + '="' + esc(String(value)) + '"';
    }
    if (found) {
      xml = xml.slice(0, found.index) + '<calcPr' + attrsText + found[2] + xml.slice(found.index + found[0].length);
    } else {
      const next = /<(oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b|<\/workbook>/.exec(xml);
      const at = next ? next.index : xml.length;
      xml = xml.slice(0, at) + '<calcPr' + attrsText + '/>' + xml.slice(at);
    }
    this.pkg.write_(this.mainPart, xml);
    return this;
  }

  /**
   * Formulas → Calculation Options, as the workbook keeps it: `calcMode` on
   * `<calcPr>` — 'manual', 'autoNoTable' (automatic except for data tables)
   * or, when the attribute is absent, automatic.
   */
  calcMode() {
    const m = /<calcPr\b([^>]*?)\/?>/.exec(this.pkg.text(this.mainPart));
    const mode = m ? /\bcalcMode="([^"]*)"/.exec(m[1])?.[1] : null;
    return mode === 'manual' || mode === 'autoNoTable' ? mode : 'auto';
  }

  /** Set the calculation mode; automatic takes the attribute away, as Excel does. */
  setCalcMode(mode) {
    if (!['auto', 'autoNoTable', 'manual'].includes(mode)) throw new Error('calculation is auto, autoNoTable or manual, not ' + mode);
    if (mode === this.calcMode()) return this;
    return this._setCalcPr({ calcMode: mode === 'auto' ? null : mode });
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

  /** A split window on a sheet (see SheetPart#splitPane), or null. */
  splitPane(sheetName) {
    return this._sheetPart(sheetName).part.splitPane();
  }

  setSplitPane(sheetName, spec) {
    this._sheetPart(sheetName).part.setSplitPane(spec);
    return this;
  }

  /**
   * How a sheet is shown — View → Normal, Page Break Preview, Page Layout —
   * from `<sheetView view>`: 'normal' (the attribute's absence),
   * 'pageBreakPreview' or 'pageLayout'.
   */
  sheetViewMode(sheetName) {
    const v = this._sheetPart(sheetName).part.sheetViewAttrs().view;
    return v === 'pageBreakPreview' || v === 'pageLayout' ? v : 'normal';
  }

  setSheetViewMode(sheetName, mode) {
    const value = mode === 'pageBreakPreview' || mode === 'pageLayout' ? mode : null;
    this._sheetPart(sheetName).part.setSheetViewAttr('view', value);
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
  /** A sheet's hyperlinks with their targets: an address from the rels, or a place in the workbook. */
  hyperlinks(sheetName) {
    const { part } = this._sheetPart(sheetName);
    const rels = new Map(this.pkg.rels(this.partNameFor(sheetName)).map((r) => [r.Id, r]));
    return part.hyperlinks().map((h) => {
      const rel = h.rId ? rels.get(h.rId) : null;
      return { ref: h.ref, href: rel ? rel.Target : null, location: h.location, display: h.display, tooltip: h.tooltip };
    });
  }

  /**
   * Put a hyperlink on a cell. An address (http, https, mailto) is a
   * relationship marked external, as Excel writes one; a place in the
   * workbook — C12, Sheet2!B4, a name — is a location and needs none.
   */
  setHyperlink(sheetName, ref, { href = null, location = null, tooltip = null, display = null } = {}) {
    if (!href && !location) throw new Error('a hyperlink needs an address or a place in the workbook');
    const { part } = this._sheetPart(sheetName);
    const rId = href ? this.pkg.addRelationshipTo(this.partNameFor(sheetName), REL_HYPERLINK, href, { external: true }) : null;
    part.setHyperlink({ ref, rId, location: href ? null : location, display, tooltip });
    return this;
  }

  removeHyperlink(sheetName, ref) {
    const { part } = this._sheetPart(sheetName);
    return part.removeHyperlink(ref);
  }

  /** The notes on a sheet — `[{ ref, author, text }]` — from its comments part. */
  comments(sheetName) {
    const found = this._sheetRelTarget(sheetName, '/comments');
    return found ? parseComments(this.pkg.text(found.name)) : [];
  }

  /** The part a sheet's relationship of one type points at, when it exists. */
  _sheetRelTarget(sheetName, typeSuffix) {
    const sheetPartName = this.partNameFor(sheetName);
    const rel = this.pkg.rels(sheetPartName).find((r) => String(r.Type).endsWith(typeSuffix));
    if (!rel) return null;
    const name = OoxmlPackage.resolveTarget(sheetPartName, rel.Target);
    return this.pkg.has(name) ? { name, rel } : null;
  }

  /**
   * A note on a cell, as Excel keeps one: the comments part carries the
   * words and who wrote them, and a VML drawing part carries the box Excel
   * draws them in — without the box, Excel calls the file damaged. Both are
   * made where the sheet has none; a note on a cell that has one replaces
   * it. A box Excel drew for another note, or any other legacy shape in the
   * drawing, is left as it was.
   */
  setComment(sheetName, ref, { author = '', text = '' } = {}) {
    const sheetPartName = this.partNameFor(sheetName);
    const { part } = this._sheetPart(sheetName);
    let comments = this._sheetRelTarget(sheetName, '/comments');
    if (!comments) {
      const n = this.pkg.nextPartNumber('xl/', 'comments');
      const name = 'xl/comments' + n + '.xml';
      this.pkg.addPart(name, commentsXml([]), CT_COMMENTS);
      this.pkg.addRelationshipTo(sheetPartName, REL_COMMENTS, '../comments' + n + '.xml');
      comments = { name };
    }
    const list = parseComments(this.pkg.text(comments.name)).filter((c) => c.ref !== ref);
    list.push({ ref, author: author || '', text: text || '' });
    this.pkg.write_(comments.name, commentsXml(list));

    let vml = this._sheetRelTarget(sheetName, '/vmlDrawing');
    if (!vml) {
      let n = 1;
      while (this.pkg.has('xl/drawings/vmlDrawing' + n + '.vml')) n += 1;
      const name = 'xl/drawings/vmlDrawing' + n + '.vml';
      this.pkg.ensureDefault('vml', CT_VML);
      this.pkg.addPart(name, VML_HEAD + '</xml>');
      const rId = this.pkg.addRelationshipTo(sheetPartName, REL_VML, '../drawings/vmlDrawing' + n + '.vml');
      part.setLegacyDrawing(rId);
      vml = { name };
    }
    this.pkg.write_(vml.name, withNoteBox(this.pkg.text(vml.name), parseRef(ref)));
    return this;
  }

  /**
   * A table (ListObject) over a range, written as Excel keeps one: the table
   * part with its columns, its autofilter and the style it asked for by
   * name, the sheet's rels pointing at it and `<tableParts>` naming it. The
   * column names are the header row's words; one that is empty or repeats
   * another is made unique the way Excel does (Column1, Qty2). Returns what
   * was written, columns included, so the caller can put a made-up name
   * into the header cell it stands for.
   */
  addTable(sheetName, ref, { name = null, style = 'TableStyleMedium2', stripes = true, headerNames = [] } = {}) {
    const sheetPartName = this.partNameFor(sheetName);
    const { part } = this._sheetPart(sheetName);
    const existing = this.tables();
    const ids = existing.map((t) => t.id).filter((n) => Number.isFinite(n));
    const id = (ids.length ? Math.max(...ids) : 0) + 1;
    const taken = new Set(existing.map((t) => String(t.name || '').toLowerCase()));
    let displayName = String(name || '').replace(/[^A-Za-z0-9_]/g, '_').replace(/^(\d)/, '_$1');
    if (!displayName || taken.has(displayName.toLowerCase())) {
      let k = id;
      do { displayName = 'Table' + k; k += 1; } while (taken.has(displayName.toLowerCase()));
    }
    const seen = new Set();
    const columns = headerNames.map((h, i) => {
      const base = String(h ?? '').trim() || 'Column' + (i + 1);
      let col = base;
      let k = 2;
      while (seen.has(col.toLowerCase())) { col = base + k; k += 1; }
      seen.add(col.toLowerCase());
      return col;
    });
    const n = this.pkg.nextPartNumber('xl/tables/', 'table');
    const partName = 'xl/tables/table' + n + '.xml';
    const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<table xmlns="' + XMLNS_MAIN + '" id="' + id + '" name="' + esc(displayName) + '" displayName="' + esc(displayName)
      + '" ref="' + esc(ref) + '" totalsRowShown="0">'
      + '<autoFilter ref="' + esc(ref) + '"/>'
      + '<tableColumns count="' + columns.length + '">'
      + columns.map((c, i) => '<tableColumn id="' + (i + 1) + '" name="' + esc(c) + '"/>').join('')
      + '</tableColumns>'
      + '<tableStyleInfo name="' + esc(style) + '" showFirstColumn="0" showLastColumn="0" showRowStripes="' + (stripes ? '1' : '0') + '" showColumnStripes="0"/>'
      + '</table>';
    this.pkg.addPart(partName, xml, CT_TABLE);
    const rId = this.pkg.addRelationshipTo(sheetPartName, REL_TABLE, '../tables/table' + n + '.xml');
    part.addTablePart(rId);
    return { part: partName, id, name: displayName, columns };
  }

  /**
   * Take the note off a cell. The parts stay, emptied: an undo puts the
   * note back into them, and a comments part with nothing in it is one
   * Excel reads without a word. Returns whether there was one.
   */
  removeComment(sheetName, ref) {
    const comments = this._sheetRelTarget(sheetName, '/comments');
    if (!comments) return false;
    const list = parseComments(this.pkg.text(comments.name));
    if (!list.some((c) => c.ref === ref)) return false;
    this.pkg.write_(comments.name, commentsXml(list.filter((c) => c.ref !== ref)));
    const vml = this._sheetRelTarget(sheetName, '/vmlDrawing');
    if (vml) this.pkg.write_(vml.name, withoutNoteBox(this.pkg.text(vml.name), parseRef(ref)));
    return true;
  }

  /* ── threaded comments ─────────────────────────────────────────────────── */

  /** The workbook's person list part, when it has one. */
  _personsPart() {
    const rel = this.pkg.rels(this.mainPart).find((r) => r.Type === REL_PERSON);
    if (!rel) return null;
    const name = OoxmlPackage.resolveTarget(this.mainPart, rel.Target);
    return this.pkg.has(name) ? name : null;
  }

  /** The people threaded comments name as authors. */
  persons() {
    const part = this._personsPart();
    return part ? parsePersons(this.pkg.text(part)) : [];
  }

  /**
   * The person a comment is signed by, added to the list (made where the
   * workbook has none) the first time they write. Matched by name.
   */
  _ensurePerson(displayName) {
    const name = String(displayName || '').trim() || 'Author';
    let part = this._personsPart();
    const list = part ? parsePersons(this.pkg.text(part)) : [];
    const known = list.find((p) => p.displayName === name);
    if (known) return known.id;
    const person = { id: newGuid(), displayName: name, userId: name, providerId: 'None' };
    list.push(person);
    if (part) {
      this.pkg.write_(part, personsXml(list));
    } else {
      part = 'xl/persons/person.xml';
      this.pkg.addPart(part, personsXml(list), CT_PERSON);
      this.pkg.addRelationshipTo(this.mainPart, REL_PERSON, 'persons/person.xml');
    }
    return person.id;
  }

  /** The threaded-comments part of a sheet, made (and related) when `create` asks. */
  _threadPart(sheetName, create = false) {
    const found = this._sheetRelTarget(sheetName, '/threadedComment');
    if (found || !create) return found?.name ?? null;
    const n = this.pkg.nextPartNumber('xl/threadedComments/', 'threadedComment');
    const name = 'xl/threadedComments/threadedComment' + n + '.xml';
    this.pkg.addPart(name, threadedCommentsXml([]), CT_THREADED);
    this.pkg.addRelationshipTo(this.partNameFor(sheetName), REL_THREADED, '../threadedComments/threadedComment' + n + '.xml');
    return name;
  }

  _threadList(sheetName) {
    const part = this._threadPart(sheetName);
    return part ? parseThreadedComments(this.pkg.text(part)) : [];
  }

  /**
   * A sheet's threads, in the order the part keeps them: `[{ ref, id, done,
   * comments: [{ id, author, personId, date, text }] }]`, the first comment
   * the one that started it and the rest its replies.
   */
  threads(sheetName) {
    const list = this._threadList(sheetName);
    if (!list.length) return [];
    const people = new Map(this.persons().map((p) => [p.id, p.displayName]));
    const byId = new Map();
    const threads = [];
    for (const c of list) {
      const comment = { id: c.id, author: people.get(c.personId) ?? '', personId: c.personId, date: c.date, text: c.text };
      if (!c.parentId) {
        const thread = { ref: c.ref, id: c.id, done: c.done, comments: [comment] };
        byId.set(c.id, thread);
        threads.push(thread);
      } else {
        byId.get(c.parentId)?.comments.push(comment);
      }
    }
    return threads;
  }

  /** The thread on a cell, or null. */
  threadAt(sheetName, ref) {
    return this.threads(sheetName).find((t) => t.ref === ref) ?? null;
  }

  /**
   * Write a sheet's comment list back, and bring the classic shadow of each
   * cell named in `refs` in line with its thread — written while it has
   * one, taken off once it has none.
   */
  _writeThreads(sheetName, list, refs) {
    const part = this._threadPart(sheetName, true);
    this.pkg.write_(part, threadedCommentsXml(list));
    for (const ref of refs) {
      const thread = this.threadAt(sheetName, ref);
      if (thread) this.setComment(sheetName, ref, { author: 'tc=' + thread.id, text: threadShadowText(thread) });
      else this.removeComment(sheetName, ref);
    }
  }

  /**
   * A comment on a cell: the start of a thread, or — when the cell has one
   * — a reply at its end. A cell with a note is refused, as Excel keeps a
   * cell to one or the other. Returns the new comment's id.
   */
  addThreadedComment(sheetName, ref, { author = '', text = '', date = new Date() } = {}) {
    const words = String(text ?? '').trim();
    if (!words) throw new Error('a comment needs some words');
    const note = this.comments(sheetName).find((c) => c.ref === ref && !String(c.author).startsWith('tc='));
    if (note) throw new Error('This cell has a note. Delete the note, or edit it, before starting a comment here.');
    const list = this._threadList(sheetName);
    const top = list.find((c) => c.ref === ref && !c.parentId);
    const comment = {
      id: newGuid(), ref, date: threadDate(date), personId: this._ensurePerson(author),
      parentId: top ? top.id : null, done: false, text: words, extra: '',
    };
    // A reply goes after its thread's last comment, where Excel puts it.
    let at = list.length;
    if (top) {
      at = list.indexOf(top) + 1;
      while (at < list.length && list[at].parentId === top.id) at += 1;
    }
    list.splice(at, 0, comment);
    this._writeThreads(sheetName, list, [ref]);
    return comment.id;
  }

  /** A comment's words changed. */
  editThreadedComment(sheetName, id, text) {
    const words = String(text ?? '').trim();
    if (!words) throw new Error('a comment needs some words');
    const list = this._threadList(sheetName);
    const c = list.find((x) => x.id === id);
    if (!c) throw new Error('no such comment');
    c.text = words;
    this._writeThreads(sheetName, list, [c.ref]);
    return this;
  }

  /** A reply taken out, or — the thread's first comment — the whole thread. */
  deleteThreadedComment(sheetName, id) {
    const list = this._threadList(sheetName);
    const c = list.find((x) => x.id === id);
    if (!c) return false;
    const kept = c.parentId ? list.filter((x) => x.id !== id) : list.filter((x) => x.id !== id && x.parentId !== id);
    this._writeThreads(sheetName, kept, [c.ref]);
    return true;
  }

  /** Resolve a cell's thread (done) or reopen it. */
  setThreadDone(sheetName, ref, done) {
    const list = this._threadList(sheetName);
    const top = list.find((c) => c.ref === ref && !c.parentId);
    if (!top) throw new Error('no comment on ' + ref);
    top.done = Boolean(done);
    this._writeThreads(sheetName, list, [ref]);
    return this;
  }

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

  /**
   * Insert one empty row before each of several rows at once — what
   * Subtotals does, a row after every group. `points` are row indices as
   * the sheet stands before any of them is inserted. The rows are
   * renumbered in one pass (one insert per point was a pass per group over
   * every row of the sheet); merges, the dimension, formulas and names are
   * adjusted a point at a time from the bottom up, so each adjustment sees
   * the sheet exactly as a single insert would.
   */
  insertRowsAt(sheetName, points, { materialize = false } = {}) {
    // A point may repeat: two rows land before the same one (a Grand Total
    // and the first group's subtotal, when summaries sit above).
    const at = points.map(Number).filter((n) => Number.isInteger(n) && n >= 0).sort((a, b) => a - b);
    if (!at.length) return this;
    const { part } = this._sheetPart(sheetName);
    // How many points lie at or above a row: that is how far it moves.
    const shift = (row) => {
      let lo = 0;
      let hi = at.length;
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        if (at[mid] <= row) lo = mid + 1;
        else hi = mid;
      }
      return lo;
    };
    for (let i = part.rows.length - 1; i >= 0; i--) {
      const row = part.rows[i];
      const by = shift(row.index);
      if (by) part._renumberRow(row, row.index + by);
    }
    // The new rows as records now, all in one sort, when the caller is about
    // to fill them: made one by one as their first cell is written, each
    // dropped the row index and the next write rebuilt it over every row.
    if (materialize) {
      at.forEach((p, i) => part.rows.push({ index: p + i, attrsStr: ' r="' + (p + i + 1) + '"', inner: '', xml: '', dirty: true }));
      part.rows.sort((a, b) => a.index - b.index);
    }
    part._rowsChanged();
    part.dirty = true;
    // From the bottom up, so each point is where it was before any insert.
    this._adjustForEdits(sheetName, 'row', 'insert', at.map((p) => ({ at: p, count: 1 })).reverse(), { rows: false });
    return this;
  }

  /** Delete several rows at once, the mirror of `insertRowsAt`: Remove All takes every subtotal row. */
  deleteRowsAt(sheetName, rows) {
    const at = [...new Set(rows.map(Number))].filter((n) => Number.isInteger(n) && n >= 0).sort((a, b) => a - b);
    if (!at.length) return this;
    const { part } = this._sheetPart(sheetName);
    const gone = new Set(at);
    const shift = (row) => {
      let lo = 0;
      let hi = at.length;
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        if (at[mid] < row) lo = mid + 1;
        else hi = mid;
      }
      return lo;
    };
    part.rows = part.rows.filter((row) => !gone.has(row.index));
    for (const row of part.rows) {
      const by = shift(row.index);
      if (by) part._renumberRow(row, row.index - by);
    }
    part._rowsChanged();
    part.dirty = true;
    this._adjustForEdits(sheetName, 'row', 'delete', at.map((p) => ({ at: p, count: 1 })).reverse(), { rows: false });
    return this;
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
    return this._adjustForEdit(sheetName, axis, op, at, count);
  }

  /**
   * Everything one structural edit moves: merges, the rows or cells and
   * `<col>` records (unless the caller has already renumbered the rows),
   * the dimension, every formula in the workbook and the defined names.
   */
  _adjustForEdit(sheetName, axis, op, at, count, { rows = true } = {}) {
    return this._adjustForEdits(sheetName, axis, op, [{ at, count }], { rows });
  }

  /**
   * The same for several edits of one kind at once, given as they apply one
   * after another — each `at` in the sheet as the edits before it left it.
   * Formulas and names are walked ONCE, each run through every edit in
   * turn: a pass over every row per edit was 600 passes over sixty thousand
   * rows for one Subtotal.
   */
  _adjustForEdits(sheetName, axis, op, edits, { rows = true } = {}) {
    const { part } = this._sheetPart(sheetName);

    for (const { at, count } of edits) {
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
        if (rows && op === 'insert') part.insertRowsShift(at, count);
        else if (rows) part.deleteRowsRange(at, count);
      } else {
        part.shiftCells(at, count, op);
        part.shiftColEntries(at, count, op);
      }
      part.shiftDimension(axis, op, at, count);
    }
    part.dirty = true;

    const through = (text, currentSheet) => {
      let t = text;
      for (const { at, count } of edits) {
        t = adjustFormula(t, { editedSheet: sheetName, currentSheet, axis, op, at, count });
      }
      return t;
    };

    // Formulas everywhere: the edited sheet's own, and every other sheet's
    // qualified references into it.
    for (const { name } of this.sheets()) {
      const { part: p } = this._sheetPart(name);
      p.adjustFormulas((f) => through(f, name));
    }

    // Defined names live in workbook.xml and are always sheet-qualified, so
    // they go through the same adjuster with no current sheet.
    this._adjustDefinedNames((t) => through(t, null));
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
