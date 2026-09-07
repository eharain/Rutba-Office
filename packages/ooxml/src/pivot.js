/**
 * Pivot tables — read, compute, refresh.
 *
 * The last thing in the capability register that was preserved and wholly
 * unmodelled. Worth being precise about what a pivot actually IS in a file,
 * because it decides what "support" can honestly mean:
 *
 *   - The pivot's numbers are MATERIALISED into ordinary sheet cells. A
 *     reader that never heard of pivots still shows the right figures, which
 *     is why a customer's pivot has always looked correct here.
 *   - Beside them sits a `pivotTableDefinition` (which field is on which
 *     axis, how each value is aggregated, where the rectangle lives) and a
 *     `pivotCacheDefinition` + `pivotCacheRecords` pair — a SNAPSHOT of the
 *     source data taken when the pivot was last refreshed.
 *
 * So the gap was never rendering. It was that editing the source changed
 * nothing: the materialised cells kept last week's totals, silently. This
 * module closes that — it reads the definition, recomputes the grid from the
 * LIVE source values, and rewrites the rectangle.
 *
 * **The cache is deliberately not rebuilt.** Rewriting `pivotCacheRecords`
 * means renumbering shared items and every `<item x="n"/>` that indexes them,
 * across parts, for no gain a person can see. Instead a refresh sets
 * `refreshOnLoad="1"` on the cache definition — the same mechanism the
 * formula layer uses with `fullCalcOnLoad` — so Excel rebuilds its own cache
 * from the source the moment it opens the file. Our numbers and Excel's agree
 * because both are computed from the same source range, not because we
 * reverse-engineered its cache format.
 *
 * Sourcing note, as everywhere in this package: this is written from
 * ECMA-376 / ISO/IEC 29500, not from reading another implementation.
 */
import { compareValues, isError } from '@rutba/formula';
import { OoxmlPackage } from './package.js';
import { parseRef, makeRef } from './workbook.js';

const attrsOf = (tag) => {
  const out = {};
  const re = /([\w:.-]+)\s*=\s*"([^"]*)"/g;
  let m;
  while ((m = re.exec(tag ?? ''))) out[m[1]] = m[2];
  return out;
};

const unesc = (s) => String(s)
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
  .replace(/&apos;/g, "'").replace(/&amp;/g, '&');

const esc = (s) => String(s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;');

/** A flag attribute: absent means the schema default, which is usually on. */
const flagOn = (v, dflt = true) => (v === undefined ? dflt : v === '1' || v === 'true');

/** `Data!$A$1:$C$9` or `A1:C9` -> {sheet, top, left, bottom, right}. */
function parseArea(ref, fallbackSheet = null) {
  const text = String(ref ?? '').replace(/\$/g, '');
  const bang = text.lastIndexOf('!');
  let sheet = fallbackSheet;
  let body = text;
  if (bang >= 0) {
    sheet = text.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'");
    body = text.slice(bang + 1);
  }
  const [a, b] = body.split(':');
  if (!/^[A-Z]+\d+$/.test(a ?? '')) return null;
  const from = parseRef(a);
  const to = b && /^[A-Z]+\d+$/.test(b) ? parseRef(b) : from;
  return {
    sheet,
    top: Math.min(from.row, to.row),
    left: Math.min(from.col, to.col),
    bottom: Math.max(from.row, to.row),
    right: Math.max(from.col, to.col),
  };
}

const areaRef = (a) => makeRef(a.top, a.left) + ':' + makeRef(a.bottom, a.right);

/** The aggregate a data field asks for. `sum` is the schema's default. */
const SUBTOTALS = new Set([
  'sum', 'count', 'countNums', 'average', 'max', 'min', 'product',
  'stdDev', 'stdDevp', 'var', 'varp',
]);

/**
 * Every pivot table in a workbook, with its cache and its layout, read from
 * the parts each sheet's relationships point at.
 *
 * A pivot this module cannot honestly recompute comes back with `unsupported`
 * set to the reason rather than being omitted — the file HAS one, and saying
 * "there is a pivot here and this is why its Refresh button is missing" is
 * the honest report. Silence would read as "no pivot".
 */
export function readPivots(wb) {
  const out = [];
  for (const { name: sheetName, part } of wb.sheets()) {
    for (const rel of wb.pkg.rels(part)) {
      if (!String(rel.Type).endsWith('/pivotTable')) continue;
      const tablePart = OoxmlPackage.resolveTarget(part, rel.Target);
      if (!wb.pkg.has(tablePart)) continue;
      const pivot = readPivotPart(wb, sheetName, tablePart);
      if (pivot) out.push(pivot);
    }
  }
  return out;
}

function readPivotPart(wb, sheetName, tablePart) {
  const xml = wb.pkg.text(tablePart);
  const head = attrsOf((/<pivotTableDefinition\b([^>]*?)>/.exec(xml) ?? [])[1] ?? '');
  const location = attrsOf((/<location\b([^>]*?)\/?>/.exec(xml) ?? [])[1] ?? '');
  const area = parseArea(location.ref, sheetName);
  if (!area) return null;

  // The cache definition sits behind this part's own relationship.
  const cacheRel = wb.pkg.rels(tablePart).find((r) => String(r.Type).endsWith('/pivotCacheDefinition'));
  const cachePart = cacheRel ? OoxmlPackage.resolveTarget(tablePart, cacheRel.Target) : null;

  const pivot = {
    name: unesc(head.name ?? head.displayName ?? 'PivotTable'),
    sheet: sheetName,
    part: tablePart,
    cachePart,
    location: {
      ...area,
      firstHeaderRow: Number(location.firstHeaderRow ?? 1),
      firstDataRow: Number(location.firstDataRow ?? 1),
      firstDataCol: Number(location.firstDataCol ?? 1),
    },
    // Layout flags. Modern Excel writes compact form by default; a file that
    // says otherwise gets its own look back, which is the whole point of
    // reading them rather than picking one.
    compact: flagOn(head.compact, true),
    outline: flagOn(head.outline, true),
    dataOnRows: flagOn(head.dataOnRows, false),
    rowGrandTotals: flagOn(head.rowGrandTotals, true),
    colGrandTotals: flagOn(head.colGrandTotals, true),
    source: null,
    cacheFields: [],
    rowFields: [],
    colFields: [],
    pageFields: [],
    dataFields: [],
    unsupported: null,
  };

  if (!cachePart || !wb.pkg.has(cachePart)) {
    pivot.unsupported = 'its cache definition is missing from the file';
    return pivot;
  }

  const cacheXml = wb.pkg.text(cachePart);
  const sourceTag = /<worksheetSource\b([^>]*?)\/?>/.exec(cacheXml);
  const cacheSource = attrsOf((/<cacheSource\b([^>]*?)\/?>/.exec(cacheXml) ?? [])[1] ?? '');
  if (!sourceTag || (cacheSource.type && cacheSource.type !== 'worksheet')) {
    pivot.unsupported = 'its data comes from outside this workbook';
    return pivot;
  }
  const sa = attrsOf(sourceTag[1]);
  // A source given as a defined name is resolved by the caller, which has the
  // name table; recorded here so the reason is specific.
  if (!sa.ref && sa.name) {
    pivot.sourceName = unesc(sa.name);
    pivot.unsupported = 'its source is the defined name "' + unesc(sa.name) + '"';
    return pivot;
  }
  pivot.source = parseArea(sa.ref, sa.sheet ? unesc(sa.sheet) : sheetName);
  if (pivot.source && sa.sheet) pivot.source.sheet = unesc(sa.sheet);
  if (!pivot.source) {
    pivot.unsupported = 'its source range could not be read';
    return pivot;
  }

  const fieldsBlock = /<cacheFields\b[^>]*>([\s\S]*?)<\/cacheFields>/.exec(cacheXml);
  if (fieldsBlock) {
    for (const m of fieldsBlock[1].matchAll(/<cacheField\b([^>]*?)(?:\/>|>([\s\S]*?)<\/cacheField>)/g)) {
      const a = attrsOf(m[1]);
      const body = m[2] ?? '';
      pivot.cacheFields.push({ name: unesc(a.name ?? ''), grouped: /<fieldGroup\b/.test(body) });
    }
  }
  if (pivot.cacheFields.some((f) => f.grouped)) {
    pivot.unsupported = 'one of its fields is grouped (dates into months, or numbers into bands)';
    return pivot;
  }

  // pivotFields are positional: the Nth is the Nth cache field.
  const pivotFieldsBlock = /<pivotFields\b[^>]*>([\s\S]*?)<\/pivotFields>/.exec(xml);
  const pivotFields = [];
  if (pivotFieldsBlock) {
    for (const m of pivotFieldsBlock[1].matchAll(/<pivotField\b([^>]*?)(?:\/>|>([\s\S]*?)<\/pivotField>)/g)) {
      pivotFields.push({ attrs: attrsOf(m[1]), body: m[2] ?? '' });
    }
  }
  if (pivotFields.some((f) => f.attrs.formula !== undefined)) {
    pivot.unsupported = 'it has a calculated field';
    return pivot;
  }

  const indices = (block, tag) => {
    const b = new RegExp('<' + block + '\\b[^>]*>([\\s\\S]*?)</' + block + '>').exec(xml);
    if (!b) return [];
    return [...b[1].matchAll(new RegExp('<' + tag + '\\b([^>]*?)\\/?>', 'g'))]
      .map((m) => Number(attrsOf(m[1]).x ?? attrsOf(m[1]).fld));
  };

  pivot.rowFields = indices('rowFields', 'field').filter((n) => Number.isInteger(n));
  pivot.colFields = indices('colFields', 'field').filter((n) => Number.isInteger(n));

  const pageBlock = /<pageFields\b[^>]*>([\s\S]*?)<\/pageFields>/.exec(xml);
  if (pageBlock) {
    for (const m of pageBlock[1].matchAll(/<pageField\b([^>]*?)\/?>/g)) {
      const a = attrsOf(m[1]);
      pivot.pageFields.push({
        fld: Number(a.fld),
        item: a.item === undefined ? null : Number(a.item),
      });
    }
  }

  const dataBlock = /<dataFields\b[^>]*>([\s\S]*?)<\/dataFields>/.exec(xml);
  if (dataBlock) {
    for (const m of dataBlock[1].matchAll(/<dataField\b([^>]*?)\/?>/g)) {
      const a = attrsOf(m[1]);
      const subtotal = a.subtotal ?? 'sum';
      if (!SUBTOTALS.has(subtotal)) {
        pivot.unsupported = 'it summarises with "' + subtotal + '", which this editor cannot compute';
        return pivot;
      }
      pivot.dataFields.push({
        fld: Number(a.fld),
        subtotal,
        name: a.name !== undefined ? unesc(a.name) : null,
      });
    }
  }

  // The `-2` field index is Excel's marker for the Values axis — where the
  // data-field names themselves sit when there is more than one. We place
  // them by rule instead, so the marker is dropped from the axis lists.
  pivot.rowFields = pivot.rowFields.filter((n) => n >= 0);
  pivot.colFields = pivot.colFields.filter((n) => n >= 0);

  if (!pivot.dataFields.length) {
    pivot.unsupported = 'it has no value field to summarise';
  }
  return pivot;
}

/** The label a data field wears: its own name, or "Sum of Qty". */
export function dataFieldLabel(pivot, df) {
  if (df.name) return df.name;
  const field = pivot.cacheFields[df.fld]?.name ?? 'field';
  const verb = {
    sum: 'Sum', count: 'Count', countNums: 'Count', average: 'Average',
    max: 'Max', min: 'Min', product: 'Product',
    stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'Varp',
  }[df.subtotal] ?? 'Sum';
  return verb + ' of ' + field;
}

/** Aggregate a list of raw cell values the way a data field asks. */
function aggregate(values, how) {
  const nums = [];
  let nonEmpty = 0;
  for (const v of values) {
    if (isError(v)) return v;
    if (v === '' || v === null || v === undefined) continue;
    nonEmpty += 1;
    if (typeof v === 'number') nums.push(v);
    else if (typeof v === 'boolean') nums.push(v ? 1 : 0);
  }
  if (how === 'count') return nonEmpty;
  if (how === 'countNums') return nums.length;
  if (!nums.length) return how === 'sum' || how === 'product' ? 0 : '';

  const sum = nums.reduce((s, n) => s + n, 0);
  switch (how) {
    case 'sum': return sum;
    case 'average': return sum / nums.length;
    case 'max': return Math.max(...nums);
    case 'min': return Math.min(...nums);
    case 'product': return nums.reduce((p, n) => p * n, 1);
    default: break;
  }
  // The spread family. Sample forms need two points; a single one has no
  // spread to report, which is #DIV/0! in a spreadsheet, not zero.
  const mean = sum / nums.length;
  const ss = nums.reduce((s, n) => s + (n - mean) ** 2, 0);
  const sample = how === 'stdDev' || how === 'var';
  if (sample && nums.length < 2) return '';
  const variance = ss / (sample ? nums.length - 1 : nums.length);
  return how === 'var' || how === 'varp' ? variance : Math.sqrt(variance);
}

/** Pivot item order: spreadsheet collation, blanks last. */
function compareItems(a, b) {
  const aBlank = a === '' || a === null || a === undefined;
  const bBlank = b === '' || b === null || b === undefined;
  if (aBlank || bBlank) return aBlank && bBlank ? 0 : (aBlank ? 1 : -1);
  const c = compareValues(a, b);
  return isError(c) ? 0 : c;
}

// Joined with the UNIT SEPARATOR control character, which no cell value
// contains — a printable separator is a collision: with a plain join,
// ['at:b'] and ['a','b'] key identically and land in one bucket.
const keyOf = (values) => values
  .map((v) => (typeof v === 'string' ? 't:' + v.toLowerCase() : typeof v + ':' + String(v)))
  .join('\u001f');

/**
 * Compute a pivot's grid from LIVE source values.
 *
 * `readCell(sheet, row, col)` is the whole interface to the data, so the
 * caller decides what "live" means — the calculation model in an open editor,
 * the stored values in a file being processed.
 *
 * Returns the rectangle as a list of `{row, col, value}` in ABSOLUTE sheet
 * coordinates, plus its size, so a caller can write it and compare shapes.
 */
export function computePivot(pivot, readCell) {
  if (pivot.unsupported) throw new Error(pivot.unsupported);
  const src = pivot.source;
  const headerRow = src.top;

  // Rows of the source, as arrays indexed by cache field.
  const rows = [];
  for (let r = headerRow + 1; r <= src.bottom; r++) {
    const line = [];
    for (let c = src.left; c <= src.right; c++) line.push(readCell(src.sheet, r, c));
    // A wholly empty line is the end of the data as a person reads it.
    if (line.every((v) => v === '' || v === null || v === undefined)) continue;
    rows.push(line);
  }

  // Page fields filter before anything is grouped. An item index refers to
  // the field's own distinct values in source order — which is how a cache
  // numbers them — so that ordering is rebuilt here rather than guessed.
  let visible = rows;
  for (const pf of pivot.pageFields) {
    if (pf.item === null || !Number.isInteger(pf.fld)) continue;
    const seen = [];
    for (const line of rows) {
      const v = line[pf.fld];
      if (!seen.some((x) => keyOf([x]) === keyOf([v]))) seen.push(v);
    }
    const chosen = seen[pf.item];
    if (chosen === undefined) continue;
    visible = visible.filter((line) => keyOf([line[pf.fld]]) === keyOf([chosen]));
  }

  const tupleOf = (line, fields) => fields.map((f) => line[f] ?? '');

  const axisKeys = (fields) => {
    if (!fields.length) return [[]];
    const map = new Map();
    for (const line of visible) {
      const t = tupleOf(line, fields);
      const k = keyOf(t);
      if (!map.has(k)) map.set(k, t);
    }
    return [...map.values()].sort((a, b) => {
      for (let i = 0; i < a.length; i++) {
        const c = compareItems(a[i], b[i]);
        if (c !== 0) return c;
      }
      return 0;
    });
  };

  const rowKeys = axisKeys(pivot.rowFields);
  const colKeys = axisKeys(pivot.colFields);

  // Bucket every source row once, then read the buckets — one pass over the
  // data rather than one per output cell, which is the difference between a
  // pivot that refreshes instantly and one that stalls on a real workbook.
  const buckets = new Map();
  for (const line of visible) {
    const rk = keyOf(tupleOf(line, pivot.rowFields));
    const ck = keyOf(tupleOf(line, pivot.colFields));
    const k = rk + '\u001e' + ck;
    let bucket = buckets.get(k);
    if (!bucket) {
      // The two halves are kept BESIDE the key rather than parsed back out
      // of it: splitting a joined key is exactly where a value containing
      // the separator would land in the wrong bucket.
      bucket = { rk, ck, lines: [] };
      buckets.set(k, bucket);
    }
    bucket.lines.push(line);
  }
  const linesFor = (rowKey, colKey) => {
    if (rowKey === null && colKey === null) return visible;
    const wantRow = rowKey === null ? null : keyOf(rowKey);
    const wantCol = colKey === null ? null : keyOf(colKey);
    const out = [];
    for (const b of buckets.values()) {
      if (wantRow !== null && b.rk !== wantRow) continue;
      if (wantCol !== null && b.ck !== wantCol) continue;
      out.push(...b.lines);
    }
    return out;
  };
  const valueAt = (rowKey, colKey, df) =>
    aggregate(linesFor(rowKey, colKey).map((line) => line[df.fld]), df.subtotal);

  // ---- layout ------------------------------------------------------------
  const multiData = pivot.dataFields.length > 1;
  // Leaf columns: every column key crossed with every data field when there
  // is more than one, which is where Excel's "Values" level comes from.
  const leaves = [];
  for (const ck of colKeys) {
    for (const df of pivot.dataFields) leaves.push({ colKey: ck, df });
  }
  const grandCols = [];
  if (pivot.rowGrandTotals && (pivot.colFields.length > 0)) {
    for (const df of pivot.dataFields) grandCols.push({ colKey: null, df, grand: true });
  }

  // Row-header columns: compact form stacks every row field into one column,
  // tabular gives each its own.
  const rowHeaderCols = pivot.rowFields.length === 0
    ? 1
    : (pivot.compact ? 1 : pivot.rowFields.length);
  // Column-header rows: one per column field, plus a Values row when needed.
  const colHeaderRows = Math.max(1, pivot.colFields.length + (multiData ? 1 : 0));
  const pageRows = pivot.pageFields.length;

  const cells = [];
  const put = (r, c, value) => {
    if (value === undefined) return;
    cells.push({ row: pivot.location.top + r, col: pivot.location.left + c, value });
  };

  let cursor = 0;
  for (const pf of pivot.pageFields) {
    const field = pivot.cacheFields[pf.fld];
    put(cursor, 0, field ? field.name : 'Filter');
    let shown = '(All)';
    if (pf.item !== null) {
      const seen = [];
      for (const line of rows) {
        const v = line[pf.fld];
        if (!seen.some((x) => keyOf([x]) === keyOf([v]))) seen.push(v);
      }
      if (seen[pf.item] !== undefined) shown = seen[pf.item];
    }
    put(cursor, 1, shown);
    cursor += 1;
  }
  // Excel leaves a blank line between the filters and the body.
  if (pageRows) cursor += 1;

  const headTop = cursor;
  const bodyTop = headTop + colHeaderRows;

  // The corner. In compact form one cell says "Row Labels"; in tabular each
  // row field names its own column on the last header row.
  if (pivot.rowFields.length === 0) {
    put(bodyTop - 1, 0, '');
  } else if (pivot.compact) {
    put(bodyTop - 1, 0, 'Row Labels');
  } else {
    pivot.rowFields.forEach((f, i) => put(bodyTop - 1, i, pivot.cacheFields[f]?.name ?? ''));
  }
  // The column field names sit above the corner where there is room for them.
  pivot.colFields.forEach((f, i) => {
    if (i < colHeaderRows - 1 || pivot.rowFields.length === 0) {
      put(headTop + i, Math.max(0, rowHeaderCols - 1), pivot.cacheFields[f]?.name ?? '');
    }
  });

  const allCols = [...leaves, ...grandCols];
  allCols.forEach((leaf, i) => {
    const c = rowHeaderCols + i;
    if (leaf.grand) {
      put(headTop, c, 'Grand Total');
      if (multiData) put(bodyTop - 1, c, dataFieldLabel(pivot, leaf.df));
      return;
    }
    leaf.colKey.forEach((v, level) => put(headTop + level, c, v));
    if (multiData) put(bodyTop - 1, c, dataFieldLabel(pivot, leaf.df));
    else if (!pivot.colFields.length) put(bodyTop - 1, c, dataFieldLabel(pivot, leaf.df));
  });

  rowKeys.forEach((rk, i) => {
    const r = bodyTop + i;
    if (pivot.compact) {
      // Compact form stacks every level into one column; the joined label is
      // the whole tuple. ONE put — the first cut emitted the leaf and then
      // the joined label at the same coordinate, and only survived because
      // every consumer happened to let the later write win.
      put(r, 0, rk.join(' — '));
    } else {
      rk.forEach((v, level) => put(r, level, v));
    }
    allCols.forEach((leaf, j) => put(r, rowHeaderCols + j, valueAt(rk, leaf.grand ? null : leaf.colKey, leaf.df)));
  });

  let height = bodyTop + rowKeys.length;
  if (pivot.colGrandTotals && pivot.rowFields.length) {
    const r = height;
    put(r, 0, 'Grand Total');
    allCols.forEach((leaf, j) => put(r, rowHeaderCols + j, valueAt(null, leaf.grand ? null : leaf.colKey, leaf.df)));
    height += 1;
  }
  const width = rowHeaderCols + allCols.length;

  return {
    cells,
    height,
    width,
    // The three offsets the stored `location` carries, as THIS layout put
    // them — so a caller updating the definition states what it actually
    // wrote rather than preserving numbers that no longer describe the grid.
    firstHeaderRow: headTop,
    firstDataRow: bodyTop,
    firstDataCol: rowHeaderCols,
    area: {
      sheet: pivot.sheet,
      top: pivot.location.top,
      left: pivot.location.left,
      bottom: pivot.location.top + height - 1,
      right: pivot.location.left + width - 1,
    },
  };
}

/**
 * Point a pivot's stored `location` at a new rectangle, and mark its cache to
 * be rebuilt by whatever opens the file next. See the note at the top: we do
 * not rewrite the records ourselves.
 */
export function updatePivotLocation(wb, pivot, area, { headerRows, dataRow, dataCol }) {
  const xml = wb.pkg.text(pivot.part);
  const next = xml.replace(/<location\b([^>]*?)\/>/, (whole, a) => {
    let out = a
      .replace(/\s+ref="[^"]*"/, '')
      .replace(/\s+firstHeaderRow="[^"]*"/, '')
      .replace(/\s+firstDataRow="[^"]*"/, '')
      .replace(/\s+firstDataCol="[^"]*"/, '');
    out += ' ref="' + esc(areaRef(area)) + '"'
      + ' firstHeaderRow="' + headerRows + '"'
      + ' firstDataRow="' + dataRow + '"'
      + ' firstDataCol="' + dataCol + '"';
    return '<location' + out + '/>';
  });
  if (next !== xml) wb.pkg.write_(pivot.part, next);

  if (pivot.cachePart && wb.pkg.has(pivot.cachePart)) {
    const cache = wb.pkg.text(pivot.cachePart);
    if (!/\brefreshOnLoad="1"/.test(cache)) {
      const marked = cache.replace(/<pivotCacheDefinition\b([^>]*?)>/, (whole, a) =>
        '<pivotCacheDefinition' + a.replace(/\s+refreshOnLoad="[^"]*"/, '') + ' refreshOnLoad="1">');
      if (marked !== cache) wb.pkg.write_(pivot.cachePart, marked);
    }
  }
  return wb;
}

// ---- creation --------------------------------------------------------------

const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const SHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const CT = {
  cacheDefinition: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
  cacheRecords: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
  table: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
  rels: 'application/vnd.openxmlformats-package.relationships+xml',
};

/**
 * Add a relationship to a part, creating its `.rels` when there is none.
 * Returns the id it was given, because the caller usually has to reference it.
 */
// Both delegates: the one implementation lives on OoxmlPackage now, shared
// with the drawing and chart inserters. The wrappers keep this module's call
// sites unchanged.
function addRelationship(pkg, fromPart, type, target) {
  return pkg.addRelationshipTo(fromPart, type, target);
}

function nextPartNumber(pkg, dir, prefix) {
  return pkg.nextPartNumber(dir, prefix);
}

/** A cache item element for one distinct source value. */
function itemXml(v) {
  if (v === '' || v === null || v === undefined) return '<m/>';
  if (typeof v === 'number') return '<n v="' + v + '"/>';
  if (typeof v === 'boolean') return '<b v="' + (v ? 1 : 0) + '"/>';
  if (isError(v)) return '<e v="' + esc(v.type) + '"/>';
  return '<s v="' + esc(String(v)) + '"/>';
}

/**
 * Create a pivot table over a source range.
 *
 * Unlike a refresh — where rebuilding the cache would mean renumbering
 * somebody else's shared items — creation AUTHORS both halves, so the cache
 * is written honestly here: shared items for the fields on an axis, literal
 * values for the rest, and records that index them. `refreshOnLoad` still
 * rides along, so Excel re-derives everything from the source on open and
 * any disagreement resolves in the source's favour rather than ours.
 *
 * @param {object} wb Workbook
 * @param {object} spec
 * @param {string} spec.name
 * @param {{sheet: string, top: number, left: number, bottom: number, right: number}} spec.source
 * @param {{sheet: string, row: number, col: number}} spec.target
 * @param {string[]} spec.rowFields column names from the source's header row
 * @param {string[]} spec.colFields
 * @param {Array<{field: string, subtotal?: string}>} spec.dataFields
 * @param {(sheet: string, row: number, col: number) => any} spec.readCell
 * @returns {object} the pivot descriptor, ready for `computePivot`
 */
export function planPivot(wb, spec) {
  const { source, readCell } = spec;
  const name = String(spec.name ?? '').trim();
  if (!/^[^\\/?*[\]:]{1,255}$/.test(name)) throw new Error('a pivot table needs a name');
  if (!wb.sheetNames().includes(source.sheet)) throw new Error('no sheet "' + source.sheet + '"');
  if (source.bottom <= source.top) throw new Error('the source needs a header row and at least one row of data');

  const headers = [];
  for (let c = source.left; c <= source.right; c++) {
    const v = readCell(source.sheet, source.top, c);
    const text = v === '' || v === null || v === undefined ? '' : String(v);
    if (!text) throw new Error('every source column needs a heading — ' + makeRef(source.top, c) + ' is empty');
    if (headers.includes(text)) throw new Error('two source columns are both called "' + text + '"');
    headers.push(text);
  }
  const indexOfField = (label) => {
    const i = headers.findIndex((h) => h.toLowerCase() === String(label ?? '').toLowerCase());
    if (i < 0) throw new Error('no source column called "' + label + '"');
    return i;
  };

  const rowFields = (spec.rowFields ?? []).map(indexOfField);
  const colFields = (spec.colFields ?? []).map(indexOfField);
  const dataFields = (spec.dataFields ?? []).map((d) => {
    const subtotal = d.subtotal ?? 'sum';
    if (!SUBTOTALS.has(subtotal)) throw new Error('"' + subtotal + '" is not a way to summarise');
    return { fld: indexOfField(d.field), subtotal, name: d.name ?? null };
  });
  if (!dataFields.length) throw new Error('a pivot needs at least one value to summarise');
  const overlap = rowFields.filter((f) => colFields.includes(f));
  if (overlap.length) throw new Error('"' + headers[overlap[0]] + '" cannot be on both axes');

  // The source rows, once — they seed both the cache and the first render.
  const rows = [];
  for (let r = source.top + 1; r <= source.bottom; r++) {
    const line = [];
    for (let c = source.left; c <= source.right; c++) line.push(readCell(source.sheet, r, c));
    rows.push(line);
  }
  return { name, headers, rowFields, colFields, dataFields, rows };
}

/**
 * Write the parts a planned pivot needs.
 *
 * Validation lives in `planPivot` so a caller can refuse BEFORE it records an
 * undo step — the house rule that a refused edit leaves no history behind.
 * By the time this runs, nothing is left to reject.
 */
export function createPivot(wb, spec, plan = planPivot(wb, spec)) {
  const { source, target } = spec;
  const { name, headers, rowFields, colFields, dataFields, rows } = plan;

  // Axis fields get shared items; a value field keeps its literals, which is
  // how Excel writes a cache and keeps the records small.
  const axisFields = new Set([...rowFields, ...colFields]);
  const shared = headers.map((_, f) => {
    if (!axisFields.has(f)) return null;
    const seen = new Map();
    for (const line of rows) {
      const v = line[f] ?? '';
      const k = keyOf([v]);
      if (!seen.has(k)) seen.set(k, v);
    }
    return [...seen.values()];
  });

  const num = nextPartNumber(wb.pkg, 'xl/pivotTables/', 'pivotTable');
  const cacheNum = nextPartNumber(wb.pkg, 'xl/pivotCache/', 'pivotCacheDefinition');
  const tablePart = 'xl/pivotTables/pivotTable' + num + '.xml';
  const cachePart = 'xl/pivotCache/pivotCacheDefinition' + cacheNum + '.xml';
  const recordsPart = 'xl/pivotCache/pivotCacheRecords' + cacheNum + '.xml';

  // A cacheId unique in the workbook, so two pivots never collide.
  const wbXml = wb.pkg.text(wb.mainPart);
  const usedIds = [...wbXml.matchAll(/<pivotCache\b[^>]*\bcacheId="(\d+)"/g)].map((m) => Number(m[1]));
  const cacheId = (usedIds.length ? Math.max(...usedIds) : 0) + 1;

  // ---- the records ------------------------------------------------------
  const recordsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<pivotCacheRecords xmlns="' + SHEET_NS + '" count="' + rows.length + '">'
    + rows.map((line) => '<r>' + headers.map((_, f) => {
      const v = line[f] ?? '';
      if (!shared[f]) return itemXml(v);
      const at = shared[f].findIndex((x) => keyOf([x]) === keyOf([v]));
      return '<x v="' + (at < 0 ? 0 : at) + '"/>';
    }).join('') + '</r>').join('')
    + '</pivotCacheRecords>';
  wb.pkg.addPart(recordsPart, recordsXml, CT.cacheRecords);

  // ---- the cache definition ---------------------------------------------
  const cacheFieldsXml = headers.map((h, f) => {
    const items = shared[f];
    const body = items
      ? '<sharedItems count="' + items.length + '">' + items.map(itemXml).join('') + '</sharedItems>'
      : '<sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1"/>';
    return '<cacheField name="' + esc(h) + '" numFmtId="0">' + body + '</cacheField>';
  }).join('');
  wb.pkg.addPart(cachePart,
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<pivotCacheDefinition xmlns="' + SHEET_NS + '" xmlns:r="' + OFFICE_REL.slice(0, -1) + '"'
    + ' refreshOnLoad="1" refreshedVersion="6" minRefreshableVersion="3" createdVersion="6"'
    + ' recordCount="' + rows.length + '">'
    + '<cacheSource type="worksheet"><worksheetSource ref="' + esc(areaRef(source)) + '" sheet="' + esc(source.sheet) + '"/></cacheSource>'
    + '<cacheFields count="' + headers.length + '">' + cacheFieldsXml + '</cacheFields>'
    + '</pivotCacheDefinition>',
    CT.cacheDefinition);
  // The records hang off the cache definition's own relationship, and the
  // r:id must name it — a cache pointing at nothing is a file Excel refuses.
  const recordsRelId = addRelationship(wb.pkg, cachePart, OFFICE_REL + 'pivotCacheRecords',
    'pivotCacheRecords' + cacheNum + '.xml');
  wb.pkg.write_(cachePart, wb.pkg.text(cachePart)
    .replace('<pivotCacheDefinition ', '<pivotCacheDefinition r:id="' + recordsRelId + '" '));

  // ---- the pivot table definition ---------------------------------------
  const axisAttr = (f) => (rowFields.includes(f) ? ' axis="axisRow"'
    : colFields.includes(f) ? ' axis="axisCol"' : '');
  const pivotFieldsXml = headers.map((_, f) => {
    const isData = dataFields.some((d) => d.fld === f);
    const items = shared[f];
    const body = items
      ? '<items count="' + (items.length + 1) + '">'
        + items.map((_v, i) => '<item x="' + i + '"/>').join('')
        + '<item t="default"/></items>'
      : '';
    const attrsText = axisAttr(f) + (isData ? ' dataField="1"' : '') + ' showAll="0"';
    return body
      ? '<pivotField' + attrsText + '>' + body + '</pivotField>'
      : '<pivotField' + attrsText + '/>';
  }).join('');

  // Item lists for each axis: every member, then the grand total line.
  const axisItems = (fields, tag) => {
    if (!fields.length) return '<' + tag + ' count="1"><i/></' + tag + '>';
    const count = shared[fields[0]]?.length ?? 0;
    const lines = [];
    for (let i = 0; i < count; i++) lines.push('<i>' + (i === 0 ? '<x/>' : '<x v="' + i + '"/>') + '</i>');
    lines.push('<i t="grand"><x/></i>');
    return '<' + tag + ' count="' + lines.length + '">' + lines.join('') + '</' + tag + '>';
  };

  const dataFieldsXml = '<dataFields count="' + dataFields.length + '">'
    + dataFields.map((d) => '<dataField name="' + esc(d.name ?? defaultDataName(headers[d.fld], d.subtotal))
      + '" fld="' + d.fld + '"'
      + (d.subtotal === 'sum' ? '' : ' subtotal="' + d.subtotal + '"')
      + ' baseField="0" baseItem="0"/>').join('')
    + '</dataFields>';

  wb.pkg.addPart(tablePart,
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<pivotTableDefinition xmlns="' + SHEET_NS + '" name="' + esc(name) + '" cacheId="' + cacheId + '"'
    + ' applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0"'
    + ' applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values"'
    + ' updatedVersion="6" minRefreshableVersion="3" createdVersion="6"'
    + ' useAutoFormatting="1" itemPrintTitles="1" indent="0" compact="0" compactData="0"'
    + ' outline="0" outlineData="0" multipleFieldFilters="0">'
    + '<location ref="' + esc(makeRef(target.row, target.col)) + ':' + esc(makeRef(target.row, target.col)) + '"'
    + ' firstHeaderRow="0" firstDataRow="1" firstDataCol="' + Math.max(1, rowFields.length) + '"/>'
    + '<pivotFields count="' + headers.length + '">' + pivotFieldsXml + '</pivotFields>'
    + (rowFields.length ? '<rowFields count="' + rowFields.length + '">'
      + rowFields.map((f) => '<field x="' + f + '"/>').join('') + '</rowFields>' : '')
    + axisItems(rowFields, 'rowItems')
    + (colFields.length ? '<colFields count="' + colFields.length + '">'
      + colFields.map((f) => '<field x="' + f + '"/>').join('') + '</colFields>' : '')
    + axisItems(colFields, 'colItems')
    + dataFieldsXml
    + '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1"'
    + ' showRowStripes="0" showColStripes="0" showLastColumn="1"/>'
    + '</pivotTableDefinition>',
    CT.table);

  addRelationship(wb.pkg, tablePart, OFFICE_REL + 'pivotCacheDefinition',
    '../pivotCache/pivotCacheDefinition' + cacheNum + '.xml');
  const targetPart = wb.partNameFor(target.sheet);
  addRelationship(wb.pkg, targetPart, OFFICE_REL + 'pivotTable',
    '../pivotTables/pivotTable' + num + '.xml');

  // The workbook lists its caches, and points at each definition. Schema
  // order puts `pivotCaches` late, so it goes at the end of the element
  // rather than beside `sheets`.
  const wbRelId = addRelationship(wb.pkg, wb.mainPart, OFFICE_REL + 'pivotCacheDefinition',
    'pivotCache/pivotCacheDefinition' + cacheNum + '.xml');
  const cacheEntry = '<pivotCache cacheId="' + cacheId + '" r:id="' + wbRelId + '"/>';
  let nextWbXml = wb.pkg.text(wb.mainPart);
  if (/<pivotCaches>/.test(nextWbXml)) {
    nextWbXml = nextWbXml.replace('</pivotCaches>', cacheEntry + '</pivotCaches>');
  } else {
    nextWbXml = nextWbXml.replace(/<\/workbook>\s*$/, '<pivotCaches>' + cacheEntry + '</pivotCaches></workbook>');
  }
  // A workbook built without one has no r: namespace declared, and an
  // undeclared prefix is a corrupt package rather than a cosmetic problem.
  if (!/xmlns:r=/.test(nextWbXml)) {
    nextWbXml = nextWbXml.replace('<workbook ', '<workbook xmlns:r="' + OFFICE_REL.slice(0, -1) + '" ');
  }
  wb.pkg.write_(wb.mainPart, nextWbXml);

  return readPivotPart(wb, target.sheet, tablePart);
}

/** "Sum of Qty" — the name Excel gives a value field nobody renamed. */
function defaultDataName(field, subtotal) {
  const verb = {
    sum: 'Sum', count: 'Count', countNums: 'Count', average: 'Average',
    max: 'Max', min: 'Min', product: 'Product',
    stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'Varp',
  }[subtotal] ?? 'Sum';
  return verb + ' of ' + field;
}

export { parseArea, areaRef, aggregate, addRelationship };
