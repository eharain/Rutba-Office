/**
 * Data → Get & Transform, in the workbook: a query's result loaded as a
 * table on a sheet of its own, the query kept in the file, Refresh running
 * it again on what its source holds now.
 *
 * The queries are kept in a part of the suite's own — customXml/
 * rutbaQueries.xml, related from the workbook — as JSON: the source (a
 * table or a range of this workbook, or a CSV file by its path) and the
 * applied steps (queries.js). Excel keeps a custom part it does not know as
 * it found it and sees the loaded tables as the tables they are; it does not
 * see the queries as its own Power Query ones, which live in a binary
 * package this suite neither reads nor writes.
 *
 * Loading and refreshing are each one undo step: the sheet, its table and
 * the queries part are put back together.
 */

import { runSteps, parseDelimited, describeStep } from './queries.js';
import { colName } from './selection.js';

const PART = 'customXml/rutbaQueries.xml';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml';

const a1 = (row, col) => colName(col) + (row + 1);
const areaText = (r) => `${a1(r.top, r.left)}:${a1(r.bottom, r.right)}`;
function parseArea(text) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(String(text || '').trim().replace(/^.*!/, ''));
  if (!m) throw new Error(`"${text}" is not a range`);
  const col = (s) => [...s.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const top = Number(m[2]) - 1;
  const left = col(m[1]);
  return { top, left, bottom: m[4] ? Number(m[4]) - 1 : top, right: m[3] ? col(m[3]) : left };
}
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** The workbook's queries, as the part keeps them. */
export function readQueries(view) {
  if (!view.pkg.has(PART)) return [];
  try {
    const xml = view.pkg.text(PART);
    const m = /<!\[CDATA\[([\s\S]*)\]\]>/.exec(xml);
    const list = JSON.parse((m ? m[1] : '[]').replace(/\]\]\]\]><!\[CDATA\[>/g, ']]>'));
    return Array.isArray(list) ? list : [];
  } catch {
    return [];
  }
}

function writeQueries(view, list) {
  const json = JSON.stringify(list).replace(/\]\]>/g, ']]]]><![CDATA[>');
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<queries xmlns="urn:rutba:queries"><![CDATA[${json}]]></queries>`;
  if (view.pkg.has(PART)) view.pkg.write_(PART, xml);
  else {
    view.pkg.addPart(PART, xml, 'application/xml');
    view.pkg.addRelationshipTo(view.workbook.mainPart, REL, '../' + PART);
  }
}

/** A query's source read now: { columns, rows }. A CSV is read with `read(path)`. */
export function sourceTable(view, source, read = null, seen = new Set()) {
  if (!source) throw new Error('A query needs a source');
  // Another query of the workbook: its result as it would load now.
  if (source.kind === 'query') {
    const q = readQueries(view).find((x) => x.id === source.id || x.name === source.name);
    if (!q) throw new Error(`There is no query called "${source.name || source.id}"`);
    if (seen.has(q.id)) throw new Error(`Query "${q.name}" reads itself`);
    const next = new Set(seen).add(q.id);
    return runSteps(sourceTable(view, q.source, read, next), q.steps, { table: (s) => sourceTable(view, s, read, next) });
  }
  if (source.kind === 'csv') {
    if (typeof read !== 'function') throw new Error('A file source is read on the main process');
    const rows = parseDelimited(read(source.path));
    const width = Math.max(1, ...rows.map((r) => r.length));
    return { columns: Array.from({ length: width }, (_, i) => `Column${i + 1}`), rows: rows.map((r) => Array.from({ length: width }, (_, i) => (r[i] === undefined ? null : r[i]))) };
  }
  let sheet;
  let area;
  let headers = 1;
  if (source.kind === 'table') {
    const t = view.workbook.tables().find((x) => String(x.name).toLowerCase() === String(source.table).toLowerCase());
    if (!t) throw new Error(`There is no table called "${source.table}"`);
    sheet = t.sheet;
    area = parseArea(t.ref);
    headers = t.headerRowCount ?? 1;
    area.bottom -= t.totalsRowCount || 0;
  } else if (source.kind === 'range') {
    sheet = source.sheet;
    if (!view.sheetNames().includes(sheet)) throw new Error(`There is no sheet called "${sheet}"`);
    area = parseArea(source.ref);
  } else {
    throw new Error(`"${source.kind}" is not a source a query reads`);
  }
  const value = (r, c) => {
    const v = view.calc.getValue(sheet, r, c);
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'object') return null; // an error value is a blank to a query
    return v;
  };
  const columns = [];
  for (let c = area.left; c <= area.right; c++) columns.push(headers ? String(view.calc.getValue(sheet, area.top, c) ?? '') : `Column${c - area.left + 1}`);
  const rows = [];
  for (let r = area.top + (headers ? 1 : 0); r <= area.bottom; r++) {
    const row = [];
    for (let c = area.left; c <= area.right; c++) row.push(value(r, c));
    rows.push(row);
  }
  return { columns, rows };
}

/** What a source is, in words: "Table Sales", "Data!A1:D7", "stock.csv". */
export function describeSource(source) {
  if (!source) return '';
  if (source.kind === 'query') return `Query ${source.name || source.id}`;
  if (source.kind === 'table') return `Table ${source.table}`;
  if (source.kind === 'range') return `${source.sheet}!${source.ref}`;
  if (source.kind === 'csv') return String(source.path || '').split(/[\\/]/).pop();
  return source.kind;
}

/** The first rows of what a query would load, for the editor: nothing changes. */
export function previewQuery(view, { source, steps = [], upTo = null, read = null, limit = 100 } = {}) {
  const t = runSteps(sourceTable(view, source, read), steps, { upTo, table: (s) => sourceTable(view, s, read) });
  return { columns: t.columns, rows: t.rows.slice(0, limit), total: t.rows.length, steps: steps.map(describeStep) };
}

const sheetNameFor = (view, name) => {
  const base = String(name || 'Query').replace(/[[\]:*?/\\]/g, ' ').replace(/^'+|'+$/g, '').trim().slice(0, 28) || 'Query';
  const taken = new Set(view.sheetNames().map((n) => n.toLowerCase()));
  let out = base;
  for (let k = 2; taken.has(out.toLowerCase()); k++) out = `${base.slice(0, 26)} ${k}`;
  return out;
};
const tableNameFor = (view, name) => {
  let base = String(name || 'Query').replace(/[^A-Za-z0-9_]+/g, '_').replace(/^_+|_+$/g, '') || 'Query';
  if (!/^[A-Za-z_]/.test(base) || /^[A-Za-z]{1,3}\d+$/.test(base)) base = `Q_${base}`;
  const taken = new Set(view.workbook.tables().map((t) => String(t.name).toLowerCase()));
  let out = base;
  for (let k = 2; taken.has(out.toLowerCase()); k++) out = `${base}_${k}`;
  return out;
};

/** A cell on any sheet set to a value as it is, not as typed. */
function put(view, sheet, row, col, value) {
  const v = value === null || value === undefined ? '' : typeof value === 'string' && value.startsWith('=') ? ` ${value}` : value;
  view.calc.setCell(sheet, row, col, v);
  view.dirtyCells.add(sheet + '!' + a1(row, col));
}

/** A query's result written over its table: the cells, and the table grown or shrunk to fit. */
function writeResult(view, q, table) {
  const sheet = q.load.sheet;
  const old = view.workbook.tables().find((t) => t.name === q.load.table && t.sheet === sheet);
  const width = Math.max(1, table.columns.length);
  const bottom = Math.max(1, table.rows.length);
  if (old) {
    const was = parseArea(old.ref);
    for (let r = was.top; r <= was.bottom; r++) for (let c = was.left; c <= was.right; c++) if (r > bottom || c >= width) put(view, sheet, r, c, null);
  }
  table.columns.forEach((name, c) => put(view, sheet, 0, c, name));
  table.rows.forEach((row, r) => row.forEach((v, c) => put(view, sheet, r + 1, c, v)));
  if (!table.rows.length) for (let c = 0; c < width; c++) put(view, sheet, 1, c, null);
  const ref = areaText({ top: 0, left: 0, bottom, right: width - 1 });
  if (!old) {
    view.workbook.addTable(sheet, ref, { name: q.load.table, headerNames: table.columns });
    return;
  }
  // The table's range and columns rewritten in place, its style and id kept.
  let xml = view.pkg.text(old.part);
  xml = xml.replace(/(<table\b[^>]*?\sref=")[^"]*"/, (m, p) => `${p}${ref}"`);
  xml = xml.replace(/(<autoFilter\b[^>]*?\sref=")[^"]*"/, (m, p) => `${p}${ref}"`);
  const cols = `<tableColumns count="${table.columns.length}">${table.columns.map((c, i) => `<tableColumn id="${i + 1}" name="${esc(c)}"/>`).join('')}</tableColumns>`;
  xml = xml.replace(/<tableColumns\b[^>]*>[\s\S]*?<\/tableColumns>|<tableColumns\b[^>]*\/>/, () => cols);
  view.pkg.write_(old.part, xml);
}

/**
 * The sheet a query's table is on now. The sheet can be renamed after the
 * query loaded, so the table, which keeps its name, says where the result
 * lives; going by the old sheet name would build a second sheet and a
 * second table beside the renamed one.
 */
function syncLoadSheet(view, q) {
  const t = view.workbook.tables().find((x) => x.name === q.load.table);
  if (t && t.sheet !== q.load.sheet) q.load.sheet = t.sheet;
}

/** Every part a query edit may change, for its one undo step. */
function partsOf(view) {
  const out = [...view.workbook.sheets().map((s) => s.part), view.workbook.mainPart, 'xl/_rels/workbook.xml.rels', '[Content_Types].xml'];
  for (const s of view.workbook.sheets()) {
    const rels = s.part.replace(/([^/]+)$/, '_rels/$1.rels');
    if (view.pkg.has(rels)) out.push(rels);
  }
  for (const t of view.workbook.tables()) out.push(t.part);
  if (view.pkg.has(PART)) out.push(PART);
  return [...new Set(out)].filter((p) => view.pkg.has(p));
}

/** Run an edit of the queries as one structural undo step, the cells written into the parts at its end. */
function queryEdit(view, label, fn) {
  return view._edit(label, null, [], () => {
    view._flushForStructure();
    const out = fn();
    view._flushForStructure();
    view._rebuildDerivedState();
    return out;
  }, { parts: partsOf(view), tracksNewParts: true, structural: true, sheetGate: false });
}

/**
 * Data → From Table/Range, From Text/CSV and the editor's Close & Load: a
 * query run and loaded as a table on a new sheet named after it, and kept.
 */
export function addQuery(view, { name = 'Query', source, steps = [], read = null } = {}) {
  view._structureGate();
  const list = readQueries(view);
  const taken = new Set(list.map((q) => q.name.toLowerCase()));
  let clean = String(name || 'Query').trim() || 'Query';
  for (let k = 2; taken.has(clean.toLowerCase()); k++) clean = `${String(name).trim() || 'Query'} (${k})`;
  const table = runSteps(sourceTable(view, source, read), steps, { table: (s) => sourceTable(view, s, read) });
  const sheet = sheetNameFor(view, clean);
  const q = { id: `q${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`, name: clean, source, steps, load: { sheet, table: tableNameFor(view, clean) }, loaded: table.rows.length, refreshed: Date.now() };
  queryEdit(view, 'load query', () => {
    view.workbook.addSheet(sheet);
    view._rebuildDerivedState();
    writeResult(view, q, table);
    writeQueries(view, [...list, q]);
  });
  view.selectSheet(sheet);
  return { ...q, sheet, rows: table.rows.length };
}

/** Data → Refresh All, or one query's Refresh: each run again on its source as it is now. */
export function refreshQueries(view, { id = null, read = null } = {}) {
  const list = readQueries(view);
  const chosen = list.filter((q) => !id || q.id === id);
  if (!chosen.length) return 0;
  for (const q of chosen) syncLoadSheet(view, q);
  const results = chosen.map((q) => ({ q, table: runSteps(sourceTable(view, q.source, read, new Set([q.id])), q.steps, { table: (s) => sourceTable(view, s, read, new Set([q.id])) }) }));
  queryEdit(view, 'refresh', () => {
    for (const { q, table } of results) {
      if (!view.sheetNames().includes(q.load.sheet)) { view.workbook.addSheet(q.load.sheet); view._rebuildDerivedState(); }
      writeResult(view, q, table);
      q.loaded = table.rows.length;
      q.refreshed = Date.now();
    }
    writeQueries(view, list);
  });
  return results.length;
}

/** The editor's Close & Load on a query already loaded: its name, source and steps changed, and it run again. */
export function editQuery(view, { id, name, source, steps, read = null } = {}) {
  const list = readQueries(view);
  const q = list.find((x) => x.id === id);
  if (!q) throw new Error('That query is no longer in this workbook');
  syncLoadSheet(view, q);
  if (name !== undefined) q.name = String(name).trim() || q.name;
  if (source !== undefined) q.source = source;
  if (steps !== undefined) q.steps = steps;
  const table = runSteps(sourceTable(view, q.source, read, new Set([q.id])), q.steps, { table: (s) => sourceTable(view, s, read, new Set([q.id])) });
  queryEdit(view, 'edit query', () => {
    if (!view.sheetNames().includes(q.load.sheet)) { view.workbook.addSheet(q.load.sheet); view._rebuildDerivedState(); }
    writeResult(view, q, table);
    q.loaded = table.rows.length;
    q.refreshed = Date.now();
    writeQueries(view, list);
  });
  return { ...q, rows: table.rows.length };
}

/** Delete a query: its loaded sheet stays, as Excel leaves it, only no longer refreshed. */
export function removeQuery(view, { id } = {}) {
  const list = readQueries(view);
  if (!list.some((q) => q.id === id)) return 0;
  queryEdit(view, 'delete query', () => writeQueries(view, list.filter((q) => q.id !== id)));
  return 1;
}
