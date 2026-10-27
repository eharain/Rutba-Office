// The other sources a Data Files table is read from (D-2026-10-22-13, step
// 5), each kept as a copy in the file and read again on a schedule: a CSV or
// spreadsheet file (.csv, .xlsx, .xls, .ods) at a path on this computer, a
// network share or a web address; a table or a query of another Data Files
// database, opened read only and, when it is protected, as the person who
// says who they are there; and a web address answering JSON. Each answers
// the records' fields and rows; base-sync.js keeps them in the file.

import fs from 'node:fs';
import path from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import { SheetView } from '@rutba/sheet-view';
import { decodeText } from '@rutba/office-formats/text';
import { readOdf } from '@rutba/office-formats/odf';
import { readXls } from '@rutba/office-formats/msxls';
import { xlsModelToXlsx } from '@rutba/office-formats/msxls-xlsx';
import { parseDelimited } from '@rutba/ooxml/mailmerge';
import { Base, tableFromRows, tableFromRecords, valueAt, sourceFieldNames } from '@rutba/database';

/** The files a source can be: what the Link box offers, and what it reads. */
export const SOURCE_FILES = ['csv', 'tsv', 'txt', 'xlsx', 'xlsm', 'xls', 'ods'];

const isWeb = (where) => /^https?:\/\//i.test(String(where || '').trim());
const TIMEOUT = 20000;

/** The address's own name, for a message: its host. */
const hostOf = (url) => { try { return new URL(url).host; } catch { return String(url); } };

/** A web address fetched, its failures said in words a person can act on. */
async function fetchFrom(url, { headers = {}, fetchImpl = fetch } = {}) {
  let res;
  try {
    res = await fetchImpl(url, { headers, redirect: 'follow', signal: AbortSignal.timeout(TIMEOUT) });
  } catch (err) {
    const m = String(err?.cause?.code || err?.code || err?.message || err);
    if (/timeout|abort/i.test(m)) throw Object.assign(new Error(`${hostOf(url)} did not answer in time.`), { unreachable: true });
    throw Object.assign(new Error(`${hostOf(url)} could not be reached.`), { unreachable: true });
  }
  if (res.status === 401 || res.status === 403) throw Object.assign(new Error(`${hostOf(url)} refused the key or token (${res.status}).`), { refused: true });
  if (!res.ok) throw new Error(`${hostOf(url)} answered ${res.status}${res.statusText ? ` ${res.statusText}` : ''}.`);
  return res;
}

/* ── files: CSV and spreadsheets ───────────────────────────────────── */

/** What a file's bytes are: 'csv', 'xlsx', 'ods' or 'xls', by their first bytes, then by their name. */
function kindOf(bytes, name) {
  const b = Buffer.from(bytes.subarray(0, 512));
  if (b[0] === 0xd0 && b[1] === 0xcf && b[2] === 0x11 && b[3] === 0xe0) return 'xls';
  if (b[0] === 0x50 && b[1] === 0x4b) return b.includes(Buffer.from('opendocument.spreadsheet')) ? 'ods' : 'xlsx';
  const ext = path.extname(String(name || '').split(/[?#]/)[0]).slice(1).toLowerCase();
  return ext === 'tsv' ? 'tsv' : 'csv';
}

/** A spreadsheet's sheets as rows of values, a date as its serial with the columns formatted as dates noted. */
function workbookRows(bytes, sheet) {
  const view = SheetView.open(Buffer.from(bytes), { viewportWidth: 800, viewportHeight: 600 });
  const sheets = view.sheetNames();
  const chosen = sheet && sheets.includes(sheet) ? sheet : sheets[0];
  view.selectSheet(chosen);
  const { maxRow, maxCol } = view.calc.usedBounds(chosen);
  const rows = [];
  const dated = {};
  for (let r = 0; r <= maxRow; r++) {
    const row = [];
    for (let c = 0; c <= maxCol; c++) {
      let v = view.calc.getValue(chosen, r, c);
      if (v && typeof v === 'object') v = v.error || v.text || String(v);
      // A sum's answer to Excel's 15 figures, as the sheet shows it: 3588.39, not 3588.3899999999994.
      if (typeof v === 'number' && Number.isFinite(v)) v = Number(v.toPrecision(15));
      if (v === '') v = null;
      if (typeof v === 'number' && r > 0) {
        const fmt = String(view.formatFor(r, c) || '');
        const bare = fmt.replace(/"[^"]*"|\[[^\]]*\]/g, '');
        if (/y/i.test(bare) && /d/i.test(bare)) dated[c] = /h/i.test(bare) ? 'datetime' : dated[c] || 'date';
      }
      row.push(v);
    }
    rows.push(row);
  }
  return { sheets, sheet: chosen, rows, dated };
}

/** An OpenDocument spreadsheet's sheet as rows of values: dates as their ISO text, numbers as numbers. */
function odsRows(bytes, sheet) {
  const odf = readOdf(new Uint8Array(bytes));
  const all = odf.sheets || [];
  const sheets = all.map((s) => s.name);
  const s = all.find((x) => x.name === sheet) || all[0];
  if (!s) throw new Error('The file has no sheets.');
  const rows = (s.rows || []).map((row) => row.map((cell) => {
    if (!cell || cell.covered) return null;
    if (['float', 'percentage', 'currency'].includes(cell.type)) return Number(cell.value ?? cell.text);
    if (cell.type === 'date' && cell.value) return String(cell.value).replace(/(T\d\d:\d\d:\d\d)\.\d+$/, '$1').replace(/T00:00:00$/, '');
    if (cell.type === 'boolean') return String(cell.value ?? cell.text).toLowerCase() === 'true';
    return cell.text ?? null;
  }));
  return { sheets, sheet: s.name, rows, dated: {} };
}

/**
 * A CSV or spreadsheet file at a path, a network share or a web address:
 * { kind, sheets, sheet, fields, rows }, the row numbered `header` naming
 * the fields (the first, unless the person said).
 */
export async function readFileSource({ where, sheet = null, header = 1 }, { fetchImpl } = {}) {
  const at = String(where || '').trim();
  if (!at) throw new Error('Say where the file is: a path, a network share or a web address.');
  let bytes;
  if (isWeb(at)) {
    bytes = new Uint8Array(await (await fetchFrom(at, { fetchImpl })).arrayBuffer());
  } else {
    try {
      bytes = new Uint8Array(await fs.promises.readFile(at));
    } catch (err) {
      if (err.code === 'ENOENT') throw Object.assign(new Error(`${path.basename(at)} is not there: it was moved, renamed or deleted, or its share is not reachable.`), { unreachable: true });
      throw Object.assign(new Error(`${path.basename(at)} could not be read: ${err.message}`), { unreachable: true });
    }
  }
  const kind = kindOf(bytes, at);
  let got;
  if (kind === 'csv' || kind === 'tsv') got = { sheets: null, sheet: null, rows: parseDelimited(decodeText(bytes), kind === 'tsv' ? '\t' : null), dated: {} };
  else if (kind === 'ods') got = odsRows(bytes, sheet);
  else if (kind === 'xls') got = workbookRows(xlsModelToXlsx(readXls(bytes)), sheet);
  else got = workbookRows(bytes, sheet);
  const { fields, rows } = tableFromRows(got.rows, { header, dated: got.dated });
  return { kind, sheets: got.sheets, sheet: got.sheet, fields, rows };
}

/* ── another Data Files database ───────────────────────────────────── */

/** Another .rbase opened to read and nothing else; a protected one as the person named, with their password. */
function openOther(file, { person = null, password = null } = {}) {
  let db;
  try {
    db = new DatabaseSync(file, { readOnly: true });
  } catch (err) {
    throw Object.assign(new Error(fs.existsSync(file) ? `${path.basename(file)} could not be opened: ${err.message}` : `${path.basename(file)} is not there: it was moved, renamed or deleted, or its share is not reachable.`), { unreachable: true });
  }
  const base = new Base(db);
  if (!base.hasDesignTable('rutba_tables')) { base.db.close(); throw new Error(`${path.basename(file)} is not a Data Files database.`); }
  if (base.isProtected()) {
    if (!person || !password) { base.db.close(); throw Object.assign(new Error(`${path.basename(file)} is protected: say who you are there.`), { who: true }); }
    try { base.signIn(person, password); } catch (err) { base.db.close(); throw Object.assign(new Error(err.message), { who: true }); }
  }
  return base;
}

/** What another database offers to read: its tables (those whose records it keeps) and its queries, as its role shows them. */
export function rbaseItems(file, who = {}) {
  const base = openOther(file, who);
  try {
    const tables = base.tables().filter((t) => !t.linked || t.mode === 'copy').map((t) => ({ kind: 'table', name: t.name }));
    const queries = base.queries().map((x) => ({ kind: 'query', name: x.name }));
    return { protected: base.isProtected(), items: [...tables, ...queries] };
  } finally {
    base.db.close();
  }
}

/** A table or query of another database read: { fields, rows }, what that person's role may see and nothing more. */
export function readRbaseSource({ where, item }, who = {}) {
  const file = String(where || '').trim();
  if (!file) throw new Error('Say where the database is.');
  const base = openOther(file, who);
  try {
    if (item?.kind === 'query') {
      if (!base.queries().some((x) => x.name === item.name)) throw new Error(`There is no query called ${item.name} there, or your role there does not show it.`);
      const out = base.queryRows({ name: item.name, limit: 10_000_000 });
      const names = sourceFieldNames(out.columns.map((c) => c.name));
      const fields = out.columns.map((c, i) => ({ name: names[i], type: c.type && c.type !== 'autonumber' ? c.type : c.type === 'autonumber' ? 'integer' : 'text' }));
      const rows = out.rows.map((r) => Object.fromEntries(out.columns.map((c, i) => [names[i], r[c.name] ?? null])));
      return { fields, rows };
    }
    const name = item?.name;
    if (!base.tables().some((t) => t.name === name)) throw new Error(`There is no table called ${name} there, or your role there does not show it.`);
    const fields = base.fields(name).map((f) => ({ name: f.name, type: f.type === 'autonumber' ? 'integer' : f.type, caption: f.caption }));
    const rows = base.rows(name, { limit: 10_000_000 }).map((r) => Object.fromEntries(fields.map((f) => [f.name, r[f.name] ?? null])));
    return { fields, rows };
  } finally {
    base.db.close();
  }
}

/* ── web addresses answering JSON ──────────────────────────────────── */

/** A page's address with its page number set. */
const withPage = (url, param, n) => { const u = new URL(url); u.searchParams.set(param, String(n)); return u.toString(); };

/** The next page's address from a Link header (rel="next"), as many APIs give it. */
function linkNext(res) {
  const h = res.headers.get('link') || '';
  const m = /<([^>]+)>\s*;\s*rel="?next"?/i.exec(h);
  return m ? m[1] : null;
}

/** The records in one answer, found where the person said (data.items), or the answer itself when it is a list. */
function recordsIn(answer, at) {
  const v = at ? valueAt(answer, at) : answer;
  if (Array.isArray(v)) return v;
  if (at) throw new Error(`There is no list of records at ${at} in the answer.`);
  throw new Error('The answer is not a list of records: say where they are in it, as data.items.');
}

/**
 * The records a web address answers in JSON, every page of them: paging by
 * a next link (`paging.next`, the answer's path to it, or the Link header)
 * or by a page number (`paging.param`, from `paging.start`). Headers are
 * sent as given, `secretOf(name)` giving those kept in the protected store;
 * those go only to the address's own origin, not to a next page elsewhere.
 * Answers { fields, rows, pages }.
 */
export async function readJsonSource({ where, headers = [], records = '', paging = { kind: 'none' } }, { secretOf = () => null, fetchImpl, maxPages = 500 } = {}) {
  const first = String(where || '').trim();
  if (!isWeb(first)) throw new Error('A web address starts with https:// or http://.');
  const sent = { Accept: 'application/json' };
  const secret = {};
  for (const h of headers || []) {
    const name = String(h?.name || '').trim();
    if (!name) continue;
    const value = h.secret ? secretOf(name) : h.value;
    if (value != null && value !== '') (h.secret ? secret : sent)[name] = String(value);
  }
  const origin = new URL(first).origin;
  const headersFor = (url) => (new URL(url).origin === origin ? { ...sent, ...secret } : sent);
  const all = [];
  let url = first;
  let pages = 0;
  let firstSize = null;
  let lastPage = null;
  const seen = new Set();
  const kind = paging?.kind || 'none';
  let n = Number(paging?.start ?? 1) || 0;
  if (kind === 'page') url = withPage(first, paging.param || 'page', n);
  while (url && pages < maxPages && !seen.has(url)) {
    seen.add(url);
    const res = await fetchFrom(url, { headers: headersFor(url), fetchImpl });
    let answer;
    try { answer = await res.json(); } catch { throw new Error(`${hostOf(url)} did not answer in JSON.`); }
    const list = recordsIn(answer, records);
    pages++;
    if (kind === 'page') {
      const same = JSON.stringify(list);
      if (!list.length || same === lastPage) break;
      lastPage = same;
    }
    all.push(...list);
    if (kind === 'next') {
      const next = paging.next ? valueAt(answer, paging.next) : linkNext(res);
      url = next ? new URL(String(next), url).toString() : null;
    } else if (kind === 'page') {
      firstSize ??= list.length;
      // A page shorter than the first is the last.
      url = list.length < firstSize ? null : withPage(first, paging.param || 'page', ++n);
    } else {
      url = null;
    }
  }
  if (!all.length) throw new Error('The answer has no records.');
  const { fields, rows } = tableFromRecords(all);
  return { fields, rows, pages };
}
