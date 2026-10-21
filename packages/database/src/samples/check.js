// What a sample database must be before it is offered: every table has
// records, every query answers something, every report has records and its
// totals add up to a number, nothing is below nought where a business would
// never see that, and every person in it is invented in the ways util.js
// says (e-mail addresses on example domains, phone numbers only from the
// ranges kept for fiction). Answers the problems found, in words; none is a
// sample ready to ship.

import { DOCUMENT_KINDS } from './index.js';

const MAIL = /^[^@\s]+@([a-z0-9-]+\.)*example(\.(com|org|net))?$/i;
const PHONE = [/^07700 900\d{3}$/, /^020 7946 0\d{3}$/, /^\(\d{3}\) 555-01\d{2}$/];

/** The problems in `base`, made from sample `spec` with its records. */
export function checkSample(base, spec) {
  const problems = [];
  const found = new Map();
  // One problem said once for each table or query and column, however many records have it.
  const say = (key, text) => { if (!found.has(key)) found.set(key, text); };
  const negativeOk = (where, column) => (spec.negative?.[where] || []).includes(column);
  const look = (where, column, v) => {
    if (typeof v === 'number') {
      if (!Number.isFinite(v)) say(`${where}|${column}|n`, `${where}: ${column} is ${v}.`);
      else if (v < 0 && !negativeOk(where, column)) say(`${where}|${column}|-`, `${where}: ${column} is ${v}, below nought.`);
    }
    if (typeof v === 'string') {
      if (/mail/i.test(column) && v && !MAIL.test(v)) say(`${where}|${column}|@`, `${where}: ${column} ${v} is not on an example domain.`);
      if (/phone|mobile/i.test(column) && v && !PHONE.some((p) => p.test(v))) say(`${where}|${column}|p`, `${where}: ${column} ${v} is not from a range kept for fiction.`);
    }
  };

  for (const t of base.tables()) {
    const rows = base.rows(t.name, { limit: 1_000_000 });
    if (!rows.length) problems.push(`The table ${t.name} has no records.`);
    for (const r of rows) for (const [k, v] of Object.entries(r)) if (k !== '_rowid') look(t.name, k, v);
  }
  problems.push(...found.values());
  found.clear();

  for (const { name } of base.queries()) {
    let out;
    try {
      out = base.runQuery({ name, limit: 1_000_000 });
    } catch (err) {
      problems.push(`The query ${name} does not run: ${err.message}`);
      continue;
    }
    if (!out.count) problems.push(`The query ${name} answers nothing.`);
    for (const c of out.columns) if (/^(Sum|Avg|Count|Min|Max) of /.test(c)) problems.push(`The query ${name}'s column ${c} wants a name of its own (as:).`);
    for (const r of out.rows) for (const [k, v] of Object.entries(r)) look(name, k, v);
    problems.push(...found.values());
    found.clear();
  }

  for (const { name, layout } of base.reports()) {
    let data;
    try {
      data = base.reportData({ name });
    } catch (err) {
      problems.push(`The report ${name} cannot be made: ${err.message}`);
      continue;
    }
    if (!data.count) problems.push(`The report ${name} has no records.`);
    const have = new Set(data.fields.map((f) => f.name));
    for (const c of layout.columns) if (!have.has(c.field)) problems.push(`The report ${name} shows ${c.field}, which its source does not have.`);
    if (layout.group && !have.has(layout.group)) problems.push(`The report ${name} is grouped by ${layout.group}, which its source does not have.`);
    if (layout.sort && !have.has(layout.sort)) problems.push(`The report ${name} is sorted by ${layout.sort}, which its source does not have.`);
    for (const t of layout.totals || []) {
      if (!have.has(t.field)) problems.push(`The report ${name} totals ${t.field}, which its source does not have.`);
      else if (data.count && data.totals[`${t.fn}:${t.field}`] == null) problems.push(`The report ${name}'s ${t.fn} of ${t.field} comes to nothing.`);
    }
  }

  // The documents it prints: each from a query that answers something, with
  // every field it places there, and a letter's [Field] places filled by its
  // query or by the Business record ([Business].[Name]).
  const business = new Set(base.tables().some((t) => t.name === 'Business') ? base.fields('Business').map((f) => f.name) : []);
  for (const d of spec.documents || []) {
    if (!DOCUMENT_KINDS.includes(d.kind)) problems.push(`The document ${d.name} is of a kind Data Files does not print: ${d.kind}.`);
    let out;
    try {
      out = base.runQuery({ name: d.query, limit: 1_000_000 });
    } catch (err) {
      problems.push(`The document ${d.name}'s query ${d.query} does not run: ${err.message}`);
      continue;
    }
    if (!out.count) problems.push(`The document ${d.name} has nothing to print: ${d.query} answers nothing.`);
    const have = new Set(out.columns);
    const named = [d.key, ...(d.header || []), ...(d.lines || []).map((c) => (typeof c === 'string' ? c : c.field)), ...(d.totals || []).map((t) => t.field)].filter(Boolean);
    for (const f of named) if (!have.has(f)) problems.push(`The document ${d.name} places ${f}, which ${d.query} does not have.`);
    if (['letter', 'agreement'].includes(d.kind)) {
      if (!d.body || d.body.length < 200) problems.push(`The ${d.kind} ${d.name} has no body of words.`);
      for (const m of String(d.body || '').matchAll(/\[([^\]]+)\](?:\.\[([^\]]+)\])?/g)) {
        const ok = m[2] ? m[1] === 'Business' && business.has(m[2]) : have.has(m[1]);
        if (!ok) problems.push(`The ${d.kind} ${d.name} has a place for ${m[0]}, which nothing fills.`);
      }
    } else if (!d.key || !(d.lines || []).length) {
      problems.push(`The document ${d.name} wants a key (one per document) and lines.`);
    }
  }
  return problems;
}
