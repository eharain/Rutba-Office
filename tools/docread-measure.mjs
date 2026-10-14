// Measures how well the fields of invoices and receipts are read, on a
// corpus from tools/docread-corpus.js:
//
//   node tools/docread-measure.mjs --corpus <dir> [--mode auto|sparse|document] [--threads 3] [--out results.json] [--only <id start>] [--misses]
//
// Each page is read through @rutba/ocr as the suite reads it, the page's
// language first, then English, as languagesFor gives for a window speaking
// it: with --mode document as recognize({ document: true }) reads it (as a
// page and as sparse text merged, its figures read again), with auto (the
// default) or sparse once in that way. What was read is kept beside the
// page (<id>.ocr.json, or <id>.ocr-<mode>.json), so a later
// run measures a change to the field reading without reading the pages
// again (--again reads them anew). Then documentFields (fields.js) runs on
// each and every field is compared with the page's truth:
//
//   amounts (subtotal, tax, total)   equal to the cent (to the yen for yen)
//   dates (date, dueDate)            the same day
//   number                           the same characters, case and spaces aside
//   supplier                         the same text, spaces and punctuation aside
//   kind                             invoice, receipt, letter or form
//   items                            the count of lines and each line's amount
//
// A field is right, wrong (a value that is not the truth) or missed (no
// value). The rates are given per field, per language, per kind of page and
// per condition, with the arithmetic check (how often a page whose sums
// agree is right, and how often a page whose sums do not agree is right
// anyway), the time a page took, and the tick boxes of the forms found and
// read (checkboxes.js).

import fs from 'node:fs';
import path from 'node:path';

const arg = (name, fallback) => {
  const at = process.argv.indexOf(`--${name}`);
  return at > 0 ? process.argv[at + 1] : fallback;
};

const UI = { eng: 'en', urd: 'ur', ara: 'ar', hin: 'hi', chi_sim: 'zh', jpn: 'ja' };

/** Read every page not read yet (or all, with --again), keeping what was read. */
async function readAll(corpus, pages, { again = false, threads = 3, mode = 'auto' } = {}) {
  const todo = pages.filter((p) => again || !fs.existsSync(path.join(corpus, cacheName(p.id, mode))));
  if (!todo.length) return;
  const { createOcrHost, languagesFor } = await import('@rutba/ocr');
  const host = createOcrHost({ threads });
  const started = Date.now();
  await Promise.all(todo.map(async (p) => {
    const r = await host.recognize(fs.readFileSync(path.join(corpus, p.picture)), { languages: languagesFor(UI[p.lang]), ...(mode === 'auto' ? {} : mode === 'document' ? { document: true } : { mode }) });
    const keep = { width: r.width, height: r.height, languages: r.languages, angle: r.angle, prepared: r.prepared, text: r.text, lines: r.lines, words: r.words, documentWords: r.document?.words, documentAgain: r.document?.again, document: r.document && { kind: r.document.kind, fields: r.document.fields, checks: r.document.checks }, ms: r.ms };
    fs.writeFileSync(path.join(corpus, cacheName(p.id, mode)), JSON.stringify(keep));
    console.log(`read ${p.id.padEnd(22)} ${String(r.words.length).padStart(4)} words  ${(r.ms / 1000).toFixed(1)} s`);
  }));
  await host.stop();
  console.log(`${todo.length} pages read in ${((Date.now() - started) / 1000).toFixed(0)} s`);
}

/** Where a page's reading is kept: one file a page and a way of reading it. */
const cacheName = (id, mode) => (mode === 'auto' ? `${id}.ocr.json` : `${id}.ocr-${mode}.json`);

const AMOUNTS = ['subtotal', 'tax', 'total'];
const DATES = ['date', 'dueDate'];
const plain = (s) => String(s ?? '').normalize('NFKC').toLowerCase().replace(/[\s.,:;'"()\-\u2013_#،]/g, '');

/** One field: 'right', 'wrong', 'missed', or null when the page has no such field. */
export function judge(key, truth, got) {
  if (truth === undefined || truth === null) return null;
  const value = got?.value;
  if (value === undefined || value === null || value === '') return 'missed';
  if (AMOUNTS.includes(key)) return Math.abs(Number(value) - truth) < 0.005 ? 'right' : 'wrong';
  if (DATES.includes(key)) return value === truth ? 'right' : 'wrong';
  return plain(value) === plain(truth) ? 'right' : 'wrong';
}

async function main() {
  const corpus = path.resolve(arg('corpus'));
  const only = arg('only', '');
  const pages = JSON.parse(fs.readFileSync(path.join(corpus, 'manifest.json'), 'utf8')).filter((p) => p.id.startsWith(only));
  const mode = arg('mode', 'auto');
  await readAll(corpus, pages, { again: process.argv.includes('--again'), threads: Number(arg('threads', 3)), mode });
  const { documentFields } = await import('@rutba/ocr/fields');
  const FIELDS = ['supplier', 'number', 'date', 'dueDate', 'subtotal', 'tax', 'total'];
  const rows = [];
  for (const p of pages) {
    const truth = JSON.parse(fs.readFileSync(path.join(corpus, `${p.id}.json`), 'utf8'));
    const read = JSON.parse(fs.readFileSync(path.join(corpus, cacheName(p.id, mode)), 'utf8'));
    // Read as a document: the fields from the words of both readings, as recognize({ document: true }) reads them.
    const got = documentFields(read.documentWords ? { ...read, words: read.documentWords } : read, { languages: read.languages, again: read.documentAgain });
    const row = { id: p.id, lang: p.lang, kind: p.kind, layout: p.layout, condition: p.condition, font: p.font, digits: p.digits, ms: read.ms, results: {}, got: {} };
    row.results.kind = got.kind === truth.kind ? 'right' : 'wrong';
    row.got.kind = got.kind;
    if (truth.kind === 'invoice' || truth.kind === 'receipt') {
      for (const k of FIELDS) {
        row.results[k] = judge(k, truth.fields[k], got.fields[k]);
        row.got[k] = got.fields[k]?.value ?? null;
      }
      const want = truth.items.map((it) => it.amount);
      const have = (got.items || []).map((it) => it.amount);
      row.results.items = want.length === have.length && want.every((a, i) => Math.abs(a - have[i]) < 0.005) ? 'right' : have.length ? 'wrong' : 'missed';
      row.checks = got.checks;
      row.truth = truth.fields;
    }
    rows.push(row);
  }

  const tally = (list, key) => {
    const t = { right: 0, wrong: 0, missed: 0 };
    for (const r of list) if (r.results[key]) t[r.results[key]]++;
    const n = t.right + t.wrong + t.missed;
    return { ...t, n, rate: n ? t.right / n : null };
  };
  const pct = (x) => (x === null ? '   - ' : `${(x * 100).toFixed(0).padStart(4)}%`);
  const KEYS = [...FIELDS, 'items', 'kind'];
  const table = (title, groups) => {
    console.log(`\n${title}`);
    console.log(`${''.padEnd(16)}${KEYS.map((k) => k.padStart(9)).join('')}`);
    for (const [name, list] of groups) console.log(`${name.padEnd(16)}${KEYS.map((k) => pct(tally(list, k).rate).padStart(9)).join('')}`);
  };
  const by = (key) => [...new Set(rows.map((r) => r[key]))].map((v) => [String(v), rows.filter((r) => r[key] === v)]);
  table('All pages', [['all', rows]]);
  table('By language', by('lang'));
  table('By kind', by('kind'));
  table('By layout', by('layout'));
  table('By condition', by('condition'));
  table('By digits', by('digits'));
  table('By font', by('font'));

  // The arithmetic as a signal: pages whose sums agreed, against pages whose three amounts were all right.
  const priced = rows.filter((r) => r.checks);
  const agreed = priced.filter((r) => r.checks.consistent);
  const allRight = (r) => AMOUNTS.every((k) => r.results[k] === 'right');
  console.log(`\nArithmetic: ${agreed.length} of ${priced.length} priced pages had sums that agree; ${agreed.filter(allRight).length} of those had all three amounts right; ${priced.filter((r) => !r.checks.consistent).filter(allRight).length} pages had all three right without sums that agree.`);
  // Time a page took to read (its own thread's time, waiting in the queue aside).
  const med = (xs) => { const s = [...xs].sort((a, b) => a - b); return s.length ? s[s.length >> 1] : 0; };
  const full = rows.filter((r) => r.kind !== 'receipt').map((r) => r.ms), rolls = rows.filter((r) => r.kind === 'receipt').map((r) => r.ms);
  console.log(`Reading time, median: an A4 page at 200 dpi ${(med(full) / 1000).toFixed(1)} s (most ${(Math.max(...full) / 1000).toFixed(1)} s), a till roll ${(med(rolls) / 1000).toFixed(1)} s.`);
  // Tick boxes on the forms: each drawn box found (its centre inside a box found), and read as ticked or not.
  const { findCheckboxes } = await import('@rutba/ocr/checkboxes');
  const { decodePng } = await import('@rutba/ocr/png');
  const { toGrey } = await import('@rutba/ocr/pixels');
  const tick = { boxes: 0, found: 0, right: 0, extra: 0, pages: 0 };
  for (const p of pages.filter((q) => q.kind === 'form')) {
    const truth = JSON.parse(fs.readFileSync(path.join(corpus, `${p.id}.json`), 'utf8'));
    const read = JSON.parse(fs.readFileSync(path.join(corpus, cacheName(p.id, mode)), 'utf8'));
    const got = findCheckboxes(toGrey(decodePng(fs.readFileSync(path.join(corpus, p.picture)))), { words: read.documentWords || read.words });
    const used = new Set();
    for (const c of truth.checkboxes) {
      tick.boxes++;
      const cx = c.box.x + c.box.width / 2, cy = c.box.y + c.box.height / 2;
      const i = got.findIndex((g, k) => !used.has(k) && cx >= g.box.x - 3 && cx <= g.box.x + g.box.width + 3 && cy >= g.box.y - 3 && cy <= g.box.y + g.box.height + 3);
      if (i < 0) { if (process.argv.includes('--misses')) console.log(`${p.id} tick box missed: ${c.label} ${JSON.stringify(c.box)}`); continue; }
      used.add(i);
      tick.found++;
      if (got[i].checked === c.checked) tick.right++;
      else if (process.argv.includes('--misses')) console.log(`${p.id} tick box ${c.label} read ${got[i].checked} (fill ${got[i].fill})`);
    }
    tick.extra += got.length - used.size;
    tick.pages++;
  }
  if (tick.pages) console.log(`Tick boxes: ${tick.found} of ${tick.boxes} found on ${tick.pages} forms, ${tick.right} of those read rightly as ticked or not; ${tick.extra} squares taken for boxes that are not.`);
  const wrongs = rows.flatMap((r) => Object.entries(r.results).filter(([, v]) => v !== 'right' && v).map(([k, v]) => `${r.id} ${k} ${v}: got ${JSON.stringify(r.got[k])} want ${JSON.stringify(k === 'kind' ? r.kind : r.truth?.[k])}`));
  if (process.argv.includes('--misses')) console.log(`\n${wrongs.join('\n')}`);
  const out = arg('out');
  if (out) fs.writeFileSync(out, `${JSON.stringify({ rows, totals: Object.fromEntries(KEYS.map((k) => [k, tally(rows, k)])) }, null, 2)}\n`);
}

if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(new URL(import.meta.url).pathname.replace(/^\/(\w:)/, '$1'))) await main();
