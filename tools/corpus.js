// The real-world corpus run: every file in a folder opened in the engines,
// written back out where the suite writes that format, opened again, and
// compared with what was read the first time.
//
//   npm run verify:corpus                        the starter corpus, tests/corpus-public
//   RUTBA_CORPUS=D:\corpus npm run verify:corpus another folder (each format in any sub-folder)
//   RUTBA_CORPUS_WINDOWS=1 ...                   also open each in the application's windows
//   RUTBA_CORPUS_RATES=rates.json ...            fail when a format opens or round-trips below its rate
//   RUTBA_CORPUS_TIMEOUT=60000 ...               ms one file may take before it counts as a hang
//   RUTBA_CORPUS_LIMIT=50 ...                    the first N files only
//   RUTBA_CORPUS_OUT=<dir> ...                   where results.json and summary.md go
//                                                (apps/desktop/build/corpus by default)
//
// The formats are the ones a person brings to an office suite: Word, Excel and
// PowerPoint in both their XML and their 97-2003 forms, OpenDocument, PDF, a
// mail message, an Outlook data file, a calendar and a contact file. Each file
// is handled in a worker of its own thread, which is stopped when the file
// takes too long, so one file that hangs is one result and not the end of the
// run. Per file the run records whether it opened, whether what was read
// survived a save and a second open, what the engines warned about, any fault
// (an error a parser threw rather than chose, in fuzz-open.js's sense), and
// how long each step took. docs/COMPATIBILITY.md describes results.json.
//
// What "round-tripped equal" means, per family:
//   documents     the text of every paragraph of the body, in order, and the
//                 number of tables and drawings;
//   workbooks     the sheet names in order, and every cell's value, or its
//                 formula where it has one;
//   decks         the slides in order, the text of each, and its pictures;
//   PDF           the pages, and the text drawn on each;
//   mail          subject, sender, recipients, the text, and each attachment's
//                 name and size;
//   calendar      each event's UID, summary, start, end, all-day flag and rule;
//   contacts      each card's name, organisation, e-mail and telephone.
// It is the suite's own reading compared with itself after a save: a file the
// engines misread the same way twice compares equal. That is what this run
// measures (does saving lose what was read?); how faithfully a file is read is
// what the window checks and the fidelity tests measure.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { Worker, isMainThread, parentPort } from 'node:worker_threads';
import { readZip } from '@rutba/ooxml/zip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

/** The schema results.json is written in; bumped when a field changes meaning. */
export const SCHEMA = 'rutba-office-corpus/1';

/**
 * Every format the run knows: the family that reads it, and the format a
 * round trip writes. `writes` is the format itself where the suite saves it,
 * the XML format a 97-2003 file is saved as (what Save does with one), and
 * null where the suite reads a format and never writes it.
 */
export const FORMATS = {
  docx: { family: 'office', kind: 'doc', writes: 'docx' },
  xlsx: { family: 'office', kind: 'sheet', writes: 'xlsx' },
  pptx: { family: 'office', kind: 'deck', writes: 'pptx' },
  odt: { family: 'office', kind: 'doc', writes: 'odt' },
  ods: { family: 'office', kind: 'sheet', writes: 'ods' },
  odp: { family: 'office', kind: 'deck', writes: 'odp' },
  doc: { family: 'office', kind: 'doc', writes: 'docx' },
  xls: { family: 'office', kind: 'sheet', writes: 'xlsx' },
  ppt: { family: 'office', kind: 'deck', writes: 'pptx' },
  pdf: { family: 'pdf', writes: 'pdf' },
  eml: { family: 'mail', writes: 'eml' },
  pst: { family: 'pst', writes: null },
  ics: { family: 'calendar', writes: 'ics' },
  vcf: { family: 'contacts', writes: 'vcf' },
};

const OOXML_EXT = { doc: 'docx', sheet: 'xlsx', deck: 'pptx' };

/** The corpus format of a file, by its extension, or null. */
export function formatOf(file) {
  const ext = path.extname(String(file)).slice(1).toLowerCase();
  return Object.hasOwn(FORMATS, ext) ? ext : null;
}

const SKIP_DIRS = new Set(['node_modules', '.git']);

/** Every file under `dir` the run knows the format of, in a stable order. */
export function collectCorpus(dir) {
  const out = [];
  const walk = (at) => {
    let entries = [];
    try {
      entries = fs.readdirSync(at, { withFileTypes: true });
    } catch {
      return;
    }
    for (const e of entries) {
      const full = path.join(at, e.name);
      if (e.isDirectory()) {
        if (!SKIP_DIRS.has(e.name)) walk(full);
      } else if (e.isFile()) {
        const format = formatOf(e.name);
        // Word's owner file beside an open document (~$name.docx) is not a document.
        if (format && !e.name.startsWith('~$')) out.push({ file: full, rel: path.relative(dir, full).split(path.sep).join('/'), format });
      }
    }
  };
  walk(dir);
  return out.sort((a, b) => a.rel.localeCompare(b.rel));
}

/* ── what a file holds, read the same way before and after ─────────────── */

const unesc = (s) =>
  String(s)
    .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');

/** Whitespace as a reader sees it: runs of it one space, none at the ends. */
const squash = (s) => String(s ?? '').replace(/\s+/g, ' ').trim();

/** A fallback a newer element carries for older readers repeats its text; it is not counted twice. */
const withoutFallbacks = (xml) => xml.replace(/<mc:Fallback\b[\s\S]*?<\/mc:Fallback>/g, '');

function partsOf(bytes) {
  const parts = new Map();
  for (const entry of readZip(Buffer.from(bytes)).entries) parts.set(entry.name.replace(/^\//, ''), entry);
  return { text: (name) => (parts.has(name) ? Buffer.from(parts.get(name).data).toString('utf8') : null), has: (name) => parts.has(name) };
}

/** A part's relationships, id to resolved part name. */
function relsOf(parts, partName) {
  const dir = path.posix.dirname(partName);
  const xml = parts.text(`${dir}/_rels/${path.posix.basename(partName)}.rels`) || '';
  const out = new Map();
  for (const m of xml.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    const id = /\bId="([^"]+)"/.exec(m[1])?.[1];
    const target = /\bTarget="([^"]+)"/.exec(m[1])?.[1];
    if (!id || !target || /TargetMode="External"/.test(m[1])) continue;
    out.set(id, target.startsWith('/') ? target.slice(1) : path.posix.normalize(`${dir}/${unesc(target)}`));
  }
  return out;
}

const textRuns = (xml, tag) => [...xml.matchAll(new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)</${tag}>`, 'g'))].map((m) => unesc(m[1])).join('');

/** A number as a value, whatever digits the writer chose to keep. */
function numberText(v) {
  const n = Number(v);
  if (v === '' || !Number.isFinite(n)) return String(v);
  return String(Number(n.toPrecision(15)));
}

/** A formula as written, less the prefixes newer functions carry in the file and the spaces. */
const formulaText = (f) => unesc(f).replace(/_xlfn\.|_xlws\./g, '').replace(/\s+/g, '');

/**
 * What a Word, Excel or PowerPoint XML file holds, as compared before and
 * after a round trip (see the header of this file).
 */
export function fingerprintOoxml(kind, bytes) {
  const parts = partsOf(bytes);
  if (kind === 'doc') {
    const xml = withoutFallbacks(parts.text('word/document.xml') || '');
    const body = /<w:body\b[\s\S]*<\/w:body>/.exec(xml)?.[0] || '';
    const paragraphs = [];
    for (const m of body.matchAll(/<w:p\b[^>]*?(?:\/>|>([\s\S]*?)<\/w:p>)/g)) {
      const text = squash(textRuns(m[1] || '', 'w:t'));
      if (text) paragraphs.push(text);
    }
    return {
      kind,
      paragraphs,
      tables: (body.match(/<w:tbl>|<w:tbl\s/g) || []).length,
      drawings: (body.match(/<w:drawing>|<w:drawing\s/g) || []).length,
    };
  }
  if (kind === 'sheet') {
    const workbook = parts.text('xl/workbook.xml') || '';
    const rels = relsOf(parts, 'xl/workbook.xml');
    const shared = [];
    for (const m of (parts.text('xl/sharedStrings.xml') || '').matchAll(/<si>([\s\S]*?)<\/si>/g)) {
      // Phonetic runs are a reading aid, not the cell's text.
      shared.push(textRuns(m[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g, ''), 't'));
    }
    const sheets = [];
    for (const m of workbook.matchAll(/<sheet\b([^>]*)\/?>/g)) {
      const name = unesc(/\bname="([^"]*)"/.exec(m[1])?.[1] || '');
      const rid = /\br:id="([^"]+)"/.exec(m[1])?.[1];
      const xml = (rid && parts.text(rels.get(rid))) || '';
      const cells = {};
      for (const c of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const ref = /\br="([A-Z]+\d+)"/.exec(c[1])?.[1];
        const inner = c[2] || '';
        if (!ref || !inner) continue;
        const type = /\bt="([^"]+)"/.exec(c[1])?.[1] || 'n';
        const f = /<f\b[^>]*>([\s\S]*?)<\/f>/.exec(inner)?.[1];
        const v = /<v>([\s\S]*?)<\/v>/.exec(inner)?.[1];
        let value;
        if (f) value = `=${formulaText(f)}`;
        else if (type === 's') value = v != null ? shared[Number(v)] ?? '' : '';
        else if (type === 'inlineStr') value = textRuns(inner, 't');
        else if (type === 'str' || type === 'e') value = v != null ? unesc(v) : '';
        else if (type === 'b') value = v === '1' ? 'TRUE' : 'FALSE';
        else value = v != null ? numberText(v) : '';
        if (value !== '') cells[ref] = value;
      }
      sheets.push({ name, cells });
    }
    return { kind, sheets };
  }
  if (kind === 'deck') {
    const presentation = parts.text('ppt/presentation.xml') || '';
    const rels = relsOf(parts, 'ppt/presentation.xml');
    const slides = [];
    for (const m of presentation.matchAll(/<p:sldId\b([^>]*)\/?>/g)) {
      const rid = /\br:id="([^"]+)"/.exec(m[1])?.[1];
      const xml = withoutFallbacks((rid && parts.text(rels.get(rid))) || '');
      const texts = [];
      for (const p of xml.matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)) {
        const t = squash(textRuns(p[1], 'a:t'));
        if (t) texts.push(t);
      }
      slides.push({ texts, pictures: (xml.match(/<p:pic>|<p:pic\s/g) || []).length });
    }
    return { kind, slides };
  }
  throw new Error(`no fingerprint for ${kind}`);
}

/**
 * Where two readings of a file part company, as sentences: at most `limit`
 * of them, the first ones in reading order. An empty list is equal.
 */
export function compareFingerprints(a, b, limit = 10) {
  const out = [];
  const say = (s) => out.length < limit && out.push(s);
  const short = (v) => {
    // Quoted, so an empty value or a space at the end can be seen.
    const s = JSON.stringify(v) ?? 'nothing';
    return s.length > 80 ? `${s.slice(0, 77)}...` : s;
  };
  const list = (what, x = [], y = [], each) => {
    if (x.length !== y.length) say(`${what}: ${x.length} before, ${y.length} after`);
    for (let i = 0; i < Math.min(x.length, y.length); i++) each(x[i], y[i], i);
  };
  const same = (what, x, y) => {
    if (JSON.stringify(x) !== JSON.stringify(y)) say(`${what}: ${short(x)} before, ${short(y)} after`);
  };
  if (a?.kind !== b?.kind) {
    say(`kind: ${a?.kind} before, ${b?.kind} after`);
    return out;
  }
  switch (a.kind) {
    case 'doc':
      list('paragraphs', a.paragraphs, b.paragraphs, (x, y, i) => same(`paragraph ${i + 1}`, x, y));
      same('tables', a.tables, b.tables);
      same('drawings', a.drawings, b.drawings);
      break;
    case 'sheet':
      list('sheets', a.sheets, b.sheets, (x, y, i) => {
        same(`sheet ${i + 1} name`, x.name, y.name);
        const refs = new Set([...Object.keys(x.cells), ...Object.keys(y.cells)]);
        for (const ref of refs) same(`${x.name}!${ref}`, x.cells[ref] ?? null, y.cells[ref] ?? null);
      });
      break;
    case 'deck':
      list('slides', a.slides, b.slides, (x, y, i) => {
        list(`slide ${i + 1} paragraphs`, x.texts, y.texts, (p, q, j) => same(`slide ${i + 1} paragraph ${j + 1}`, p, q));
        same(`slide ${i + 1} pictures`, x.pictures, y.pictures);
      });
      break;
    case 'pdf':
      list('pages', a.pages, b.pages, (x, y, i) => same(`page ${i + 1} text`, x, y));
      break;
    case 'mail':
      for (const k of ['subject', 'from', 'to', 'cc', 'text', 'attachments']) same(k, a[k], b[k]);
      break;
    case 'calendar':
      list('events', a.events, b.events, (x, y, i) => same(`event ${i + 1}`, x, y));
      break;
    case 'contacts':
      list('cards', a.cards, b.cards, (x, y, i) => same(`card ${i + 1}`, x, y));
      break;
    default:
      same('content', a, b);
  }
  return out;
}

/** How much a reading held, for the results: a round trip of nothing is not a proof of anything. */
export function factsOf(print) {
  switch (print?.kind) {
    case 'doc':
      return { paragraphs: print.paragraphs.length, tables: print.tables, drawings: print.drawings };
    case 'sheet':
      return { sheets: print.sheets.length, cells: print.sheets.reduce((n, s) => n + Object.keys(s.cells).length, 0) };
    case 'deck':
      return { slides: print.slides.length, paragraphs: print.slides.reduce((n, s) => n + s.texts.length, 0), pictures: print.slides.reduce((n, s) => n + s.pictures, 0) };
    case 'pdf':
      return { pages: print.pages.length, characters: print.pages.reduce((n, p) => n + p.length, 0) };
    case 'mail':
      return { recipients: print.to.length + print.cc.length, characters: print.text.length, attachments: print.attachments.length };
    case 'calendar':
      return { events: print.events.length };
    case 'contacts':
      return { cards: print.cards.length };
    default:
      return null;
  }
}

/* ── the numbers ────────────────────────────────────────────────────────── */

const percentile = (values, p) => {
  if (!values.length) return 0;
  const sorted = [...values].sort((x, y) => x - y);
  return sorted[Math.min(sorted.length - 1, Math.floor((sorted.length - 1) * p))];
};
const rate = (n, d) => (d ? Math.round((n / d) * 10000) / 10000 : null);

/**
 * The summary per format, from the per-file results.
 *
 * A file protected by a password is counted apart and left out of the open
 * rate: asking for the password is the right answer, and the corpus does not
 * carry passwords. A round trip is counted only where one was tried.
 */
export function summarise(files) {
  const formats = {};
  for (const r of files) {
    const s = (formats[r.format] ??= {
      files: 0, opened: 0, locked: 0, refused: 0, faults: 0, timeouts: 0,
      roundTrips: 0, roundTripEqual: 0, withWarnings: 0, windowShown: 0, windowNotShown: 0, openRate: null, roundTripRate: null,
      ms: { median: 0, p90: 0, max: 0 }, times: [],
    });
    s.files += 1;
    if (r.status === 'opened') s.opened += 1;
    else if (r.status === 'locked') s.locked += 1;
    else if (r.status === 'refused') s.refused += 1;
    else if (r.status === 'timeout') s.timeouts += 1;
    if (r.faults?.length) s.faults += 1;
    if (r.roundTrip && r.roundTrip.equal != null) {
      s.roundTrips += 1;
      if (r.roundTrip.equal) s.roundTripEqual += 1;
    }
    if (r.warnings?.length) s.withWarnings += 1;
    // The window run, where there was one: shown, or refused, hung or crashed there. Skipped and locked are neither.
    if (r.window?.status === 'ok') s.windowShown += 1;
    else if (['refused', 'timeout', 'hung', 'crashed'].includes(r.window?.status)) s.windowNotShown += 1;
    if (r.status === 'opened') s.times.push(r.ms.total);
  }
  for (const s of Object.values(formats)) {
    s.openRate = rate(s.opened, s.files - s.locked);
    s.roundTripRate = rate(s.roundTripEqual, s.roundTrips);
    s.ms = { median: percentile(s.times, 0.5), p90: percentile(s.times, 0.9), max: percentile(s.times, 1) };
    delete s.times;
  }
  return Object.fromEntries(Object.keys(formats).sort().map((k) => [k, formats[k]]));
}

/**
 * The rates a format falls short of, as sentences. `rates` is
 * { "<format>": { "open": 0.95, "roundTrip": 0.9 } }, with "*" for every
 * format not named.
 */
export function checkRates(formats, rates = {}) {
  const out = [];
  for (const [format, s] of Object.entries(formats)) {
    const want = rates[format] || rates['*'];
    if (!want) continue;
    if (want.open != null && s.openRate != null && s.openRate < want.open) out.push(`${format}: opened ${s.openRate} of its files, below the agreed ${want.open}`);
    if (want.roundTrip != null && s.roundTripRate != null && s.roundTripRate < want.roundTrip) out.push(`${format}: round-tripped ${s.roundTripRate} equal, below the agreed ${want.roundTrip}`);
  }
  return out;
}

/** The per-format table, as Markdown, for a person and for the site. */
export function summaryMarkdown(results) {
  const lines = [
    `# Corpus run: ${results.corpus.files} files`,
    '',
    `Rutba Office ${results.suite} (${results.commit || 'no commit'}), ${results.platform}, ${results.ranAt}.`,
    `Corpus: ${results.corpus.root}.`,
    '',
    '| Format | Files | Opened | Password | Refused | Faults | Hangs | Round trips equal | Open rate | Median ms | p90 ms |',
    '|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|',
  ];
  for (const [format, s] of Object.entries(results.formats)) {
    const pct = (r) => (r == null ? 'n/a' : `${(r * 100).toFixed(1)}%`);
    lines.push(`| ${format} | ${s.files} | ${s.opened} | ${s.locked} | ${s.refused} | ${s.faults} | ${s.timeouts} | ${s.roundTrips ? `${s.roundTripEqual}/${s.roundTrips}` : 'n/a'} | ${pct(s.openRate)} | ${s.ms.median} | ${s.ms.p90} |`);
  }
  const notable = results.files.filter((r) => r.status !== 'opened' && r.status !== 'locked' || r.faults.length || r.roundTrip?.equal === false);
  if (notable.length) {
    lines.push('', '## Files to look at', '');
    for (const r of notable) {
      const why = r.faults[0] || r.refused || (r.roundTrip?.equal === false ? `round trip: ${r.roundTrip.differences[0] || r.roundTrip.error}` : r.status);
      lines.push(`- ${r.file} (${r.status}): ${why}`);
    }
  }
  if (results.rateFailures?.length) lines.push('', '## Below the agreed rates', '', ...results.rateFailures.map((s) => `- ${s}`));
  return lines.join('\n') + '\n';
}

/* ── one file, in the worker ────────────────────────────────────────────── */

/** Errors a parser throws without meaning to (fuzz-open.js's judgement, copied as fuzz-pdf.js copies it). */
const FAULT_NAMES = new Set(['RangeError', 'TypeError', 'ReferenceError', 'EvalError', 'InternalError', 'AggregateError', 'SyntaxError']);
const FAULT_WORDS = /Maximum call stack|Invalid array length|Array buffer allocation|out of memory|outside the bounds|Offset is outside|is not a function|is not iterable|Cannot read propert|Cannot set propert|undefined is not|null is not/i;

/** A thrown error: a refusal in a sentence, or a fault. */
export function judge(err) {
  const links = [];
  for (let e = err, n = 0; e && n < 8; e = e.cause, n++) links.push({ name: e?.name || typeof e, message: String(e?.message ?? e) });
  const fault = links.find((l) => FAULT_NAMES.has(l.name) || FAULT_WORDS.test(l.message));
  const top = links[0]?.message || '';
  const sentence = /[a-z]{3}.*[ .]/.test(top);
  return fault || !sentence
    ? { outcome: 'fault', message: `${(fault || links[0]).name}: ${(fault || links[0]).message}`.slice(0, 300) }
    : { outcome: 'refused', message: top.slice(0, 300) };
}

let services = null;
async function load() {
  if (services) return services;
  const { createDocumentService } = await import('../apps/desktop/main/documents.js');
  const pdfReader = (await import('@rutba/pdf/reader')).default;
  const pdfAssemble = (await import('@rutba/pdf/assemble')).default;
  const { parseMessage } = await import('@rutba/mailbox/mime');
  const { toRfc822 } = await import('@rutba/mailbox/rfc822');
  const { PstFile } = await import('@rutba/mailbox/pst');
  const { fileReader } = await import('@rutba/mailbox/import');
  const { readCalendar, writeCalendar } = await import('@rutba/calendar/ical');
  const { readVCards, writeVCards } = await import('@rutba/contacts/vcard');
  services = {
    doc: createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) }),
    pdfReader, pdfAssemble, parseMessage, toRfc822, PstFile, fileReader, readCalendar, writeCalendar, readVCards, writeVCards,
  };
  return services;
}

const now = () => Number(process.hrtime.bigint() / 1000000n);

/** What each family reads, and how a round trip writes it back. */
const FAMILIES = {
  async office(s, file, format, tmp, result) {
    const { kind, writes } = FORMATS[format];
    const ooxml = OOXML_EXT[kind];
    const opened = s.doc.open({ path: file });
    if (opened.locked) return 'locked';
    // Not just opened: drawn as the window would draw it.
    s.doc.model({ id: opened.id });
    if (opened.converted) result.converted = opened.converted.format || opened.converted.from;
    result.ms.open = now() - result.started;

    const started = now();
    result.roundTrip = { target: writes, equal: null, differences: [] };
    try {
      // Before: the file itself where it is the XML format, else the suite's first reading of it.
      let before;
      if (format === ooxml) before = fingerprintOoxml(kind, fs.readFileSync(file));
      else {
        const first = path.join(tmp, `first.${ooxml}`);
        s.doc.save({ id: opened.id, path: first });
        before = fingerprintOoxml(kind, fs.readFileSync(first));
      }
      result.facts = factsOf(before);
      const written = path.join(tmp, `round-trip.${writes}`);
      if (writes === ooxml) s.doc.save({ id: opened.id, path: written });
      else s.doc.export({ id: opened.id, format: writes, path: written });
      s.doc.close({ id: opened.id });

      const again = s.doc.open({ path: written });
      s.doc.model({ id: again.id });
      const second = path.join(tmp, `second.${ooxml}`);
      s.doc.save({ id: again.id, path: second });
      s.doc.close({ id: again.id });
      const after = fingerprintOoxml(kind, fs.readFileSync(second));
      result.roundTrip.differences = compareFingerprints(before, after);
      result.roundTrip.equal = result.roundTrip.differences.length === 0;
    } catch (err) {
      roundTripFailed(result, err);
    }
    result.ms.roundTrip = now() - started;
    return 'opened';
  },

  async pdf(s, file, format, tmp, result) {
    const { openPdf } = s.pdfReader;
    const bytes = fs.readFileSync(file);
    const pdf = openPdf(bytes);
    const read = (p) => ({ kind: 'pdf', pages: Array.from({ length: p.pageCount }, (_, i) => squash(p.pageText(i))) });
    const before = read(pdf);
    result.facts = factsOf(before);
    result.ms.open = now() - result.started;
    if (pdf.encrypted) {
      result.warnings.push('protected with a password: the pages were counted, the text and a round trip were not tried');
      return 'opened';
    }
    const started = now();
    result.roundTrip = { target: 'pdf', equal: null, differences: [] };
    try {
      // Save as the PDF app saves a changed file: every page assembled into a new one.
      const out = s.pdfAssemble.assemble(before.pages.map((_, page) => ({ file: pdf, page })));
      const after = read(openPdf(out));
      result.roundTrip.differences = compareFingerprints(before, after);
      result.roundTrip.equal = result.roundTrip.differences.length === 0;
    } catch (err) {
      roundTripFailed(result, err);
    }
    result.ms.roundTrip = now() - started;
    return 'opened';
  },

  async mail(s, file, format, tmp, result) {
    const read = (bytes) => {
      const m = s.parseMessage(bytes);
      const people = (list) => (list || []).map((p) => (p.address || p.name || '').toLowerCase()).sort();
      return {
        message: m,
        print: {
          kind: 'mail',
          subject: squash(m.subject),
          from: people(m.from),
          to: people(m.to),
          cc: people(m.cc),
          text: squash(m.text),
          attachments: (m.attachments || []).map((a) => `${a.filename}:${a.size}`).sort(),
        },
      };
    };
    const before = read(fs.readFileSync(file));
    result.facts = factsOf(before.print);
    if (!before.message.headers?.length) throw new Error('This file has no message headers, so it is not a mail message.');
    result.ms.open = now() - result.started;
    if (before.message.smimeRaw) result.warnings.push('signed or encrypted: the round trip writes the parts, so the signature does not survive it');
    const started = now();
    result.roundTrip = { target: 'eml', equal: null, differences: [] };
    try {
      const after = read(s.toRfc822(before.message));
      result.roundTrip.differences = compareFingerprints(before.print, after.print);
      result.roundTrip.equal = result.roundTrip.differences.length === 0;
    } catch (err) {
      roundTripFailed(result, err);
    }
    result.ms.roundTrip = now() - started;
    return 'opened';
  },

  async pst(s, file, format, tmp, result) {
    // Read from the disk, not into memory: a real archive runs to gigabytes.
    const pst = s.PstFile.open(s.fileReader(file));
    try {
      const info = pst.info;
      if (info.readable === false) throw new Error(`This Outlook data file is protected (${info.encoding}), so its messages cannot be read.`);
      // Every folder, the hidden ones too: a message in one is still in the file.
      const folders = pst.folders({ includeHidden: true, force: true }).list;
      // Every message's headers, and the first two hundred read whole, as the Mail app reads one opened.
      let messages = 0;
      let readWhole = 0;
      for (const folder of folders) {
        const nids = pst.messageNids(folder.nid);
        messages += nids.length;
        for (const nid of nids) {
          if (readWhole >= 200) break;
          pst.message(nid, { withBytes: false });
          readWhole += 1;
        }
      }
      result.facts = { kind: info.kind, folders: folders.length, messages, readWhole };
      if (!messages) result.warnings.push('no messages found in any folder');
    } finally {
      pst.close?.();
    }
    result.ms.open = now() - result.started;
    return 'opened';
  },

  async calendar(s, file, format, tmp, result) {
    const at = (dt) => (dt ? (dt.allDay ? `${dt.at}d` : String(dt.at)) : null);
    const read = (text) => {
      const cal = s.readCalendar(text);
      return {
        cal,
        print: {
          kind: 'calendar',
          events: cal.events.map((e) => ({ uid: e.uid, summary: squash(e.summary), start: at(e.start), end: at(e.end), allDay: e.allDay, rrule: e.rrule ? JSON.stringify(e.rrule) : null })),
        },
      };
    };
    const text = fs.readFileSync(file, 'utf8');
    if (!/BEGIN:VCALENDAR/i.test(text)) throw new Error('This file has no BEGIN:VCALENDAR, so it is not a calendar.');
    const before = read(text);
    result.facts = factsOf(before.print);
    result.ms.open = now() - result.started;
    if (!before.cal.events.length) result.warnings.push('no events found');
    const started = now();
    result.roundTrip = { target: 'ics', equal: null, differences: [] };
    try {
      const after = read(s.writeCalendar({ events: before.cal.events, name: before.cal.name }));
      result.roundTrip.differences = compareFingerprints(before.print, after.print);
      result.roundTrip.equal = result.roundTrip.differences.length === 0;
    } catch (err) {
      roundTripFailed(result, err);
    }
    result.ms.roundTrip = now() - started;
    return 'opened';
  },

  async contacts(s, file, format, tmp, result) {
    const read = (text) => {
      const cards = s.readVCards(text);
      return {
        cards,
        print: {
          kind: 'contacts',
          cards: cards.map((c) => ({
            name: squash(c.name?.full),
            org: c.org || null,
            emails: (c.emails || []).map((e) => e.value.toLowerCase()).sort(),
            phones: (c.phones || []).map((p) => p.value.replace(/[^\d+]/g, '')).sort(),
          })),
        },
      };
    };
    const text = fs.readFileSync(file, 'utf8');
    if (!/BEGIN:VCARD/i.test(text)) throw new Error('This file has no BEGIN:VCARD, so it is not a contact file.');
    const before = read(text);
    result.facts = factsOf(before.print);
    result.ms.open = now() - result.started;
    if (!before.cards.length) result.warnings.push('no cards found');
    const started = now();
    result.roundTrip = { target: 'vcf', equal: null, differences: [] };
    try {
      const after = read(s.writeVCards(before.cards));
      result.roundTrip.differences = compareFingerprints(before.print, after.print);
      result.roundTrip.equal = result.roundTrip.differences.length === 0;
    } catch (err) {
      roundTripFailed(result, err);
    }
    result.ms.roundTrip = now() - started;
    return 'opened';
  },
};

/** A round trip that threw: a fault when the engine broke, else the sentence it gave. */
function roundTripFailed(result, err) {
  const j = judge(err);
  result.roundTrip.equal = false;
  result.roundTrip.error = j.message;
  if (j.outcome === 'fault') result.faults.push(`round trip: ${j.message}`);
  else result.warnings.push(`round trip refused: ${j.message}`);
}

/** Open one file and record what happened. Runs in the worker. */
export async function examine(file, format) {
  const s = await load();
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-corpus-file-'));
  const result = { status: 'opened', converted: null, roundTrip: null, warnings: [], faults: [], refused: null, facts: null, ms: { open: 0, roundTrip: 0, total: 0 }, started: now() };
  // What the engines say on the console while they work on this file is a warning about it.
  const said = [];
  const keep = { warn: console.warn, error: console.error };
  console.warn = (...args) => said.push(args.map(String).join(' ').slice(0, 300));
  console.error = console.warn;
  try {
    result.status = await FAMILIES[FORMATS[format].family](s, file, format, tmp, result);
  } catch (err) {
    const j = judge(err);
    if (j.outcome === 'fault') {
      result.status = 'fault';
      result.faults.push(j.message);
    } else {
      result.status = 'refused';
      result.refused = j.message;
    }
  } finally {
    console.warn = keep.warn;
    console.error = keep.error;
    fs.rmSync(tmp, { recursive: true, force: true });
  }
  for (const line of said) if (result.warnings.length < 20) result.warnings.push(line);
  result.ms.total = now() - result.started;
  if (result.status === 'opened' && !result.ms.open) result.ms.open = result.ms.total;
  delete result.started;
  return result;
}

if (!isMainThread) {
  parentPort.on('message', async ({ file, format }) => {
    try {
      parentPort.postMessage(await examine(file, format));
    } catch (err) {
      parentPort.postMessage({ status: 'fault', faults: [`the run itself failed: ${err?.message || err}`], warnings: [], roundTrip: null, refused: null, ms: { open: 0, roundTrip: 0, total: 0 } });
    }
  });
}

/* ── the run ────────────────────────────────────────────────────────────── */

/** The commit the suite is at, or null outside a checkout. */
function commitOf() {
  const r = spawnSync('git', ['rev-parse', '--short', 'HEAD'], { cwd: root, encoding: 'utf8' });
  return r.status === 0 ? r.stdout.trim() : null;
}

/**
 * Run the corpus: every file through the engines, each in the worker with a
 * deadline, and, with `windows`, through the application's windows as well.
 * Answers the results object (the shape docs/COMPATIBILITY.md describes) and
 * writes results.json and summary.md into `out`.
 */
export async function runCorpus({ corpus, out, timeout = 60000, limit = 0, rates = null, windows = false, log = console.log } = {}) {
  let files = collectCorpus(corpus);
  if (limit > 0) files = files.slice(0, limit);
  fs.mkdirSync(out, { recursive: true });
  log(`corpus: ${files.length} files under ${corpus}, ${timeout} ms each at most`);

  let worker = null;
  const fresh = () => {
    worker = new Worker(fileURLToPath(import.meta.url));
    worker.unref();
  };
  const ask = (file, format) =>
    new Promise((resolve) => {
      const timer = setTimeout(() => {
        worker.removeAllListeners('message');
        worker.removeAllListeners('error');
        worker.terminate();
        fresh();
        resolve({ status: 'timeout', faults: [`took longer than ${timeout / 1000} s`], warnings: [], roundTrip: null, refused: null, ms: { open: 0, roundTrip: 0, total: timeout } });
      }, timeout);
      worker.once('message', (answer) => {
        clearTimeout(timer);
        worker.removeAllListeners('error');
        resolve(answer);
      });
      worker.once('error', (err) => {
        clearTimeout(timer);
        worker.removeAllListeners('message');
        fresh();
        resolve({ status: 'fault', faults: [`the worker died: ${err?.message || err}`.slice(0, 300)], warnings: [], roundTrip: null, refused: null, ms: { open: 0, roundTrip: 0, total: 0 } });
      });
      worker.postMessage({ file, format });
    });
  fresh();

  const results = [];
  let n = 0;
  for (const { file, rel, format } of files) {
    n += 1;
    const bytes = fs.readFileSync(file);
    const r = await ask(file, format);
    const entry = {
      file: rel,
      format,
      bytes: bytes.length,
      sha256: crypto.createHash('sha256').update(bytes).digest('hex'),
      status: r.status,
      opened: r.status === 'opened',
      converted: r.converted ?? null,
      roundTrip: r.roundTrip ?? null,
      warnings: r.warnings || [],
      faults: r.faults || [],
      refused: r.refused ?? null,
      facts: r.facts ?? null,
      ms: r.ms,
      window: null,
    };
    results.push(entry);
    const rt = entry.roundTrip ? (entry.roundTrip.equal ? 'equal' : 'DIFFERS') : '-';
    log(`${String(n).padStart(4)}/${files.length} ${entry.status.padEnd(8)} rt ${rt.padEnd(7)} ${String(entry.ms.total).padStart(6)} ms  ${rel}${entry.faults.length ? `  ${entry.faults[0]}` : entry.refused ? `  ${entry.refused}` : ''}`);
  }
  await worker.terminate();

  if (windows) {
    runWindows(files, out, log);
    mergeWindows(results, files, out);
  }

  const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  const formats = summarise(results);
  const report = {
    schema: SCHEMA,
    suite: pkg.version,
    commit: commitOf(),
    ranAt: new Date().toISOString(),
    platform: `${process.platform}-${process.arch}`,
    node: process.version,
    corpus: { root: path.relative(root, corpus).split(path.sep).join('/') || corpus, files: results.length },
    windows: Boolean(windows),
    formats,
    files: results,
    rateFailures: rates ? checkRates(formats, rates) : [],
  };
  fs.writeFileSync(path.join(out, 'results.json'), JSON.stringify(report, null, 2) + '\n');
  fs.writeFileSync(path.join(out, 'summary.md'), summaryMarkdown(report));
  return report;
}

/** The window run (apps/desktop/build/verify-corpus-run.js) over the files a window opens. */
function runWindows(files, out, log) {
  const listFile = path.join(out, 'window-files.txt');
  fs.writeFileSync(listFile, files.map((f) => f.file).join('\n') + '\n');
  const desktop = path.join(root, 'apps', 'desktop');
  log('windows: building the bundle and opening each file in the application');
  const env = { ...process.env, RUTBA_CORPUS_LIST: listFile, RUTBA_CORPUS_OUT: path.join(out, 'windows') };
  for (const script of ['build/bundle.js', 'build/verify-corpus-run.js']) {
    const r = spawnSync(process.execPath, [script], { cwd: desktop, env, stdio: 'inherit' });
    // The window run's own verdict is its report; a refusal there is a result, not a stop.
    if (script === 'build/bundle.js' && r.status !== 0) throw new Error('the bundle did not build, so the windows could not be opened');
  }
}

function mergeWindows(results, files, out) {
  let report = [];
  try {
    report = JSON.parse(fs.readFileSync(path.join(out, 'windows', 'report.json'), 'utf8'));
  } catch {
    return;
  }
  const byFile = new Map(report.map((w) => [path.resolve(w.file), w]));
  files.forEach(({ file }, i) => {
    const w = byFile.get(path.resolve(file));
    if (w) results[i].window = { status: w.status, ms: w.ms, reason: w.reason ?? null };
  });
}

async function main() {
  const corpus = path.resolve(process.env.RUTBA_CORPUS || path.join(root, 'tests', 'corpus-public'));
  if (!fs.existsSync(corpus)) throw new Error(`there is no corpus at ${corpus}; set RUTBA_CORPUS to a folder of files`);
  const out = path.resolve(process.env.RUTBA_CORPUS_OUT || path.join(root, 'apps', 'desktop', 'build', 'corpus'));
  const rates = process.env.RUTBA_CORPUS_RATES ? JSON.parse(fs.readFileSync(process.env.RUTBA_CORPUS_RATES, 'utf8')) : null;
  const report = await runCorpus({
    corpus,
    out,
    timeout: Number(process.env.RUTBA_CORPUS_TIMEOUT || 60000),
    limit: Number(process.env.RUTBA_CORPUS_LIMIT || 0),
    rates,
    windows: process.env.RUTBA_CORPUS_WINDOWS === '1',
  });
  console.log('');
  console.log(summaryMarkdown(report));
  console.log(`results: ${path.join(out, 'results.json')}`);
  const faults = report.files.filter((f) => f.faults.length).length;
  // Faults and hangs fail the run; a refusal counts against its format's rate.
  process.exit(faults || report.rateFailures.length ? 1 : 0);
}

if (isMainThread && process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main().catch((err) => {
    console.error(`the corpus run stopped: ${err?.stack || err}`);
    process.exit(1);
  });
}
