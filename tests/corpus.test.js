/**
 * The corpus run (tools/corpus.js): what it calls a format, how it reads a
 * file before and after a round trip, how it judges an error, the numbers it
 * gives per format, and a whole run over a small folder in a temporary
 * directory. The starter corpus in tests/corpus-public is held to its record
 * of sources and licences.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation';
import { writeCalendar, dateTimeAt } from '@rutba/calendar/ical';
import { writeVCards } from '@rutba/contacts/vcard';
import { toRfc822 } from '@rutba/mailbox/rfc822';
import {
  FORMATS, SCHEMA, formatOf, collectCorpus, fingerprintOoxml, compareFingerprints, factsOf, summarise, checkRates, judge, summaryMarkdown, runCorpus,
} from '../tools/corpus.js';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-corpus-test-'));

test('the fourteen formats of the plan are the ones the run knows, each with what a round trip writes', () => {
  assert.deepEqual(Object.keys(FORMATS).sort(), ['doc', 'docm', 'docx', 'dotx', 'eml', 'ics', 'odp', 'ods', 'odt', 'pdf', 'potx', 'ppt', 'pptm', 'pptx', 'pst', 'vcf', 'xls', 'xlsm', 'xlsx', 'xltx']);
  assert.equal(FORMATS.doc.writes, 'docx', 'a 97-2003 document is saved as the XML format, as Save does');
  assert.equal(FORMATS.odt.writes, 'odt');
  assert.equal(FORMATS.pst.writes, null, 'the suite reads an Outlook data file and never writes one');
  assert.equal(FORMATS.docm.writes, 'docm', 'a macro-enabled file is saved under its own name');
  assert.equal(FORMATS.potx.writes, 'potx', 'a template is saved as a template');
  assert.equal(formatOf('C:/x/Report.DOCX'), 'docx');
  assert.equal(formatOf('notes.txt'), null);
});

test('a corpus folder is walked in every sub-folder, in a stable order, skipping what is not a document', () => {
  const dir = tmp();
  try {
    fs.mkdirSync(path.join(dir, 'b', 'deep'), { recursive: true });
    fs.writeFileSync(path.join(dir, 'b', 'deep', 'z.xlsx'), 'x');
    fs.writeFileSync(path.join(dir, 'a.docx'), 'x');
    fs.writeFileSync(path.join(dir, '~$a.docx'), 'x');
    fs.writeFileSync(path.join(dir, 'readme.txt'), 'x');
    fs.mkdirSync(path.join(dir, 'node_modules'));
    fs.writeFileSync(path.join(dir, 'node_modules', 'c.pdf'), 'x');
    assert.deepEqual(collectCorpus(dir).map((f) => [f.rel, f.format]), [['a.docx', 'docx'], ['b/deep/z.xlsx', 'xlsx']]);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a document is read as its paragraphs in order, its tables and its drawings', () => {
  const print = fingerprintOoxml('doc', buildDocx({ paragraphs: ['First   line', '', 'یہ اردو ہے & more'] }));
  assert.deepEqual(print.paragraphs, ['First line', 'یہ اردو ہے & more'], 'spaces squashed, empty paragraphs left out, entities read');
  assert.equal(print.tables, 0);
  assert.deepEqual(factsOf(print), { paragraphs: 2, tables: 0, drawings: 0 });
});

test('a workbook is read as its sheets in order, each cell its value or its formula', () => {
  const print = fingerprintOoxml('sheet', buildXlsx({ sheets: [{ name: 'Data', rows: [['Name', 'Value'], ['a', 1.5], ['b', '=B2*2']] }, { name: 'Other', rows: [[true]] }] }));
  assert.deepEqual(print.sheets.map((s) => s.name), ['Data', 'Other']);
  assert.equal(print.sheets[0].cells.A1, 'Name');
  assert.equal(print.sheets[0].cells.B2, '1.5');
  assert.equal(print.sheets[0].cells.B3, '=B2*2');
  assert.deepEqual(factsOf(print), { sheets: 2, cells: 7 });
});

test('a deck is read as its slides in order, the text of each', () => {
  const print = fingerprintOoxml('deck', buildPptx({ title: 'T', slides: [{ layout: 'title', title: 'Hello', body: 'World' }, { layout: 'obj', title: 'Two', body: ['a', 'b'] }] }));
  assert.equal(print.slides.length, 2);
  assert.deepEqual(print.slides[0].texts, ['Hello', 'World']);
  assert.deepEqual(print.slides[1].texts, ['Two', 'a', 'b']);
});

test('two readings that part company say where, in reading order, and no more than the limit', () => {
  const a = fingerprintOoxml('doc', buildDocx({ paragraphs: ['One', 'Two', 'Three'] }));
  const b = fingerprintOoxml('doc', buildDocx({ paragraphs: ['One', 'Too', 'Three'] }));
  assert.deepEqual(compareFingerprints(a, a), []);
  assert.deepEqual(compareFingerprints(a, b), ['paragraph 2: "Two" before, "Too" after']);
  const s1 = fingerprintOoxml('sheet', buildXlsx({ sheets: [{ name: 'S', rows: [[1, 2, 3, 4]] }] }));
  const s2 = fingerprintOoxml('sheet', buildXlsx({ sheets: [{ name: 'S', rows: [[1, 9, 9, 9]] }] }));
  assert.equal(compareFingerprints(s1, s2, 2).length, 2, 'the limit holds');
  assert.match(compareFingerprints(s1, s2)[0], /^S!B1: "2" before, "9" after$/);
  const d1 = { kind: 'deck', slides: [{ texts: ['a'], pictures: 0 }] };
  assert.deepEqual(compareFingerprints(d1, { kind: 'deck', slides: [] }), ['slides: 1 before, 0 after']);
  assert.deepEqual(compareFingerprints(d1, { kind: 'doc', paragraphs: [] }), ['kind: deck before, doc after']);
});

test('an error a parser threw is a fault; one it chose to say in a sentence is a refusal', () => {
  assert.equal(judge(new TypeError('Cannot read properties of undefined')).outcome, 'fault');
  assert.equal(judge(new Error('x is not iterable')).outcome, 'fault');
  assert.equal(judge(new Error('This file is not a Word document.')).outcome, 'refused');
  assert.equal(judge(new Error('Not a document.', { cause: new RangeError('Offset is outside the bounds') })).outcome, 'fault', 'a fault under a sentence is still a fault');
});

test('the numbers per format: a password-protected file is not counted against the open rate, a round trip only where one was tried', () => {
  const r = (format, status, extra = {}) => ({ format, status, faults: [], warnings: [], roundTrip: null, ms: { total: 10 }, ...extra });
  const formats = summarise([
    r('docx', 'opened', { roundTrip: { equal: true }, ms: { total: 10 } }),
    r('docx', 'opened', { roundTrip: { equal: false }, ms: { total: 30 }, warnings: ['w'] }),
    r('docx', 'locked'),
    r('docx', 'refused'),
    r('pst', 'opened'),
    r('pdf', 'timeout', { faults: ['took too long'] }),
  ]);
  assert.deepEqual(Object.keys(formats), ['docx', 'pdf', 'pst']);
  assert.equal(formats.docx.files, 4);
  assert.equal(formats.docx.openRate, 0.6667, 'two of the three that could open did');
  assert.equal(formats.docx.roundTrips, 2);
  assert.equal(formats.docx.roundTripRate, 0.5);
  assert.equal(formats.docx.withWarnings, 1);
  assert.deepEqual(formats.docx.ms, { median: 10, p90: 10, max: 30 });
  assert.equal(formats.pst.roundTripRate, null, 'never tried is not zero');
  assert.equal(formats.pdf.timeouts, 1);
  assert.equal(formats.pdf.faults, 1);
  assert.deepEqual(checkRates(formats, { docx: { open: 0.9, roundTrip: 0.5 }, '*': { open: 0.5 } }), [
    'docx: opened 0.6667 of its files, below the agreed 0.9',
    'pdf: opened 0 of its files, below the agreed 0.5',
  ]);
});

test('a whole run over a folder: every file a result, results.json in its schema, and summary.md per format', async () => {
  const dir = tmp();
  const out = tmp();
  try {
    fs.writeFileSync(path.join(dir, 'letter.docx'), buildDocx({ paragraphs: ['Dear Bo,', 'The paper arrives on Thursday.'] }));
    fs.writeFileSync(path.join(dir, 'broken.docx'), Buffer.from('this is not a zip file at all, only words'));
    fs.writeFileSync(path.join(dir, 'book.xlsx'), buildXlsx({ sheets: [{ name: 'S', rows: [['a', 1], ['b', '=B1+1']] }] }));
    fs.writeFileSync(path.join(dir, 'notes.pdf'), Buffer.from('%PDF-1.4 nothing after this'));
    fs.writeFileSync(path.join(dir, 'mail.eml'), toRfc822({ subject: 'Hello', from: [{ address: 'a@example.com' }], to: [{ address: 'b@example.com' }], text: 'Body text.', date: '2026-01-01T00:00:00Z' }));
    fs.writeFileSync(path.join(dir, 'cal.ics'), writeCalendar({ events: [{ uid: 'u1', summary: 'Meet', start: dateTimeAt(Date.UTC(2026, 0, 2, 9)), end: dateTimeAt(Date.UTC(2026, 0, 2, 10)) }] }));
    fs.writeFileSync(path.join(dir, 'people.vcf'), writeVCards([{ name: { full: 'Bo Lindqvist', given: 'Bo', family: 'Lindqvist' }, emails: [{ value: 'bo@example.com' }] }]));

    const said = [];
    const report = await runCorpus({ corpus: dir, out, timeout: 60000, log: (s) => said.push(s), rates: { '*': { open: 0.5 } } });
    const saved = JSON.parse(fs.readFileSync(path.join(out, 'results.json'), 'utf8'));
    assert.deepEqual(saved, JSON.parse(JSON.stringify(report)));
    assert.equal(saved.schema, SCHEMA);
    assert.equal(saved.corpus.files, 7);
    const by = Object.fromEntries(saved.files.map((f) => [f.file, f]));

    assert.equal(by['letter.docx'].status, 'opened');
    assert.equal(by['letter.docx'].roundTrip.target, 'docx');
    assert.equal(by['letter.docx'].roundTrip.equal, true, JSON.stringify(by['letter.docx'].roundTrip));
    assert.deepEqual(by['letter.docx'].facts, { paragraphs: 2, tables: 0, drawings: 0 });
    assert.equal(by['book.xlsx'].roundTrip.equal, true, JSON.stringify(by['book.xlsx'].roundTrip));
    assert.equal(by['broken.docx'].status, 'refused');
    assert.ok(by['broken.docx'].refused.length > 10, 'a refusal carries its sentence');
    assert.deepEqual(by['broken.docx'].faults, []);
    assert.equal(by['notes.pdf'].status, 'refused');
    for (const name of ['mail.eml', 'cal.ics', 'people.vcf']) assert.equal(by[name].roundTrip.equal, true, `${name}: ${JSON.stringify(by[name])}`);
    assert.deepEqual(by['cal.ics'].facts, { events: 1 });

    for (const f of saved.files) {
      for (const key of ['file', 'format', 'bytes', 'sha256', 'status', 'opened', 'roundTrip', 'warnings', 'faults', 'refused', 'ms', 'window']) assert.ok(key in f, `${f.file} has ${key}`);
      assert.ok(f.ms.total >= 0);
    }
    assert.equal(saved.formats.docx.files, 2);
    assert.equal(saved.formats.docx.openRate, 0.5);
    assert.deepEqual(saved.rateFailures, ['pdf: opened 0 of its files, below the agreed 0.5']);
    const md = fs.readFileSync(path.join(out, 'summary.md'), 'utf8');
    assert.match(md, /\| docx \| 2 \| 1 \| 0 \| 1 \| 0 \| 0 \| 1\/1 \| 50\.0% \|/);
    assert.match(md, /broken\.docx \(refused\)/);
    assert.equal(md, summaryMarkdown(saved));
    assert.ok(said.some((l) => /letter\.docx/.test(l)), 'each file is said as it goes');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
    fs.rmSync(out, { recursive: true, force: true });
  }
});

test('a file that takes past the limit is a hang, and the run goes on to the next', async () => {
  const dir = tmp();
  const out = tmp();
  try {
    fs.writeFileSync(path.join(dir, 'a.vcf'), 'BEGIN:VCARD\r\nFN:A\r\nEND:VCARD\r\n');
    fs.writeFileSync(path.join(dir, 'b.vcf'), 'BEGIN:VCARD\r\nFN:B\r\nEND:VCARD\r\n');
    // One millisecond is less than a worker takes to start: each file is stopped and the next given a fresh one.
    const report = await runCorpus({ corpus: dir, out, timeout: 1, log: () => {} });
    assert.deepEqual(report.files.map((f) => f.status), ['timeout', 'timeout']);
    assert.match(report.files[0].faults[0], /took longer than/);
    assert.equal(report.formats.vcf.timeouts, 2);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
    fs.rmSync(out, { recursive: true, force: true });
  }
});

test('the starter corpus holds every format, and every file in it is named with its source, licence and hash', () => {
  const dir = path.join(root, 'tests', 'corpus-public');
  const record = JSON.parse(fs.readFileSync(path.join(dir, 'SOURCES.json'), 'utf8'));
  const named = new Map(record.files.map((f) => [f.file, f]));
  const present = collectCorpus(dir);
  for (const { rel, file } of present) {
    const entry = named.get(rel);
    assert.ok(entry, `${rel} is not in SOURCES.json`);
    assert.ok(entry.source && entry.source.length > 10, `${rel} names its source`);
    assert.equal(entry.licence, 'AGPL-3.0-or-later', `${rel} carries the repository's licence`);
    assert.equal(crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex'), entry.sha256, `${rel} is the file recorded`);
  }
  assert.deepEqual([...named.keys()].sort(), present.map((p) => p.rel).sort(), 'SOURCES.json names no file that is missing');
  const other = fs.readdirSync(dir, { recursive: true }).map((p) => String(p).split(path.sep).join('/')).filter((p) => fs.statSync(path.join(dir, p)).isFile() && !formatOf(p));
  assert.deepEqual(other.sort(), ['README.md', 'SOURCES.json'], 'nothing else lives in the folder');
  assert.deepEqual([...new Set(present.map((p) => p.format))].sort(), Object.keys(FORMATS).sort(), 'one file at least of every format');
});
