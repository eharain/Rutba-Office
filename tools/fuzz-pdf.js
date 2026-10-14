// The PDF reader against damaged files, as tools/fuzz-open.js holds the
// document readers: node tools/fuzz-pdf.js [rounds] [seed]
//
// The PDF app reads files it did not write, on the main process, so a file
// that makes the reader throw a fault, loop or run out of memory is a window
// that hangs. Each damaged file is opened in a worker of its own with a
// two-second limit, its pages listed, each page's text read, and its pages
// assembled into a new file; its comments, form fields, bookmarks, page
// labels and links read, a comment added as an update and an area of its
// first page redacted; anything but a result or a refusal in a sentence is
// a fault. Among the sources are a file with comments, a form and a book
// with bookmarks, labels and links, so their readers meet damage too.

import fs from 'node:fs';
import { Worker, isMainThread, parentPort } from 'node:worker_threads';
import { fileURLToPath, pathToFileURL } from 'node:url';
import path from 'node:path';

// fuzz-open.js's judgement, copied rather than imported: importing that file
// from a worker would start its own worker side too.
const FAULT_NAMES = new Set(['RangeError', 'TypeError', 'ReferenceError', 'EvalError', 'InternalError', 'AggregateError']);
const FAULT_WORDS = /Maximum call stack|Invalid array length|Array buffer allocation|out of memory|outside the bounds|Offset is outside|is not a function|is not iterable|Cannot read propert|Cannot set propert|undefined is not|null is not/i;
function judge(err) {
  const links = [];
  for (let e = err, n = 0; e && n < 8; e = e.cause, n++) links.push({ name: e?.name || typeof e, message: String(e?.message ?? e) });
  const fault = links.find((l) => FAULT_NAMES.has(l.name) || FAULT_WORDS.test(l.message));
  const top = links[0]?.message || '';
  return fault || !/[a-z]{3}.*[ .]/.test(top)
    ? { outcome: 'FAULT', message: `${(fault || links[0]).name}: ${(fault || links[0]).message}`.slice(0, 200) }
    : { outcome: 'refused', message: top.slice(0, 200) };
}

if (!isMainThread) {
  const { openPdf } = await import('@rutba/pdf/reader');
  const { assemble } = await import('@rutba/pdf/assemble');
  const { readAnnotations, annotate } = await import('@rutba/pdf/annotations');
  const { readFields } = await import('@rutba/pdf/forms');
  const { outlinePages, readLinks, pageLabels } = await import('@rutba/pdf/navigation');
  const { applyRedactions } = await import('@rutba/pdf/redact');
  // A refusal in a sentence from one step does not stop the others being tried.
  const step = (fn) => { try { fn(); } catch (err) { const j = judge(err); if (j.outcome === 'FAULT') throw err; } };
  parentPort.on('message', ({ bytes }) => {
    try {
      const pdf = openPdf(Buffer.from(bytes));
      const pages = pdf.pages();
      for (let i = 0; i < Math.min(pages.length, 20); i++) pdf.pageText(i);
      if (pages.length) assemble(pages.slice(0, 20).map((_, i) => ({ file: pdf, page: i, rotate: 90 })), { title: 'fuzz' });
      step(() => readAnnotations(pdf));
      step(() => readFields(pdf));
      step(() => { outlinePages(pdf); readLinks(pdf); pageLabels(pdf); });
      if (pages.length) {
        step(() => annotate(Buffer.from(bytes), { add: [{ page: 0, type: 'Text', rect: [10, 10, 30, 30], contents: 'fuzz' }] }));
        step(() => applyRedactions(Buffer.from(bytes), { areas: [{ page: 0, rects: [[0, 300, 400, 800]] }] }));
      }
      parentPort.postMessage({ outcome: 'opened' });
    } catch (err) {
      parentPort.postMessage(judge(err));
    }
  });
} else if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const { PdfDocument } = await import('@rutba/pdf');
  const rounds = Number(process.argv[2] || 300);
  let state = (Number(process.argv[3] || Date.now() % 100000) >>> 0) || 1;
  const seed = state;
  const rand = () => { state ^= state << 13; state ^= state >>> 17; state ^= state << 5; return ((state >>> 0) % 1000000) / 1000000; };
  const pick = (n) => Math.floor(rand() * n);

  const written = (count, compress) => {
    const doc = new PdfDocument({ title: 'Fuzz', compress });
    for (let i = 0; i < count; i++) {
      const page = doc.addPage();
      page.text(`Page ${i + 1}`, 72, 100, { font: 'Helvetica', size: 24 });
      page.text('Body text on the page, for the reader to find.', 72, 140, { font: 'Courier', size: 12 });
    }
    return doc.toBuffer();
  };
  const SOURCES = { 'three pages': written(3, true), 'one page': written(1, false), 'twelve pages': written(12, true) };
  // A file with comments, a form, and a book with bookmarks, labels and links.
  const { annotate } = await import('@rutba/pdf/annotations');
  const tests = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'packages', 'pdf', 'test');
  const { formPdf } = (await import(pathToFileURL(path.join(tests, 'form-fixture.js')).href)).default;
  const { book } = (await import(pathToFileURL(path.join(tests, 'nav-fixture.js')).href)).default;
  const commented = annotate(written(2, true), { add: [
    { page: 0, type: 'Highlight', quads: [[72, 700, 200, 700, 72, 686, 200, 686]], contents: 'fuzz' },
    { page: 0, type: 'FreeText', rect: [72, 400, 300, 450], contents: 'A box' },
    { page: 1, type: 'Ink', ink: [[100, 100, 120, 130, 150, 120]] },
  ] }).bytes;
  SOURCES['with comments'] = commented;
  SOURCES['a form'] = formPdf();
  SOURCES['a book'] = book('Page');
  const fixtures = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'tests', 'fixtures');
  for (const name of fs.readdirSync(fixtures).filter((n) => n.toLowerCase().endsWith('.pdf'))) SOURCES[name] = fs.readFileSync(path.join(fixtures, name));

  const DAMAGE = {
    'flip a byte': (b) => { const o = Buffer.from(b); o[pick(o.length)] ^= 1 << pick(8); return o; },
    'zero a run': (b) => { const o = Buffer.from(b); const at = pick(o.length); o.fill(0, at, Math.min(o.length, at + 1 + pick(64))); return o; },
    'fill with ones': (b) => { const o = Buffer.from(b); const at = pick(o.length); o.fill(0xff, at, Math.min(o.length, at + 1 + pick(16))); return o; },
    'a digit changed': (b) => {
      // Numbers carry a PDF's structure (offsets, lengths, object numbers):
      // change one to another digit, which a byte flip rarely does.
      const o = Buffer.from(b);
      const digits = [];
      for (let i = 0; i < o.length; i++) if (o[i] >= 48 && o[i] <= 57) digits.push(i);
      if (digits.length) o[digits[pick(digits.length)]] = 48 + pick(10);
      return o;
    },
    'cut the tail': (b) => Buffer.from(b).subarray(0, 1 + pick(b.length)),
    'cut the head': (b) => Buffer.from(b).subarray(pick(b.length)),
    'cut a chunk out': (b) => { const at = pick(b.length); const to = Math.min(b.length, at + 1 + pick(256)); return Buffer.concat([b.subarray(0, at), b.subarray(to)]); },
    'splice in noise': (b) => { const at = pick(b.length); const n = Buffer.alloc(1 + pick(128)); for (let i = 0; i < n.length; i++) n[i] = pick(256); return Buffer.concat([b.subarray(0, at), n, b.subarray(at)]); },
    'double it': (b) => Buffer.concat([b, b]),
  };
  const names = Object.keys(DAMAGE);

  const run = (bytes) => new Promise((resolve) => {
    const w = new Worker(fileURLToPath(import.meta.url), { resourceLimits: { maxOldGenerationSizeMb: 512 } });
    const timer = setTimeout(() => { w.terminate(); resolve({ outcome: 'FAULT', message: 'did not finish in two seconds' }); }, 2000);
    w.once('message', (m) => { clearTimeout(timer); w.terminate(); resolve(m); });
    w.once('error', (err) => { clearTimeout(timer); resolve({ outcome: 'FAULT', message: String(err?.message || err).slice(0, 200) }); });
    w.postMessage({ bytes });
  });

  console.log(`fuzz-pdf: ${rounds} mutations of each of ${Object.keys(SOURCES).length} sources, seed ${seed}`);
  const counts = { opened: 0, refused: 0, FAULT: 0 };
  const faults = [];
  for (const [source, bytes] of Object.entries(SOURCES)) {
    for (let i = 0; i < rounds; i++) {
      const how = names[pick(names.length)];
      const damaged = DAMAGE[how](bytes);
      const r = await run(damaged);
      counts[r.outcome] = (counts[r.outcome] || 0) + 1;
      if (r.outcome === 'FAULT') faults.push({ source, how, message: r.message, bytes: damaged });
    }
  }
  const total = counts.opened + counts.refused + counts.FAULT;
  console.log(`${total} damaged PDFs: opened ${counts.opened}, refused in a sentence ${counts.refused}, faults ${counts.FAULT}`);
  if (faults.length) {
    const dir = fs.mkdtempSync(path.join((await import('node:os')).tmpdir(), 'rutba-fuzz-pdf-'));
    faults.forEach((f, i) => fs.writeFileSync(path.join(dir, `${i}.pdf`), f.bytes));
    console.log(`faults (files kept in ${dir}):`);
    for (const f of faults.slice(0, 20)) console.log(`  ${f.source}, ${f.how}: ${f.message}`);
    process.exit(1);
  }
  console.log('every damaged PDF was read or refused in a sentence.');
}
