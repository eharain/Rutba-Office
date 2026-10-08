// Open damaged files until something breaks.
//
//   node tools/fuzz-open.js            300 mutations of each source
//   node tools/fuzz-open.js 2000       more of them
//   node tools/fuzz-open.js 300 4242   the same run again, from a seed
//
// The corpus run opens files people have; this opens files nobody has, made
// by damaging good ones: a byte flipped, a run of bytes zeroed, a chunk cut
// out, the tail removed, two files spliced together. The sources are a new
// document, workbook and deck, and every old-format file in
// tests/fixtures/binary — Word 97-2003 and 95, Excel 97-2003 and 95,
// PowerPoint 97-2003, and the password-protected ones — since those readers
// take the most hostile input there is. It drives the same document service
// the windows use, in a worker of its own, and asks one question of every
// answer: was it a document, or a refusal in a sentence? Anything else is a
// fault — an error a parser threw rather than chose (a RangeError from a
// DataView, a stack overflow, a TypeError), or a file that takes longer than
// two seconds, which the worker is stopped for — because a person's folder
// will eventually contain the file that finds it.
//
// A refusal is a pass. This is not a test of what damaged files mean; it is a
// test that damage cannot reach further than the sentence.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Worker, isMainThread, parentPort } from 'node:worker_threads';

/** Errors a parser throws without meaning to: each names a place in the engine, not something a person can do. */
const FAULT_NAMES = new Set(['RangeError', 'TypeError', 'ReferenceError', 'EvalError', 'InternalError', 'AggregateError']);
const FAULT_WORDS = /Maximum call stack|Invalid array length|Array buffer allocation|out of memory|outside the bounds|Offset is outside|is not a function|is not iterable|Cannot read propert|Cannot set propert|undefined is not|null is not/i;

/** Every error in the chain an answer came wrapped in: the sentence, and what it was made from. */
function chain(err) {
  const out = [];
  for (let e = err, n = 0; e && n < 8; e = e.cause, n++) out.push({ name: e?.name || typeof e, message: String(e?.message ?? e) });
  return out;
}

/** Was this thrown error a refusal in a sentence, or a fault? */
export function judge(err) {
  const links = chain(err);
  const fault = links.find((l) => FAULT_NAMES.has(l.name) || FAULT_WORDS.test(l.message));
  const top = links[0]?.message || '';
  const sentence = /[a-z]{3}.*[ .]/.test(top);
  return fault || !sentence
    ? { outcome: 'FAULT', message: `${(fault || links[0]).name}: ${(fault || links[0]).message}`.slice(0, 200) }
    : { outcome: 'refused', message: top.slice(0, 200) };
}

if (!isMainThread) {
  // The worker: opens each file it is handed, as a window would, and says how it went.
  const { createDocumentService } = await import('../apps/desktop/main/documents.js');
  const service = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
  parentPort.on('message', ({ file }) => {
    try {
      const opened = service.open({ path: file });
      // A password-protected file asks for its password: an answer.
      if (!opened.locked) {
        // Not just opened: drawn. A model that throws on the way to the
        // window is the same blank window as a parser that throws.
        service.model({ id: opened.id });
        service.close({ id: opened.id });
      }
      parentPort.postMessage({ outcome: 'opened' });
    } catch (err) {
      parentPort.postMessage(judge(err));
    }
  });
} else if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  await main();
}

async function main() {
  const { buildDocx, buildXlsx } = await import('@rutba/ooxml/build');
  const { buildPptx } = await import('@rutba/presentation');
  const rounds = Number(process.argv[2] || 300);
  const seed = Number(process.argv[3] || Date.now() % 100000);
  const TIMEOUT_MS = 2000;

  /** A repeatable pseudo-random source, so a failure can be run again. */
  let state = seed >>> 0 || 1;
  const rand = () => {
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    return ((state >>> 0) % 1000000) / 1000000;
  };
  const pick = (n) => Math.floor(rand() * n);

  const lorem = 'The supplier shall deliver the goods described in Schedule 2 within thirty days of the order.';
  const SOURCES = {
    'new.docx': buildDocx({ title: 'Agreement', paragraphs: [lorem, 'Second paragraph.', 'Third.'] }),
    'new.xlsx': buildXlsx({ sheets: [{ name: 'S', rows: [['Region', 'Total'], ['North', 12], ['South', 8], [null, '=SUM(B2:B3)']] }] }),
    'new.pptx': buildPptx({ title: 'Deck', slides: [{ layout: 'title', title: 'Rutba' }, { layout: 'titleAndContent', title: 'Two', bullets: ['a', 'b'] }] }),
  };
  const binary = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'tests', 'fixtures', 'binary');
  for (const name of fs.existsSync(binary) ? fs.readdirSync(binary).sort() : []) SOURCES[name] = fs.readFileSync(path.join(binary, name));

  /** The ways a file gets damaged in the world, and a few nastier ones. */
  const DAMAGE = {
    'flip a byte': (b) => {
      const out = Buffer.from(b);
      out[pick(out.length)] ^= 1 << pick(8);
      return out;
    },
    'zero a run': (b) => {
      const out = Buffer.from(b);
      const at = pick(out.length);
      out.fill(0, at, Math.min(out.length, at + 1 + pick(64)));
      return out;
    },
    'fill with ones': (b) => {
      // 0xFF over a run: a sector number, a count or a length read as the
      // largest it can be — what a 32-bit field does to a reader that trusts it.
      const out = Buffer.from(b);
      const at = pick(out.length);
      out.fill(0xff, at, Math.min(out.length, at + 1 + pick(16)));
      return out;
    },
    'cut the tail': (b) => Buffer.from(b).subarray(0, 1 + pick(b.length)),
    'cut the head': (b) => Buffer.from(b).subarray(pick(b.length)),
    'cut a chunk out': (b) => {
      const at = pick(b.length);
      const to = Math.min(b.length, at + 1 + pick(256));
      return Buffer.concat([Buffer.from(b).subarray(0, at), Buffer.from(b).subarray(to)]);
    },
    'splice in noise': (b) => {
      const at = pick(b.length);
      const noise = Buffer.alloc(1 + pick(128));
      for (let i = 0; i < noise.length; i++) noise[i] = pick(256);
      return Buffer.concat([Buffer.from(b).subarray(0, at), noise, Buffer.from(b).subarray(at)]);
    },
    'double it': (b) => Buffer.concat([Buffer.from(b), Buffer.from(b)]),
  };

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-fuzz-'));
  const names = Object.keys(DAMAGE);
  const sources = Object.keys(SOURCES);
  const counts = { opened: 0, refused: 0, faults: 0 };
  const slowest = [];
  const faults = [];

  // One worker opens the files; one that takes too long is stopped and
  // another started, and the file is a fault.
  let worker = null;
  const fresh = () => {
    worker = new Worker(fileURLToPath(import.meta.url));
    worker.unref();
  };
  const ask = (file) => new Promise((resolve) => {
    const timer = setTimeout(() => {
      worker.removeAllListeners('message');
      worker.removeAllListeners('error');
      worker.terminate();
      fresh();
      resolve({ outcome: 'FAULT', message: `took longer than ${TIMEOUT_MS / 1000} s` });
    }, TIMEOUT_MS);
    worker.once('message', (answer) => { clearTimeout(timer); worker.removeAllListeners('error'); resolve(answer); });
    worker.once('error', (err) => { clearTimeout(timer); worker.removeAllListeners('message'); fresh(); resolve({ outcome: 'FAULT', message: `the worker died: ${err?.message || err}`.slice(0, 200) }); });
    worker.postMessage({ file });
  });
  fresh();
  // The first file waits for the worker to load the service; that is not the file's time.
  {
    const warm = path.join(dir, 'warm.docx');
    fs.writeFileSync(warm, SOURCES['new.docx']);
    await new Promise((resolve) => { worker.once('message', resolve); worker.postMessage({ file: warm }); });
  }

  console.log(`fuzz: ${rounds} mutations of each of ${sources.length} sources (${sources.length - 3} old-format files), seed ${seed}`);

  for (let round = 0; round < rounds; round++) {
    for (const source of sources) {
      const how = names[pick(names.length)];
      const bytes = DAMAGE[how](SOURCES[source]);
      if (!bytes.length) continue;
      const ext = path.extname(source);
      const file = path.join(dir, `fuzz${ext}`);
      fs.writeFileSync(file, bytes);

      const started = Date.now();
      const answer = await ask(file);
      const ms = Date.now() - started;
      if (answer.outcome === 'opened') counts.opened += 1;
      else if (answer.outcome === 'refused') counts.refused += 1;
      else {
        counts.faults += 1;
        faults.push({ source, how, round, message: answer.message, bytes: bytes.length });
        fs.writeFileSync(path.join(dir, `fault-${counts.faults}${ext}`), bytes);
      }
      slowest.push({ ms, source, how, outcome: answer.outcome });
    }
  }
  await worker.terminate();

  slowest.sort((a, b) => b.ms - a.ms);
  const total = counts.opened + counts.refused + counts.faults;
  console.log(`\n${total} damaged files: opened ${counts.opened}, refused in a sentence ${counts.refused}, faults ${counts.faults}`);
  console.log('slowest five:');
  for (const s of slowest.slice(0, 5)) console.log(`  ${String(s.ms).padStart(6)} ms  ${s.source.padEnd(18)} ${s.how.padEnd(16)} ${s.outcome}`);

  if (faults.length) {
    console.log(`\nfaults (files kept in ${dir}):`);
    const seen = new Set();
    for (const f of faults) {
      const key = `${path.extname(f.source)}:${f.message}`;
      if (seen.has(key)) continue;
      seen.add(key);
      console.log(`  ${f.source} — ${f.how} — ${f.message}`);
    }
    console.log(`\n${faults.length} of ${total} damaged files answered with something other than a sentence.`);
    process.exit(1);
  }

  fs.rmSync(dir, { recursive: true, force: true });
  console.log('\nevery damaged file was drawn, asked its password, or refused in a sentence.');
}
