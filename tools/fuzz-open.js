// Open damaged files until something breaks.
//
//   node tools/fuzz-open.js            300 mutations of each format
//   node tools/fuzz-open.js 2000       more of them
//   node tools/fuzz-open.js 300 4242   the same run again, from a seed
//
// The corpus run opens files people have; this opens files nobody has, made
// by damaging good ones: a byte flipped, a run of bytes zeroed, a chunk cut
// out, the tail removed, two files spliced together. It drives the same
// document service the windows use, in node, and asks one question of every
// answer: was it a document, or a refusal in a sentence? Anything else — a
// stack from deep inside a parser, a hang, a heap that never comes back — is
// a fault, because a person's folder will eventually contain the file that
// finds it.
//
// A refusal is a pass. This is not a test of what damaged files mean; it is a
// test that damage cannot reach further than the sentence.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createDocumentService } from '../apps/desktop/main/documents.js';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation';

const rounds = Number(process.argv[2] || 300);
const seed = Number(process.argv[3] || Date.now() % 100000);

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
  docx: buildDocx({ title: 'Agreement', paragraphs: [lorem, 'Second paragraph.', 'Third.'] }),
  xlsx: buildXlsx({ sheets: [{ name: 'S', rows: [['Region', 'Total'], ['North', 12], ['South', 8], [null, '=SUM(B2:B3)']] }] }),
  pptx: buildPptx({ title: 'Deck', slides: [{ layout: 'title', title: 'Rutba' }, { layout: 'titleAndContent', title: 'Two', bullets: ['a', 'b'] }] }),
};

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

const service = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-fuzz-'));
const names = Object.keys(DAMAGE);
const kinds = Object.keys(SOURCES);
const counts = { opened: 0, refused: 0, faults: 0 };
const slowest = [];
const faults = [];

console.log(`fuzz: ${rounds} mutations per format, seed ${seed}`);

for (let round = 0; round < rounds; round++) {
  for (const kind of kinds) {
    const how = names[pick(names.length)];
    const bytes = DAMAGE[how](SOURCES[kind]);
    if (!bytes.length) continue;
    const file = path.join(dir, `fuzz.${kind}`);
    fs.writeFileSync(file, bytes);

    const started = Date.now();
    let outcome;
    try {
      const opened = service.open({ path: file });
      // Not just opened: drawn. A model that throws on the way to the window
      // is the same blank window as a parser that throws.
      service.model({ id: opened.id });
      service.close({ id: opened.id });
      outcome = 'opened';
      counts.opened += 1;
    } catch (err) {
      // A sentence is an answer. A stack out of a parser is not: an error
      // whose message is a type error, an assertion, or a code names a place
      // in the engine rather than something a person can do.
      const message = String(err?.message ?? err);
      const isSentence = /[a-z]{3}.*[ .]/.test(message) && !/^[A-Za-z]*Error:/.test(message) && !/undefined|not a function|Cannot read/i.test(message);
      if (isSentence) {
        outcome = 'refused';
        counts.refused += 1;
      } else {
        outcome = 'FAULT';
        counts.faults += 1;
        faults.push({ kind, how, round, message: message.slice(0, 200), bytes: bytes.length });
        fs.writeFileSync(path.join(dir, `fault-${counts.faults}.${kind}`), bytes);
      }
    }
    const ms = Date.now() - started;
    slowest.push({ ms, kind, how, outcome });
  }
}

slowest.sort((a, b) => b.ms - a.ms);
console.log(`\nopened ${counts.opened}, refused in a sentence ${counts.refused}, faults ${counts.faults}`);
console.log('slowest five:');
for (const s of slowest.slice(0, 5)) console.log(`  ${String(s.ms).padStart(6)} ms  ${s.kind.padEnd(5)} ${s.how.padEnd(16)} ${s.outcome}`);

if (faults.length) {
  console.log(`\nfaults (files kept in ${dir}):`);
  const seen = new Set();
  for (const f of faults) {
    const key = `${f.kind}:${f.message}`;
    if (seen.has(key)) continue;
    seen.add(key);
    console.log(`  ${f.kind} — ${f.how} — ${f.message}`);
  }
  console.log(`\n${faults.length} of ${counts.opened + counts.refused + faults.length} damaged files answered with something other than a sentence.`);
  process.exit(1);
}

fs.rmSync(dir, { recursive: true, force: true });
console.log('\nevery damaged file was drawn or refused in a sentence.');
