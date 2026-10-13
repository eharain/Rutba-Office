// Measures text recognition on a corpus from tools/ocr-corpus.js:
//
//   node tools/ocr-measure.mjs --corpus <dir> --core <dir> --models <dir>
//        [--build simd-lstm] [--label fast] [--out results.json]
//   node tools/ocr-measure.mjs --corpus <dir> --ocr [--out results.json]
//
// --ocr reads through @rutba/ocr instead, as the suite does: the shipped
// engine and models, with our own analysis first (deskew, scaling, and
// ink and paper for unevenly lit pages), on the package's thread pool.
//
// --core is a directory holding tesseract-core-<build>.js and .wasm (the
// engine tesseract.js 7 runs, tesseract.js-core 7), --models one holding
// <lang>.traineddata. Each page is read in a process of its own, twice: the
// first read includes starting the engine and loading the model, the second
// is the page alone, and the process's peak memory is what it reached over
// both. Error rates compare the text read with the text drawn: the
// character error rate is the edit distance over the drawn text's length in
// characters, the word error rate the same over words. Chinese and Japanese
// are written without spaces, so their character rate ignores spaces and
// they have no word rate.

import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const arg = (name, fallback) => {
  const at = process.argv.indexOf(`--${name}`);
  return at > 0 ? process.argv[at + 1] : fallback;
};

/** One page read through @rutba/ocr, our own analysis included, on its pool of one thread. */
async function oneThroughPackage() {
  const { createOcrHost } = await import('@rutba/ocr');
  const host = createOcrHost({ threads: 1 });
  const picture = fs.readFileSync(arg('picture').replace(/\.jpg$/, '.png'));
  const started = performance.now();
  const r = await host.recognize(picture, { languages: arg('lang').split('+') });
  const first = performance.now();
  await host.recognize(picture, { languages: arg('lang').split('+') });
  const second = performance.now();
  await host.stop();
  process.stdout.write(`${JSON.stringify({ text: r.text, angle: r.angle, prepared: r.prepared, loadMs: 0, firstMs: Math.round(first - started), pageMs: Math.round(second - first), peakMB: Math.round(process.resourceUsage().maxRSS / 1024) })}\n`);
}

/** One page, in this process: prints a JSON line and exits. */
async function one() {
  if (process.argv.includes('--ocr')) return oneThroughPackage();
  const core = path.resolve(arg('core'));
  const build = arg('build', 'simd-lstm');
  const started = performance.now();
  const require = createRequire(import.meta.url);
  const Core = require(path.join(core, `tesseract-core-${build}.js`));
  const engine = await Core({ wasmBinary: fs.readFileSync(path.join(core, `tesseract-core-${build}.wasm`)), print() {}, printErr() {} });
  const lang = arg('lang');
  for (const l of lang.split('+')) engine.FS.writeFile(`/${l}.traineddata`, fs.readFileSync(path.join(arg('models'), `${l}.traineddata`)));
  const api = new engine.TessBaseAPI();
  if (api.Init('/', lang, 1) !== 0) throw new Error(`the ${lang} model did not load`);
  const loaded = performance.now();
  const read = () => {
    engine.FS.writeFile('/input', fs.readFileSync(arg('picture')));
    api.SetImageFile(1, 0);
    api.Recognize(null);
    return api.GetUTF8Text();
  };
  const text = read();
  const first = performance.now();
  read();
  const second = performance.now();
  api.End();
  process.stdout.write(`${JSON.stringify({
    text,
    loadMs: Math.round(loaded - started),
    firstMs: Math.round(first - started),
    pageMs: Math.round(second - first),
    peakMB: Math.round(process.resourceUsage().maxRSS / 1024),
  })}\n`);
}

/** Edit distance between two arrays. */
function distance(a, b) {
  let prev = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    const row = [i];
    for (let j = 1; j <= b.length; j++) row[j] = Math.min(prev[j] + 1, row[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    prev = row;
  }
  return prev[b.length];
}

const NO_SPACES = new Set(['chi_sim', 'jpn']);
const normal = (text, lang) => {
  const t = text.normalize('NFC').replace(/\s+/g, ' ').trim();
  return NO_SPACES.has(lang) ? t.replace(/ /g, '') : t;
};

export function errorRates(truth, read, lang) {
  const a = normal(truth, lang);
  const b = normal(read, lang);
  const cer = distance([...a], [...b]) / Math.max(1, [...a].length);
  if (NO_SPACES.has(lang)) return { cer, wer: null };
  const wa = a.split(' ');
  const wb = b ? b.split(' ') : [];
  return { cer, wer: distance(wa, wb) / Math.max(1, wa.length) };
}

async function all() {
  const corpus = path.resolve(arg('corpus'));
  const label = arg('label', 'fast');
  const build = arg('build', 'simd-lstm');
  const pages = JSON.parse(fs.readFileSync(path.join(corpus, 'manifest.json'), 'utf8'));
  const only = arg('only');
  const results = [];
  for (const p of pages) {
    if (only && !p.id.startsWith(only)) continue;
    const run = spawnSync(process.execPath, [fileURLToPath(import.meta.url), '--one', ...(process.argv.includes('--ocr') ? ['--ocr'] : []),
      '--core', arg('core', ''), '--models', arg('models', ''), '--build', build, '--lang', p.lang, '--picture', path.join(corpus, p.picture)],
    { encoding: 'utf8', maxBuffer: 1 << 24 });
    if (run.status !== 0) {
      console.log(`${p.id}: failed ${run.stderr.split('\n').slice(0, 3).join(' ')}`);
      results.push({ ...p, label, build, failed: true });
      continue;
    }
    const got = JSON.parse(run.stdout.trim().split('\n').pop());
    const truth = fs.readFileSync(path.join(corpus, `${p.id}.txt`), 'utf8');
    const rates = errorRates(truth, got.text, p.lang);
    results.push({ ...p, label, build, ...rates, loadMs: got.loadMs, firstMs: got.firstMs, pageMs: got.pageMs, peakMB: got.peakMB, text: got.text });
    const pct = (x) => (x === null ? '  -  ' : `${(x * 100).toFixed(1).padStart(5)}%`);
    console.log(`${p.id.padEnd(24)} CER ${pct(rates.cer)}  WER ${pct(rates.wer)}  page ${String(got.pageMs).padStart(6)} ms  first ${String(got.firstMs).padStart(6)} ms  peak ${got.peakMB} MB`);
  }
  const out = arg('out');
  if (out) fs.writeFileSync(out, `${JSON.stringify(results, null, 2)}\n`);
}

if (process.argv.includes('--one')) await one();
else if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) await all();
