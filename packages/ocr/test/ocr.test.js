// @rutba/ocr: the six languages read from pages the suite drew itself, the
// pool of threads (never the caller's), its queue and cancelling, and a
// recognition run that makes no network request.
//
// The fixtures are 150 dpi pages from tools/ocr-corpus.js with the text
// drawn on each beside it.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { threadId } from 'node:worker_threads';
import { createOcrHost, checkLanguages, languagesFor, LANGUAGES, OcrLanguageError } from '@rutba/ocr';
import { decodePng, isPng } from '@rutba/ocr/png';
import { toGrey, pgm } from '@rutba/ocr/pixels';
import { fromEngine, joinWords, directionOf } from '@rutba/ocr/results';

const here = path.dirname(fileURLToPath(import.meta.url));
const fixture = (name) => fs.readFileSync(path.join(here, 'fixtures', name));

/** Character error rate, spaces ignored where `unspaced`. */
function cer(truth, read, unspaced = false) {
  const norm = (s) => {
    const t = s.normalize('NFC').replace(/\s+/g, ' ').trim();
    return [...(unspaced ? t.replace(/ /g, '') : t)];
  };
  const a = norm(truth), b = norm(read);
  let prev = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    const row = [i];
    for (let j = 1; j <= b.length; j++) row[j] = Math.min(prev[j] + 1, row[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    prev = row;
  }
  return prev[b.length] / a.length;
}

test('languages: the six shipped, Urdu marked beta, unknown codes refused, a window language suggests its own script', () => {
  assert.deepEqual(LANGUAGES.map((l) => l.code), ['eng', 'urd', 'ara', 'hin', 'chi_sim', 'jpn']);
  assert.equal(LANGUAGES.find((l) => l.code === 'urd').beta, true);
  assert.equal(LANGUAGES.filter((l) => l.beta).length, 1);
  assert.deepEqual(checkLanguages('urd+eng+urd'), ['urd', 'eng']);
  assert.deepEqual(checkLanguages([]), ['eng']);
  assert.throws(() => checkLanguages(['deu']), OcrLanguageError);
  assert.deepEqual(languagesFor('ur'), ['urd', 'eng']);
  assert.deepEqual(languagesFor('en-GB'), ['eng']);
  for (const l of LANGUAGES) {
    assert.ok(fs.existsSync(new URL(`../vendor/tessdata/${l.model}.traineddata`, import.meta.url)), `${l.code} model shipped`);
    assert.ok(l.ui && l.script && ['ltr', 'rtl'].includes(l.direction), `${l.code}: a full row`);
  }
  // Every model shipped belongs to a row: the table is the one list.
  const shipped = fs.readdirSync(new URL('../vendor/tessdata/', import.meta.url)).filter((f) => f.endsWith('.traineddata')).map((f) => f.replace(/\.traineddata$/, ''));
  assert.deepEqual(shipped.sort(), [...new Set(LANGUAGES.map((l) => l.model))].sort());
  // Each language the suite's windows speak, other than English, suggests its own model first.
  for (const l of LANGUAGES.filter((x) => x.code !== 'eng')) assert.deepEqual(languagesFor(l.ui), [l.code, 'eng']);
});

/** A PNG made here, row filters and all, to read back. */
function makePng(width, height, colour, depth, rows, filters) {
  const chunk = (type, body) => {
    const len = Buffer.alloc(4); len.writeUInt32BE(body.length);
    const crcBuf = Buffer.alloc(4); crcBuf.writeUInt32BE(zlib.crc32(Buffer.concat([Buffer.from(type), body])) >>> 0);
    return Buffer.concat([len, Buffer.from(type), body, crcBuf]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0); ihdr.writeUInt32BE(height, 4); ihdr[8] = depth; ihdr[9] = colour;
  const raw = Buffer.concat(rows.map((r, i) => Buffer.concat([Buffer.from([filters[i]]), Buffer.from(r)])));
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

test('png: RGB rows under the Sub, Up, Average and Paeth filters, grey at one bit, and a page from the corpus decode exactly', () => {
  // 2x4 RGB; each row's bytes already filtered by hand for its filter.
  const plain = [[10, 20, 30, 40, 50, 60], [11, 22, 33, 44, 55, 66], [200, 100, 50, 10, 10, 10], [1, 2, 3, 4, 5, 6]];
  const sub = [10, 20, 30, 30, 30, 30];
  const up = plain[1].map((v, i) => (v - plain[0][i]) & 255);
  const avg = plain[2].map((v, i) => (v - (((i >= 3 ? plain[2][i - 3] : 0) + plain[1][i]) >> 1)) & 255);
  const paeth = plain[3].map((v, i) => {
    const a = i >= 3 ? plain[3][i - 3] : 0, b = plain[2][i], c = i >= 3 ? plain[2][i - 3] : 0;
    const p = a + b - c, pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c);
    return (v - (pa <= pb && pa <= pc ? a : pb <= pc ? b : c)) & 255;
  });
  const png = makePng(2, 4, 2, 8, [sub, up, avg, paeth], [1, 2, 3, 4]);
  const out = decodePng(png);
  assert.equal(out.width, 2);
  const rgb = [];
  for (let i = 0; i < out.data.length; i += 4) rgb.push(out.data[i], out.data[i + 1], out.data[i + 2]);
  assert.deepEqual(rgb, plain.flat());
  const one = decodePng(makePng(8, 1, 0, 1, [[0b10100000]], [0]));
  assert.deepEqual([...one.data].filter((_, i) => i % 4 === 0), [255, 0, 255, 0, 0, 0, 0, 0]);
  const page = decodePng(fixture('eng-serif.png'));
  assert.equal(page.width, 1125);
  assert.ok(isPng(fixture('eng-serif.png')) && !isPng(Buffer.from('nope')));
});

test('pixels: grey from RGBA reads transparency over white, and a PGM header the engine reads', () => {
  const g = toGrey({ width: 2, height: 1, data: new Uint8Array([0, 0, 0, 255, 0, 0, 0, 0]) });
  assert.deepEqual([...g.data], [0, 255]);
  const p = pgm({ width: 2, height: 1, data: new Uint8Array([7, 9]) });
  assert.equal(Buffer.from(p.subarray(0, 11)).toString('latin1'), 'P5\n2 1\n255\n');
  assert.deepEqual([...p.subarray(11)], [7, 9]);
});

test('results: words joined with spaces except between Chinese or Japanese characters, line direction by its letters, boxes clamped to the picture', () => {
  assert.equal(joinWords(['今日', 'は', 'Rutba', 'Office']), '今日は Rutba Office');
  assert.equal(directionOf('رتبہ آفس 12'), 'rtl');
  assert.equal(directionOf('Invoice 12 رقم'), 'ltr');
  const json = { blocks: [{ paragraphs: [{ lines: [{ bbox: { x0: -5, y0: 2, x1: 50, y1: 12 }, words: [
    { text: 'Hello', confidence: 91.6, bbox: { x0: -5, y0: 2, x1: 20, y1: 12 } }, { text: ' ', confidence: 0, bbox: { x0: 21, y0: 2, x1: 22, y1: 12 } }, { text: 'there', confidence: 88, bbox: { x0: 25, y0: 2, x1: 50, y1: 12 } }] }] }] }] };
  const r = fromEngine(json, { width: 40, height: 20, languages: ['eng'] });
  assert.equal(r.text, 'Hello there');
  assert.deepEqual(r.words[0].box, { x: 0, y: 2, width: 20, height: 10 });
  assert.equal(r.words[1].box.width, 15);
  assert.deepEqual(r.lines[0].words, [0, 2]);
  assert.equal(r.lines[0].confidence, 90);
});

test('results: a sliver takes its line\'s height, a small word on a line merged with a large heading keeps its own box, on a page turned before reading too', () => {
  const line = (words, box) => ({ bbox: box, words: words.map(([text, b]) => ({ text, confidence: 90, bbox: b })) });
  const json = { blocks: [{ paragraphs: [{ lines: [
    // A line of ordinary type, 20 high, with an Arabic word boxed as its baseline alone, 4 high.
    line([['كلمة', { x0: 10, y0: 26, x1: 60, y1: 30 }], ['أخرى', { x0: 70, y0: 10, x1: 120, y1: 30 }]], { x0: 10, y0: 10, x1: 120, y1: 30 }),
    // A small address, 16 high, the engine put on one line with a heading 60 high.
    line([['Leeds', { x0: 10, y0: 60, x1: 70, y1: 76 }], ['INVOICE', { x0: 300, y0: 40, x1: 600, y1: 100 }]], { x0: 10, y0: 40, x1: 600, y1: 100 }),
    line([['two', { x0: 10, y0: 120, x1: 50, y1: 140 }]], { x0: 10, y0: 120, x1: 50, y1: 140 }),
  ] }] }] };
  const r = fromEngine(json, { width: 700, height: 200, languages: ['ara', 'eng'] });
  const box = (t) => r.words.find((w) => w.text === t).box;
  assert.deepEqual(box('كلمة'), { x: 10, y: 10, width: 50, height: 20 }, 'the sliver takes its line\'s height');
  assert.deepEqual(box('Leeds'), { x: 10, y: 60, width: 60, height: 16 }, 'the small word keeps its own');
  // Mapped onto a page turned by a few degrees, every box grows taller; the words are judged as the engine read them.
  const grow = (b) => ({ x: b.x, y: b.y - b.width * 0.05, width: b.width, height: b.height + b.width * 0.1 });
  const t = fromEngine(json, { width: 700, height: 200, languages: ['ara', 'eng'], map: grow });
  assert.equal(t.words.find((w) => w.text === 'Leeds').box.height, 22, 'not stretched to the heading\'s line');
});

const PAGES = [
  { file: 'eng-serif', lang: 'eng', most: 0.02 },
  { file: 'urd-naskh', lang: 'urd', most: 0.08 },
  { file: 'ara-naskh', lang: 'ara', most: 0.04 },
  { file: 'hin', lang: 'hin', most: 0.02 },
  { file: 'chi_sim', lang: 'chi_sim', most: 0.04, unspaced: true },
  { file: 'jpn', lang: 'jpn', most: 0.02, unspaced: true },
];

test('recognize: each of the six languages read from a page the suite drew, on threads of the pool, never the caller\'s', async () => {
  const host = createOcrHost({ threads: 2 });
  try {
    for (const p of PAGES) {
      const truth = fixture(`${p.file}.txt`).toString('utf8');
      const r = await host.recognize(fixture(`${p.file}.png`), { languages: [p.lang] });
      const rate = cer(truth, r.text, p.unspaced);
      assert.ok(rate <= p.most, `${p.file}: CER ${(rate * 100).toFixed(1)}% over ${(p.most * 100).toFixed(0)}%: ${r.text.slice(0, 120)}`);
      assert.notEqual(r.thread, threadId, 'read on a thread of its own');
      assert.ok(r.words.length > 10 && r.lines.length >= 4, `${p.file}: ${r.words.length} words, ${r.lines.length} lines`);
      for (const w of r.words) {
        assert.ok(w.box.x >= 0 && w.box.y >= 0 && w.box.x + w.box.width <= r.width && w.box.y + w.box.height <= r.height, `${p.file}: ${w.text} inside the page`);
      }
      if (p.lang === 'urd' || p.lang === 'ara') assert.equal(r.lines[0].direction, 'rtl');
    }
  } finally {
    await host.stop();
  }
});

test('recognize: words carry boxes where the words are and a confidence; pixels read as well as the PNG they came from', async () => {
  const host = createOcrHost({ threads: 1 });
  try {
    const page = decodePng(fixture('eng-serif.png'));
    const fromPixels = await host.recognize(page, { languages: 'eng' });
    const fromPng = await host.recognize(fixture('eng-serif.png'), { languages: 'eng' });
    assert.equal(fromPixels.text, fromPng.text);
    const first = fromPixels.words[0];
    assert.equal(first.text, 'Rutba');
    assert.ok(first.confidence > 80, `confidence ${first.confidence}`);
    // The first word's box holds ink and the space left of it holds none.
    const grey = toGrey(page);
    const ink = (x0, y0, x1, y1) => { let n = 0; for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) if (grey.data[y * grey.width + x] < 128) n++; return n; };
    const b = first.box;
    assert.ok(ink(b.x, b.y, b.x + b.width, b.y + b.height) > b.width * b.height * 0.1, 'ink inside the box');
    assert.equal(ink(Math.max(0, b.x - 30), b.y, b.x - 2, b.y + b.height), 0, 'paper before it');
    assert.equal(page.data[0], 255, 'the caller keeps its pixels');
  } finally {
    await host.stop();
  }
});

test('pool: at most its threads at once, the rest queued in order; a job cancelled while waiting or reading stops and the others finish', async () => {
  const host = createOcrHost({ threads: 2 });
  try {
    const png = fixture('eng-serif.png');
    const order = [];
    const jobs = [0, 1, 2, 3].map((i) => host.recognize(png, { languages: ['eng'] }).then((r) => { order.push(i); return r; }));
    await new Promise((r) => setTimeout(r, 50));
    const s = host.stats;
    assert.equal(s.busy, 2);
    assert.equal(s.queued, 2);
    assert.ok(s.threads <= 2);
    const results = await Promise.all(jobs);
    assert.equal(new Set(results.map((r) => r.text)).size, 1);
    assert.ok(new Set(results.map((r) => r.thread)).size <= 2);

    const waiting = new AbortController();
    const reading = new AbortController();
    const a = assert.rejects(host.recognize(png, { languages: ['eng'], signal: reading.signal }), { code: 'OCR_CANCELLED' });
    const b = host.recognize(png, { languages: ['eng'] });
    const c = assert.rejects(host.recognize(png, { languages: ['eng'], signal: waiting.signal }), { code: 'OCR_CANCELLED' });
    const d = host.recognize(png, { languages: ['eng'] });
    waiting.abort();
    await new Promise((r) => setTimeout(r, 30));
    reading.abort();
    await a;
    await c;
    assert.match((await b).text, /Rutba Office/);
    assert.match((await d).text, /Rutba Office/);
    assert.equal(host.stats.cancelled, 2);
    await assert.rejects(host.recognize(png, { languages: ['xxx'] }), { code: 'OCR_LANGUAGE' });
    await assert.rejects(host.recognize(new Uint8Array([1, 2, 3]), {}), { code: 'OCR_IMAGE' });
  } finally {
    await host.stop();
  }
});

test('recognize makes no network request: every way out is refused in the process and its threads, and none is tried', () => {
  const run = spawnSync(process.execPath, ['--require', path.join(here, 'no-network.cjs'), path.join(here, 'no-network-run.mjs'), path.join(here, 'fixtures', 'urd-naskh.png')], { encoding: 'utf8' });
  assert.equal(run.status, 0, run.stderr);
  const got = JSON.parse(run.stdout.trim().split('\n').pop());
  assert.deepEqual({ guarded: got.guarded, inMain: got.inMain, attempts: got.attempts, failed: got.failed }, { guarded: true, inMain: true, attempts: 0, failed: 0 });
  assert.ok(got.words > 10);
});

test('recognize: a page turned two and a half degrees is turned back first, its words boxed where they are on the turned page; tiny text is scaled up to be read', async () => {
  const { rotate, scale } = await import('@rutba/ocr/analyse');
  const host = createOcrHost({ threads: 1 });
  try {
    const truth = fixture('eng-serif.txt').toString('utf8');
    const grey = toGrey(decodePng(fixture('eng-serif.png')));
    const turned = rotate(grey, 2.5);
    const r = await host.recognize(turned, { languages: ['eng'] });
    assert.ok(Math.abs(r.angle - 2.5) <= 0.15, `measured ${r.angle}`);
    assert.ok(cer(truth, r.text) <= 0.02, r.text);
    const ink = (b) => { let n = 0; for (let y = b.y; y < b.y + b.height; y++) for (let x = b.x; x < b.x + b.width; x++) if (turned.data[y * turned.width + x] < 128) n++; return n / (b.width * b.height); };
    for (const w of r.words.slice(0, 12)) assert.ok(ink(w.box) > 0.05, `${w.text} boxed over ink`);
    const plain = await host.recognize(grey, { languages: ['eng'] });
    assert.ok(Math.abs(plain.angle) <= 0.1, `a level page measured ${plain.angle}`);
    assert.equal(plain.prepared.turned, 0);
    assert.equal(plain.prepared.factor, 1);
    assert.ok(Math.abs(r.prepared.turned - 2.5) <= 0.15, 'turned back before reading');

    const tiny = scale(grey, 0.42);
    const small = await host.recognize(tiny, { languages: ['eng'] });
    assert.ok(small.prepared.factor > 1, JSON.stringify(small.prepared));
    assert.ok(cer(truth, small.text) <= 0.05, small.text);
    assert.ok(small.words.every((w) => w.box.x + w.box.width <= tiny.width && w.box.y + w.box.height <= tiny.height));
  } finally {
    await host.stop();
  }
});

test('recognize: heavy figures, prices in Arial Black, Impact, Segoe UI Black and others, small and large, on white, on a dark ground and on paper lit from one side, are read exactly', async () => {
  // Each a picture of one line, as Copy Text is handed a price cut from a label or a screenshot
  // (drawn by Windows' own text renderer). Read as they came, a dark ground was taken for paper lit
  // unevenly and each light letter drawn as its outline, and a heavy stroke on uneven paper was
  // left hollow: "£1.65" came back as "", "21.05" or "£1°65", a "1" as "/" and "1".
  const host = createOcrHost({ threads: 2 });
  try {
    const figures = JSON.parse(fixture('figures.json').toString('utf8'));
    const read = await Promise.all(figures.map((f) => host.recognize(fixture(f.file), { languages: ['eng'] })));
    const wrong = figures
      .map((f, i) => ({ f, got: read[i].text.replace(/\s+/g, ' ').trim() }))
      .filter(({ f, got }) => got !== f.text)
      .map(({ f, got }) => `${f.face} ${f.px}px, ${f.ground}: "${f.text}" read as "${got}"`);
    assert.deepEqual(wrong, []);
    figures.forEach((f, i) => assert.equal(read[i].prepared.inverted, /^(white|light) on/.test(f.ground), `${f.file}: read as its negative only when light on dark`));
  } finally {
    await host.stop();
  }
});

test('recognize: Urdu and Arabic keep their own comma where the model reads a Latin one', async () => {
  const host = createOcrHost({ threads: 1 });
  try {
    const r = await host.recognize(fixture('urd-naskh.png'), { languages: ['urd'] });
    assert.ok(r.text.includes('\u060c'), r.text);
    assert.ok(!/[\u0600-\u06ff],/.test(r.text), r.text);
  } finally {
    await host.stop();
  }
});

test('recognize with table: a ruled table and one set with space alone come back as their rows and cells, numbers and all', async () => {
  const host = createOcrHost({ threads: 1 });
  try {
    const truth = JSON.parse(fixture('table.json').toString('utf8'));
    for (const name of ['table-ruled', 'table-plain']) {
      const r = await host.recognize(fixture(`${name}.png`), { languages: ['eng'], table: true });
      assert.deepEqual(r.table.cells, truth, name);
      assert.equal(r.table.ruled, name === 'table-ruled');
      assert.ok(r.table.confidence.flat().every((c) => c > 50), JSON.stringify(r.table.confidence));
    }
    const plain = await host.recognize(fixture('table-plain.png'), { languages: ['eng'] });
    assert.equal(plain.table, undefined, 'a table only when asked for');
  } finally {
    await host.stop();
  }
});
