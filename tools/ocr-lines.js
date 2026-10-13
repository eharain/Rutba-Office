// Lines of text drawn by the suite's own renderer, for training and testing
// a text recogniser of our own (docs/OCR.md, "A model of our own"):
//
//   npx electron tools/ocr-lines.js --fonts <dir> (--catalogue ur | --source <file>)
//        --font NotoNastaliqUrdu.ttf --language urd --out <dir>
//        [--count 2000] [--test 0.05] [--seed 1] [--clean]
//
// Text comes from a UTF-8 file (one paragraph a line) or from the suite's
// own catalogue for a language (`--catalogue ur`, our own words, so nothing
// of anyone else's is trained on), cut into lines of three to twelve
// words. Each line is set in the font named (an OFL font from --fonts) at a
// size, weight and colour drawn at random, and spoiled as print and scans
// spoil it: a little turned, softened, on a paper grain or a shadow, drawn
// small and enlarged as a low-resolution scan is, and saved as a JPEG of a
// random quality. Everything random comes from --seed, so a set can be made
// again exactly. --clean draws the lines unspoiled.
//
// Written: <out>/train/ and <out>/test/ with <id>.png beside <id>.gt.txt
// (the line's text in logical order, as tesstrain and most line recognisers
// read ground truth), and <out>/manifest.tsv with each line's settings.

import { app, BrowserWindow, nativeImage } from 'electron';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const arg = (name, fallback) => {
  const at = process.argv.indexOf(`--${name}`);
  return at > 0 && at + 1 < process.argv.length ? process.argv[at + 1] : fallback;
};
const flag = (name) => process.argv.includes(`--${name}`);

const out = path.resolve(arg('out', path.join(here, '..', 'ocr-lines')));
const fonts = path.resolve(arg('fonts', path.join(here, 'fonts')));
const fontFile = arg('font', 'NotoNastaliqUrdu.ttf');
const lang = arg('language', 'urd');
const count = Number(arg('count', 2000));
const testShare = Number(arg('test', 0.05));
const clean = flag('clean');
const RTL = new Set(['urd', 'ara', 'fas', 'heb']);

/** A small fast generator of numbers in [0, 1), from a seed. */
function random(seed) {
  let a = seed >>> 0 || 1;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
const rand = random(Number(arg('seed', 1)));
const between = (a, b) => a + (b - a) * rand();
const pick = (list) => list[Math.floor(rand() * list.length)];

/**
 * Paragraphs of text: a catalogue's translations with their placeholders
 * filled, or a file's lines. (A value like "catalogue:ur" followed by more
 * arguments stops Electron on Windows before the script runs, so the
 * catalogue has an option of its own.)
 */
async function paragraphs({ catalogue, source }) {
  if (catalogue) {
    const code = catalogue;
    const file = path.join(here, '..', 'packages', 'office-ui', 'src', 'catalogues', `${code}.js`);
    const { default: words } = await import(pathToFileURL(file).href);
    const fill = ['12', '2026', 'رتبہ', '3.50', '7', 'اکتوبر'];
    const all = [];
    for (const value of Object.values(words)) {
      for (const text of typeof value === 'string' ? [value] : Object.values(value)) {
        all.push(String(text).replace(/\{[a-z]+\}/gi, () => pick(fill)));
      }
    }
    return all;
  }
  return fs.readFileSync(path.resolve(source), 'utf8').split(/\r?\n/);
}

/** Lines of three to twelve words, not crossing a paragraph, without repeats. */
function linesOf(paras) {
  const seen = new Set();
  const out = [];
  for (const p of paras) {
    const words = p.replace(/[\u200b-\u200f\u2060\ufeff]/g, '').split(/\s+/).filter(Boolean);
    let i = 0;
    while (i < words.length) {
      const n = Math.min(words.length - i, 3 + Math.floor(rand() * 10));
      const line = words.slice(i, i + n).join(' ');
      i += n;
      if (line.length >= 4 && !seen.has(line)) { seen.add(line); out.push(line); }
    }
  }
  return out;
}

/** One line's settings: everything drawn at random once. */
function settings() {
  const s = {
    size: Math.round(between(26, 46)),
    weight: pick([400, 400, 400, 500, 600, 700]),
    ink: Math.round(between(0, 70)),
    paper: Math.round(between(clean ? 255 : 222, 255)),
    turn: clean ? 0 : Math.round(between(-1.4, 1.4) * 10) / 10,
    blur: clean ? 0 : Math.round(pick([0, 0, 0.3, 0.5, 0.8, 1.1]) * 10) / 10,
    grain: clean ? 0 : Math.round(pick([0, 0.15, 0.25, 0.35, 0.5]) * 100) / 100,
    shadow: clean ? 0 : Math.round(pick([0, 0, 0, 0.12, 0.22]) * 100) / 100,
    small: clean ? 1 : pick([1, 1, 0.8, 0.65, 0.5]),
    jpeg: clean ? 0 : pick([0, 0, 40, 55, 70, 85]),
  };
  return s;
}

/** A page of lines, each in a box of its own, laid out down the page. */
function page(items) {
  const face = `@font-face { font-family: Line; src: url("${pathToFileURL(path.join(fonts, fontFile)).href}"); font-weight: 100 900; }`;
  const grain = `<svg width="0" height="0" style="position:absolute"><filter id="grain"><feTurbulence type="fractalNoise" baseFrequency="0.85" numOctaves="2" seed="5"/><feColorMatrix values="0 0 0 0 0.4  0 0 0 0 0.38  0 0 0 0 0.35  0 0 0 1 0"/></filter></svg>`;
  const rows = items.map(({ text, s }, i) => `<div class="row" id="r${i}"><div class="line" style="font-size:${s.size}px;font-weight:${s.weight};color:rgb(${s.ink},${s.ink},${s.ink});background:rgb(${s.paper},${s.paper - 2},${s.paper - 6});transform:rotate(${s.turn}deg);filter:blur(${s.blur}px)">${text.replace(/&/g, '&amp;').replace(/</g, '&lt;')}${s.grain ? `<i class="g" style="opacity:${s.grain}"></i>` : ''}${s.shadow ? `<i class="sh" style="opacity:${s.shadow}"></i>` : ''}</div></div>`).join('');
  return `<!doctype html><html lang="${lang}" dir="${RTL.has(lang) ? 'rtl' : 'ltr'}"><meta charset="utf-8"><style>
${face}
html, body { margin: 0; background: #fff; overflow: hidden; }
.row { padding: 14px 18px; width: max-content; }
.line { position: relative; font-family: Line; line-height: ${fontFile.includes('Nastaliq') ? 2.15 : 1.5}; padding: 0.15em 0.5em; white-space: nowrap; overflow: hidden; }
.g { position: absolute; inset: 0; filter: url(#grain); background: #fff; mix-blend-mode: multiply; }
.sh { position: absolute; inset: 0; background: linear-gradient(90deg, #000, transparent 60%); }
</style>${grain}${rows}</html>`;
}

app.commandLine.appendSwitch('force-device-scale-factor', '1');
app.whenReady().then(async () => {
  setTimeout(() => { console.log('gave up after forty minutes'); app.exit(2); }, 40 * 60 * 1000);
  const all = linesOf(await paragraphs({ catalogue: arg('catalogue', arg('source') ? null : 'ur'), source: arg('source') }));
  // Shuffled once, then the first `count`, so the set is a fair draw from the text.
  for (let i = all.length - 1; i > 0; i--) { const j = Math.floor(rand() * (i + 1)); [all[i], all[j]] = [all[j], all[i]]; }
  const chosen = all.slice(0, count);
  for (const split of ['train', 'test']) fs.mkdirSync(path.join(out, split), { recursive: true });
  const manifest = ['id\tsplit\ttext\tfont\tsize\tweight\tink\tpaper\tturn\tblur\tgrain\tshadow\tsmall\tjpeg'];
  const win = new BrowserWindow({ show: false, width: 1600, height: 1200, enableLargerThanScreen: true, webPreferences: { offscreen: true } });
  win.webContents.setFrameRate(10);
  let frame = null;
  win.webContents.on('paint', (_e, _d, image) => { frame = image; });
  const scratch = path.join(out, '.page.html');
  const BATCH = 16;
  let done = 0;
  for (let at = 0; at < chosen.length; at += BATCH) {
    const items = chosen.slice(at, at + BATCH).map((text) => ({ text, s: settings() }));
    fs.writeFileSync(scratch, page(items));
    await win.loadFile(scratch);
    await win.webContents.executeJavaScript('document.fonts.ready.then(() => 0)');
    const boxes = await win.webContents.executeJavaScript(`[...document.querySelectorAll('.row')].map((r) => { const b = r.getBoundingClientRect(); return { x: Math.floor(b.left), y: Math.floor(b.top), width: Math.ceil(b.width), height: Math.ceil(b.height) }; })`);
    const width = Math.max(...boxes.map((b) => b.x + b.width));
    const height = Math.max(...boxes.map((b) => b.y + b.height));
    win.setContentSize(Math.max(200, width), Math.max(200, height));
    await new Promise((r) => setTimeout(r, 250));
    frame = null;
    win.webContents.invalidate();
    for (let i = 0; i < 80 && !(frame && frame.getSize().height >= height); i++) await new Promise((r) => setTimeout(r, 100));
    if (!frame) throw new Error('the page of lines never painted');
    items.forEach((item, k) => {
      const box = boxes[k];
      let shot = frame.crop(box);
      const { s } = item;
      // A low-resolution scan: drawn small, then brought back to the size the line was.
      if (s.small < 1) shot = shot.resize({ width: Math.max(8, Math.round(box.width * s.small)), quality: 'good' }).resize({ width: box.width, height: box.height, quality: 'good' });
      const id = `${lang}-${String(at + k + 1).padStart(6, '0')}`;
      const split = rand() < testShare ? 'test' : 'train';
      // A JPEG's loss kept, in a PNG, so every line is one kind of file.
      fs.writeFileSync(path.join(out, split, `${id}.png`), s.jpeg ? nativeImage.createFromBuffer(shot.toJPEG(s.jpeg)).toPNG() : shot.toPNG());
      fs.writeFileSync(path.join(out, split, `${id}.gt.txt`), `${item.text}\n`);
      manifest.push([id, split, item.text, fontFile, s.size, s.weight, s.ink, s.paper, s.turn, s.blur, s.grain, s.shadow, s.small, s.jpeg].join('\t'));
    });
    done += items.length;
    if (process.env.OCR_LINES_TRACE || done % 400 < BATCH) console.log(`${done} of ${chosen.length} lines`);
  }
  fs.rmSync(scratch, { force: true });
  fs.writeFileSync(path.join(out, 'manifest.tsv'), `${manifest.join('\n')}\n`);
  console.log(`${chosen.length} lines (${all.length} to draw from) into ${out}`);
  win.destroy();
  app.exit(0);
});
