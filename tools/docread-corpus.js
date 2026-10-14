// A corpus of invoices, receipts, letters and forms with known fields, for
// measuring how well the fields of a document are read from its picture:
//
//   npx electron tools/docread-corpus.js --fonts <dir> --out <dir> [--only <id start>]
//
// tools/docread-docs.mjs decides every page (its wording, font, figures and
// layout, from a seed) and writes its HTML; this script draws each page with
// the suite's own renderer at 200 dots per inch and spoils it as
// tools/ocr-corpus.js does: "clean" as drawn, "scan" turned about a degree,
// softened, on a grain of paper and saved as a JPEG, "photo" turned more,
// lit unevenly and softened more. Beside each picture is <id>.json with what
// the page holds (its kind, its fields, its items, its tick boxes and where
// they are), and manifest.json lists them all for tools/docread-measure.mjs.
//
// Fonts: the Noto files under --fonts are OFL and not kept in the
// repository (docs/OCR.md lists them); the other faces are Windows' own.

import { app, BrowserWindow, nativeImage } from 'electron';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { makeDocument, documentHtml, corpusPlan } from './docread-docs.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
const arg = (name, fallback) => {
  const at = process.argv.indexOf(`--${name}`);
  return at > 0 ? process.argv[at + 1] : fallback;
};
const fonts = path.resolve(arg('fonts', path.join(here, 'fonts')));
const out = path.resolve(arg('out', path.join(here, '..', 'docread-corpus')));
const only = arg('only', '');
const texts = JSON.parse(fs.readFileSync(path.join(here, 'ocr-texts.json'), 'utf8'));
const DPI = 200;

const FACES = [['NotoNaskh', 'NotoNaskhArabic.ttf'], ['NotoNastaliq', 'NotoNastaliqUrdu.ttf'], ['NotoDevanagari', 'NotoSansDevanagari.ttf']];
const fontFaces = FACES.map(([family, file]) => `@font-face { font-family: ${family}; src: url("${pathToFileURL(path.join(fonts, file)).href}"); }`).join('\n');

/** How a condition spoils the drawn page: CSS on the page and over it. */
function spoil(condition, seed) {
  const sign = seed % 2 ? 1 : -1;
  if (condition === 'scan') {
    return {
      paper: '#f4f1ea', ink: '#222', page: `transform: rotate(${(sign * (0.8 + (seed % 5) / 10)).toFixed(2)}deg); filter: blur(0.55px);`,
      over: `<svg class="over" xmlns="http://www.w3.org/2000/svg"><filter id="n"><feTurbulence type="fractalNoise" baseFrequency="0.9" numOctaves="2" seed="${seed % 97}"/><feColorMatrix values="0 0 0 0 0.45  0 0 0 0 0.42  0 0 0 0 0.38  0 0 0 0.55 0"/></filter><rect width="100%" height="100%" filter="url(#n)"/></svg>`,
      jpeg: 55,
    };
  }
  if (condition === 'photo') {
    return {
      paper: '#ece8df', ink: '#2a2a2a', page: `transform: rotate(${(sign * (2.4 + (seed % 4) / 10)).toFixed(2)}deg); filter: blur(0.8px);`,
      // Light from one side: the page darkens towards the far corner, as a phone's photo of it does.
      over: `<div class="over" style="background: linear-gradient(${110 + (seed % 60)}deg, rgba(255,255,255,0) 30%, rgba(40,30,20,0.42) 100%)"></div>`,
      jpeg: 60,
    };
  }
  return { paper: '#fff', ink: '#111', page: '', over: '', jpeg: 0 };
}

app.commandLine.appendSwitch('force-device-scale-factor', '1');
app.whenReady().then(async () => {
  setTimeout(() => { console.log('gave up after fifteen minutes'); app.exit(2); }, 900000);
  fs.mkdirSync(out, { recursive: true });
  const scratch = path.join(out, '.pages');
  fs.mkdirSync(scratch, { recursive: true });
  const manifest = [];
  const win = new BrowserWindow({ show: false, width: 800, height: 600, enableLargerThanScreen: true, webPreferences: { offscreen: true } });
  win.webContents.setFrameRate(10);
  let frame = null;
  win.webContents.on('paint', (_e, _dirty, image) => { frame = image; });
  for (const p of corpusPlan()) {
    if (only && !p.id.startsWith(only)) continue;
    const doc = makeDocument(p, { texts });
    const how = spoil(p.condition, p.seed);
    let html = documentHtml(doc, { dpi: DPI, fontFaces });
    html = html.replace('</style>', `html, body { background: ${how.paper}; overflow: hidden; } main { color: ${how.ink}; ${how.page} } .over { position: fixed; inset: 0; width: 100%; height: 100%; pointer-events: none; }</style>`).replace('</main>', `</main>${how.over}`);
    const file = path.join(scratch, `${p.id}.html`);
    fs.writeFileSync(file, html);
    await win.loadFile(file);
    await win.webContents.executeJavaScript('document.fonts.ready.then(() => 0)');
    const size = await win.webContents.executeJavaScript('(() => { const m = document.querySelector("main"); return { w: m.offsetWidth, h: m.offsetHeight }; })()');
    win.setContentSize(size.w, size.h);
    await new Promise((r) => setTimeout(r, 400));
    // Where each tick box was drawn, in the picture's pixels.
    const boxes = await win.webContents.executeJavaScript('[...document.querySelectorAll(".cb")].map((e) => { const r = e.getBoundingClientRect(); return { x: Math.round(r.x), y: Math.round(r.y), width: Math.round(r.width), height: Math.round(r.height), checked: e.dataset.checked === "1" }; })');
    frame = null;
    win.webContents.invalidate();
    for (let i = 0; i < 100 && !(frame && frame.getSize().width === size.w && frame.getSize().height === size.h); i++) await new Promise((r) => setTimeout(r, 100));
    if (!frame) throw new Error(`${p.id}: no frame`);
    const shot = frame;
    if (how.jpeg) {
      const jpg = shot.toJPEG(how.jpeg);
      fs.writeFileSync(path.join(out, `${p.id}.jpg`), jpg);
      fs.writeFileSync(path.join(out, `${p.id}.png`), nativeImage.createFromBuffer(jpg).toPNG());
    } else fs.writeFileSync(path.join(out, `${p.id}.png`), shot.toPNG());
    const truth = { id: p.id, ...doc.truth, condition: p.condition, seed: p.seed, dpi: DPI, width: size.w, height: size.h };
    truth.checkboxes = truth.checkboxes.map((c, i) => ({ ...c, box: boxes[i] ? { x: boxes[i].x, y: boxes[i].y, width: boxes[i].width, height: boxes[i].height } : null }));
    fs.writeFileSync(path.join(out, `${p.id}.json`), `${JSON.stringify(truth, null, 2)}\n`);
    manifest.push({ id: p.id, picture: `${p.id}.png`, kind: p.kind, layout: p.layout, lang: p.lang, condition: p.condition, font: truth.font, digits: truth.digits, width: size.w, height: size.h });
    console.log(`${p.id.padEnd(22)} ${p.condition.padEnd(6)} ${truth.font.padEnd(16)} ${size.w}x${size.h}`);
  }
  fs.rmSync(scratch, { recursive: true, force: true });
  fs.writeFileSync(path.join(out, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`);
  console.log(`${manifest.length} pages into ${out}`);
  win.destroy();
  app.exit(0);
});
