// A corpus of pages for measuring text recognition, drawn by the suite's own
// renderer so every page's text is known exactly:
//
//   npx electron tools/ocr-corpus.js --fonts <dir> --out <dir>
//
// Each language's sentences (tools/ocr-texts.json) are set as one page in a
// font that draws the script, at 150 and 300 dots per inch, and once more at
// 300 as a simulated scan: turned a little, softened, on a grain of paper
// and saved as a JPEG (and that JPEG decoded again as a PNG). Beside each picture is its text, and manifest.json
// lists them all for tools/ocr-measure.mjs.
//
// Fonts: the Noto files under --fonts are OFL and are not kept in the
// repository (docs/OCR.md lists them); the Latin, Chinese and Japanese
// pages use fonts Windows already has.

import { app, BrowserWindow, nativeImage } from 'electron';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const arg = (name, fallback) => {
  const at = process.argv.indexOf(`--${name}`);
  return at > 0 ? process.argv[at + 1] : fallback;
};
const fonts = path.resolve(arg('fonts', path.join(here, 'fonts')));
const out = path.resolve(arg('out', path.join(here, '..', 'ocr-corpus')));
const texts = JSON.parse(fs.readFileSync(path.join(here, 'ocr-texts.json'), 'utf8'));

/** Every page: which sentences, in which face, which way. */
const VARIANTS = [
  { id: 'eng-serif', lang: 'eng', family: '"Times New Roman"', dir: 'ltr' },
  { id: 'eng-sans', lang: 'eng', family: 'Arial', dir: 'ltr' },
  { id: 'urd-naskh', lang: 'urd', family: 'NotoNaskh', file: 'NotoNaskhArabic.ttf', dir: 'rtl', script: 'Naskh' },
  { id: 'urd-nastaliq', lang: 'urd', family: 'NotoNastaliq', file: 'NotoNastaliqUrdu.ttf', dir: 'rtl', script: 'Nastaliq', leading: 2.1 },
  { id: 'ara-naskh', lang: 'ara', family: 'NotoNaskh', file: 'NotoNaskhArabic.ttf', dir: 'rtl', script: 'Naskh' },
  { id: 'hin', lang: 'hin', family: 'NotoDevanagari', file: 'NotoSansDevanagari.ttf', dir: 'ltr' },
  { id: 'chi_sim', lang: 'chi_sim', family: 'SimSun', dir: 'ltr' },
  { id: 'jpn', lang: 'jpn', family: '"Yu Gothic"', dir: 'ltr' },
];

const KINDS = [
  { kind: 'clean', dpi: 150 },
  { kind: 'clean', dpi: 300 },
  { kind: 'scan', dpi: 300 },
];

/** A page of sentences: a 6.5 inch column of 12 point type with half-inch margins. */
function page(v, { kind, dpi }) {
  const px = (inches) => Math.round(inches * dpi);
  const face = v.file ? `@font-face { font-family: ${v.family}; src: url("${pathToFileURL(path.join(fonts, v.file)).href}"); }` : '';
  // The grain: fractal noise at a low opacity, as paper and a scanner's sensor give.
  const grain = kind === 'scan'
    ? `<svg class="grain" xmlns="http://www.w3.org/2000/svg"><filter id="n"><feTurbulence type="fractalNoise" baseFrequency="0.9" numOctaves="2" seed="7"/><feColorMatrix values="0 0 0 0 0.45  0 0 0 0 0.42  0 0 0 0 0.38  0 0 0 0.55 0"/></filter><rect width="100%" height="100%" filter="url(#n)"/></svg>`
    : '';
  const tilt = kind === 'scan' ? 'transform: rotate(1.2deg); filter: blur(0.55px);' : '';
  return `<!doctype html><html lang="${v.lang}" dir="${v.dir}"><meta charset="utf-8"><style>
${face}
html, body { margin: 0; overflow: hidden; background: ${kind === 'scan' ? '#f4f1ea' : '#fff'}; }
main { box-sizing: border-box; width: ${px(7.5)}px; padding: ${px(0.5)}px; font: ${(12 / 72) * dpi}px/${v.leading || 1.45} ${v.family}; color: ${kind === 'scan' ? '#222' : '#000'}; ${tilt} }
p { margin: 0 0 ${px(0.12)}px; }
.grain { position: fixed; inset: 0; width: 100%; height: 100%; }
</style><main>${texts[v.lang].map((s) => `<p>${s}</p>`).join('')}</main>${grain}</html>`;
}

app.commandLine.appendSwitch('force-device-scale-factor', '1');
app.whenReady().then(async () => {
  // A page that never paints must not hold the machine's one window run.
  setTimeout(() => { console.log('gave up after ten minutes'); app.exit(2); }, 600000);
  fs.mkdirSync(out, { recursive: true });
  const scratch = path.join(out, '.pages');
  fs.mkdirSync(scratch, { recursive: true });
  const manifest = [];
  const win = new BrowserWindow({ show: false, width: 800, height: 600, enableLargerThanScreen: true, webPreferences: { offscreen: true } });
  win.webContents.setFrameRate(10);
  let frame = null;
  win.webContents.on('paint', (_e, _dirty, image) => { frame = image; });
  for (const v of VARIANTS) {
    for (const k of KINDS) {
      const id = `${v.id}-${k.kind}-${k.dpi}`;
      const file = path.join(scratch, `${id}.html`);
      fs.writeFileSync(file, page(v, k));
      await win.loadFile(file);
      await win.webContents.executeJavaScript('document.fonts.ready.then(() => 0)');
      const height = await win.webContents.executeJavaScript('Math.ceil(document.querySelector("main").getBoundingClientRect().height)');
      const width = Math.round(7.5 * k.dpi);
      win.setContentSize(width, height);
      await new Promise((r) => setTimeout(r, 400));
      frame = null;
      win.webContents.invalidate();
      for (let i = 0; i < 100 && !(frame && frame.getSize().width === width && frame.getSize().height === height); i++) await new Promise((r) => setTimeout(r, 100));
      if (!frame) throw new Error(id + ': no frame');
      const shot = frame;
      const picture = k.kind === 'scan' ? `${id}.jpg` : `${id}.png`;
      fs.writeFileSync(path.join(out, picture), k.kind === 'scan' ? shot.toJPEG(55) : shot.toPNG());
      // The scan decoded again as a PNG too, for readers that take pixels or PNG only.
      if (k.kind === 'scan') fs.writeFileSync(path.join(out, `${id}.png`), nativeImage.createFromBuffer(fs.readFileSync(path.join(out, picture))).toPNG());
      const truth = texts[v.lang].join('\n');
      fs.writeFileSync(path.join(out, `${id}.txt`), `${truth}\n`);
      const size = shot.getSize();
      manifest.push({ id, picture, lang: v.lang, script: v.script || null, font: v.family.replace(/"/g, ''), dpi: k.dpi, kind: k.kind, width: size.width, height: size.height });
      console.log(`${id}  ${size.width}x${size.height}`);
    }
  }
  fs.rmSync(scratch, { recursive: true, force: true });
  fs.writeFileSync(path.join(out, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`);
  console.log(`${manifest.length} pages into ${out}`);
  win.destroy();
  app.exit(0);
});
