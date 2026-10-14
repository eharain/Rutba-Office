// The pictures the document-reading window checks use, drawn as the corpus
// draws its clean pages (tools/docread-docs.mjs, the suite's renderer):
//
//   npx electron tools/docread-fixtures.js [--out packages/ocr/test/fixtures]
//
// Each fixture is the first page from a seed on that holds what its check
// needs: Western digits (Arabic-Indic for the one that shows they are not read), a face Windows ships (the Noto files are not kept
// in the repository), and for the invoice a page whose labels sit beside
// their values; the same invoice again with its total misprinted, so its
// sums do not agree. Beside each picture is <name>.json with what the page
// holds.

import { app, BrowserWindow } from 'electron';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { makeDocument, documentHtml, formatMoney } from './docread-docs.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
const arg = (name, fallback) => {
  const at = process.argv.indexOf(`--${name}`);
  return at > 0 ? process.argv[at + 1] : fallback;
};
const out = path.resolve(arg('out', path.join(here, '..', 'packages', 'ocr', 'test', 'fixtures')));
const DPI = 150;
const WINDOWS_FACES = /^(Arial|"?Times New Roman"?|Calibri|Georgia|Tahoma|Consolas|"?Courier New"?|"?Nirmala UI"?|SimSun|"?Microsoft YaHei"?|"?Yu Gothic"?|"?MS Gothic"?)$/;

const WANTED = [
  { name: 'invoice-eng', kind: 'invoice', layout: 'word', lang: 'eng', from: 7001, ok: (d) => d.currency.code === 'GBP' && !d.upper },
  // The total printed ten pounds over what its parts make: the reading must mark the figures to check.
  { name: 'invoice-eng-wrong', kind: 'invoice', layout: 'word', lang: 'eng', from: 7001, ok: (d) => d.currency.code === 'GBP' && !d.upper, wrong: true },
  // A till receipt in each of the other five languages, and one printed in Arabic-Indic digits, which are not read yet.
  { name: 'receipt-ara', kind: 'receipt', layout: 'receipt', lang: 'ara', from: 7201, ok: () => true },
  { name: 'receipt-urd', kind: 'receipt', layout: 'receipt', lang: 'urd', from: 7301, ok: () => true },
  { name: 'receipt-hin', kind: 'receipt', layout: 'receipt', lang: 'hin', from: 7401, ok: () => true },
  { name: 'receipt-chi_sim', kind: 'receipt', layout: 'receipt', lang: 'chi_sim', from: 7501, ok: () => true },
  { name: 'receipt-jpn', kind: 'receipt', layout: 'receipt', lang: 'jpn', from: 7601, ok: () => true },
  { name: 'receipt-ara-digits', kind: 'receipt', layout: 'receipt', lang: 'ara', from: 7701, digits: 'arabic', ok: () => true },
  { name: 'form-eng', kind: 'form', layout: 'form', lang: 'eng', from: 7101, ok: (d) => d.mark === 'tick' && d.truth.checkboxes.filter((c) => c.checked).length === 2 },
];

app.commandLine.appendSwitch('force-device-scale-factor', '1');
app.whenReady().then(async () => {
  setTimeout(() => { console.log('gave up after five minutes'); app.exit(2); }, 300000);
  fs.mkdirSync(out, { recursive: true });
  const scratch = fs.mkdtempSync(path.join(out, '.pages-'));
  const win = new BrowserWindow({ show: false, width: 800, height: 600, enableLargerThanScreen: true, webPreferences: { offscreen: true } });
  win.webContents.setFrameRate(10);
  let frame = null;
  win.webContents.on('paint', (_e, _dirty, image) => { frame = image; });
  try {
    for (const want of WANTED) {
      let doc = null;
      for (let seed = want.from; seed < want.from + 400 && !doc; seed++) {
        const d = makeDocument({ kind: want.kind, layout: want.layout, lang: want.lang, seed });
        if (d.digits === (want.digits || 'latin') && WINDOWS_FACES.test(d.font) && want.ok(d)) doc = d;
      }
      if (!doc) throw new Error(`${want.name}: no seed holds what it needs`);
      let html = documentHtml(doc, { dpi: DPI });
      if (want.wrong) {
        const total = formatMoney(doc.truth.fields.total, doc);
        if (!html.includes(total)) throw new Error(`${want.name}: the total is not on the page`);
        doc.truth.fields.total = Math.round((doc.truth.fields.total + 10) * 100) / 100;
        html = html.split(total).join(formatMoney(doc.truth.fields.total, doc));
      }
      html = html.replace('</style>', 'html, body { background: #fff; overflow: hidden; } main { color: #111; }</style>');
      const file = path.join(scratch, `${want.name}.html`);
      fs.writeFileSync(file, html);
      await win.loadFile(file);
      await win.webContents.executeJavaScript('document.fonts.ready.then(() => 0)');
      const size = await win.webContents.executeJavaScript('(() => { const m = document.querySelector("main"); return { w: m.offsetWidth, h: m.offsetHeight }; })()');
      win.setContentSize(size.w, size.h);
      await new Promise((r) => setTimeout(r, 400));
      const boxes = await win.webContents.executeJavaScript('[...document.querySelectorAll(".cb")].map((e) => { const r = e.getBoundingClientRect(); return { x: Math.round(r.x), y: Math.round(r.y), width: Math.round(r.width), height: Math.round(r.height) }; })');
      frame = null;
      win.webContents.invalidate();
      for (let i = 0; i < 100 && !(frame && frame.getSize().width === size.w && frame.getSize().height === size.h); i++) await new Promise((r) => setTimeout(r, 100));
      if (!frame) throw new Error(`${want.name}: no frame`);
      fs.writeFileSync(path.join(out, `${want.name}.png`), frame.toPNG());
      const truth = { name: want.name, ...doc.truth, seed: doc.seed, dpi: DPI, width: size.w, height: size.h };
      truth.checkboxes = truth.checkboxes.map((c, i) => ({ ...c, box: boxes[i] || null }));
      fs.writeFileSync(path.join(out, `${want.name}.json`), `${JSON.stringify(truth, null, 2)}\n`);
      console.log(`${want.name.padEnd(14)} seed ${doc.seed} ${truth.font.padEnd(16)} ${size.w}x${size.h}`);
    }
  } finally {
    fs.rmSync(scratch, { recursive: true, force: true });
    win.destroy();
  }
  app.exit(0);
});
