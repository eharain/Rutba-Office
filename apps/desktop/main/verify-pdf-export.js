// PDF: pages out as pictures, and a smaller PDF, pressed in the window and
// read back from the files written. A two-page PDF made here carries a
// large photograph shown small. Export as Pictures writes this page as a
// PNG at 72 dots an inch and every page as a JPEG at 150, each the size its
// resolution asks, named as Acrobat names them; Reduce File Size writes a
// new file a fraction of the size, its photograph drawn at the detail the
// page needs and its words as they were, and leaves the document alone.
// Run alone with RUTBA_VERIFY_ONLY=pdf-export.

import fs from 'node:fs';
import path from 'node:path';
import pdfWriter from '@rutba/pdf';
import pdfReader from '@rutba/pdf/reader';

const { PdfDocument } = pdfWriter;
const { openPdf, Stream } = pdfReader;

function makeAlbum(file) {
  const doc = new PdfDocument({ title: 'Album', created: '2026-10-14T09:00:00Z' });
  const big = 1600;
  const rgb = Buffer.alloc(big * big * 3);
  // A photograph's worth of detail: gradients with grain, which deflate cannot shrink much.
  let s = 11;
  for (let y = 0; y < big; y++) {
    for (let x = 0; x < big; x++) {
      s = (s * 1103515245 + 12345) & 0x7fffffff;
      const grain = (s >> 26) - 16;
      const o = (y * big + x) * 3;
      rgb[o] = Math.max(0, Math.min(255, (x * 255) / big + grain));
      rgb[o + 1] = Math.max(0, Math.min(255, (y * 255) / big + grain));
      rgb[o + 2] = Math.max(0, Math.min(255, 180 + grain));
    }
  }
  const photo = doc.addImage({ width: big, height: big, colorSpace: 'DeviceRGB', data: rgb });
  const one = doc.addPage();
  one.text('Holiday album', 72, 90, { font: 'Helvetica-Bold', size: 24 });
  one.image(photo, 72, 120, 216, 216); // three inches: 533 dots an inch
  doc.addPage().text('The second page has words only.', 72, 90, { font: 'Helvetica', size: 14 });
  fs.writeFileSync(file, doc.toBuffer());
}

/** A PNG's or a JPEG's width and height, read from its header. */
function pictureSize(file) {
  const b = fs.readFileSync(file);
  if (b[0] === 0x89) return { kind: 'png', width: b.readUInt32BE(16), height: b.readUInt32BE(20) };
  for (let i = 2; i + 9 < b.length;) {
    if (b[i] !== 0xff) { i += 1; continue; }
    const marker = b[i + 1];
    const len = b.readUInt16BE(i + 2);
    if (marker >= 0xc0 && marker <= 0xc3) return { kind: 'jpeg', height: b.readUInt16BE(i + 5), width: b.readUInt16BE(i + 7) };
    i += 2 + len;
  }
  return null;
}

const answer = (kind, value) => {
  (globalThis.__rutbaCheckDialogAnswers ||= []).push({ kind, answer: value });
};

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the files are made (a folder of this block's own is made in it)
 */
export async function verifyPdfExport(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const root = path.join(dir, 'pdf-export');
  fs.mkdirSync(root, { recursive: true });
  const file = path.join(root, 'album.pdf');
  makeAlbum(file);
  const original = fs.readFileSync(file);

  const win = await open('pdf', file);
  const wc = win.webContents;
  const js = (code) => wc.executeJavaScript(code);
  const capture = async (name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE) return;
    wc.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
  };
  const role = async (r) => {
    const sel = `[data-role="${r}"]`;
    await until(() => js(`Boolean(document.querySelector(${JSON.stringify(sel)}) && !document.querySelector(${JSON.stringify(sel)}).disabled)`), `the ${r} button`, 5000).catch(() => {});
    return js(`(() => { const b = document.querySelector(${JSON.stringify(sel)}); if (!b) return 'no button ${r}'; if (b.disabled) return 'disabled ${r}'; b.click(); return 'clicked'; })()`);
  };
  const choose = (sel, value) => js(`(() => {
    const el = document.querySelector(${JSON.stringify(sel)});
    if (!el) return false;
    Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(el, ${JSON.stringify(String(value))});
    el.dispatchEvent(new Event('change', { bubbles: true }));
    return true;
  })()`);
  await until(() => js(`document.querySelectorAll('.pdf-thumb').length === 2 && document.querySelectorAll('.pdf-page canvas[data-drawn]').length >= 1`), 'two pages drawn', 20000).catch(() => {});
  await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Pages')?.click(), 1`);

  // This page as a PNG at 72 dots an inch: an A4 page, 595 by 842 pixels.
  const pngDir = path.join(root, 'png');
  fs.mkdirSync(pngDir, { recursive: true });
  await role('export-pictures');
  await until(() => js(`Boolean(document.querySelector('[data-role="export-dialog"]'))`), 'the export dialog', 4000).catch(() => {});
  await js(`(document.querySelectorAll('[name="pdf-export-scope"]')[2].click(), 1)`);
  await choose('[data-role="export-dpi"]', 72);
  await capture('pdf-export-01-dialog.png');
  answer('open', [pngDir]);
  await role('export-go');
  const png = path.join(pngDir, 'album_Page_1.png');
  const wrotePng = await until(() => fs.existsSync(png), 'the PNG', 15000).then(() => true).catch(() => false);
  const pngSize = wrotePng ? pictureSize(png) : null;
  check('pdf-export: Export as Pictures writes this page as a PNG at 72 dots an inch, the page\'s size in pixels, named as Acrobat names it', wrotePng && pngSize?.kind === 'png' && pngSize.width === 595 && pngSize.height === 841,
    JSON.stringify({ wrotePng, pngSize, files: fs.readdirSync(pngDir) }));

  // Every page as a JPEG at 150.
  const jpgDir = path.join(root, 'jpg');
  fs.mkdirSync(jpgDir, { recursive: true });
  await role('export-pictures');
  await until(() => js(`Boolean(document.querySelector('[data-role="export-dialog"]'))`), 'the export dialog', 4000).catch(() => {});
  await choose('[data-role="export-format"]', 'jpeg');
  await choose('[data-role="export-dpi"]', 150);
  answer('open', [jpgDir]);
  await role('export-go');
  const jpgs = [1, 2].map((n) => path.join(jpgDir, `album_Page_${n}.jpg`));
  const wroteJpg = await until(() => jpgs.every((p) => fs.existsSync(p) && fs.statSync(p).size > 0), 'the JPEGs', 20000).then(() => true).catch(() => false);
  const jpgSizes = wroteJpg ? jpgs.map(pictureSize) : [];
  check('pdf-export: Export as Pictures writes every page as a JPEG at 150 dots an inch', wroteJpg && jpgSizes.every((s) => s?.kind === 'jpeg' && s.width === 1240 && s.height === 1753),
    JSON.stringify({ wroteJpg, jpgSizes }));

  // Reduce File Size, standard: a new file, a fraction of the size.
  const reduced = path.join(root, 'album_Reduced.pdf');
  await role('compress');
  await until(() => js(`Boolean(document.querySelector('[data-role="compress-dialog"]'))`), 'the dialog', 4000).catch(() => {});
  await capture('pdf-export-02-compress.png');
  answer('save', reduced);
  await role('compress-go');
  const wroteSmall = await until(() => fs.existsSync(reduced), 'the smaller file', 30000).then(() => true).catch(() => false);
  let back = {};
  if (wroteSmall) {
    const bytes = fs.readFileSync(reduced);
    const pdf = openPdf(bytes);
    const xo = pdf.resolve(pdf.resolve(pdf.pages()[0].dict.get('Resources')).get('XObject'));
    const pic = [...xo.keys()].map((k) => pdf.resolve(xo.get(k))).find((x) => x instanceof Stream);
    back = {
      before: original.length,
      after: bytes.length,
      pages: pdf.pageCount,
      words: /Holiday album/.test(pdf.pageText(0)) && /words only/.test(pdf.pageText(1)),
      width: pic?.dict.get('Width'),
      filter: pic?.dict.get('Filter')?.name,
      unchanged: fs.readFileSync(file).equals(original),
    };
  }
  await until(() => js(`[...document.querySelectorAll('.rw-toast')].some((n) => /album_Reduced/.test(n.textContent))`), 'the toast', 6000).catch(() => {});
  await capture('pdf-export-03-reduced.png');
  check('pdf-export: Reduce File Size writes a new file a fraction of the size, its photograph drawn at the page\'s need as JPEG, its words as they were, and leaves this one alone', wroteSmall
    && back.after < back.before / 5 && back.pages === 2 && back.words && back.width === 450 && back.filter === 'DCTDecode' && back.unchanged, JSON.stringify(back));
  const errors = await errorsIn(win);
  check('pdf-export: no error was shown', !errors.length, errors.join(' | '));
}
