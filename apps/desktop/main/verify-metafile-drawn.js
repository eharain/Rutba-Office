// Documents, Worksheets and Presentations: a picture kept as an EMF or a WMF
// (a chart or a diagram pasted from another program, clip art from Word 97)
// is drawn, as an SVG, where the app showed an empty frame. Each app opens a
// file holding one of each; the check reads that the picture handed to the
// window is an SVG, that the window decodes it, and what it says. Run alone
// with RUTBA_VERIFY_ONLY=metafile-drawn.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { SheetView } from '@rutba/sheet-view';
import { Deck, buildPptx } from '@rutba/presentation';
import { gradientPng } from './sample-picture.js';
import { showcaseEmf, showcaseWmf, withMetafiles } from '../../../tests/fixtures/metafiles.js';

const PNG = () => gradientPng(40, 30, [60, 160, 90], [200, 50, 50]);

// In the window: each picture's href, whether it is an SVG, its size once the
// window has decoded it, and the words it carries.
const READ = (selector) => `(async () => {
  const nodes = [...document.querySelectorAll(${JSON.stringify(selector)})];
  const out = [];
  for (const n of nodes) {
    const href = n.getAttribute('src') || n.getAttribute('href') || n.getAttribute('xlink:href') || '';
    if (!href) continue;
    let type = '';
    let words = '';
    try {
      const res = await fetch(href);
      type = (res.headers.get('content-type') || '').split(';')[0];
      if (type === 'image/svg+xml') words = [...new DOMParser().parseFromString(await res.text(), 'image/svg+xml').querySelectorAll('text')].map((t) => t.textContent).join(' | ');
    } catch (e) { type = 'unfetched: ' + e.message; }
    const img = new Image();
    img.src = href;
    let w = 0;
    let h = 0;
    try { await img.decode(); w = img.naturalWidth; h = img.naturalHeight; } catch {}
    out.push({ type, w, h, words: words.slice(0, 200) });
  }
  return out;
})()`;

const drawnBoth = (seen) => Array.isArray(seen) && seen.length >= 2
  && seen.every((s) => s.type === 'image/svg+xml' && s.w > 0 && s.h > 0)
  && seen.some((s) => /Quarterly sales/.test(s.words)) && seen.some((s) => /Rutba drawing/.test(s.words));

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the files are written
 */
export async function verifyMetafileDrawn({ open, check, until, wait, errorsIn }, { dir }) {
  const metafiles = [showcaseEmf(), showcaseWmf()];
  const shot = async (wc, name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE) return;
    wc.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
  };

  // Documents.
  try {
    const file = path.join(dir, 'metafile-drawn.docx');
    const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Pictures kept as drawing calls', style: 'Heading1' }, { text: 'A chart pasted as an EMF:' }, { text: 'Clip art kept as a WMF:' }, { text: 'The end.' }] }));
    view.setSelection({ block: 1, offset: 0 });
    view.insertImage({ name: 'chart', contentType: 'image/png', data: PNG(), widthPx: 480, heightPx: 300 });
    view.setSelection({ block: 3, offset: 0 });
    view.insertImage({ name: 'clip', contentType: 'image/png', data: PNG(), widthPx: 288, heightPx: 192 });
    fs.writeFileSync(file, await withMetafiles(view.save(), metafiles));
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.wd-page img.wd-image').length >= 2 && [...document.querySelectorAll('.wd-page img.wd-image')].every((i) => i.complete && i.naturalWidth > 0)`), 'the two pictures', 10000).catch(() => {});
    const seen = await js(READ('.wd-page img.wd-image'));
    await shot(wc, 'metafile-word.png');
    check('word: an EMF and a WMF picture in a document are drawn as SVG on the page, their words in place', drawnBoth(seen), JSON.stringify(seen));
    const complaints = await errorsIn(win);
    check('word: the metafile pictures report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the metafile checks ran', false, err?.message || String(err));
  }

  // Worksheets.
  try {
    const file = path.join(dir, 'metafile-drawn.xlsx');
    const view = new SheetView(buildXlsx({ sheets: [{ name: 'Sales', rows: [['Quarter', 'Sales'], ['Q1', 120], ['Q2', 260], ['Q3', 180]] }] }));
    view.select(1, 3);
    view.insertPicture({ name: 'chart', contentType: 'image/png', data: PNG(), widthPx: 480, heightPx: 300 });
    view.select(17, 3);
    view.insertPicture({ name: 'clip', contentType: 'image/png', data: PNG(), widthPx: 288, heightPx: 192 });
    fs.writeFileSync(file, await withMetafiles(view.save(), metafiles));
    const win = await open('sheets', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.sh-drawing image').length >= 2`), 'the two pictures', 10000).catch(() => {});
    const seen = await js(READ('.sh-drawing image'));
    await shot(wc, 'metafile-sheet.png');
    check('sheets: an EMF and a WMF picture in a worksheet are drawn as SVG over the cells, their words in place', drawnBoth(seen), JSON.stringify(seen));
    const complaints = await errorsIn(win);
    check('sheets: the metafile pictures report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the metafile checks ran', false, err?.message || String(err));
  }

  // Presentations.
  try {
    const file = path.join(dir, 'metafile-drawn.pptx');
    const d = Deck.open(buildPptx({ title: 'Metafiles', slides: [{ layout: 'blank' }] }));
    d.addPicture(0, { data: PNG(), contentType: 'image/png', name: 'Chart', x: 40, y: 60, w: 480, h: 300 });
    d.addPicture(0, { data: PNG(), contentType: 'image/png', name: 'Clip', x: 580, y: 120, w: 324, h: 216 });
    fs.writeFileSync(file, await withMetafiles(d.save(), metafiles));
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.sl-slide image').length >= 2`), 'the two pictures', 10000).catch(() => {});
    const seen = await js(READ('.sl-slide image'));
    await shot(wc, 'metafile-slides.png');
    check('slides: an EMF and a WMF picture on a slide are drawn as SVG, their words in place', drawnBoth(seen), JSON.stringify(seen));
    const complaints = await errorsIn(win);
    check('slides: the metafile pictures report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('slides: the metafile checks ran', false, err?.message || String(err));
  }
}
