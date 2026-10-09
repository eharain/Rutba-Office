// Word: Design → Watermark → Picture watermark… — a picture from this
// computer, drawn washed out behind the words on the page, and saved in the
// header as Word saves one. Run alone with RUTBA_VERIFY_ONLY=picwatermark.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { gradientPng } from './sample-picture.js';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn
 * @param {{ dir: string }} args where the files are written
 */
export async function verifyWordWatermarkPicture({ open, check, until, wait, press, errorsIn }, { dir }) {
  const file = path.join(dir, 'watermark-picture.docx');
  const picture = path.join(dir, 'logo.png');
  try {
    fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [{ text: 'Watermarked', style: 'Heading1' }, { text: 'These words stand in front of the picture.' }] }));
    fs.writeFileSync(picture, gradientPng(600, 300, [40, 90, 170], [230, 200, 120]));
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="1"]'))`), 'the paragraphs', 8000);
    await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Design')?.click(); return 1; })()`);
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Watermark')))`), 'the Watermark button', 4000);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Watermark')); b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 1; })()`);
    await until(() => js(`[...document.querySelectorAll('.rw-menu button')].some((b) => b.textContent.trim() === 'Picture watermark…')`), 'the Picture watermark item', 3000);
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [picture] });
    await js(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'Picture watermark…').click(), 1`);
    const drawn = await until(() => js(`(() => { const img = document.querySelector('.wd-watermark-pic'); return Boolean(img && img.complete && img.naturalWidth > 0 && img.classList.contains('washout') && Number(getComputedStyle(img).opacity) < 0.5); })()`), 'the picture behind the page', 8000).then(() => true).catch(() => false);
    const place = await js(`(() => { const img = document.querySelector('.wd-watermark-pic'); const page = document.querySelector('.wd-page'); if (!img || !page) return null; const a = img.getBoundingClientRect(); const p = page.getBoundingClientRect(); return { centreOff: Math.round(Math.abs((a.left + a.width / 2) - (p.left + p.width / 2))), w: Math.round(a.width), behind: getComputedStyle(img).zIndex }; })()`);
    if (process.env.RUTBA_VERIFY_CAPTURE) { wc.invalidate(); await wait(500); fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-watermark-picture.png'), (await wc.capturePage()).toPNG()); }
    check('word: Design → Watermark → Picture watermark puts the picture behind the words, washed out and centred on the page',
      drawn && place && place.centreOff < 30 && place.w > 100, JSON.stringify({ drawn, place }));

    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(wc, 's', { modifiers: ['control'] });
    const saved = await until(() => {
      try {
        const d = openDocx(fs.readFileSync(file)).doc.doc;
        const mark = d.headerFooters().watermark;
        return Boolean(mark?.picturePart && mark.washout && d.pkg.has(mark.picturePart));
      } catch { return false; }
    }, 'the save', 6000).then(() => true).catch(() => false);
    check('word: the saved document keeps the picture watermark in its header, as Word writes one', saved, saved ? 'WordPictureWatermark in the header' : 'not in the file');
    const complaints = await errorsIn(win);
    check('word: the picture watermark reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the picture watermark checks ran', false, err?.message || String(err));
  }
}
