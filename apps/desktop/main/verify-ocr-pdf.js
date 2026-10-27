// PDF: Recognise Text, pressed in the window. A three-page PDF made here
// holds an English page that is only a picture of text, a page that already
// has text, and an Urdu page that is only a picture; Recognise Text's dialog
// offers the six languages with Urdu marked beta, reads the two scanned
// pages on the main process's threads (the main process's own loop never
// held up while it does) and leaves the text page alone; then Find finds
// words on both scans, English and Urdu, the recognised words can be
// selected and copied, the pages look as they did, Undo takes the layer away
// and Redo puts it back, and Save As writes a file that begins with every
// byte of the original, the layer appended after it.
// Run alone with RUTBA_VERIFY_ONLY=ocr-pdf.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import pdfWriter from '@rutba/pdf';
import pdfReader from '@rutba/pdf/reader';
import { decodePng } from '@rutba/ocr/png';

const { PdfDocument } = pdfWriter;
const { openPdf } = pdfReader;

/** A page the suite drew, from @rutba/ocr's own test pages, as RGB samples. */
function scan(name) {
  const file = fileURLToPath(new URL(`../test/fixtures/${name}.png`, import.meta.resolve('@rutba/ocr')));
  const { width, height, data } = decodePng(fs.readFileSync(file));
  const rgb = Buffer.alloc(width * height * 3);
  for (let i = 0, o = 0; i < data.length; i += 4, o += 3) { rgb[o] = data[i]; rgb[o + 1] = data[i + 1]; rgb[o + 2] = data[i + 2]; }
  return { width, height, data: rgb };
}

function makeScannedPdf(file) {
  const doc = new PdfDocument({ title: 'Scanned letter', created: '2026-10-13T09:00:00Z' });
  const eng = scan('eng-serif');
  const urd = scan('urd-naskh');
  const engRef = doc.addImage({ ...eng, colorSpace: 'DeviceRGB' });
  const urdRef = doc.addImage({ ...urd, colorSpace: 'DeviceRGB' });
  // 150 dots per inch: each picture as wide on the page as it was drawn.
  doc.addPage().image(engRef, 36, 48, (eng.width / 150) * 72, (eng.height / 150) * 72);
  const typed = doc.addPage();
  typed.text('This page was typed, not scanned.', 72, 110, { font: 'Helvetica-Bold', size: 20 });
  doc.addPage().image(urdRef, 36, 48, (urd.width / 150) * 72, (urd.height / 150) * 72);
  fs.writeFileSync(file, doc.toBuffer());
}

function driver(win, { until, wait }) {
  const wc = win.webContents;
  const js = (code) => wc.executeJavaScript(code);
  const capture = async (name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE) return;
    wc.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
  };
  const tip = (label) => js(`(() => {
    const b = [...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith(${JSON.stringify(label)}));
    if (!b) return 'no button ${label}';
    if (b.disabled) return 'disabled ${label}';
    b.click();
    return 'clicked';
  })()`);
  const role = (r) => js(`(() => { const b = document.querySelector('[data-role="${r}"]'); if (!b) return 'no ${r}'; if (b.disabled) return 'disabled ${r}'; b.click(); return 'clicked'; })()`);
  const setInput = (sel, value) => js(`(() => {
    const el = document.querySelector(${JSON.stringify(sel)});
    if (!el) return false;
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, ${JSON.stringify(String(value))});
    el.dispatchEvent(new Event('input', { bubbles: true }));
    return true;
  })()`);
  /**
   * Find's count for a word, once the search for that word has finished: the
   * count says which words its hits are for (data-for). Read 400 ms after any
   * count showed, a busy machine gave the word before's count, or none.
   */
  const find = async (word) => {
    await setInput('[data-role="find-input"]', word);
    await until(() => js(`document.querySelector('.pdf-find-count')?.dataset.for === ${JSON.stringify(word)}`), `find ${word}`, 30000).catch(() => {});
    return js(`Number(document.querySelector('.pdf-find-count')?.dataset.count || 0)`);
  };
  return { wc, js, capture, tip, role, setInput, find };
}

const answer = (kind, value) => {
  (globalThis.__rutbaCheckDialogAnswers ||= []).push({ kind, answer: value });
};

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the files are made
 */
export async function verifyOcrPdf(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const root = path.join(dir, 'ocr-pdf');
  fs.mkdirSync(root, { recursive: true });
  const file = path.join(root, 'scanned-letter.pdf');
  makeScannedPdf(file);
  const original = fs.readFileSync(file);

  const win = await open('pdf', file);
  const d = driver(win, { until, wait });
  try {
    await until(() => d.js(`document.querySelectorAll('.pdf-thumb').length === 3 && document.querySelectorAll('.pdf-page canvas[data-drawn]').length > 0`), 'three pages drawn', 20000).catch(() => {});
    // The typed page's own words come in after its picture, and in a busy
    // full run they were counted before they arrived.
    await until(() => d.js(`(document.querySelectorAll('.pdf-page .textLayer')[1]?.textContent.length || 0) > 0`), 'the typed page’s text', 15000).catch(() => {});
    const before = await d.js(`[...document.querySelectorAll('.pdf-page .textLayer')].map((l) => l.textContent.length)`);

    // The dialog: three ways to choose pages, six languages, Urdu marked beta and explained when ticked.
    const pressed = await d.tip('Recognise Text');
    const dialog = await until(() => d.js(`Boolean(document.querySelector('[data-role="ocr-dialog"]'))`), 'the Recognise Text dialog', 5000).then(() => true).catch(() => false);
    const offered = await d.js(`(() => {
      const langs = [...document.querySelectorAll('.ocr-lang')].map((n) => ({ code: n.dataset.code, on: n.querySelector('input').checked, beta: Boolean(n.querySelector('.ocr-beta')) }));
      return { langs, scopes: document.querySelectorAll('[name="ocr-scope"]').length };
    })()`);
    check('ocr-pdf: Recognise Text opens a dialog offering all pages, the selected ones or this one, and six languages with only Urdu marked beta', pressed === 'clicked' && dialog && offered.scopes === 3
      && offered.langs.map((l) => l.code).join() === 'eng,urd,ara,hin,chi_sim,jpn' && offered.langs.filter((l) => l.beta).map((l) => l.code).join() === 'urd', JSON.stringify({ pressed, dialog, offered }));
    // English and Urdu, whatever was chosen last time.
    await d.js(`(() => {
      for (const n of document.querySelectorAll('.ocr-lang')) {
        const want = n.dataset.code === 'eng' || n.dataset.code === 'urd';
        const box = n.querySelector('input');
        if (box.checked !== want) box.click();
      }
      return 1;
    })()`);
    await until(() => d.js(`Boolean(document.querySelector('[data-role="ocr-urdu-note"]'))`), 'the Urdu note', 3000).catch(() => {});
    const note = await d.js(`document.querySelector('[data-role="ocr-urdu-note"]')?.textContent || ''`);
    await d.capture('ocr-pdf-01-dialog.png');
    check('ocr-pdf: ticking Urdu says it is beta, good on Naskh and not yet reliable on Nastaliq print', /beta/i.test(note) && /Nastaliq/.test(note), note);

    // Reading: the main process's loop is timed all along; a hold of more than a fraction of a second would show.
    let worst = 0;
    let last = Date.now();
    const ticker = setInterval(() => { const now = Date.now(); worst = Math.max(worst, now - last - 25); last = now; }, 25);
    const started = Date.now();
    const go = await d.role('ocr-go');
    const progressSeen = await until(() => d.js(`Boolean(document.querySelector('[data-role="ocr-progress"]'))`), 'the progress', 5000).then(() => true).catch(() => false);
    await d.capture('ocr-pdf-02-progress.png');
    const finished = await until(() => d.js(`!document.querySelector('[data-role="ocr-progress"]') && !document.querySelector('[data-role="ocr-dialog"]')`), 'recognition to finish', 90000).then(() => true).catch(() => false);
    clearInterval(ticker);
    const seconds = (Date.now() - started) / 1000;
    const stats = await d.js(`window.rutbaOffice.ocr.stats()`);
    const toasts = () => d.js(`[...document.querySelectorAll('.rw-toast')].map((n) => n.textContent).join(' | ')`);
    await until(async () => /recognised/i.test(await toasts()), 'the toast', 5000).catch(() => {});
    const toast = await toasts();
    check('ocr-pdf: the two scanned pages are read on threads of the pool, the typed page left alone, and the main process is never held', go === 'clicked' && finished && stats.done >= 2 && stats.threads >= 1 && /2 pages/.test(toast) && worst < 400,
      JSON.stringify({ go, progressSeen, finished, seconds, stats, toast, worstMainHoldMs: worst }));

    // Find, on the English scan and on the Urdu one.
    await d.tip('Find');
    await until(() => d.js(`Boolean(document.querySelector('[data-role="find-input"]'))`), 'the find box', 4000).catch(() => {});
    const invoice = await d.find('invoice');
    const record = await d.find('records');
    const urdu = await d.find('کمپیوٹر');
    const toUrdu = await until(() => d.js(`document.querySelector('.pdf-thumb.current')?.dataset.page === '2' && (globalThis.CSS?.highlights?.get('pdf-find-now')?.size || 0) === 1`), 'the Urdu match shown on its page', 20000).then(() => true).catch(() => false);
    await d.capture('ocr-pdf-03-find-urdu.png');
    const english = await d.find('Nothing is sent');
    await until(() => d.js(`(globalThis.CSS?.highlights?.get('pdf-find-now')?.size || 0) === 1`), 'the match marked', 20000).catch(() => {});
    await d.capture('ocr-pdf-04-find.png');
    check('ocr-pdf: Find finds words on the scanned pages, English and Urdu, and marks them where they are', invoice >= 1 && record >= 1 && urdu >= 1 && toUrdu && english === 1,
      JSON.stringify({ invoice, record, urdu, toUrdu, english }));

    // The words over the first scan selected and copied, as on any PDF with text.
    await d.js(`document.querySelector('[data-tip^="Close Find"]')?.click(), 1`);
    // The first scan's words laid on its page again once Find has closed, and
    // selected and copied in the same turn as they are read: on a busy machine
    // the page, back in view, could draw its words again between a read that
    // found them and the copy that followed. What each page holds is counted in
    // that same turn too, for the same reason.
    const selectCopy = `(() => {
      const counts = [...document.querySelectorAll('.pdf-page .textLayer')].map((l) => l.textContent.length);
      const layer = document.querySelector('.pdf-page[data-index="0"] .textLayer[data-ready]');
      if (!layer) return { text: '', counts };
      const r = document.createRange();
      r.selectNodeContents(layer);
      getSelection().removeAllRanges();
      getSelection().addRange(r);
      const text = getSelection().toString();
      getSelection().removeAllRanges();
      return { text, counts };
    })()`;
    const inOrder = (text) => /Rutba Office reads the words/.test(text.replace(/\s+/g, ' ')) && /delivery/.test(text);
    let copied = '';
    let after = [];
    await until(async () => {
      const read = (await d.js(selectCopy)) || {};
      copied = String(read.text || '');
      after = read.counts || [];
      return inOrder(copied);
    }, 'the first scan\'s words selected and copied', 60000).catch(() => {});
    check('ocr-pdf: the words recognised on a scan are there to select and copy, in order', inOrder(copied), copied.slice(0, 160));
    check('ocr-pdf: the scanned pages had no text before and have it now; the typed page has what it had', before[0] === 0 && after[0] > 200 && before[1] === after[1] && before[1] > 0, JSON.stringify({ before, after }));

    // Undo takes the layer away, Redo puts it back.
    await d.tip('Undo');
    await d.tip('Find');
    const undone = await until(async () => (await d.find('invoice')) === 0, 'no match after Undo', 8000).then(() => true).catch(() => false);
    await d.js(`document.querySelector('[data-tip^="Close Find"]')?.click(), 1`);
    await d.tip('Redo');
    await d.tip('Find');
    const redone = await until(async () => (await d.find('invoice')) >= 1, 'the match back after Redo', 8000).then(() => true).catch(() => false);
    await d.js(`document.querySelector('[data-tip^="Close Find"]')?.click(), 1`);
    check('ocr-pdf: Undo takes the recognised text away and Redo puts it back', undone && redone, JSON.stringify({ undone, redone }));

    // Save As: the original's bytes first, the layer after them, read back with its words.
    const saved = path.join(root, 'scanned-letter-recognised.pdf');
    answer('save', saved);
    await d.tip('Save As');
    const wrote = await until(() => fs.existsSync(saved) && fs.statSync(saved).size > original.length, 'the saved file', 10000).then(() => true).catch(() => false);
    let readBack = {};
    if (wrote) {
      const bytes = fs.readFileSync(saved);
      const pdf = openPdf(bytes);
      readBack = {
        prefix: bytes.subarray(0, original.length).equals(original),
        pages: pdf.pageCount,
        repaired: pdf.repaired,
        layer: /3 Tr/.test(pdf.pageContent(0).toString('latin1')) && /3 Tr/.test(pdf.pageContent(2).toString('latin1')),
        typedUntouched: !/3 Tr/.test(pdf.pageContent(1).toString('latin1')),
      };
    }
    check('ocr-pdf: Save As writes the original bytes unchanged with the text layer appended as an incremental update', wrote && readBack.prefix && readBack.pages === 3 && !readBack.repaired && readBack.layer && readBack.typedUntouched,
      JSON.stringify({ wrote, ...readBack }));
    await d.capture('ocr-pdf-05-saved.png');
    const errors = await errorsIn(win);
    check('ocr-pdf: no error was shown', !errors.length, errors.join(' | '));
  } finally {
    try { win.webContents.executeJavaScript(`localStorage.removeItem('rutba-ocr-languages'), 1`).catch(() => {}); } catch { /* closed */ }
  }
}
