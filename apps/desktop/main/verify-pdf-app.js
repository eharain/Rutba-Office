// PDF: the suite's tenth app, pressed in its window and read back from the
// files it writes. A five-page PDF made here opens with five thumbnails;
// page 3 is dragged to the front, turned right, page 5 deleted, the pages
// of a second PDF inserted after page 1, and Undo and Redo take the insert
// back and put it back; Save As writes a new file, read back here for its
// page count, the order of its pages (by the words on each) and the turn;
// it opens again the same; two files are merged into one in an empty
// window; the six pages are split into files of two; a word is found and
// marked on its page; Pictures offers the PDF to the PDF app; and the
// window stands right to left in Urdu with the pages left as they are.
// Run alone with RUTBA_VERIFY_ONLY=pdf-app.

import fs from 'node:fs';
import path from 'node:path';
import pdfWriter from '@rutba/pdf';
import pdfReader from '@rutba/pdf/reader';

const { PdfDocument } = pdfWriter;
const { openPdf } = pdfReader;

const WORDS = ['one', 'two', 'three', 'four', 'five'];

/** A PDF of pages each saying which it is, the fourth with a word to find. */
function makePdf(file, titles, extra = {}) {
  const doc = new PdfDocument({ title: path.basename(file, '.pdf'), created: '2026-10-13T09:00:00Z' });
  titles.forEach((title, i) => {
    const page = doc.addPage(i % 2 && extra.mixed ? 'Letter' : null);
    page.text(title, 72, 110, { font: 'Helvetica-Bold', size: 30 });
    page.text(extra.body?.[i] || 'The quick brown fox jumps over the lazy dog.', 72, 160, { font: 'Helvetica', size: 13 });
    page.rect(72, 190, 300, 120, { fill: ['#dbe7ff', '#ffe2d6', '#e1f5e4', '#fff3c4', '#eadcff'][i % 5], stroke: '#556070', width: 1 });
  });
  fs.writeFileSync(file, doc.toBuffer());
}

/** The first line of words on each page of a file on disk, and each page's turn. */
function readBack(file) {
  const pdf = openPdf(fs.readFileSync(file));
  return {
    count: pdf.pageCount,
    titles: Array.from({ length: pdf.pageCount }, (_, i) => pdf.pageText(i).split('\n')[0]),
    rotate: Array.from({ length: pdf.pageCount }, (_, i) => pdf.pageInfo(i).rotate),
  };
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
  const thumbs = () => js(`[...document.querySelectorAll('.pdf-thumb')].map((n) => ({ page: Number(n.dataset.page), from: n.dataset.from, rotate: Number(n.dataset.rotate), selected: n.classList.contains('selected'), drawn: Boolean(n.querySelector('canvas[data-drawn]')) }))`);
  const count = () => js(`document.querySelectorAll('.pdf-thumb').length`);
  const tab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 1`);
  const role = async (r) => {
    const sel = `[data-role="${r}"]`;
    await until(() => js(`Boolean(document.querySelector(${JSON.stringify(sel)}) && !document.querySelector(${JSON.stringify(sel)}).disabled)`), `the ${r} button`, 5000).catch(() => {});
    return js(`(() => { const b = document.querySelector(${JSON.stringify(sel)}); if (!b) return 'no button ${r}'; if (b.disabled) return 'disabled ${r}'; b.click(); return 'clicked'; })()`);
  };
  /** A press and a release on thumbnail `i`: a click, with Ctrl or Shift when asked. */
  const pick = (i, { ctrl = false, shift = false } = {}) => js(`(() => {
    const el = document.querySelectorAll('.pdf-thumb')[${i}];
    if (!el) return 'no thumbnail';
    const r = el.getBoundingClientRect();
    const at = { clientX: r.left + r.width / 2, clientY: r.top + r.height / 2, bubbles: true, cancelable: true, button: 0, ctrlKey: ${ctrl}, shiftKey: ${shift} };
    el.dispatchEvent(new MouseEvent('mousedown', at));
    el.dispatchEvent(new MouseEvent('mouseup', at));
    return 'picked';
  })()`);
  /** Thumbnail `from` dragged to just above thumbnail `to`. */
  const drag = (from, to) => js(`(async () => {
    const all = document.querySelectorAll('.pdf-thumb');
    const a = all[${from}].getBoundingClientRect();
    const b = all[${to}].getBoundingClientRect();
    const x = a.left + a.width / 2;
    const ev = (y) => ({ clientX: x, clientY: y, bubbles: true, cancelable: true, button: 0, buttons: 1 });
    all[${from}].dispatchEvent(new MouseEvent('mousedown', ev(a.top + a.height / 2)));
    const steps = 8;
    for (let i = 1; i <= steps; i++) {
      const y = a.top + a.height / 2 + ((b.top + 4) - (a.top + a.height / 2)) * (i / steps);
      window.dispatchEvent(new MouseEvent('mousemove', ev(y)));
      await new Promise((r) => setTimeout(r, 30));
    }
    const line = Boolean(document.querySelector('.pdf-drop-line'));
    window.dispatchEvent(new MouseEvent('mouseup', ev(b.top + 4)));
    return line;
  })()`);
  const setInput = (sel, value) => js(`(() => {
    const el = document.querySelector(${JSON.stringify(sel)});
    if (!el) return false;
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, ${JSON.stringify(String(value))});
    el.dispatchEvent(new Event('input', { bubbles: true }));
    return true;
  })()`);
  return { js, capture, thumbs, count, tab, role, pick, drag, setInput };
}

const answer = (kind, value) => {
  (globalThis.__rutbaCheckDialogAnswers ||= []).push({ kind, answer: value });
};

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the files are made
 */
export async function verifyPdfApp(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const root = path.join(dir, 'pdf-app');
  fs.mkdirSync(root, { recursive: true });
  const main = path.join(root, 'report.pdf');
  const extra = path.join(root, 'appendix.pdf');
  makePdf(main, WORDS.map((w) => `Page ${w}`), { body: ['The opening page.', 'Second page words.', 'The third page.', 'Here a zebra crosses the page.', 'The last page.'] });
  makePdf(extra, ['Appendix A', 'Appendix B'], { mixed: true });

  const win = await open('pdf', main);
  const d = driver(win, { until, wait });
  const five = await until(async () => (await d.count()) === 5 && (await d.js(`document.querySelectorAll('.pdf-page canvas[data-drawn]').length`)) > 0, 'five thumbnails and the first page drawn', 20000).then(() => true).catch(() => false);
  await until(async () => (await d.thumbs()).filter((t) => t.drawn).length === 5, 'every thumbnail drawn', 15000).catch(() => {});
  const look = await d.js(`(() => {
    const r = (s) => { const n = document.querySelector(s); if (!n) return null; const b = n.getBoundingClientRect(); return { left: b.left, right: b.right, top: b.top, width: b.width, height: b.height }; };
    return { side: r('.pdf-side'), view: r('.pdf-view'), page: r('.pdf-page'), title: document.title, status: document.querySelector('.rw-status')?.textContent || '' };
  })()`);
  await d.capture('pdf-01-open.png');
  check('pdf: a five-page PDF opens with five thumbnails beside its first page, drawn', five && look.side && look.view && look.side.right <= look.view.left + 2 && look.page?.width > 300,
    JSON.stringify({ five, look }));
  check('pdf: the status bar says which page of how many', /Page 1 of 5/.test(look.status), look.status);

  // Page 3 dragged to the front.
  const lineShown = await d.drag(2, 0);
  const moved = await until(async () => (await d.thumbs()).map((t) => t.page).join() === '2,0,1,3,4', 'page 3 at the front', 6000).then(() => true).catch(() => false);
  await d.capture('pdf-02-dragged.png');
  check('pdf: page 3 dragged to the front of the strip, a line showing where it lands', moved && lineShown, JSON.stringify({ order: (await d.thumbs()).map((t) => t.page), lineShown }));

  // Turned right: the page moved is the page picked.
  await d.tab('Pages');
  await wait(200);
  const rotatePress = await d.role('rotate-right');
  const turned = await until(async () => (await d.thumbs())[0]?.rotate === 90, 'the first page turned', 5000).then(() => true).catch(() => false);
  const pageTurned = await until(() => d.js(`(() => { const p = document.querySelector('.pdf-page'); return p && p.dataset.rotate === '90' && p.getBoundingClientRect().width > p.getBoundingClientRect().height; })()`), 'the page drawn on its side', 5000).then(() => true).catch(() => false);
  check('pdf: Rotate Right turns the selected page a quarter turn, in the strip and the page view', turned && pageTurned, JSON.stringify({ rotatePress, turned, pageTurned }));

  // Page five deleted.
  await d.pick(4);
  await until(async () => (await d.thumbs())[4]?.selected, 'page five picked', 3000).catch(() => {});
  const delPress = await d.role('delete-pages');
  const deleted = await until(async () => (await d.thumbs()).map((t) => t.page).join() === '2,0,1,3', 'four pages', 5000).then(() => true).catch(() => false);
  check('pdf: Delete takes the selected page out', deleted, JSON.stringify({ delPress, order: (await d.thumbs()).map((t) => t.page) }));

  // The appendix inserted after page 2 of the list (the old page 1).
  await d.pick(1);
  await until(async () => (await d.thumbs())[1]?.selected, 'the second thumbnail picked', 3000).catch(() => {});
  answer('open', [extra]);
  const insPress = await d.role('insert-file');
  const inserted = await until(async () => (await d.thumbs()).map((t) => `${t.from}:${t.page}`).join() === 'report.pdf:2,report.pdf:0,appendix.pdf:0,appendix.pdf:1,report.pdf:1,report.pdf:3', 'the appendix inserted', 10000).then(() => true).catch(() => false);
  await until(async () => (await d.thumbs()).every((t) => t.drawn), 'the inserted pages drawn', 10000).catch(() => {});
  await d.capture('pdf-03-inserted.png');
  check('pdf: Insert from File puts another PDF\'s pages after the selected page', inserted, JSON.stringify({ insPress, order: (await d.thumbs()).map((t) => `${t.from}:${t.page}`) }));

  // Undo takes the insert back, Redo puts it back.
  await d.js(`document.querySelector('.rw-quick .rw-btn[data-tip^="Undo"]')?.click(), 1`);
  const undone = await until(async () => (await d.count()) === 4, 'the insert undone', 4000).then(() => true).catch(() => false);
  await d.js(`document.querySelector('.rw-quick .rw-btn[data-tip^="Redo"]')?.click(), 1`);
  const redone = await until(async () => (await d.count()) === 6, 'the insert redone', 4000).then(() => true).catch(() => false);
  check('pdf: Undo takes the insert back and Redo puts it back', undone && redone, JSON.stringify({ undone, redone }));

  // Save As, and the file read back.
  const out = path.join(root, 'report-arranged.pdf');
  answer('save', out);
  await d.tab('Home');
  await wait(150);
  await d.js(`[...document.querySelectorAll('.rw-groups .rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Save As'))?.click(), 1`);
  const saved = await until(() => fs.existsSync(out) && fs.statSync(out).size > 0, 'the file saved', 10000).then(() => true).catch(() => false);
  const back = saved ? readBack(out) : null;
  const wantTitles = ['Page three', 'Page one', 'Appendix A', 'Appendix B', 'Page two', 'Page four'];
  check('pdf: Save As writes the pages in their new order, the third page turned, read back from the file', Boolean(back) && back.count === 6 && back.titles.join('|') === wantTitles.join('|') && back.rotate.join() === '90,0,0,0,0,0',
    JSON.stringify(back));
  const clean = await until(() => d.js(`!document.querySelector('.rw-titlebar [data-dirty="true"], .rw-title .dirty')`), 'the window saved', 3000).then(() => true).catch(() => false);
  check('pdf: the original file is left as it was', readBack(main).titles.join('|') === WORDS.map((w) => `Page ${w}`).join('|'), JSON.stringify({ clean }));

  // Opened again: the same six pages, the first on its side.
  const again = await open('pdf', out);
  const a = driver(again, { until, wait });
  const reopened = await until(async () => (await a.count()) === 6, 'the saved file reopened', 15000).then(() => true).catch(() => false);
  const firstTurned = reopened && (await a.thumbs())[0].rotate === 90;
  check('pdf: the saved file opens again with its six pages and the turn kept', reopened && firstTurned, JSON.stringify(await a.thumbs()));

  // Find: a word on page six (the old page four), marked where it is.
  await a.js(`[...document.querySelectorAll('.rw-groups .rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Find'))?.click(), 1`);
  await until(() => a.js(`Boolean(document.querySelector('[data-role="find-input"]'))`), 'the find box', 4000).catch(() => {});
  await a.setInput('[data-role="find-input"]', 'zebra');
  const found = await until(() => a.js(`document.querySelector('.pdf-find-count')?.dataset.count === '1'`), 'one match', 10000).then(() => true).catch(() => false);
  const marked = await until(() => a.js(`(globalThis.CSS?.highlights?.get('pdf-find-now')?.size || 0) === 1 && document.querySelector('.pdf-thumb.current')?.dataset.page === '5'`), 'the match marked on its page', 10000).then(() => true).catch(() => false);
  const findWords = await a.js(`document.querySelector('.pdf-find-count')?.textContent || ''`);
  await a.capture('pdf-04-find.png');
  check('pdf: Find finds a word on the sixth page and marks it there', found && marked && /1 of 1/.test(findWords), JSON.stringify({ found, marked, findWords }));
  // A word that is not there says so.
  await a.setInput('[data-role="find-input"]', 'giraffe');
  const none = await until(() => a.js(`document.querySelector('.pdf-find-count')?.textContent === 'No matches'`), 'no matches', 6000).then(() => true).catch(() => false);
  check('pdf: Find says when a word is not in the document', none);

  // Save to the file it came from asks first: Cancel leaves the file as it was, Replace writes it.
  await a.js(`document.querySelector('.pdf-find .rw-btn[data-tip^="Close Find"]')?.click(), 1`);
  await a.pick(1);
  await a.tab('Pages');
  await wait(150);
  await a.role('rotate-right');
  await until(async () => (await a.thumbs())[1]?.rotate === 90, 'the second page turned', 4000).catch(() => {});
  await a.tab('Home');
  await wait(150);
  const stamp = fs.statSync(out).mtimeMs;
  answer('message', { response: 2 });
  await a.js(`[...document.querySelectorAll('.rw-groups .rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Save:'))?.click(), 1`);
  await wait(900);
  const kept = fs.statSync(out).mtimeMs === stamp && readBack(out).rotate[1] === 0;
  answer('message', { response: 1 });
  await a.js(`[...document.querySelectorAll('.rw-groups .rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Save:'))?.click(), 1`);
  const replaced = await until(() => { try { return readBack(out).rotate.join() === '90,90,0,0,0,0'; } catch { return false; } }, 'the file replaced', 8000).then(() => true).catch(() => false);
  check('pdf: Save asks before replacing the file the pages came from: Cancel keeps it, Replace the Original writes it', kept && replaced, JSON.stringify({ kept, replaced }));

  // Split the six pages into files of two.
  const splitDir = path.join(root, 'split');
  fs.mkdirSync(splitDir, { recursive: true });
  await d.tab('Pages');
  await wait(150);
  const splitPress = await d.role('split');
  await until(() => d.js(`Boolean(document.querySelector('[data-role="split-every"]'))`), 'the Split dialog', 4000).catch(() => {});
  await d.setInput('[data-role="split-every"]', '2');
  await wait(150);
  await d.capture('pdf-05-split.png');
  answer('open', [splitDir]);
  await d.role('split-go');
  const parts = ['report-arranged_Part1.pdf', 'report-arranged_Part2.pdf', 'report-arranged_Part3.pdf'].map((n) => path.join(splitDir, n));
  const split = await until(() => parts.every((p) => fs.existsSync(p)), 'three files', 10000).then(() => true).catch(() => false);
  const partTitles = split ? parts.map((p) => readBack(p).titles.join('|')) : [];
  check('pdf: Split writes files of two pages each, in order', split && partTitles.join(' / ') === 'Page three|Page one / Appendix A|Appendix B / Page two|Page four',
    JSON.stringify({ splitPress, partTitles }));

  // Merge Files in an empty window: two PDFs combined, then saved.
  const empty = await open('pdf', null);
  const m = driver(empty, { until, wait });
  await until(() => m.js(`Boolean(document.querySelector('.pdf-empty'))`), 'the empty window', 6000).catch(() => {});
  await m.capture('pdf-06-empty.png');
  answer('open', [main, extra]);
  const mergePress = await m.js(`[...document.querySelectorAll('.pdf-empty .rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Merge Files'))?.click(), 'clicked'`);
  const merged = await until(async () => (await m.count()) === 7, 'seven pages merged', 15000).then(() => true).catch(() => false);
  const mergedFile = path.join(root, 'combined.pdf');
  answer('save', mergedFile);
  await m.js(`[...document.querySelectorAll('.rw-groups .rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Save As'))?.click(), 1`);
  const mergedSaved = await until(() => fs.existsSync(mergedFile), 'the combined file', 10000).then(() => true).catch(() => false);
  const mergedBack = mergedSaved ? readBack(mergedFile) : null;
  check('pdf: Merge Files combines two PDFs into one, saved with all seven pages in order', merged && mergedBack?.count === 7 && mergedBack.titles.join('|') === [...WORDS.map((w) => `Page ${w}`), 'Appendix A', 'Appendix B'].join('|'),
    JSON.stringify({ mergePress, merged, mergedBack }));

  const errs = [...(await errorsIn(win)), ...(await errorsIn(again)), ...(await errorsIn(empty))];
  check('pdf: no errors in the windows', errs.length === 0, errs.slice(0, 3).join(' | '));

  // Pictures offers a PDF it shows to the PDF app.
  const pics = await open('pictures', out);
  const offered = await until(() => pics.webContents.executeJavaScript(`Boolean(document.querySelector('[data-role="open-in-pdf"]'))`), 'In PDF in Pictures', 8000).then(() => true).catch(() => false);
  check('pdf: Pictures, showing a PDF, offers to open it in the PDF app', offered);

  // In Urdu: the frame right to left, the strip at the right, the pages as they are.
  const was = await d.js(`window.rutbaOffice.store.get({ key: 'language', fallback: null })`);
  try {
    await d.js(`window.rutbaOffice.store.set({ key: 'language', value: 'ur' }).then(() => 1)`);
    const ur = await open('pdf', out);
    const u = driver(ur, { until, wait });
    await until(async () => (await u.count()) === 6, 'the PDF in Urdu', 15000).catch(() => {});
    await until(async () => (await u.thumbs()).every((t) => t.drawn), 'its thumbnails', 10000).catch(() => {});
    await u.tab('صفحات');
    await wait(300);
    const frame = await u.js(`(() => {
      const r = (s) => document.querySelector(s)?.getBoundingClientRect();
      const side = r('.pdf-side');
      const view = r('.pdf-view');
      return {
        dir: document.documentElement.dir,
        right: side && view ? side.left >= view.right - 2 : false,
        pagesDir: getComputedStyle(document.querySelector('.pdf-view')).direction,
        tabs: [...document.querySelectorAll('.rw-tab')].map((n) => n.textContent.trim()),
        words: [...document.querySelectorAll('.rw-groups .rw-btn span')].map((n) => n.textContent.trim()).filter(Boolean),
      };
    })()`);
    await u.capture('pdf-07-urdu.png');
    const english = frame.words.filter((w) => /^(Rotate Left|Rotate Right|Delete|Extract|Insert from File|Split|Merge Files)$/.test(w));
    check('pdf: in Urdu the window runs right to left with the strip at the right, its words in Urdu, the pages left to right',
      frame.dir === 'rtl' && frame.right && frame.pagesDir === 'ltr' && english.length === 0 && frame.words.length >= 6 && !frame.tabs.includes('Pages'),
      JSON.stringify({ ...frame, english }));
    ur.destroy();
  } finally {
    await d.js(`window.rutbaOffice.store.set({ key: 'language', value: ${JSON.stringify(was ?? null)} }).then(() => 1)`);
  }
}
