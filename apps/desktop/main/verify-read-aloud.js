// Read Aloud in Documents and Presentations, pressed in the window. Its bar
// is a strip over the page or the slide, as wide as that column and one row
// high, and the page keeps its width: it once stood as a column beside the
// page, half the window wide. In Documents the sentence being read is shown
// on the page, and Forward a Sentence moves what is shown to the next one.
// Run alone with RUTBA_VERIFY_ONLY=read-aloud.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation';
import { labelKeyed } from './verify-keys.js';

const SENTENCES = ['The first sentence is read here.', 'A second one follows it.', 'The next paragraph holds the third.'];

function kit({ wait }, dir) {
  const root = path.join(dir, 'read-aloud');
  fs.mkdirSync(root, { recursive: true });

  const tools = (wc) => {
    const js = (code) => wc.executeJavaScript(code);
    const tab = (label) => js(`((list) => list.find((t) => ${labelKeyed(label)}(t)) || list.find((t) => t.textContent.trim() === ${JSON.stringify(label)}))([...document.querySelectorAll('.rw-tab')])?.click(), 1`);
    const ribbon = (label) => js(`(() => { const b = ((list) => list.find((n) => ${labelKeyed(label)}(n)) || list.find((n) => n.textContent.trim() === ${JSON.stringify(label)}))([...document.querySelectorAll('.rw-ribbon .rw-btn')]); if (!b) return 'no button'; b.click(); return 'clicked'; })()`);
    // The bar's own buttons, by the tooltip's key or its English words.
    const inBar = (label) => js(`(() => { const b = ((list) => list.find((n) => ${labelKeyed(label, { tip: true })}(n)) || list.find((n) => (n.title || n.dataset.tip || n.getAttribute('aria-label') || '') === ${JSON.stringify(label)}))([...document.querySelectorAll('.read-aloud-bar button')]); if (!b) return 'no button'; b.click(); return 'clicked'; })()`);
    const box = (sel) => js(`(() => { const r = document.querySelector(${JSON.stringify(sel)})?.getBoundingClientRect(); return r ? { x: Math.round(r.left), y: Math.round(r.top), w: Math.round(r.width), h: Math.round(r.height) } : null; })()`);
    const capture = async (name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      wc.invalidate();
      await wait(500);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
    };
    return { js, tab, ribbon, inBar, box, capture };
  };
  // A strip: over its column from edge to edge, one row high, and the column under it as wide as before.
  const strip = (bar, column, before) => Boolean(bar && column && before) && bar.h <= 48 && Math.abs(bar.w - column.w) <= 2 && Math.abs(column.w - before.w) <= 2 && bar.y <= column.y;
  return { root, tools, strip };
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the files are made (a folder of this block's own is made in it)
 */
export async function verifyReadAloudDocuments({ open, check, until, wait, errorsIn }, { dir }) {
  const { root, tools, strip } = kit({ wait }, dir);
  try {
    const file = path.join(root, 'read.docx');
    fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [{ text: `${SENTENCES[0]} ${SENTENCES[1]}` }, { text: SENTENCES[2] }] }));
    const win = await open('word', file);
    const { js, tab, ribbon, inBar, box, capture } = tools(win.webContents);
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="1"]'))`), 'the paragraphs', 8000);
    const before = await box('.wd-scroll');
    await tab('Review');
    const pressed = await until(async () => (await ribbon('Read Aloud')) === 'clicked', 'the Read Aloud button', 5000).then(() => true).catch(() => false);
    await until(() => js(`Boolean(document.querySelector('.read-aloud-bar'))`), 'the bar', 5000).catch(() => {});
    await inBar('Pause');
    await wait(300);
    const bar = await box('.read-aloud-bar');
    const column = await box('.wd-scroll');
    await capture('read-aloud-01-documents.png');
    check('read-aloud: in Documents the bar is a strip over the page, and the page keeps its width', pressed && strip(bar, column, before),
      JSON.stringify({ pressed, bar, column, before }));

    // What the page shows as being read: the words selected on it, drawn where they stand.
    const shown = () => js(`(() => {
      const sel = window.getSelection();
      const text = sel && sel.rangeCount ? sel.toString() : '';
      const r = sel && sel.rangeCount ? sel.getRangeAt(0).getBoundingClientRect() : null;
      const where = document.querySelector('.read-aloud-where')?.textContent || '';
      const marked = [...document.querySelectorAll('.wd-reading')].map((n) => n.textContent).join('');
      return { text, drawn: Boolean(r && r.width > 0 && r.height > 0), inPage: Boolean(sel && sel.anchorNode && document.querySelector('.wd-page')?.contains(sel.anchorNode)), where, marked };
    })()`);
    const at = (s) => Number((/\d+/.exec(s.where) || [0])[0]);
    await until(async () => Boolean((await shown()).text), 'the first sentence shown', 5000).catch(() => {});
    const first = await shown();
    await inBar('Forward a sentence');
    await inBar('Pause');
    await until(async () => { const s = await shown(); return at(s) !== at(first) && Boolean(s.text); }, 'the next sentence shown', 5000).catch(() => {});
    const second = await shown();
    await capture('read-aloud-02-sentence.png');
    const reads = (s) => s.drawn && s.inPage && at(s) >= 1 && s.text.trim() === SENTENCES[at(s) - 1];
    check('read-aloud: Documents shows the sentence being read on the page, and Forward a Sentence shows the next', reads(first) && reads(second) && at(second) === at(first) + 1,
      JSON.stringify({ first, second }));
    await inBar('Stop reading');
    const gone = await until(() => js(`!document.querySelector('.read-aloud-bar')`), 'the bar to close', 4000).then(() => true).catch(() => false);
    const after = await box('.wd-scroll');
    check('read-aloud: Stop Reading takes the bar away and the page is as it was', gone && Boolean(after) && Math.abs(after.w - before.w) <= 2 && Math.abs(after.h - before.h) <= 2, JSON.stringify({ gone, before, after }));
    const errors = await errorsIn(win);
    check('read-aloud: Documents reported nothing', !errors.length, errors.join(' | '));
  } catch (err) {
    check('read-aloud: the Documents checks ran', false, err?.stack || String(err));
  }
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the files are made
 */
export async function verifyReadAloudPresentations({ open, check, until, wait, errorsIn }, { dir }) {
  const { root, tools, strip } = kit({ wait }, dir);
  try {
    const file = path.join(root, 'read.pptx');
    fs.writeFileSync(file, buildPptx({ title: 'Read', slides: [{ layout: 'title', title: SENTENCES[0] }, { layout: 'title', title: SENTENCES[1] }] }));
    const win = await open('slides', file);
    const { js, tab, ribbon, box, capture } = tools(win.webContents);
    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slides', 8000);
    const before = await box('.sl-stage');
    await tab('Review');
    const pressed = await until(async () => (await ribbon('Read Aloud')) === 'clicked', 'the Read Aloud button', 5000).then(() => true).catch(() => false);
    // The button opens a short menu: this slide, its notes, or both.
    const chose = await until(() => js(`(() => { const m = ((list) => list.find((n) => ${labelKeyed('This slide')}(n)) || list.find((n) => n.textContent.trim() === 'This slide'))([...document.querySelectorAll('.rw-menu button')]); if (!m) return false; m.click(); return true; })()`), 'This slide', 4000).then(() => true).catch(() => false);
    await until(() => js(`Boolean(document.querySelector('.read-aloud-bar'))`), 'the bar', 5000).catch(() => {});
    await wait(300);
    const bar = await box('.read-aloud-bar');
    const column = await box('.sl-stage');
    await capture('read-aloud-03-presentations.png');
    check('read-aloud: in Presentations the bar is a strip over the slide, and the slide keeps its width', pressed && chose && strip(bar, column, before),
      JSON.stringify({ pressed, chose, bar, column, before }));
    const errors = await errorsIn(win);
    check('read-aloud: Presentations reported nothing', !errors.length, errors.join(' | '));
  } catch (err) {
    check('read-aloud: the Presentations checks ran', false, err?.stack || String(err));
  }
}
