// Read Aloud in Documents and Presentations, pressed in the window. Its bar
// is a strip over the page or the slide, as wide as that column and one row
// high, and the page keeps its width: it once stood as a column beside the
// page, half the window wide. In Documents the sentence being read is shown
// on the page, and Forward a Sentence moves what is shown to the next one.
// That sentence is the selection, so a key pressed while listening would
// put itself in the sentence's place: the words stay as they are, Space
// pauses and reads on, and Escape stops. A click in the page makes the caret
// the person's own: what they type next stops the reading and goes there.
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
 * @param {object} h the harness: open, check, until, wait, press, errorsIn
 * @param {{ dir: string }} args where the files are made (a folder of this block's own is made in it)
 */
export async function verifyReadAloudDocuments({ open, check, until, wait, press, errorsIn }, { dir }) {
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
    const reads = (s) => s.drawn && s.inPage && at(s) >= 1 && s.text.trim() === SENTENCES[at(s) - 1];
    await until(async () => reads(await shown()), 'the first sentence shown', 5000).catch(() => {});
    const first = await shown();
    await inBar('Forward a sentence');
    await inBar('Pause');
    // The bar counts on at once and the page shows the sentence a moment after: the wait is for the page.
    await until(async () => { const s = await shown(); return at(s) !== at(first) && reads(s); }, 'the next sentence shown', 5000).catch(() => {});
    const second = await shown();
    await capture('read-aloud-02-sentence.png');
    check('read-aloud: Documents shows the sentence being read on the page, and Forward a Sentence shows the next', reads(first) && reads(second) && at(second) === at(first) + 1,
      JSON.stringify({ first, second }));
    // A key pressed while listening. The page has the keyboard and the sentence is selected on it,
    // which is told first: keys sent to a page that held neither would prove nothing.
    const words = () => js(`[...document.querySelectorAll('.wd-page [data-block]')].map((b) => b.textContent).join('\\n')`);
    const reading = () => js(`Boolean(((list) => list.find((n) => ${labelKeyed('Pause', { tip: true })}(n)) || list.find((n) => (n.title || n.dataset.tip || n.getAttribute('aria-label') || '') === 'Pause'))([...document.querySelectorAll('.read-aloud-bar button')]))`);
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    const held = await shown();
    const keyboard = await js(`document.activeElement === document.querySelector('.wd-page')`);
    const wordsBefore = await words();
    await press(win.webContents, 'x', { char: true });
    await press(win.webContents, 'Backspace');
    // Space reads on, and is seen to: by then the two keys before it have been taken too.
    await press(win.webContents, ' ', { char: true });
    const readOn = await until(() => reading(), 'Space to read on', 4000).then(() => true).catch(() => false);
    await press(win.webContents, ' ', { char: true });
    const paused = await until(async () => !(await reading()), 'Space to pause', 4000).then(() => true).catch(() => false);
    await wait(400);
    const wordsAfter = await words();
    await capture('read-aloud-02b-key.png');
    check('read-aloud: a key pressed while Documents reads leaves the words as they are, and Space pauses and reads on',
      keyboard && reads(held) && wordsBefore.includes(SENTENCES[1]) && wordsAfter === wordsBefore && readOn && paused,
      JSON.stringify({ keyboard, held, readOn, paused, wordsBefore, wordsAfter }));
    await inBar('Stop reading');
    const gone = await until(() => js(`!document.querySelector('.read-aloud-bar')`), 'the bar to close', 4000).then(() => true).catch(() => false);
    const after = await box('.wd-scroll');
    check('read-aloud: Stop Reading takes the bar away and the page is as it was', gone && Boolean(after) && Math.abs(after.w - before.w) <= 2 && Math.abs(after.h - before.h) <= 2, JSON.stringify({ gone, before, after }));
    // Escape, with the keyboard in the page, stops the reading as the bar's own button does. The
    // reading is paused first: one left to run would come to its end and take the bar away itself.
    await ribbon('Read Aloud');
    await until(() => js(`Boolean(document.querySelector('.read-aloud-bar'))`), 'the bar again', 5000).catch(() => {});
    await inBar('Pause');
    await until(async () => !(await reading()), 'the reading paused', 4000).catch(() => {});
    await wait(600);
    const again = await js(`Boolean(document.querySelector('.read-aloud-bar'))`) && !(await reading());
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(win.webContents, 'Escape');
    const stopped = await until(() => js(`!document.querySelector('.read-aloud-bar')`), 'Escape to stop the reading', 4000).then(() => true).catch(() => false);
    check('read-aloud: Escape in the page stops the reading in Documents', again && stopped, JSON.stringify({ again, stopped }));
    // A click in the words, then a letter: the reading stops and the letter is typed where the click was.
    await ribbon('Read Aloud');
    await until(() => js(`Boolean(document.querySelector('.read-aloud-bar'))`), 'the bar a third time', 5000).catch(() => {});
    await inBar('Pause');
    await until(async () => !(await reading()), 'the reading paused again', 4000).catch(() => {});
    const spot = await js(`(() => { const r = document.querySelector('.wd-page [data-block="1"]')?.getBoundingClientRect(); return r ? { x: Math.round(r.left + 40), y: Math.round(r.top + r.height / 2) } : null; })()`);
    const third = () => js(`document.querySelector('.wd-page [data-block="1"]')?.textContent || ''`);
    const thirdBefore = await third();
    if (spot) {
      win.webContents.sendInputEvent({ type: 'mouseDown', x: spot.x, y: spot.y, button: 'left', clickCount: 1 });
      win.webContents.sendInputEvent({ type: 'mouseUp', x: spot.x, y: spot.y, button: 'left', clickCount: 1 });
      await wait(300);
      await press(win.webContents, 'q', { char: true });
    }
    const typed = await until(async () => { const s = await third(); return s.length === thirdBefore.length + 1 && s.includes('q'); }, 'the letter typed at the click', 5000).then(() => true).catch(() => false);
    const closed = await until(() => js(`!document.querySelector('.read-aloud-bar')`), 'the reading to stop for the edit', 4000).then(() => true).catch(() => false);
    const first2 = await js(`document.querySelector('.wd-page [data-block="0"]')?.textContent || ''`);
    check('read-aloud: a click in the page and a letter typed stop the reading in Documents, and the letter goes where the click was',
      Boolean(spot) && thirdBefore === SENTENCES[2] && typed && closed && first2 === `${SENTENCES[0]} ${SENTENCES[1]}`,
      JSON.stringify({ spot, thirdBefore, thirdAfter: await third(), typed, closed, first2 }));
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
