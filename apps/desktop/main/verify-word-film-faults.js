// Word: four faults the how-to films found in 2.24.0. Tab at the start of a
// numbered item typed a tab two letters in, because the "2." was counted as
// words; typing over a selected word with Track Changes on drew the deleted
// word twice; a table of contents took the colour off the headings and Modify
// would no longer open; and a long contents line wrapped its page number.
// Run alone with RUTBA_VERIFY_ONLY=wordfilmfaults.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { labelKeyed } from './verify-keys.js';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the fixture is written
 */
export async function verifyWordFilmFaults({ open, check, until, wait, press, errorsIn, doc, sessionFor }, { dir }) {
  const file = path.join(dir, 'film-faults.docx');
  const long = 'What the landlord and the tenant each agree to do about repairs to the outside walls of the shop';
  fs.writeFileSync(file, buildDocx({
    styles: true,
    paragraphs: [
      'Opening jobs',
      'Unlock the door',
      'look for anything left out',
      'Open the shop at nine',
      { text: 'First heading', style: 'Heading1' },
      'Words under the first heading.',
      { text: long, style: 'Heading2' },
      'Words under the second heading.',
    ],
  }));
  let win = null;
  try {
    win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const id = () => sessionFor('doc').id;
    const model = () => doc.model({ id: id() });
    const text = (block) => model().blocks[block].text;
    // A place in the words themselves, as a click leaves it: never inside the
    // number the page draws beside them.
    const place = (block, from, to = from) => js(`(() => {
      const el = document.querySelector('.wd-page [data-block="${block}"]');
      if (!el) return false;
      const nodes = [];
      const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
      for (let n = walker.nextNode(); n; n = walker.nextNode()) if (!n.parentElement.closest('.wd-marker, .wd-del')) nodes.push(n);
      const find = (left) => { for (const n of nodes) { if (left <= n.nodeValue.length) return [n, left]; left -= n.nodeValue.length; } return null; };
      const a = find(${from}), b = find(${to});
      if (!a || !b) return false;
      document.querySelector('.wd-page').focus();
      const r = document.createRange(); r.setStart(a[0], a[1]); r.setEnd(b[0], b[1]);
      const s = getSelection(); s.removeAllRanges(); s.addRange(r);
      document.querySelector('.wd-page').dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      return true;
    })()`);
    const typeChars = async (words) => {
      for (const ch of words) {
        wc.sendInputEvent({ type: 'keyDown', keyCode: ch });
        wc.sendInputEvent({ type: 'char', keyCode: ch });
        wc.sendInputEvent({ type: 'keyUp', keyCode: ch });
        await wait(60);
      }
    };
    const clickTab = (name) => js(`((list) => list.find((t) => ${labelKeyed(name)}(t)) || list.find((t) => t.textContent.trim() === ${JSON.stringify(name)}))([...document.querySelectorAll('.rw-tab')])?.click(), 'tab'`);
    const clickRibbon = (title) => js(`(() => {
      const b = ((list) => list.find((n) => ${labelKeyed(title, { tip: true, start: true })}(n) && !n.disabled) || list.find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled))([...document.querySelectorAll('.rw-ribbon .rw-btn')]);
      if (!b) return 'no button';
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);

    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="7"]'))`), 'the paragraphs', 8000);

    // 1. A numbered list: Tab at the start, and typing in the middle of a line.
    await doc.apply({ id: id(), ops: [{ op: 'setSelection', anchor: { block: 1, offset: 0 }, focus: { block: 2, offset: 3 } }, { op: 'setParagraphFormat', delta: { list: 'number' } }] });
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="2"] .wd-marker'))`), 'the numbers', 4000).catch(() => false);
    const indent = () => model().listLabels?.[2]?.indentPx ?? null;
    const start = indent();
    await place(2, 0);
    await until(() => model().selection?.focus?.block === 2 && model().selection?.focus?.offset === 0, 'the caret at the start of the second item', 3000).catch(() => false);
    const caretRead = model().selection?.focus?.offset;
    await press(wc, 'Tab');
    const inward = await until(() => indent() > start, 'a level in', 3000).then(() => true).catch(() => false);
    const noTab = text(2) === 'look for anything left out';
    check('word: Tab at the very start of a numbered item moves it a level in and types nothing, the number not counted as words',
      caretRead === 0 && inward && noTab, JSON.stringify({ caretRead, start, indent: indent(), inward, text: text(2) }));
    await press(wc, 'Tab', { modifiers: ['shift'] });
    await until(() => indent() === start, 'a level out', 3000).catch(() => false);
    await place(2, 5);
    await until(() => model().selection?.focus?.offset === 5, 'the caret five letters in', 3000).catch(() => false);
    await typeChars('Z');
    const typedAt = await until(() => text(2).includes('Z'), 'the letter', 3000).then(() => text(2).indexOf('Z')).catch(() => -1);
    await place(1, 0);
    await until(() => model().selection?.focus?.block === 1, 'the first item', 3000).catch(() => false);
    await typeChars('Q');
    const typedFirst = await until(() => text(1).includes('Q'), 'the letter', 3000).then(() => text(1).indexOf('Q')).catch(() => -1);
    check('word: letters typed on a numbered line land where the caret is, at its start and part way along', typedAt === 5 && typedFirst === 0, JSON.stringify({ typedAt, typedFirst, two: text(2), one: text(1) }));

    // 2. Track Changes: a selected word typed over.
    await doc.apply({ id: id(), ops: [{ op: 'toggleTrackChanges', on: true, author: 'Kim' }] });
    await clickTab('Review');
    await wait(200);
    await clickRibbon('Display for Review');
    await until(() => js(`Boolean(((list) => list.find((b) => ${labelKeyed("All Markup")}(b)) || list.find((b) => b.textContent.trim() === 'All Markup'))([...document.querySelectorAll('.rw-menu button')]))`), 'the markup menu', 3000).catch(() => false);
    await js(`((list) => list.find((b) => ${labelKeyed("All Markup")}(b)) || list.find((b) => b.textContent.trim() === 'All Markup'))([...document.querySelectorAll('.rw-menu button')])?.click(), 1`);
    await wait(300);
    await place(3, 9, 13);
    await until(() => { const s = model().selection; return s?.anchor?.offset === 9 && s?.focus?.offset === 13; }, 'the word selected', 3000).catch(() => false);
    await typeChars('store');
    await until(() => text(3).includes('store'), 'the new word', 4000).catch(() => false);
    await wait(400);
    const drawn = await js(`(() => {
      const el = document.querySelector('.wd-page [data-block="3"]');
      return { all: el.textContent, struck: [...el.querySelectorAll('.wd-del')].map((n) => n.textContent) };
    })()`);
    const times = (drawn.all.match(/shop/g) || []).length;
    check('word: typing over a selected word with Track Changes on shows the word struck out once and the new word beside it',
      times === 1 && drawn.struck.join('|') === 'shop' && /store/.test(drawn.all) && text(3) === 'Open the store at nine', JSON.stringify({ drawn, text: text(3) }));
    // The caret after a struck word still counts only the words.
    await place(3, 15);
    await until(() => model().selection?.focus?.offset === 15, 'the caret after the struck word', 3000).catch(() => false);
    const afterStruck = model().selection?.focus?.offset;
    check('word: a struck-out word is not counted when the caret is placed after it', afterStruck === 15, JSON.stringify({ afterStruck }));
    // A second word typed over on the line that already holds a struck word.
    await place(3, 18, 22);
    await until(() => { const s = model().selection; return s?.anchor?.offset === 18 && s?.focus?.offset === 22; }, 'the second word selected', 3000).catch(() => false);
    await typeChars('ten');
    await until(() => text(3).includes('ten'), 'the second new word', 4000).catch(() => false);
    await wait(400);
    const second = await js(`(() => {
      const el = document.querySelector('.wd-page [data-block="3"]');
      return { all: el.textContent, struck: [...el.querySelectorAll('.wd-del')].map((n) => n.textContent) };
    })()`);
    check('word: a second word typed over on the same line is struck once too, and the words land where they were typed',
      second.struck.join('|') === 'shop|nine' && text(3) === 'Open the store at ten' && (second.all.match(/shop/g) || []).length === 1 && (second.all.match(/nine/g) || []).length === 1, JSON.stringify({ second, text: text(3) }));
    await doc.apply({ id: id(), ops: [{ op: 'toggleTrackChanges', on: false }] });

    // 3 and 4. A table of contents: the headings keep their colour, Modify opens, each line keeps its number.
    const headingColour = () => js(`(() => { const el = [...document.querySelectorAll('.wd-page [data-style="Heading1"]')].find((n) => n.textContent.includes('First heading')); const run = el?.querySelector('span') || el; return run ? getComputedStyle(run).color : null; })()`);
    const before = await headingColour();
    // In a heading, with another heading after it: every heading gets its bookmark as the table is built.
    await place(4, 5);
    await until(() => model().selection?.focus?.block === 4, 'the caret in the first heading', 3000).catch(() => false);
    await clickTab('References');
    await wait(200);
    const pressed = await clickRibbon('Table of Contents');
    await until(() => (model().tableOfContents?.entries || []).length === 2, 'the two entries', 6000).catch(() => false);
    await wait(600);
    const after = await headingColour();
    const kept = model().blocks.filter((b) => /^Heading/.test(b.style || '')).map((b) => b.text.slice(0, 13));
    check('word: a table of contents put in with the caret in a heading leaves the headings where they were, in their colour', pressed === 'clicked' && Boolean(before) && before !== 'rgb(0, 0, 0)' && after === before && kept.join('|') === 'First heading|What the land', JSON.stringify({ pressed, before, after, kept }));

    await clickTab('Home');
    await wait(200);
    await js(`(() => { const b = document.querySelector('.rw-ribbon .wd-style[data-style="Heading1"]'); if (!b) return false; const r = b.getBoundingClientRect(); b.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: r.left + 8, clientY: r.top + 8 })); return true; })()`);
    await until(() => js(`[...document.querySelectorAll('.rw-menu button')].some((b) => b.textContent.trim().startsWith('Modify'))`), 'the style menu', 3000).catch(() => false);
    await js(`((list) => list.find((b) => ${labelKeyed("Modify", { start: true })}(b)) || list.find((b) => b.textContent.trim().startsWith('Modify')))([...document.querySelectorAll('.rw-menu button')])?.click(), 1`);
    const modifies = await until(() => js(`Boolean(document.querySelector('.wd-modify-style'))`), 'the Modify Style dialog', 4000).then(() => true).catch(() => false);
    const shownColour = modifies ? await js(`document.querySelector('.wd-modify-style [data-role="style-size"]')?.value`) : null;
    check('word: Modify still opens on a heading style after a table of contents went in', modifies && shownColour === '16', JSON.stringify({ modifies, shownColour }));
    await press(wc, 'Escape');
    await wait(300);

    const lines = await js(`[...document.querySelectorAll('.wd-page [data-style^="TOC"]')].map((p) => {
      const tab = p.querySelector('.wd-tab');
      if (!tab) return { text: p.textContent, tab: false };
      const range = document.createRange();
      range.setStartAfter(tab); range.setEnd(p, p.childNodes.length);
      const num = range.getBoundingClientRect();
      const t = tab.getBoundingClientRect();
      const page = p.closest('.wd-page').getBoundingClientRect();
      const pad = parseFloat(getComputedStyle(p.closest('.wd-page')).paddingRight) || 0;
      return { text: p.textContent.slice(0, 30), number: range.toString(), sameLine: num.top < t.bottom && num.bottom > t.top, insideMargin: num.right <= page.right - pad + 1, tabWidth: t.width };
    })`);
    const entries = lines.filter((l) => l.text);
    check('word: every contents line keeps its page number on the line with the leader, inside the margin', entries.length === 2 && entries.every((l) => l.tab !== false && l.sameLine && l.insideMargin && /^\d+$/.test(l.number)), JSON.stringify(lines));

    const complaints = await errorsIn(win);
    check('word: none of that reported an error', complaints.length === 0, JSON.stringify(complaints));
  } catch (err) {
    check('word: the film faults check ran to its end', false, err?.stack || String(err));
  } finally {
    if (win && !win.isDestroyed()) win.destroy();
    fs.rmSync(file, { force: true });
  }
}
