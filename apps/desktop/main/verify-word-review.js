// Word: comments on the words they are about, and tracked changes one at a
// time (2.18.0, A3.5 and A3.6). In 2.17.1 a new comment hung on the end of
// its paragraph unsigned, with nothing marked; Next and Previous went to a
// paragraph's start; This Change took every change in the paragraph, and
// Accept was greyed out unless the caret stood in a changed line. Now the
// selected words are marked and the comment signed as a tracked change is;
// Next selects each change, This Change takes that one only, and Accept
// All works from anywhere. Run alone with RUTBA_VERIFY_ONLY=wordreview.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the fixture is written
 */
export async function verifyWordReview({ open, check, until, wait, press, errorsIn, doc, sessionFor }, { dir }) {
  const file = path.join(dir, 'review.docx');
  fs.writeFileSync(file, buildDocx({ paragraphs: ['Count the till before you lock the door.', 'Open the shop at nine', 'Lock the back door'] }));
  const author = process.env.RUTBA_SCREENS_AUTHOR || os.userInfo().username || 'Rutba Office user';
  let win = null;
  try {
    win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const id = () => sessionFor('doc').id;
    const model = () => doc.model({ id: id() });
    // Select from one place on the page to another, then release the mouse: how the editor reads it.
    const selectOnPage = (block, from, to) => js(`(() => {
      const el = document.querySelector('.wd-page [data-block="${block}"]');
      if (!el) return false;
      const at = (n) => { const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT); let left = n; for (let t = w.nextNode(); t; t = w.nextNode()) { if (left <= t.nodeValue.length) return [t, left]; left -= t.nodeValue.length; } return [el, 0]; };
      document.querySelector('.wd-page').focus();
      const [a, ao] = at(${from}); const [b, bo] = at(${to});
      const r = document.createRange(); r.setStart(a, ao); r.setEnd(b, bo);
      const s = getSelection(); s.removeAllRanges(); s.addRange(r);
      document.querySelector('.wd-page').dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      return true;
    })()`);
    // Review has two Next buttons: the comments' and the tracked changes'.
    const nextChange = () => js(`(() => { const b = document.querySelector('.rw-ribbon .rw-btn[title^="Next tracked change"]') || [...document.querySelectorAll('.rw-ribbon .rw-btn')].filter((n) => n.textContent.trim() === 'Next').pop(); b?.click(); return Boolean(b); })()`);
    const tab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 1`);
    const ribbon = (label) => js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!b) return false; b.click(); return !b.disabled; })()`);
    const menuPick = (label) => js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!b || b.disabled) return false; b.click(); return true; })()`);
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="2"]'))`), 'the paragraphs', 8000);

    // A comment over "till".
    await selectOnPage(0, 10, 14);
    await until(() => { const s = model().selection; return s && s.from.offset === 10 && s.to.offset === 14; }, 'the selection read', 3000).catch(() => false);
    await tab('Review');
    await wait(200);
    await ribbon('New Comment');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog textarea'))`), 'the comment box', 4000);
    await js(`(() => { const el = document.querySelector('.rw-dialog textarea'); Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(el, 'Both tills?'); el.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await wait(100);
    await js(`[...document.querySelectorAll('.rw-dialog button')].find((b) => b.textContent.trim() === 'Add')?.click(), 1`);
    let marked = null;
    const shown = await until(async () => {
      marked = await js(`[...document.querySelectorAll('.wd-page .wd-commented')].map((n) => n.textContent).join('')`);
      return marked === 'till';
    }, 'the words marked', 5000).then(() => true).catch(() => false);
    const c = model().comments?.[0] || null;
    check('word: a new comment over selected words marks those words on the page and is signed as a tracked change is',
      shown && c?.author === author && c?.text === 'Both tills?' && model().blocks[0].structural === false,
      JSON.stringify({ marked, comment: c && { author: c.author, text: c.text }, expected: author }));

    // Track Changes on, then two changes in the second line and one in the
    // third, made on the page: "back " taken out, " sharp" typed, "shop " taken out.
    await ribbon('Track Changes');
    await until(() => model().trackRevisions === true, 'recording on', 3000).catch(() => false);
    const typeChars = async (text) => {
      for (const ch of text) {
        wc.sendInputEvent({ type: 'keyDown', keyCode: ch });
        wc.sendInputEvent({ type: 'char', keyCode: ch });
        wc.sendInputEvent({ type: 'keyUp', keyCode: ch });
        await wait(50);
      }
    };
    await selectOnPage(2, 9, 14);
    await wait(200);
    await press(wc, 'Backspace');
    await until(() => model().blocks[2].runs.some((r) => r.del), 'the third line\'s deletion', 3000).catch(() => false);
    await selectOnPage(1, 21, 21);
    await wait(200);
    await typeChars(' sharp');
    await until(() => model().blocks[1].text === 'Open the shop at nine sharp', 'the words typed', 3000).catch(() => false);
    await selectOnPage(1, 9, 14);
    await wait(200);
    await press(wc, 'Backspace');
    await until(() => model().blocks[1].text === 'Open the at nine sharp', 'the second line\'s deletion', 3000).catch(() => false);
    const made = model().blocks.map((b) => b.text);
    // The caret back at the start of the first line, which has no change.
    await selectOnPage(0, 0, 0);
    await wait(200);

    const sel = () => { const s = model().selection; return s ? [s.from.block, s.from.offset, s.to.offset] : null; };
    await tab('Review');
    await wait(200);
    await nextChange();
    const firstSel = await until(() => JSON.stringify(sel()) === JSON.stringify([1, 9, 9]), 'the first change selected', 4000).then(() => sel()).catch(() => sel());
    await nextChange();
    const secondSel = await until(() => JSON.stringify(sel()) === JSON.stringify([1, 16, 22]), 'the second change selected', 4000).then(() => sel()).catch(() => sel());
    const highlighted = await js(`getSelection().toString()`);
    check('word: Next selects each tracked change in turn, the inserted words selected on the page',
      JSON.stringify(firstSel) === '[1,9,9]' && JSON.stringify(secondSel) === '[1,16,22]' && highlighted === ' sharp',
      JSON.stringify({ made, firstSel, secondSel, highlighted }));

    // Accept This Change: the insertion only.
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Accept'); b?.click(); return 1; })()`);
    await until(() => js(`[...document.querySelectorAll('.rw-menu button')].some((b) => b.textContent.trim() === 'Accept This Change')`), 'the Accept menu', 3000).catch(() => false);
    await menuPick('Accept This Change');
    const left = await until(() => model().blocks.filter((b) => b.tracked).length === 2 && !model().blocks[1].runs.some((r) => r.ins), 'the insertion accepted', 4000)
      .then(() => true).catch(() => false);
    const stillDeleted = model().blocks[1].runs.some((r) => r.del) && model().blocks[2].runs.some((r) => r.del);
    check('word: Accept This Change keeps the selected insertion and leaves both deletions waiting', left && stillDeleted,
      JSON.stringify(model().blocks.slice(1).map((b) => b.runs.map((r) => (r.ins ? '+' : r.del ? '-' : '') + (r.text || r.del?.text || '')))));

    // Accept All with the caret in the first line, which has no change.
    await selectOnPage(0, 3, 3);
    await wait(300);
    const enabled = await js(`!([...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Accept')?.disabled)`);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Accept'); b?.click(); return 1; })()`);
    await until(() => js(`[...document.querySelectorAll('.rw-menu button')].some((b) => b.textContent.trim() === 'Accept All Changes')`), 'the Accept menu', 3000).catch(() => false);
    const thisOff = await js(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'Accept This Change')?.disabled === true`);
    await menuPick('Accept All Changes');
    const allGone = await until(() => !model().blocks.some((b) => b.tracked), 'every change accepted', 4000).then(() => true).catch(() => false);
    check('word: Accept All works with the caret in a line with no change (This Change greyed there)',
      enabled && thisOff && allGone && model().blocks[2].text === 'Lock the door',
      JSON.stringify({ enabled, thisOff, allGone, texts: model().blocks.map((b) => b.text) }));

    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(wc, 's', { modifiers: ['control'] });
    let xml = '';
    const saved = await until(() => {
      try {
        xml = OoxmlPackage.read(fs.readFileSync(file)).text('word/document.xml');
        return /<w:commentRangeStart w:id="1"\/><w:r>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t[^>]*>till<\/w:t><\/w:r><w:commentRangeEnd w:id="1"\/>/.test(xml);
      } catch { return false; }
    }, 'the save', 6000).then(() => true).catch(() => false);
    check('word: the saved file carries Word\'s comment range round "till"', saved, xml.slice(xml.indexOf('<w:body>'), xml.indexOf('<w:body>') + 420));

    const complaints = await errorsIn(win);
    check('word: the review checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the review checks ran', false, err?.message || String(err));
  }
}
