// Word: Tab and Shift+Tab on the page (2.18.0, A3.2). In 2.17.1 Tab moved
// the keyboard off the page. Now Tab at the start of a list item moves it a
// level in and Shift+Tab a level out; in a table Tab walks to the next
// cell, Shift+Tab back, and Tab in the last cell adds a row; the keyboard
// stays on the page throughout. Run alone with RUTBA_VERIFY_ONLY=wordtab.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the fixture is written
 */
export async function verifyWordTabKeys({ open, check, until, wait, press, errorsIn, doc, sessionFor }, { dir }) {
  const file = path.join(dir, 'tab-keys.docx');
  fs.writeFileSync(file, buildDocx({ paragraphs: ['Opening jobs', 'Unlock the door', 'Turn the till on', { table: { rows: [['', '']] } }, 'Closing jobs'] }));
  let win = null;
  try {
    win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const id = () => sessionFor('doc').id;
    const model = () => doc.model({ id: id() });
    // A list item's level shows as its marker's indent: a level in, a larger indent.
    const level = (block) => model().listLabels?.[block]?.indentPx ?? null;
    const onPage = () => js(`Boolean(document.activeElement?.closest('.wd-page'))`);
    // The caret put at a place on the page, then the mouse released: how the editor reads it.
    const caretAt = (block, char) => js(`(() => {
      const el = document.querySelector('.wd-page [data-block="${block}"]');
      if (!el) return false;
      const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
      let left = ${char};
      let node = el, at = 0;
      for (let n = walker.nextNode(); n; n = walker.nextNode()) { if (left <= n.nodeValue.length) { node = n; at = left; break; } left -= n.nodeValue.length; }
      document.querySelector('.wd-page').focus();
      const r = document.createRange(); r.setStart(node, at); r.collapse(true);
      const s = getSelection(); s.removeAllRanges(); s.addRange(r);
      document.querySelector('.wd-page').dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      return true;
    })()`);
    const typeChars = async (text) => {
      for (const ch of text) {
        wc.sendInputEvent({ type: 'keyDown', keyCode: ch });
        wc.sendInputEvent({ type: 'char', keyCode: ch });
        wc.sendInputEvent({ type: 'keyUp', keyCode: ch });
        await wait(50);
      }
    };
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="2"]'))`), 'the paragraphs', 8000);
    await doc.apply({ id: id(), ops: [{ op: 'setSelection', anchor: { block: 1, offset: 0 }, focus: { block: 2, offset: 3 } }, { op: 'setParagraphFormat', delta: { list: 'bullet' } }] });
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="2"] .wd-marker'))`), 'the bullets', 4000).catch(() => false);

    const start = level(2);
    await caretAt(2, 0);
    await until(() => model().selection?.focus?.block === 2, 'the caret in the second item', 3000).catch(() => false);
    await press(wc, 'Tab');
    const inward = await until(() => level(2) > start, 'a level in', 3000).then(() => true).catch(() => false);
    const stayed = await onPage();
    const words = model().blocks[2].text;
    await press(wc, 'Tab', { modifiers: ['shift'] });
    const outward = await until(() => level(2) === start, 'a level out', 3000).then(() => true).catch(() => false);
    check('word: Tab at the start of a list item moves it a level in, Shift+Tab a level out, the keyboard still on the page',
      start != null && inward && outward && stayed && words === 'Turn the till on', JSON.stringify({ start, inward, outward, stayed, words, level: level(2) }));

    // A table of one row: Tab from cell to cell, a row from the last.
    const first = model().blocks.findIndex((b) => b.container);
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="${first + 1}"]'))`), 'the table', 4000).catch(() => false);
    const placed = await caretAt(first, 0);
    await typeChars('Tea');
    await press(wc, 'Tab');
    const next = await until(() => model().selection?.focus?.block === first + 1, 'the next cell', 3000).then(() => true).catch(() => false);
    await typeChars('2.40');
    const before = model().blocks.length;
    await press(wc, 'Tab');
    const grown = await until(() => model().blocks.length === before + 2 && model().selection?.focus?.block === first + 2, 'a new row', 4000).then(() => true).catch(() => false);
    await press(wc, 'Tab', { modifiers: ['shift'] });
    const back = await until(() => model().selection?.focus?.block === first + 1, 'back a cell', 3000).then(() => true).catch(() => false);
    const cells = model().blocks.slice(first, first + 2).map((b) => b.text);
    check('word: in a table Tab moves to the next cell, Tab in the last cell adds a row, Shift+Tab goes back a cell',
      next && grown && back && JSON.stringify(cells) === '["Tea","2.40"]' && (await onPage()), JSON.stringify({ first, placed, next, grown, back, cells }));

    const complaints = await errorsIn(win);
    check('word: the Tab key checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the Tab key checks ran', false, err?.message || String(err));
  }
}
