// Word: Insert → Quick Parts, building blocks.
//
// Two paragraphs selected and saved to the Quick Part Gallery under a name;
// put in again from the menu at the end of the document, a heading still a
// heading and bold still bold. Words inside a paragraph saved to AutoText
// go in at the caret as words. The Building Blocks Organizer lists both and
// deletes one; the file keeps what went in. Run alone with
// RUTBA_VERIFY_ONLY=blocks.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the fixture is written
 */
export async function verifyWordBlocks(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, doc, sessionFor } = h;
  const file = path.join(dir, 'blocks.docx');
  fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [
    { text: 'Terms of payment', style: 'Heading1' },
    { runs: [{ text: 'Invoices are due within ' }, { text: 'thirty days', bold: true }, { text: ' of their date.' }] },
    { text: 'Another paragraph.' },
  ] }));
  let win = null;
  try {
    win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const model = () => doc.model({ id: sessionFor('doc').id });
    const texts = () => model().blocks.map((b) => b.text);
    const clickTab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 'tab'`);
    // A DOM range on the page, then the mouse released — how the editor reads a selection.
    const select = (fromBlock, fromChar, toBlock, toChar) => js(`(() => {
      const at = (block, char) => {
        const el = document.querySelector('.wd-page [data-block="' + block + '"]');
        const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
        let left = char;
        for (let n = walker.nextNode(); n; n = walker.nextNode()) {
          if (left <= n.nodeValue.length) return [n, left];
          left -= n.nodeValue.length;
        }
        return [el, el.childNodes.length];
      };
      const first = document.querySelector('.wd-page [data-block="${fromBlock}"]');
      first.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
      const r = document.createRange();
      r.setStart(...at(${fromBlock}, ${fromChar}));
      r.setEnd(...at(${toBlock}, ${toChar}));
      const s = getSelection(); s.removeAllRanges(); s.addRange(r);
      document.querySelector('.wd-page').dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      return true;
    })()`);
    const quickParts = async (label) => {
      await clickTab('Insert');
      await wait(150);
      await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Quick Parts'); b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); b.click(); return 'open'; })()`);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}))`), `"${label}" in Quick Parts`, 3000);
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); if (!b || b.disabled) return 'missing'; b.click(); return 'picked'; })()`);
    };
    const nameAndOk = async (name) => {
      await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Create New Building Block"] .wd-bb-name'))`), 'Create New Building Block', 4000);
      const suggested = await js(`document.querySelector('.wd-bb-name').value`);
      await js(`(() => { const el = document.querySelector('.wd-bb-name'); const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set; setter.call(el, ${JSON.stringify(name)}); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
      await js(`[...document.querySelectorAll('.rw-dialog[aria-label="Create New Building Block"] .rw-dialog-foot button')].find((b) => b.textContent.trim() === 'OK')?.click(), 'ok'`);
      await until(() => js(`!document.querySelector('.rw-dialog[aria-label="Create New Building Block"]')`), 'the dialog to close', 4000);
      return suggested;
    };
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="2"]'))`), 'the paragraphs', 8000);
    await js(`window.rutbaOffice.store.set({ key: 'word.buildingBlocks', value: [] })`);

    // Two whole paragraphs, to the Quick Part Gallery.
    await select(0, 0, 1, model().blocks[1].text.length);
    await until(() => model().selection?.focus?.block === 1, 'the two paragraphs selected', 4000).catch(() => false);
    await quickParts('Save Selection to Quick Part Gallery…');
    const suggested = await nameAndOk('Payment terms');
    const kept = await js(`window.rutbaOffice.store.get({ key: 'word.buildingBlocks', fallback: [] })`);
    check('word: Save Selection to Quick Part Gallery keeps the paragraphs under a name, its first words offered',
      kept.length === 1 && kept[0].name === 'Payment terms' && kept[0].gallery === 'quickParts' && kept[0].paragraphs.length === 2 && /^Terms of payment/.test(suggested),
      `${JSON.stringify(kept.map((k) => [k.name, k.gallery, k.paragraphs.length]))}; offered "${suggested}"`);

    // Put in again at the end of the document, from the menu.
    await select(2, 'Another paragraph.'.length, 2, 'Another paragraph.'.length);
    await until(() => model().selection?.focus?.block === 2, 'the caret at the end', 4000).catch(() => false);
    const picked = await quickParts('Payment terms');
    const inserted = await until(() => texts().length === 5, 'the block in the document', 5000).catch(() => false);
    const after = model().blocks;
    check('word: a Quick Part from the menu goes in after the caret\'s paragraph, a heading still a heading and bold still bold',
      picked === 'picked' && inserted === true && after[3].text === 'Terms of payment' && after[3].style === 'Heading1' && after[4].runs.some((r) => r.text === 'thirty days' && r.bold),
      `${picked}; ${JSON.stringify(texts())}; style ${after[3]?.style}`);

    // Words inside a paragraph, to AutoText, and in again at the caret.
    const start = 'Invoices are due within '.length;
    await select(1, start, 1, start + 'thirty days'.length);
    await until(() => model().selection?.focus?.offset === start + 'thirty days'.length, 'the two words selected', 4000).catch(() => false);
    await quickParts('Save Selection to AutoText Gallery…');
    await nameAndOk('Due');
    await select(2, 'Another paragraph.'.length, 2, 'Another paragraph.'.length);
    await until(() => model().selection?.focus?.block === 2 && model().selection?.focus?.offset === 'Another paragraph.'.length, 'the caret at the end of the third', 4000).catch(() => false);
    await quickParts('Due');
    const words = await until(() => model().blocks[2].text === 'Another paragraph.thirty days', 'the words at the caret', 5000).catch(() => false);
    check('word: words saved to AutoText go in at the caret as words, bold still',
      words === true && model().blocks[2].runs.some((r) => r.text === 'thirty days' && r.bold), JSON.stringify(model().blocks[2].text));

    // The Organizer lists both and deletes one.
    await quickParts('Building Blocks Organizer…');
    await until(() => js(`document.querySelectorAll('.rw-dialog[aria-label="Building Blocks Organizer"] .wd-bb-row:not(.head)').length === 2`), 'both blocks listed', 4000);
    const rows = await js(`[...document.querySelectorAll('.wd-bb-row:not(.head)')].map((r) => r.textContent)`);
    await js(`[...document.querySelectorAll('.wd-bb-row:not(.head)')].find((r) => r.textContent.startsWith('Due'))?.click(), 'picked'`);
    await js(`[...document.querySelectorAll('.rw-dialog[aria-label="Building Blocks Organizer"] .rw-dialog-foot button')].find((b) => b.textContent.trim() === 'Delete')?.click(), 'deleted'`);
    const left = await until(() => js(`document.querySelectorAll('.wd-bb-row:not(.head)').length === 1`), 'one left', 3000).catch(() => false);
    await js(`[...document.querySelectorAll('.rw-dialog[aria-label="Building Blocks Organizer"] .rw-dialog-foot button')].find((b) => b.textContent.trim() === 'Close')?.click(), 'closed'`);
    const stored = await js(`window.rutbaOffice.store.get({ key: 'word.buildingBlocks', fallback: [] })`);
    check('word: the Building Blocks Organizer lists every block with its gallery, and deletes one',
      rows.length === 2 && rows.some((r) => /Payment terms\s*Quick Parts\s*General/.test(r)) && rows.some((r) => /Due\s*AutoText/.test(r)) && left === true && stored.length === 1 && stored[0].name === 'Payment terms',
      `${JSON.stringify(rows)}; left ${stored.map((s) => s.name)}`);

    // Saved: what went in is in the file.
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(wc, 's', { modifiers: ['control'] });
    await wait(1400);
    const saved = openDocx(fs.readFileSync(file)).render({ pages: false }).blocks;
    check('word: the file keeps what the building blocks put in',
      saved.length === 5 && saved[3].style === 'Heading1' && saved[2].runs.map((r) => r.text).join('') === 'Another paragraph.thirty days',
      JSON.stringify(saved.map((b) => b.runs.map((r) => r.text).join(''))));

    const complaints = await errorsIn(win);
    check('word: building blocks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the building block checks ran', false, err.message);
  } finally {
    if (win) await win.webContents.executeJavaScript(`window.rutbaOffice.store.set({ key: 'word.buildingBlocks', value: [] })`).catch(() => {});
  }
}
