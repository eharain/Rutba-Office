// Word: References → Table of Contents — a real field, not text.
//
// Insert reads each heading's on-screen page from the page it is actually
// drawn on (pages.js); the entries and their pages show up in the model at
// once. Adding a heading afterwards and pressing Update Table rebuilds the
// list; the field lands in the saved file exactly as the engine writes it.
// Run alone with RUTBA_VERIFY_ONLY=toc.
import fs from 'node:fs';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 */
export async function verifyWordToc(h, { file }) {
  const { open, check, until, wait, press, errorsIn, capture, doc, sessionFor } = h;
  if (!file) return check('word: the table of contents fixture was made', false, 'no fixture');
  try {
    const win = await open('word', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('doc');
    const model = () => doc.model({ id: session.id });

    const clickRibbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no button ' + ${JSON.stringify(title)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const clickTab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 'tab'`);
    const caretIn = (block, atEnd) => js(`(() => {
      const b = document.querySelector('.wd-page [data-block="${block}"]');
      if (!b) return false;
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
      const r = document.createRange(); r.selectNodeContents(b); r.collapse(${atEnd ? 'false' : 'true'});
      const s = getSelection(); s.removeAllRanges(); s.addRange(r);
      document.querySelector('.wd-page').dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      return true;
    })()`);

    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="0"]'))`), 'the title paragraph', 8000);

    // The caret at the end of the title, then References → Table of Contents.
    await caretIn(0, true);
    await wait(200);
    await clickTab('References');
    await wait(200);
    const pressed = await clickRibbon('Table of Contents');

    const inserted = await until(
      () => (model().tableOfContents?.entries || []).length === 3,
      'the table of contents to hold three entries',
      6000
    ).catch(() => false);
    const entries = model().tableOfContents?.entries || [];
    check(
      'word: References → Table of Contents lists the headings, each with a page number read off the screen',
      pressed === 'clicked' && inserted === true && entries.map((e) => e.text).join('|') === 'Introduction|Background|Method' && entries.every((e) => e.page),
      JSON.stringify(entries)
    );
    check(
      'word: "Method" landed on a later page than "Introduction" — the page numbers on the entries are real, not all "1"',
      Boolean(entries[0]) && Boolean(entries[2]) && entries[0].page !== entries[2].page,
      JSON.stringify(entries.map((e) => e.page))
    );

    // The entries are drawn on the page too, not just held in the model.
    const onPage = await js(`[...document.querySelectorAll('.wd-link')].map((l) => l.closest('[data-block]')?.textContent || '').filter(Boolean)`);
    check(
      'word: the entries are drawn on the page, as clickable text with a page number',
      onPage.some((t) => t.includes('Introduction')) && onPage.some((t) => t.includes('Method')),
      JSON.stringify(onPage)
    );

    // An off-screen window hands back its last painted frame: ask for a
    // fresh one first, or the capture shows the page before the table.
    win.webContents.invalidate();
    await wait(700);
    await capture(win, 'word-toc.png');

    // A new heading, added the same way the picture check adds a picture:
    // through the window's own bridge but without going through the ribbon,
    // so it lands in the model and the file — then Ctrl+Z/Ctrl+Y so the
    // window's own React state, which never heard about an apply it did not
    // dispatch itself, catches up before the next button press reads it.
    const before = model();
    const lastBlock = before.blocks.length - 1;
    const lastLen = (before.blocks[lastBlock]?.text || '').length;
    await js(`(async () => {
      const all = await window.rutbaOffice.doc.sessions({});
      const mine = all.filter((s) => s.kind === 'doc').pop();
      await window.rutbaOffice.doc.apply({ id: mine.id, ops: [
        { op: 'setSelection', anchor: { block: ${lastBlock}, offset: ${lastLen} }, focus: { block: ${lastBlock}, offset: ${lastLen} } },
        { op: 'splitParagraph' },
        { op: 'setParagraphFormat', delta: { styleId: 'Heading1' } },
        { op: 'insertText', text: 'Results' },
      ] });
      return 1;
    })()`);
    await js(`document.querySelector('.wd-page')?.focus(), 'focused'`);
    await press(win.webContents, 'z', { modifiers: ['control'] });
    await wait(300);
    await press(win.webContents, 'y', { modifiers: ['control'] });
    await until(() => (model().blocks || []).some((b) => b.text === 'Results' && b.style === 'Heading1'), 'the new heading in the model', 6000).catch(() => {});

    await clickTab('References');
    await wait(200);
    const updated = await clickRibbon('Update Table');
    const afterUpdate = await until(
      () => (model().tableOfContents?.entries || []).map((e) => e.text).join('|') === 'Introduction|Background|Method|Results',
      'Update Table to add the new heading',
      6000
    ).catch(() => false);
    check(
      'word: Update Table adds an entry for the heading added afterwards',
      updated === 'clicked' && afterUpdate === true,
      JSON.stringify(model().tableOfContents?.entries)
    );

    await clickRibbon('Save');
    const tocInFile = () => {
      try {
        return openDocx(fs.readFileSync(file)).doc.doc.tableOfContents();
      } catch {
        return null;
      }
    };
    await until(() => tocInFile()?.entries?.length === 4, 'the table of contents to land in the file', 8000).catch(() => {});
    const savedToc = tocInFile();
    check(
      'word: the saved file writes the field, with all four entries — a table of contents field, not frozen text',
      Boolean(savedToc) && savedToc.entries.map((e) => e.text).join('|') === 'Introduction|Background|Method|Results',
      JSON.stringify(savedToc)
    );

    const complaints = await errorsIn(win);
    check('word: the table of contents checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    // toc.docx is this check's own fixture, rebuilt fresh by makeFixtures on
    // every run — unlike report.docx there is nothing shared to restore.
  } catch (err) {
    check('word: the table of contents checks ran', false, err.message);
  }
}
