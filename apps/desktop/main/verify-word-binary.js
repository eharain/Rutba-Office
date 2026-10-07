// Word: a Word 97-2003 document, opened in full.
//
// A .doc Word itself wrote (tests/fixtures/binary/structure.doc) opens as
// its pages — the title, the table with its bold header row, the link, the
// footnote, the picture, the header and footer, the page break — not as the
// words scraped out of it. Run alone with RUTBA_VERIFY_ONLY=binary.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const FIXTURE = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..', '..', 'tests', 'fixtures', 'binary', 'structure.doc');

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the copy is opened from
 */
export async function verifyWordBinary(h, { dir }) {
  const { open, check, until, errorsIn, doc, sessionFor } = h;
  const file = path.join(dir, 'structure.doc');
  try {
    fs.copyFileSync(FIXTURE, file);
    const win = await open('word', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('doc');
    await until(() => js(`document.querySelectorAll('.wd-page table.wd-table').length > 0`), 'the table to be drawn', 8000).catch(() => {});
    if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-binary.png'), (await win.webContents.capturePage()).toPNG());

    const shown = await js(`(() => {
      const page = document.querySelector('.wd-page');
      const table = page?.querySelector('table.wd-table');
      const first = table?.querySelector('tr td, tr th');
      return {
        text: page ? page.innerText : '',
        cells: table ? [...table.querySelectorAll('tr:first-child > *')].map((c) => c.innerText.trim()) : [],
        boldHeader: first ? [...first.querySelectorAll('*')].some((n) => Number(getComputedStyle(n).fontWeight) >= 600) : false,
        pictures: page ? page.querySelectorAll('img').length : 0,
        links: page ? page.querySelectorAll('a[href^="https://example.com"], [data-href^="https://example.com"]').length : 0,
      };
    })()`);
    const model = doc.model({ id: session.id });
    check('word: a Word 97-2003 document opens as its pages, not as scraped words',
      session.converted?.from === 'doc' && !session.converted?.partial && /Structure/.test(shown.text) && /Second page/.test(shown.text),
      `from ${session.converted?.from}${session.converted?.partial ? ' (partial)' : ''}; ${shown.text.replace(/\s+/g, ' ').slice(0, 160)}`);
    check('word: the .doc\'s table is a table, its header row bold',
      shown.cells.join('|') === 'Region|Q1|Q2' && shown.boldHeader,
      `${shown.cells.join('|') || 'no table'}; bold ${shown.boldHeader}`);
    check('word: the .doc\'s picture is drawn',
      shown.pictures >= 1, `${shown.pictures} picture(s)`);
    const everything = JSON.stringify(model) + shown.text;
    const header = everything.includes('Structure header');
    const footnote = everything.includes('own words');
    check('word: the .doc\'s header and footnote came with it', header && footnote, `header ${header}; footnote ${footnote}`);

    const complaints = await errorsIn(win);
    check('word: opening a .doc reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the .doc checks ran', false, err.message);
  }
}
