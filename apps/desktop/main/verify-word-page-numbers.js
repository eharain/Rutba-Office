// Word: page numbers from the header and footer box (2.18.0, A3.10, A3.8).
// In 2.17.1 the box wrote "{PAGE} of {PAGES}" into the file as words, Tab
// left the box, and "Page number" then "Of total" gave "{PAGE} {PAGE} of
// {PAGES}". Now Tab types a tab, Of total after a page number finishes it,
// each page's footer draws its own number and the count, and the saved file
// carries Word's PAGE and NUMPAGES fields. Run alone with
// RUTBA_VERIFY_ONLY=wordpagenumbers.
import fs from 'node:fs';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn
 * @param {{ dir: string }} args where the fixture is written
 */
export async function verifyWordPageNumbers({ open, check, until, wait, press, errorsIn }, { dir }) {
  const file = path.join(dir, 'page-numbers.docx');
  fs.writeFileSync(file, buildDocx({ paragraphs: Array.from({ length: 120 }, (_, i) => `Clause ${i + 1}. The shop opens at nine and closes at half past five, every day but Sunday, and the stock is counted on the last Friday of each month.`) }));
  let win = null;
  try {
    win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const box = `document.querySelector('.rw-dialog [data-role="band-text"]')`;
    const value = () => js(`${box}?.value ?? null`);
    const click = (label) => js(`(() => { const b = [...document.querySelectorAll('.rw-dialog button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); b?.click(); return Boolean(b); })()`);
    await until(() => js(`document.querySelectorAll('.wd-sheet').length >= 3 || document.querySelectorAll('[data-page]').length >= 3`), 'three pages', 10000).catch(() => false);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click(), 1`);
    await wait(200);
    await js(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((b) => b.textContent.trim() === 'Footer')?.click(), 1`);
    await until(() => js(`Boolean(${box})`), 'the Footer box', 4000);

    // "Larkspur", Tab, Page number, Of total.
    await js(`${box}.focus(), 1`);
    for (const ch of 'Larkspur') {
      wc.sendInputEvent({ type: 'keyDown', keyCode: ch });
      wc.sendInputEvent({ type: 'char', keyCode: ch });
      wc.sendInputEvent({ type: 'keyUp', keyCode: ch });
      await wait(40);
    }
    await press(wc, 'Tab');
    const tabbed = await value();
    const stillIn = await js(`document.activeElement === ${box}`);
    await click('Page number');
    await click('Of total');
    await wait(150);
    const written = await value();
    check('word: in the Footer box Tab types a tab, and Page number then Of total reads "{PAGE} of {PAGES}" with one space',
      tabbed === 'Larkspur\t' && stillIn && written === 'Larkspur\t{PAGE} of {PAGES}', JSON.stringify({ tabbed, stillIn, written }));

    await click('Apply');
    await until(() => js(`!${box}`), 'the box to close', 4000).catch(() => false);
    const footers = async () => js(`[...document.querySelectorAll('.wd-band.wd-footer')].map((f) => f.textContent.replace(/\\s+/g, ' ').trim())`);
    await until(async () => (await footers()).length >= 3 && /1 of \d/.test((await footers())[0]), 'the numbers drawn', 6000).catch(() => false);
    const drawn = await footers();
    const count = drawn.length;
    const cells = await js(`(() => { const line = document.querySelector('.wd-band.wd-footer .wd-band-tabbed'); return line ? [...line.children].map((c) => [c.textContent, getComputedStyle(c).textAlign]) : null; })()`);
    check('word: each page\'s footer draws its own number and the count, the number at the centre stop',
      count >= 3 && drawn.every((f, k) => f === `Larkspur${k + 1} of ${count}`) && cells?.[1]?.[0] === `1 of ${count}` && cells?.[1]?.[1] === 'center',
      JSON.stringify({ drawn: drawn.slice(0, 3), cells }));

    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(wc, 's', { modifiers: ['control'] });
    let detail = '';
    const saved = await until(() => {
      try {
        const pkg = OoxmlPackage.read(fs.readFileSync(file));
        const name = pkg.partNames().find((n) => /^word\/footer\d+\.xml$/.test(n));
        const xml = name ? pkg.text(name) : '';
        detail = xml.slice(xml.indexOf('<w:p'), xml.indexOf('<w:p') + 400);
        return /<w:fldSimple w:instr=" PAGE ">/.test(xml) && /<w:fldSimple w:instr=" NUMPAGES ">/.test(xml) && !xml.includes('{PAGE') && /<w:tab w:val="center"/.test(xml);
      } catch { return false; }
    }, 'the save', 6000).then(() => true).catch(() => false);
    check('word: the saved footer carries Word\'s PAGE and NUMPAGES fields and its centre and right stops, not the braces', saved, detail);

    const complaints = await errorsIn(win);
    check('word: the page number checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the page number checks ran', false, err?.message || String(err));
  }
}
