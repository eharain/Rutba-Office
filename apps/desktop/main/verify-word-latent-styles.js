// Word: headings in a file whose styles part has only Normal (2.18.0, A3.3).
// In 2.17.1 the gallery offered only Normal and References > Add Text's
// levels failed. Now Word's built-in Title and Heading 1 to 3 are offered
// after the file's own styles; picking Heading 1 in the gallery, or Level 2
// under Add Text, makes the paragraph that heading, and the saved file has
// the style written in, its Normal left as it was. Run alone with
// RUTBA_VERIFY_ONLY=wordlatentstyles.
import fs from 'node:fs';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';

const NORMAL = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="22"/></w:rPr></w:style>';
const ONLY_NORMAL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${NORMAL}</w:styles>`;

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn
 * @param {{ dir: string }} args where the fixture is written
 */
export async function verifyWordLatentStyles({ open, check, until, wait, press, errorsIn }, { dir }) {
  const file = path.join(dir, 'only-normal.docx');
  const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: ['Opening jobs', 'Unlock the door and turn the till on.', 'Closing jobs'] }));
  pkg.write_('word/styles.xml', ONLY_NORMAL);
  fs.writeFileSync(file, pkg.write());
  try {
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const look = (i) => js(`(() => { const el = document.querySelector('.wd-page [data-block="${i}"]'); const run = el?.querySelector('span') || el; const s = run ? getComputedStyle(run) : null; return s ? { size: parseFloat(s.fontSize), weight: Number(s.fontWeight) } : null; })()`);
    const caretIn = (i) => js(`(() => {
      const el = document.querySelector('.wd-page [data-block="${i}"]');
      const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
      const first = walker.nextNode();
      document.querySelector('.wd-page').focus();
      const r = document.createRange(); r.setStart(first, 1); r.collapse(true);
      const s = getSelection(); s.removeAllRanges(); s.addRange(r);
      document.querySelector('.wd-page').dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      return 1;
    })()`);
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="2"]'))`), 'the page', 8000);
    const offered = await until(() => js(`['Normal', 'Title', 'Heading1', 'Heading2', 'Heading3'].every((id) => document.querySelector('.rw-ribbon .wd-style[data-style="' + id + '"]'))`), 'the styles gallery', 6000).then(() => true).catch(() => false);
    check('word: a file whose styles part has only Normal is offered Title and Heading 1 to 3 in the gallery', offered,
      String(await js(`[...document.querySelectorAll('.rw-ribbon .wd-style')].map((b) => b.dataset.style).join(', ')`)));

    // Heading 1 from the gallery, on the first paragraph.
    const plain = await look(0);
    await caretIn(0);
    await wait(150);
    await js(`document.querySelector('.rw-ribbon .wd-style[data-style="Heading1"]')?.click(), 1`);
    let h1 = null;
    const heading = await until(async () => { h1 = await look(0); return h1 && h1.size > plain.size + 4 && h1.weight >= 600; }, 'Heading 1 on the page', 5000).then(() => true).catch(() => false);
    check('word: Heading 1 picked from the gallery makes the paragraph a heading', heading, JSON.stringify({ plain, h1 }));

    // References > Add Text > Level 2, on the third paragraph.
    await caretIn(2);
    await wait(150);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'References')?.click(), 1`);
    await wait(250);
    await js(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((b) => b.textContent.trim() === 'Add Text')?.click(), 1`);
    await until(() => js(`[...document.querySelectorAll('.rw-menu button')].some((b) => b.textContent.trim().startsWith('Level 2'))`), 'the Add Text menu', 3000).catch(() => {});
    await js(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith('Level 2'))?.click(), 1`);
    let h2 = null;
    const level = await until(async () => { h2 = await look(2); return h2 && h2.size > plain.size + 2 && h2.weight >= 600; }, 'Heading 2 on the page', 5000).then(() => true).catch(() => false);
    const failed = await js(`[...document.querySelectorAll('.rw-toast')].map((n) => n.textContent).join(' | ')`);
    check('word: References > Add Text > Level 2 makes the paragraph Heading 2, with no error', level && !/style/i.test(failed), JSON.stringify({ h2, failed }));

    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(wc, 's', { modifiers: ['control'] });
    let detail = '';
    const saved = await until(() => {
      try {
        const out = OoxmlPackage.read(fs.readFileSync(file));
        const xml = out.text('word/styles.xml');
        const body = out.text('word/document.xml');
        detail = xml.slice(0, 600);
        return xml.includes(NORMAL) && /w:styleId="Heading1"/.test(xml) && /w:styleId="Heading2"/.test(xml) && !/w:styleId="Heading3"/.test(xml)
          && /<w:pStyle w:val="Heading1"\/>/.test(body) && /<w:pStyle w:val="Heading2"\/>/.test(body);
      } catch { return false; }
    }, 'the save', 6000).then(() => true).catch(() => false);
    check('word: the saved file has Heading 1 and 2 written into its styles, only those, its Normal as it was', saved, detail);

    const complaints = await errorsIn(win);
    check('word: the latent style checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the latent style checks ran', false, err?.message || String(err));
  }
}
