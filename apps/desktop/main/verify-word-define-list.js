// Word: Home → Multilevel List → Define New Multilevel List — the first
// level made Chapter I. and linked to Heading 1, the second legal 1.1
// linked to Heading 2: the preview says so as it is typed, and on OK every
// heading of either style is numbered by its style, the body left alone.
// The saved file carries the list and the links as Word writes them. Run
// alone with RUTBA_VERIFY_ONLY=worddefinelist.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn
 * @param {{ dir: string }} args where the fixture is written
 */
export async function verifyWordDefineList({ open, check, until, wait, press, errorsIn }, { dir }) {
  const file = path.join(dir, 'word-define-list.docx');
  fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [
    { text: 'Intro', style: 'Heading1' },
    { text: 'A body paragraph, in no list.' },
    { text: 'Scope', style: 'Heading2' },
    { text: 'Terms', style: 'Heading2' },
    { text: 'Method', style: 'Heading1' },
    { text: 'Steps', style: 'Heading2' },
  ] }));
  const WANT = ['Chapter I.', null, '1.1', '1.2', 'Chapter II.', '2.1'];
  try {
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const markers = () => js(`[0, 1, 2, 3, 4, 5].map((i) => document.querySelector('.wd-page [data-block="' + i + '"] .wd-marker')?.textContent.trim() || null)`);
    const setField = (role, value) => js(`(() => {
      const el = document.querySelector('.wd-define-list [data-role="${role}"]');
      if (!el) return false;
      const proto = el.tagName === 'SELECT' ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(value)});
      el.dispatchEvent(new Event(el.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true }));
      return true;
    })()`);
    const label = (level) => js(`document.querySelector('.wd-define-preview [data-level="${level}"] .wd-define-label')?.textContent || null`);

    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="0"]'))`), 'the paragraphs', 8000);
    const at = await js(`(() => { const r = document.querySelector('.wd-page [data-block="0"]').getBoundingClientRect(); return { x: Math.round(r.left + 12), y: Math.round(r.top + r.height / 2) }; })()`);
    wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    await wait(300);
    await js(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Home'); t?.click(); return 1; })()`);
    await wait(200);

    // Multilevel List → Define New Multilevel List…
    await js(`(() => { const b = document.querySelector('.wd-multilevel'); b?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b?.click(); return 1; })()`);
    await until(() => js(`[...document.querySelectorAll('.rw-menu button')].some((b) => b.textContent.includes('Define New Multilevel List'))`), 'the menu item', 3000).catch(() => {});
    await js(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.includes('Define New Multilevel List'))?.click(), 1`);
    const opened = await until(() => js(`Boolean(document.querySelector('.wd-define-list [data-role="define-format"]'))`), 'the dialog', 4000).then(() => true).catch(() => false);

    // Level 1: I, II, III as "Chapter %1.", linked to Heading 1.
    await setField('define-format', 'upperRoman');
    await setField('define-text', 'Chapter %1.');
    await setField('define-style', 'Heading1');
    await wait(150);
    // Level 2: legal, "%1.%2", linked to Heading 2.
    await js(`document.querySelectorAll('.wd-define-list .wd-define-level')[1]?.click(), 1`);
    await wait(150);
    await setField('define-text', '%1.%2');
    await js(`(() => { const c = document.querySelector('.wd-define-list .wd-define-legal'); if (c && !c.checked) c.click(); return 1; })()`);
    await setField('define-style', 'Heading2');
    await setField('define-name', 'Chapters');
    await wait(200);
    const preview = [await label(0), await label(1)];
    const previewed = preview[0] === 'Chapter I.' && preview[1] === '1.1';
    if (process.env.RUTBA_VERIFY_CAPTURE) { await wait(300); fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-define-list.png'), (await wc.capturePage()).toPNG()); }
    check('word: Define New Multilevel List opens from the gallery, and its preview numbers each level as it is set (Chapter I., legal 1.1)',
      opened && previewed, JSON.stringify({ opened, preview }));

    await js(`document.querySelector('.wd-define-list-ok')?.click(), 1`);
    const numbered = await until(async () => JSON.stringify(await markers()) === JSON.stringify(WANT), 'the headings numbered', 6000).then(() => true).catch(() => false);
    const shown = await markers();
    check('word: on OK every Heading 1 reads Chapter I., II. and every Heading 2 1.1, 1.2, 2.1 by its style, the body paragraph in no list',
      numbered, JSON.stringify(shown));

    // Saved as Word writes it: the list named, the styles linked.
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(wc, 's', { modifiers: ['control'] });
    let detail = 'not saved';
    const saved = await until(() => {
      try {
        const reopened = openDocx(fs.readFileSync(file));
        const frame = reopened.render({ pages: false });
        const labels = frame.blocks.map((b) => frame.listLabels?.[b.index]?.label ?? null);
        const numbering = reopened.doc.doc.pkg.text('word/numbering.xml');
        const styles = reopened.doc.doc.pkg.text('word/styles.xml');
        detail = JSON.stringify(labels);
        return JSON.stringify(labels) === JSON.stringify(WANT)
          && /<w:name w:val="Chapters"\/>/.test(numbering) && /<w:pStyle w:val="Heading1"\/>/.test(numbering) && /<w:isLgl\/>/.test(numbering)
          && /w:styleId="Heading1"[\s\S]*?<w:numPr><w:numId w:val="\d+"\/><\/w:numPr>/.test(styles);
      } catch { return false; }
    }, 'the save', 6000).then(() => true).catch(() => false);
    check('word: the saved document carries the list, named, its levels linked to the heading styles as Word writes them', saved, detail);

    const complaints = await errorsIn(win);
    check('word: Define New Multilevel List reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the Define New Multilevel List checks ran', false, err?.message || String(err));
  }
}
