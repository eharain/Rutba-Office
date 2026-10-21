// Presentations: Insert → Symbol puts its character where the caret is.
//
// The box stays open while the ribbon and the Symbol dialog are used, as in
// PowerPoint: the check opens a text box, puts the caret after its second
// character, presses Symbol on the Insert tab, picks a character, reads the
// editor's words and caret, then leaves the box and reads the slide's words
// in the engine.

import fs from 'node:fs';
import path from 'node:path';
import { buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, doc
 */
export async function verifyDeckSymbol(h, { dir }) {
  const { open, check, until, wait, doc } = h;
  const file = path.join(dir, 'symbol.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Symbols', slides: [{ layout: 'title', title: 'Price', body: 'Subtitle' }] }));
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.sl-hit[data-shape]').length > 0`), 'the slide\'s shapes', 8000);
    const session = doc.sessions().filter((s) => s.kind === 'deck' && s.path === file).pop();
    const titleText = () => {
      const shapes = doc.model({ id: session.id }).slide?.shapes || [];
      return shapes.map((s) => (s.text?.paragraphs || []).map((p) => p.plain).join('\n')).find((t) => t.startsWith('Pr')) ?? null;
    };

    // Open the title's box, put the caret after "Pr", and let the box go.
    const opened = await js(`(() => {
      const hits = [...document.querySelectorAll('.sl-hit[data-shape]')];
      const hit = hits[0];
      if (!hit) return 'no shape';
      hit.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
      return 'opened';
    })()`);
    await until(() => js(`Boolean(document.querySelector('.sl-editor'))`), 'the editor', 4000).catch(() => {});
    await js(`(() => { const el = document.querySelector('.sl-editor'); el.focus(); const t = document.createTreeWalker(el, NodeFilter.SHOW_TEXT).nextNode(); const r = document.createRange(); r.setStart(t, 2); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); return 1; })()`);
    await wait(150);

    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click(), 'tab'`);
    await wait(250);
    const pressed = await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Symbol'); if (!b) return 'no button'; b.click(); return 'clicked'; })()`);
    await until(() => js(`document.querySelectorAll('.wd-symbol').length > 10`), 'the Symbol dialog', 4000).catch(() => {});
    const grid = await js(`(() => { const g = document.querySelector('.wd-symbols'); return g ? getComputedStyle(g).display : null; })()`);
    await js(`(() => { [...document.querySelectorAll('.rw-dialog .ml-filter')].find((n) => n.textContent.trim() === 'Currency')?.click(); return 1; })()`);
    await until(() => js(`[...document.querySelectorAll('.wd-symbol')].some((n) => n.textContent === '€')`), 'the currency symbols', 3000).catch(() => {});
    const picked = await js(`(() => { const b = [...document.querySelectorAll('.wd-symbol')].find((n) => n.textContent === '€'); if (!b) return 'no euro'; b.click(); return 'picked'; })()`);
    await until(() => js(`document.querySelector('.sl-editor')?.innerText.trim() === 'Pr€ice'`), 'the euro in the editor', 4000).catch(() => {});
    const caret = await js(`(() => { const el = document.querySelector('.sl-editor'); const s = getSelection(); return el ? { value: el.innerText.trim(), at: s.anchorOffset } : null; })()`);
    // Leaving the box keeps its words.
    await js(`(() => { document.querySelector('.sl-editor')?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true })); return 1; })()`);
    const landed = await until(() => titleText() === 'Pr€ice', 'the euro in the title', 5000).then(() => true).catch(() => false);
    check('presentations: Insert → Symbol puts its character where the caret stands in the box being typed in, the caret just past it, and leaving the box keeps it',
      opened === 'opened' && pressed === 'clicked' && picked === 'picked' && landed === true && caret?.value === 'Pr€ice' && caret?.at === 3,
      JSON.stringify({ opened, pressed, picked, text: titleText(), caret }));
    check('presentations: the Symbol dialog lays its characters in a grid in this window too', grid === 'grid', `display ${grid}`);
  } catch (err) {
    check('presentations: the Symbol checks ran', false, err.message);
  }
}
