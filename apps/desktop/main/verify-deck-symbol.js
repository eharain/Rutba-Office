// Presentations: Insert → Symbol puts its character where the caret was.
//
// A ribbon press takes the focus from the box being edited — which commits
// the words and closes the editor — so the check does the same: it opens a
// text box, puts the caret after its second character, lets the box go as a
// press of the ribbon would, presses Symbol, picks a character, and reads the
// slide's words in the engine and the reopened editor's caret.

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
    await js(`(() => { const el = document.querySelector('.sl-editor'); el.focus(); el.setSelectionRange(2, 2); el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'ArrowRight' })); el.blur(); return 1; })()`);
    await until(() => js(`!document.querySelector('.sl-editor')`), 'the editor to close', 3000).catch(() => {});

    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click(), 'tab'`);
    await wait(250);
    const pressed = await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Symbol'); if (!b) return 'no button'; b.click(); return 'clicked'; })()`);
    await until(() => js(`document.querySelectorAll('.wd-symbol').length > 10`), 'the Symbol dialog', 4000).catch(() => {});
    const grid = await js(`(() => { const g = document.querySelector('.wd-symbols'); return g ? getComputedStyle(g).display : null; })()`);
    await js(`(() => { [...document.querySelectorAll('.rw-dialog .ml-filter')].find((n) => n.textContent.trim() === 'Currency')?.click(); return 1; })()`);
    await until(() => js(`[...document.querySelectorAll('.wd-symbol')].some((n) => n.textContent === '€')`), 'the currency symbols', 3000).catch(() => {});
    const picked = await js(`(() => { const b = [...document.querySelectorAll('.wd-symbol')].find((n) => n.textContent === '€'); if (!b) return 'no euro'; b.click(); return 'picked'; })()`);
    const landed = await until(() => titleText() === 'Pr€ice', 'the euro in the title', 5000).then(() => true).catch(() => false);
    await until(() => js(`Boolean(document.querySelector('.sl-editor'))`), 'the editor to reopen', 3000).catch(() => {});
    const caret = await js(`(() => { const el = document.querySelector('.sl-editor'); return el ? { value: el.value, at: el.selectionStart } : null; })()`);
    check('presentations: Insert → Symbol puts its character where the caret stood in the box, and the box reopens with the caret just past it',
      opened === 'opened' && pressed === 'clicked' && picked === 'picked' && landed === true && caret?.value === 'Pr€ice' && caret?.at === 3,
      JSON.stringify({ opened, pressed, picked, text: titleText(), caret }));
    check('presentations: the Symbol dialog lays its characters in a grid in this window too', grid === 'grid', `display ${grid}`);
    await js(`(() => { document.querySelector('.sl-editor')?.blur(); return 1; })()`);
  } catch (err) {
    check('presentations: the Symbol checks ran', false, err.message);
  }
}
