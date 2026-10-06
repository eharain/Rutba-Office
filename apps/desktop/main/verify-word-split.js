// Documents: View → Split, pressed on the ribbon — the document in two
// panes, one over the other: the lower pane is the same document, scrolled
// on its own; what is typed above shows below; a click on a paragraph in
// the lower pane puts the caret there above; pressed again, one pane.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait, doc, sessionFor
 */
export async function verifyWordSplit(h, { dir }) {
  const { open, check, until, wait, doc } = h;
  const file = path.join(dir, 'split.docx');
  try {
    const paragraphs = Array.from({ length: 60 }, (_, i) => ({ text: `Paragraph ${i + 1}: the northern region grew fastest in absolute terms, and the margin it opened held through the autumn.` }));
    fs.writeFileSync(file, buildDocx({ styles: true, paragraphs }));
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'doc' && s.path === file).pop();
    const model = () => doc.model({ id: session.id });
    const press = () => js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'View')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Split');
      if (!b || b.disabled) return 'no button'; b.click(); return 'clicked';
    })()`);

    await until(() => js(`document.querySelectorAll('.wd-page [data-block]').length >= 60`), 'the page', 8000);
    const pressed = await press();
    const copied = await until(() => js(`document.querySelectorAll('.wd-split-pane [data-block]').length >= 60`), 'the lower pane', 5000).then(() => true).catch(() => false);
    // Each pane scrolls on its own.
    const own = await js(`(() => {
      const top = document.querySelector('.wd-scroll.split-top');
      const low = document.querySelector('.wd-split-pane');
      top.scrollTop = 0;
      low.scrollTop = low.scrollHeight;
      return { top: top.scrollTop, low: low.scrollTop, topH: Math.round(top.getBoundingClientRect().height), lowH: Math.round(low.getBoundingClientRect().height) };
    })()`);
    check('documents: View → Split shows the document in two panes, one over the other, each scrolled on its own',
      pressed === 'clicked' && copied && own.top === 0 && own.low > 1000 && own.topH > 100 && own.lowH > 100, JSON.stringify({ pressed, copied, own }));

    // Typed above, shown below.
    // Typed in the upper pane, at the start of the first paragraph.
    const at = await js(`(() => { const b = document.querySelector('.wd-scroll .wd-page [data-block="0"]'); b.scrollIntoView({ block: 'center' }); const r = b.getBoundingClientRect(); return { x: Math.round(r.left + 3), y: Math.round(r.top + 8) }; })()`);
    wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    await until(() => model().selection?.focus?.block === 0, 'the caret in paragraph 1', 4000).catch(() => {});
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Home' }); wc.sendInputEvent({ type: 'keyUp', keyCode: 'Home' });
    await wait(150);
    wc.insertText('Typed above. ');
    const shown = await until(() => js(`(document.querySelector('.wd-split-pane [data-block="0"]')?.textContent || '').includes('Typed above')`), 'the change below', 6000).then(() => true).catch(() => false);
    // A click on paragraph 55 below puts the caret there above.
    await js(`(() => { const b = document.querySelector('.wd-split-pane [data-block="54"]'); b?.scrollIntoView({ block: 'center' }); b?.click(); return Boolean(b); })()`);
    const went = await until(() => model().selection?.focus?.block === 54, 'the caret at paragraph 55', 5000).then(() => true).catch(() => false);
    check('documents: what is typed in the upper pane shows in the lower one, and a click on a paragraph below puts the caret there above',
      shown && went, JSON.stringify({ shown, went, focus: model().selection?.focus }));

    await press();
    const one = await until(() => js(`!document.querySelector('.wd-split-pane')`), 'one pane again', 3000).then(() => true).catch(() => false);
    check('documents: Split pressed again puts the document back in one pane', one, one ? 'one pane' : 'still split');
    await wait(50);
  } catch (err) {
    check('documents: the Split checks ran', false, err.message);
  }
}
