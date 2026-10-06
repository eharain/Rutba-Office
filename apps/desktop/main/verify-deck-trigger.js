// Presentations: Animations → Trigger, pressed on the ribbon — the answer
// fades in from the gallery, Trigger puts it on a click of the button, and
// in the show the answer waits hidden, a click on the button brings it in
// without moving the show on, and a click anywhere else still does.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, doc
 */
export async function verifyDeckTrigger(h, { dir }) {
  const { open, check, until, wait, doc } = h;
  const file = path.join(dir, 'trigger.pptx');
  try {
    const deck = Deck.open(buildPptx({ title: 'Quiz', slides: [{ layout: 'title', title: 'What is two and two?', body: 'Four' }, { layout: 'obj', title: 'Next question' }] }));
    deck.addTextBox(0, { x: 40, y: 40, w: 220, h: 60, paragraphs: [{ runs: [{ text: 'Show the answer' }] }] });
    fs.writeFileSync(file, deck.save());
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'deck' && s.path === file).pop();
    const model = () => doc.model({ id: session.id, slide: 0 }).slide;
    const shapes = model().shapes;
    const answer = String(shapes.find((s) => (s.text?.paragraphs || []).some((p) => p.plain === 'Four')).id);
    const button = shapes.find((s) => (s.text?.paragraphs || []).some((p) => p.plain === 'Show the answer'));
    const bar = () => js(`document.querySelector('.sl-present-bar')?.textContent.trim().split(' ')[0] || null`);

    // Fade on the answer, from the gallery; then Trigger → On Click of the button.
    await until(() => js(`document.querySelectorAll('.sl-hit[data-shape]').length >= 3`), 'the shapes', 8000);
    await js(`(() => { const el = document.querySelector('.sl-hit[data-shape="${answer}"]'); el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true })); el.click(); return 1; })()`);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Animations')?.click(), 1`);
    await until(() => js(`document.querySelectorAll('.sl-an-pick').length >= 8`), 'the gallery', 4000).catch(() => {});
    await js(`(() => { const b = document.querySelector('.sl-an-pick[data-effect="fade"]'); b?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); b?.click(); return 1; })()`);
    await until(() => (model().animations || []).length === 1, 'the Fade', 5000).catch(() => {});
    await wait(300);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Trigger'); b?.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 400, clientY: 110 })); return 1; })()`);
    const item = `On Click of ${button.name || `Shape ${button.id}`}`;
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(item)}))`), 'the Trigger menu', 3000).catch(() => {});
    const picked = await js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(item)}); if (!b) return 'no item'; b.click(); return 'picked'; })()`);
    const put = await until(() => String(model().animations?.[0]?.triggerShape) === String(button.id), 'the trigger kept', 4000).then(() => true).catch(() => false);
    check('presentations: Animations → Trigger puts the picked effect on a click of another shape',
      picked === 'picked' && put, JSON.stringify({ picked, animations: model().animations }));

    // The show: hidden at first, in on a click of the button, the show still on slide 1.
    await js(`(async () => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Slide Show')?.click(); await new Promise((r) => setTimeout(r, 200)); [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'From Beginning')?.click(); return 1; })()`);
    await until(async () => (await bar()) === '1', 'the show', 5000).catch(() => {});
    const visible = () => js(`(() => { const g = document.querySelector('.sl-present [data-shape="${answer}"]'); return g ? getComputedStyle(g).visibility : null; })()`);
    await until(async () => (await visible()) != null, 'the answer drawn', 4000).catch(() => {});
    const before = await visible();
    await js(`(() => { document.querySelector('.sl-present [data-shape="${button.id}"]')?.dispatchEvent(new MouseEvent('click', { bubbles: true })); return 1; })()`);
    const shown = await until(async () => (await visible()) === 'visible', 'the answer in', 4000).then(() => true).catch(() => false);
    await wait(700);
    const stayed = await bar();
    await js(`document.querySelector('.sl-present')?.click(), 1`);
    const moved = await until(async () => (await bar()) === '2', 'the next slide', 4000).then(() => true).catch(() => false);
    win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' });
    win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' });
    check('presentations: in the show a triggered effect waits, plays on a click of its trigger without moving the show on, and any other click moves it on',
      before === 'hidden' && shown && stayed === '1' && moved, JSON.stringify({ before, shown, stayed, moved }));
  } catch (err) {
    check('presentations: the Trigger checks ran', false, err.message);
  }
}
