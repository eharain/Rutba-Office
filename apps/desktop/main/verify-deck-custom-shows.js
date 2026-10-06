// Presentations: Slide Show → Custom Slide Show, pressed on the ribbon — a
// show made in the Custom Shows dialog from slides 4 and 2, in that order,
// played from the menu: the show starts on slide 4, goes to 2 and stops
// there; and the saved deck keeps the show as PowerPoint keeps one.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, doc
 */
export async function verifyDeckCustomShows(h, { dir }) {
  const { open, check, until, wait, doc } = h;
  const file = path.join(dir, 'custom-shows.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Shows', slides: [1, 2, 3, 4, 5].map((n) => ({ layout: 'obj', title: `Slide ${n}` })) }));
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'deck' && s.path === file).pop();
    const shows = () => doc.model({ id: session.id }).customShows || [];
    const bar = () => js(`document.querySelector('.sl-present-bar')?.textContent.trim().split(' ')[0] || null`);
    const openMenu = () => js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Slide Show')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Custom Slide Show');
      if (!b) return 'no button'; b.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 300, clientY: 110 })); return 'clicked';
    })()`);
    const pick = (label) => js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'no item'; b.click(); return 'picked'; })()`);
    const click = (sel) => js(`(() => { const b = document.querySelector(${JSON.stringify(sel)}); if (!b) return 'missing ' + ${JSON.stringify(sel)}; b.click(); return 'clicked'; })()`);

    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 5`), 'the slides', 8000);
    const pressed = await openMenu();
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === 'Custom Shows…'))`), 'the menu', 3000).catch(() => {});
    await pick('Custom Shows…');
    await until(() => js(`Boolean(document.querySelector('.sl-cs-new'))`), 'the Custom Shows dialog', 3000).catch(() => {});
    await click('.sl-cs-new');
    await until(() => js(`Boolean(document.querySelector('.sl-cs-name'))`), 'Define Custom Show', 3000).catch(() => {});
    await js(`(() => { const el = document.querySelector('.sl-cs-name'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, 'Short version'); el.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    // Slides 2 and 4 added, then 4 moved first.
    await click('.sl-cs-all [data-slide="1"]');
    await click('.sl-cs-all [data-slide="3"]');
    await click('.sl-cs-add');
    await wait(100);
    await js(`(() => { document.querySelectorAll('.sl-cs-in .sl-cs-item')[1]?.click(); return 1; })()`);
    await wait(100);
    await js(`(() => { [...document.querySelectorAll('.sl-cs-order .rw-btn')][0]?.click(); return 1; })()`);
    await wait(100);
    await click('.sl-cs-ok');
    const made = await until(() => shows().length === 1 && shows()[0].name === 'Short version' && JSON.stringify(shows()[0].slides) === '[3,1]', 'the show in the deck', 4000).then(() => true).catch(() => false);
    check('presentations: Custom Shows makes a named show of chosen slides, in an order of its own',
      pressed === 'clicked' && made, JSON.stringify({ pressed, shows: shows() }));

    // Played from the dialog's Show: 4, then 2, then it stops there.
    await click('.sl-cs-play');
    const started = await until(async () => (await bar()) === '4', 'the show on slide 4', 5000).then(() => true).catch(() => false);
    win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Right' });
    win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Right' });
    const second = await until(async () => (await bar()) === '2', 'slide 2', 4000).then(() => true).catch(() => false);
    win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Right' });
    win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Right' });
    await wait(600);
    const stayed = await bar();
    win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' });
    win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' });
    await until(async () => (await bar()) === null, 'the show to end', 4000).catch(() => {});
    check('presentations: a custom show plays only its slides, in its order, and ends on its last',
      started && second && stayed === '2', JSON.stringify({ started, second, stayed }));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => { try { const s = Deck.open(fs.readFileSync(file)).customShows(); return s.length === 1 && JSON.stringify(s[0].slides) === '[3,1]'; } catch { return false; } }, 'the saved show', 8000).then(() => true).catch(() => false);
    check('presentations: the saved deck keeps the custom show as PowerPoint keeps one', saved, saved ? 'custShowLst in the file' : 'not in the file');
  } catch (err) {
    check('presentations: the Custom Slide Show checks ran', false, err.message);
  }
}
