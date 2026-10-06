// Presentations: Insert → Action, pressed on the ribbon — the title is told
// to go to slide 3 in the Action Settings dialog, a click on it in the show
// goes there while a click anywhere else still moves the show on, and End
// Show set the same way ends the show.

import fs from 'node:fs';
import path from 'node:path';
import { buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, doc
 */
export async function verifyDeckAction(h, { dir }) {
  const { open, check, until, wait, doc } = h;
  const file = path.join(dir, 'action.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Actions', slides: [1, 2, 3, 4].map((n) => ({ layout: 'title', title: `Slide ${n}`, body: 'Words' })) }));
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.sl-hit[data-shape]').length > 0`), 'the slide\'s shapes', 8000);
    const session = doc.sessions().filter((s) => s.kind === 'deck' && s.path === file).pop();
    const title = () => (doc.model({ id: session.id, slide: 0 }).slide?.shapes || []).find((s) => s.placeholder?.type === 'title' || s.placeholder?.type === 'ctrTitle');
    const id = title().id;
    const ribbon = (tab, label) => js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(tab)})?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)});
      if (!b) return 'no button'; if (b.disabled) return 'disabled'; b.click(); return 'clicked';
    })()`);
    const setAction = async (kind, slide = null) => {
      await js(`(() => { document.querySelector('.sl-hit[data-shape="${id}"]')?.click(); return 1; })()`);
      await until(() => js(`Boolean(document.querySelector('.sl-hit.selected'))`), 'the title selected', 3000).catch(() => {});
      const pressed = await ribbon('Insert', 'Action');
      await until(() => js(`Boolean(document.querySelector('.sl-action'))`), 'the Action Settings dialog', 4000).catch(() => {});
      await js(`(() => {
        const pick = (sel, v) => { const el = document.querySelector(sel); Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(el, String(v)); el.dispatchEvent(new Event('change', { bubbles: true })); };
        document.querySelector('.sl-action-link').click();
        pick('.sl-action-kind', ${JSON.stringify(kind)});
        return 1;
      })()`);
      if (slide != null) {
        await until(() => js(`Boolean(document.querySelector('.sl-action-slide'))`), 'the slide list', 2000).catch(() => {});
        await js(`(() => { const el = document.querySelector('.sl-action-slide'); Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(el, '${slide}'); el.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
      }
      await wait(100);
      await js(`document.querySelector('.sl-action-ok')?.click(), 1`);
      return pressed;
    };
    const bar = () => js(`document.querySelector('.sl-present-bar')?.textContent.trim().split(' ')[0] || null`);
    const startShow = async () => {
      await ribbon('Slide Show', 'From Beginning');
      await until(async () => (await bar()) === '1', 'the show on slide 1', 5000).catch(() => {});
      await until(() => js(`Boolean(document.querySelector('.sl-present [data-shape="${id}"]'))`), 'the title in the show', 4000).catch(() => {});
    };
    const clickTitle = () => js(`(() => { const g = document.querySelector('.sl-present [data-shape="${id}"]'); if (!g) return 'no shape'; g.dispatchEvent(new MouseEvent('click', { bubbles: true })); return getComputedStyle(g).cursor; })()`);

    // 1. Go to slide 3.
    const pressed = await setAction('slide', 2);
    const kept = await until(() => title()?.action?.kind === 'slide' && title()?.action?.slide === 2, 'the action on the title', 4000).then(() => true).catch(() => false);
    await startShow();
    const cursor = await clickTitle();
    const jumped = await until(async () => (await bar()) === '3', 'slide 3', 4000).then(() => true).catch(() => false);
    await js(`document.querySelector('.sl-present')?.click(), 1`);
    const advanced = await until(async () => (await bar()) === '4', 'slide 4 on an ordinary click', 4000).then(() => true).catch(() => false);
    win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' });
    win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' });
    await until(async () => (await bar()) === null, 'the show to end', 4000).catch(() => {});
    check('presentations: Insert → Action sends a click on the title to slide 3 in the show, and a click anywhere else still moves the show on',
      pressed === 'clicked' && kept && cursor === 'pointer' && jumped && advanced, JSON.stringify({ pressed, kept, action: title()?.action, cursor, jumped, advanced }));

    // 2. End Show.
    await js(`(() => { [...document.querySelectorAll('.sl-thumb, .sl-sorter button')][0]?.click(); return 1; })()`);
    await setAction('end');
    await until(() => title()?.action?.kind === 'end', 'End Show on the title', 4000).catch(() => {});
    await startShow();
    await clickTitle();
    const ended = await until(async () => (await bar()) === null, 'the show to end on the click', 4000).then(() => true).catch(() => false);
    check('presentations: an End Show action ends the show on the click', title()?.action?.kind === 'end' && ended, JSON.stringify({ action: title()?.action, ended }));
  } catch (err) {
    check('presentations: the Action checks ran', false, err.message);
  }
}
