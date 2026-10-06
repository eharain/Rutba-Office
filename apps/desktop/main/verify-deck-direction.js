// Presentations: View → View Direction, pressed on the ribbon — right to
// left puts the slides strip on the right of the slide, the panes mirrored
// and each one's own content as it was; pressed again, left to right.

import fs from 'node:fs';
import path from 'node:path';
import { buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait
 */
export async function verifyDeckDirection(h, { dir }) {
  const { open, check, until, wait } = h;
  const file = path.join(dir, 'direction.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Direction', slides: [{ layout: 'title', title: 'One' }, { layout: 'obj', title: 'Two' }] }));
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const sides = () => js(`(() => {
      const strip = document.querySelector('.sl-thumb')?.closest('.rw-panel, aside, section') || document.querySelector('.sl-thumb');
      const stage = document.querySelector('.sl-stage');
      if (!strip || !stage) return null;
      return { strip: Math.round(strip.getBoundingClientRect().left), stage: Math.round(stage.getBoundingClientRect().left), thumbDir: getComputedStyle(document.querySelector('.sl-thumb')).direction };
    })()`);
    const press = () => js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'View')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'View Direction');
      if (!b || b.disabled) return 'no button'; b.click(); return 'clicked';
    })()`);
    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slides', 8000);
    const before = await sides();
    const pressed = await press();
    const flipped = await until(async () => { const s = await sides(); return Boolean(s) && s.strip > s.stage; }, 'the strip on the right', 4000).then(() => true).catch(() => false);
    const during = await sides();
    await press();
    const back = await until(async () => { const s = await sides(); return Boolean(s) && s.strip < s.stage; }, 'the strip on the left again', 4000).then(() => true).catch(() => false);
    check('presentations: View → View Direction puts the slides strip on the right of the slide, each pane\'s content as it was, and back again',
      before?.strip < before?.stage && pressed === 'clicked' && flipped && during?.thumbDir === 'ltr' && back,
      JSON.stringify({ before, pressed, flipped, during, back }));
    await wait(50);
  } catch (err) {
    check('presentations: the View Direction checks ran', false, err.message);
  }
}
