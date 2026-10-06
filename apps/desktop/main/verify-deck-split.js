// Presentations: View → Move Split — the arrow keys move the slide pane's
// edge and the top of the notes, and Enter puts the mode away, after which
// the arrows are the slides' own again.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait
 */
export async function verifyDeckSplit(h, { dir }) {
  const { open, check, until, wait } = h;
  const file = path.join(dir, 'split.pptx');
  try {
    const deck = Deck.open(buildPptx({ title: 'Split', slides: [1, 2].map((n) => ({ layout: 'title', title: `Slide ${n}` })) }));
    deck.setNotes(0, 'What to say on the first slide.');
    fs.writeFileSync(file, deck.save());
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const key = async (keyCode, n = 1) => {
      for (let i = 0; i < n; i++) {
        wc.sendInputEvent({ type: 'keyDown', keyCode });
        wc.sendInputEvent({ type: 'keyUp', keyCode });
        await wait(80);
      }
    };
    const sizes = () => js(`({ rail: document.querySelector('.rw-panel:not(.right)')?.offsetWidth ?? null, notes: document.querySelector('.sl-notes')?.offsetHeight ?? null })`);
    await until(async () => (await sizes()).notes != null, 'the notes under the slide', 8000);
    const before = await sizes();

    const pressed = await js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'View')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Move Split');
      if (!b) return 'no button'; if (b.disabled) return 'disabled'; b.click(); return 'clicked';
    })()`);
    await js(`document.activeElement?.blur?.(), 1`);
    await key('Right', 3);
    await key('Up', 2);
    await until(async () => { const s = await sizes(); return s.rail === before.rail + 48 && s.notes === before.notes + 32; }, 'the splits to move', 3000).catch(() => {});
    const moved = await sizes();
    check('presentations: View → Move Split moves the slide pane\'s edge and the top of the notes with the arrow keys',
      pressed === 'clicked' && moved.rail === before.rail + 48 && moved.notes === before.notes + 32, JSON.stringify({ pressed, before, moved }));

    await key('Enter');
    await key('Right', 2);
    await wait(200);
    const after = await sizes();
    check('presentations: Enter puts Move Split away, and the arrows no longer move the splits', after.rail === moved.rail && after.notes === moved.notes, JSON.stringify({ moved, after }));
  } catch (err) {
    check('presentations: the Move Split checks ran', false, err.message);
  }
}
