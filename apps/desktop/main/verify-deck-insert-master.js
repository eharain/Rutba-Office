// Presentations: Slide Master → Insert Slide Master, pressed on the ribbon —
// the strip gains a second master with as many layouts as the first, the
// new master is the one being edited, and the saved deck carries it.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait
 */
export async function verifyDeckInsertMaster(h, { dir }) {
  const { open, check, until, wait } = h;
  const file = path.join(dir, 'masters.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Masters', slides: [{ layout: 'title', title: 'One' }, { layout: 'obj', title: 'Two' }] }));
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const strip = () => js(`[...document.querySelectorAll('.sl-masterstrip .sl-mthumb')].map((b) => ({ kind: b.classList.contains('sl-mthumb-master') ? 'master' : 'layout', part: b.dataset.part, active: b.classList.contains('active') }))`);
    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slides', 8000);
    await js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'View')?.click();
      await new Promise((r) => setTimeout(r, 200));
      [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Slide Master')?.click();
      return 1;
    })()`);
    await until(async () => (await strip()).length > 1, 'the master strip', 5000).catch(() => {});
    const before = await strip();
    const pressed = await js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Slide Master')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Insert Slide Master');
      if (!b) return 'no button'; if (b.disabled) return 'disabled'; b.click(); return 'clicked';
    })()`);
    await until(async () => (await strip()).filter((s) => s.kind === 'master').length === 2, 'a second master', 5000).catch(() => {});
    const after = await strip();
    const masters = after.filter((s) => s.kind === 'master');
    const layoutsOf = (list, master) => { const at = list.findIndex((s) => s.part === master); let n = 0; for (let i = at + 1; i < list.length && list[i].kind === 'layout'; i++) n++; return n; };
    const active = after.find((s) => s.active)?.part;
    check('presentations: Slide Master → Insert Slide Master adds a second master with as many layouts as the first, and opens it for editing',
      pressed === 'clicked' && masters.length === 2 && layoutsOf(after, masters[1].part) === layoutsOf(before, before[0].part) && active === masters[1].part,
      JSON.stringify({ pressed, before: before.length, after: after.length, masters: masters.map((m) => m.part), active }));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => { try { return Deck.open(fs.readFileSync(file)).masterParts().length === 2; } catch { return false; } }, 'the saved masters', 8000).then(() => true).catch(() => false);
    check('presentations: the saved deck carries both masters', saved, saved ? 'two masters' : 'one master');
    await wait(100);
  } catch (err) {
    check('presentations: the Insert Slide Master checks ran', false, err.message);
  }
}
