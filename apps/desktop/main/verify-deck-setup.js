// Presentations: Slide Show → Set Up Slide Show, set in its dialog and then
// obeyed by the show — it starts on the range's first slide, loops inside
// the range, and as a kiosk takes no clicks — and kept in the saved file.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, doc
 */
export async function verifyDeckSetup(h, { dir }) {
  const { open, check, until, wait, doc } = h;
  const file = path.join(dir, 'setup-show.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Show', slides: [1, 2, 3, 4].map((n) => ({ layout: 'title', title: `Slide ${n}` })) }));
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'deck' && s.path === file).pop();
    const settings = () => doc.model({ id: session.id }).showSettings || {};
    const pressRibbon = (label) => js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Slide Show')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)});
      if (!b) return 'no button'; if (b.disabled) return 'disabled'; b.click(); return 'clicked';
    })()`);
    const bar = () => js(`document.querySelector('.sl-present-bar')?.textContent.trim().split(' ')[0] || null`);
    const key = (k) => js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: ${JSON.stringify(k)}, bubbles: true })), 1`);

    // 1. The dialog: slides 2 to 3, looping, without animation.
    const opened = await pressRibbon('Set Up Slide Show');
    await until(() => js(`Boolean(document.querySelector('.sl-setup'))`), 'the Set Up Show dialog', 4000).catch(() => {});
    await js(`(() => {
      const setNum = (sel, v) => { const el = document.querySelector(sel); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, String(v)); el.dispatchEvent(new Event('input', { bubbles: true })); };
      document.querySelector('.sl-setup-some').click();
      setNum('.sl-setup-from', 2);
      setNum('.sl-setup-to', 3);
      document.querySelector('.sl-setup-loop').click();
      document.querySelector('.sl-setup-noanim').click();
      document.querySelector('.sl-setup-ok').click();
      return 1;
    })()`);
    const set = await until(() => settings().loop === true && settings().range?.from === 2 && settings().range?.to === 3 && settings().animation === false, 'the settings in the deck', 4000).then(() => true).catch(() => false);
    check('presentations: Set Up Slide Show writes slides 2 to 3, looping, without animation, into the deck',
      opened === 'clicked' && set === true, JSON.stringify({ opened, settings: settings() }));

    // 2. From Beginning starts on slide 2, moves to 3, and loops back to 2.
    await pressRibbon('From Beginning');
    const started = await until(async () => (await bar()) === '2', 'the show on slide 2', 5000).then(() => true).catch(() => false);
    await key('ArrowRight');
    const third = await until(async () => (await bar()) === '3', 'slide 3', 4000).then(() => true).catch(() => false);
    await key('ArrowRight');
    const looped = await until(async () => (await bar()) === '2', 'back to slide 2', 4000).then(() => true).catch(() => false);
    await key('Escape');
    await until(async () => (await bar()) === null, 'the show to end', 4000).catch(() => {});
    check('presentations: the show starts on the range\'s first slide and loops inside the range',
      started === true && third === true && looped === true, JSON.stringify({ started, third, looped }));

    // 3. A kiosk: a click does not move the show.
    await pressRibbon('Set Up Slide Show');
    await until(() => js(`Boolean(document.querySelector('.sl-setup'))`), 'the dialog again', 4000).catch(() => {});
    await js(`(() => { document.querySelector('input[name="sl-setup-type"][value="kiosk"]').click(); document.querySelector('.sl-setup-ok').click(); return 1; })()`);
    await until(() => settings().type === 'kiosk', 'kiosk in the deck', 4000).catch(() => {});
    await pressRibbon('From Beginning');
    await until(async () => (await bar()) === '2', 'the kiosk show', 5000).catch(() => {});
    await js(`document.querySelector('.sl-present')?.click(), 1`);
    await key('ArrowRight');
    await wait(500);
    const stayed = await bar();
    await key('Escape');
    await until(async () => (await bar()) === null, 'the kiosk show to end on Escape', 4000).catch(() => {});
    const ended = (await bar()) === null;
    check('presentations: as a kiosk the show takes no click or arrow — only Escape ends it',
      settings().type === 'kiosk' && stayed === '2' && ended === true, JSON.stringify({ type: settings().type, stayed, ended }));

    // 4. Saved, the file carries the kiosk.
    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try { return Deck.open(fs.readFileSync(file)).showSettings().type === 'kiosk'; } catch { return false; }
    }, 'the saved show settings', 8000).then(() => true).catch(() => false);
    check('presentations: the saved deck carries the show settings, as PowerPoint keeps them', saved === true, saved ? 'kiosk, slides 2 to 3' : 'not in the file');
  } catch (err) {
    check('presentations: the Set Up Slide Show checks ran', false, err.message);
  }
}
