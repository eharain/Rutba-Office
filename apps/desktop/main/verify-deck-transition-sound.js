// Presentations: Transitions → Sound, pressed on the ribbon — the menu
// lists No Sound, Stop Previous Sound, the suite's own sounds, Other Sound…
// and Loop Until Next Sound; Chime on slide 2 is heard as it is picked and
// written to the deck; Loop Until Next Sound marks it; in the show the
// sound plays as slide 2 comes in, and not on slide 1; and the saved deck
// keeps it as PowerPoint keeps a transition sound.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, doc, sessionFor
 */
export async function verifyDeckTransitionSound(h, { dir }) {
  const { open, check, until, wait, doc, sessionFor } = h;
  const file = path.join(dir, 'transition-sound.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Sound', slides: [{ layout: 'title', title: 'One' }, { layout: 'obj', title: 'Two' }] }));
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const model = (slide) => doc.model({ id: sessionFor('deck').id, slide });
    const openMenu = () => js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Transitions')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Sound');
      if (!b || b.disabled) return 'no button';
      b.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 700, clientY: 110 }));
      await new Promise((r) => setTimeout(r, 200));
      return [...document.querySelectorAll('.rw-menu button')].map((n) => n.textContent.trim());
    })()`);
    const pick = (label) => js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'no item'; b.click(); return 'picked'; })()`);

    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slides', 8000);
    // Every play() the page makes, recorded.
    await js(`(() => { window.__plays = []; const play = HTMLMediaElement.prototype.play; HTMLMediaElement.prototype.play = function () { window.__plays.push({ src: this.src, loop: this.loop, at: Date.now() }); return play.call(this); }; return 1; })()`);
    await js(`(() => { document.querySelectorAll('.sl-thumb')[1]?.click(); return 1; })()`);
    await until(() => js(`document.querySelectorAll('.sl-thumb')[1]?.classList.contains('active')`), 'slide 2', 4000).catch(() => {});
    const items = await openMenu();
    await pick('Chime');
    const written = await until(() => model(1).slide?.transitionSound?.name === 'chime.wav', 'the sound in the deck', 5000).then(() => true).catch(() => false);
    const heard = await js(`window.__plays.length`);
    check('presentations: Transitions → Sound lists the sounds, and Chime on slide 2 is heard as it is picked and written to the deck',
      Array.isArray(items) && ['[No Sound]', '[Stop Previous Sound]', 'Chime', 'Drum Roll', 'Other Sound…', 'Loop Until Next Sound'].every((l) => items.includes(l)) && written && heard >= 1,
      JSON.stringify({ items, written, heard, sound: model(1).slide?.transitionSound }));

    await openMenu();
    await pick('Loop Until Next Sound');
    const looped = await until(() => model(1).slide?.transitionSound?.loop === true, 'the sound looping', 5000).then(() => true).catch(() => false);

    // The show: nothing on slide 1, the chime as slide 2 comes in, looping.
    await js(`(() => { window.__plays = []; [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Slide Show')?.click(); return 1; })()`);
    await wait(200);
    await js(`(() => { [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'From Beginning')?.click(); return 1; })()`);
    await until(() => js(`document.querySelector('.sl-present-bar')?.textContent.trim().startsWith('1')`), 'the show on slide 1', 6000).catch(() => {});
    await wait(500);
    const onFirst = await js(`window.__plays.length`);
    win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Right' });
    win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Right' });
    const played = await until(() => js(`window.__plays.some((p) => p.loop && /^rutba:/.test(p.src))`), 'the chime in the show', 6000).then(() => true).catch(() => false);
    win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' });
    win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' });
    await until(() => js(`!document.querySelector('.sl-present-bar')`), 'the show to end', 4000).catch(() => {});
    check('presentations: Loop Until Next Sound marks it, and in the show the sound plays as slide 2 comes in, looping, and not on slide 1',
      looped && onFirst === 0 && played, JSON.stringify({ looped, onFirst, played }));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const deck = Deck.open(fs.readFileSync(file));
        const s = deck.transitionSound(1);
        return s?.name === 'chime.wav' && s.loop && deck.media(s.part)?.length > 1000 && !deck.transitionSound(0);
      } catch { return false; }
    }, 'the saved sound', 8000).then(() => true).catch(() => false);
    check('presentations: the saved deck keeps the transition sound on slide 2 as PowerPoint keeps one', saved, saved ? 'sndAc with the WAV in ppt/media' : 'not in the file');
  } catch (err) {
    check('presentations: the transition sound checks ran', false, err.message);
  }
}
