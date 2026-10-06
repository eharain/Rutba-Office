// Presentations: Slide Show → Rehearse Timings, Use Timings and Play
// Narrations — a rehearsal timed slide by slide at three paces, its timings
// kept when the show ends (the question answered as a person would), each
// slide then carrying its own time, and the two flags pressed off and on.

import fs from 'node:fs';
import path from 'node:path';
import { buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, doc
 */
export async function verifyDeckRehearse(h, { dir }) {
  const { open, check, until, wait, doc } = h;
  const file = path.join(dir, 'rehearse.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Rehearse', slides: [1, 2, 3].map((n) => ({ layout: 'title', title: `Slide ${n}` })) }));
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'deck' && s.path === file).pop();
    const model = () => doc.model({ id: session.id });
    const after = (i) => doc.model({ id: session.id, slide: i }).slide?.transition?.advanceAfter ?? null;
    const key = (keyCode) => { wc.sendInputEvent({ type: 'keyDown', keyCode }); wc.sendInputEvent({ type: 'keyUp', keyCode }); };
    const press = (label) => js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Slide Show')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)});
      if (!b) return 'no button'; if (b.disabled) return 'disabled'; b.click(); return 'clicked';
    })()`);
    const pressedState = (label) => js(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)})?.getAttribute('aria-pressed')`);
    const bar = () => js(`document.querySelector('.sl-present-bar')?.textContent.trim().split(' ')[0] || null`);

    // 1. Rehearse: 1.5 s on the first slide, 0.8 s on the second, 0.4 s on the third.
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'message', answer: { response: 0, checked: false } });
    const pressed = await press('Rehearse Timings');
    await until(async () => (await bar()) === '1', 'the rehearsal on slide 1', 5000).catch(() => {});
    const clock = await until(() => js(`Boolean(document.querySelector('.sl-rehearse-total'))`), 'the clock', 3000).then(() => true).catch(() => false);
    await wait(1500);
    key('Right');
    await until(async () => (await bar()) === '2', 'slide 2', 3000).catch(() => {});
    await wait(800);
    key('Right');
    await until(async () => (await bar()) === '3', 'slide 3', 3000).catch(() => {});
    await wait(400);
    key('Escape');
    const kept = await until(() => after(0) != null && after(1) != null && after(2) != null && model().showSettings?.useTimings === true, 'the timings kept', 6000).then(() => true).catch(() => false);
    const times = [after(0), after(1), after(2)];
    const near = (t, want) => t != null && t >= want - 0.15 && t <= want + 0.9;
    check('presentations: Rehearse Timings times each slide as the show leaves it and, kept, gives each its own time with Use Timings on',
      pressed === 'clicked' && clock && kept && near(times[0], 1.5) && near(times[1], 0.8) && near(times[2], 0.4) && times[0] > times[1] && times[1] > times[2],
      JSON.stringify({ pressed, clock, kept, times, useTimings: model().showSettings?.useTimings }));

    // 2. Use Timings and Play Narrations, pressed off and read back.
    await until(() => js(`!document.querySelector('.sl-present')`), 'the editor back', 4000).catch(() => {});
    const timingsOn = await pressedState('Use Timings');
    await press('Use Timings');
    const timingsOff = await until(() => model().showSettings?.useTimings === false, 'Use Timings off', 4000).then(() => true).catch(() => false);
    await press('Play Narrations');
    const narrationOff = await until(() => model().showSettings?.narration === false, 'Play Narrations off', 4000).then(() => true).catch(() => false);
    await until(async () => (await pressedState('Play Narrations')) === 'false', 'the button to read off', 3000).catch(() => {});
    const shown = await pressedState('Play Narrations');
    check('presentations: Use Timings and Play Narrations read the show\'s settings and set them, one flag each',
      timingsOn === 'true' && timingsOff && narrationOff && shown === 'false' && model().showSettings?.loop === false,
      JSON.stringify({ timingsOn, timingsOff, narrationOff, shown, settings: model().showSettings }));
  } catch (err) {
    check('presentations: the Rehearse Timings checks ran', false, err.message);
  }
}
