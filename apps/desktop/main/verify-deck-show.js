// Presentations: Record → Save as Show writes a copy PowerPoint opens as a show.
//
// The Save dialog is answered by the check (ipc.js's check-run queue), the
// button pressed as a person presses it, and the file on disk read: a .ppsx
// whose main part says it is a show, beside a deck still named and labelled
// as the presentation it was.

import fs from 'node:fs';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml/package';

/**
 * @param {object} h the harness: open, check, until, wait
 */
export async function verifyDeckShow(h, { file }) {
  const { open, check, until, wait } = h;
  const deck = path.join(path.dirname(file), 'show-source.pptx');
  const show = path.join(path.dirname(file), 'show-source.ppsx');
  try {
    fs.copyFileSync(file, deck);
    fs.rmSync(show, { force: true });
    const win = await open('slides', deck);
    const js = (code) => win.webContents.executeJavaScript(code);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Record')?.click(), 'tab'`);
    await wait(250);
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: show });
    const pressed = await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Save as Show'); if (!b) return 'no button'; if (b.disabled) return 'disabled'; b.click(); return 'clicked'; })()`);
    const written = await until(() => fs.existsSync(show), 'the show to be written', 8000).then(() => true).catch(() => false);
    const type = written ? OoxmlPackage.read(fs.readFileSync(show)).contentTypeOf('ppt/presentation.xml') : null;
    const deckType = OoxmlPackage.read(fs.readFileSync(deck)).contentTypeOf('ppt/presentation.xml');
    const title = win.getTitle();
    check('presentations: Record → Save as Show writes a .ppsx labelled as a show, and the deck stays the .pptx it was',
      pressed === 'clicked' && written === true && /slideshow\.main\+xml$/.test(type || '') && /presentation\.main\+xml$/.test(deckType) && /show-source\.pptx/.test(title),
      JSON.stringify({ pressed, written, type, deckType, title }));
  } catch (err) {
    check('presentations: the Save as Show check ran', false, err.message);
  }
  await verifyPresenterClock(h, { file });
}

/**
 * The presenter view's Pause and Reset (A4.1) are the clock's alone, as in
 * PowerPoint: a show on its second slide, Pause holds the clock and the show
 * goes on there; Reset puts the clock back to 0:00 and the slide stays;
 * Resume takes the time up again.
 */
async function verifyPresenterClock({ open, check, until, wait }, { file }) {
  const deck = path.join(path.dirname(file), 'presenter-clock.pptx');
  let js = null;
  try {
    fs.copyFileSync(file, deck);
    const win = await open('slides', deck);
    js = (code) => win.webContents.executeJavaScript(code);
    const { id } = await js(`window.rutbaOffice.doc.open({ path: ${JSON.stringify(deck)} })`);
    await js(`window.rutbaOffice.present.set({ id: ${JSON.stringify(id)}, index: 1, step: 0, running: true, blank: false })`);
    const pw = await open('slides', null, { presenter: id });
    const pjs = (code) => pw.webContents.executeJavaScript(code);
    await until(() => pjs(`Boolean(document.querySelector('[data-role="presenter-pause"]'))`), 'the presenter view', 10000);
    const clock = () => pjs(`document.querySelector('[data-role="presenter-clock"]')?.textContent.trim() || ''`);
    const state = () => js('window.rutbaOffice.present.state()');
    const press = (role) => pjs(`(() => { const b = document.querySelector('[data-role="${role}"]'); if (!b) return 'no ' + ${JSON.stringify(role)}; b.click(); return 'clicked'; })()`);
    await wait(1300);
    const paused = await press('presenter-pause');
    await until(() => pjs(`Boolean(document.querySelector('.pv-clock.paused'))`), 'the clock paused', 3000).catch(() => {});
    const held = await clock();
    await wait(2200);
    const later = await clock();
    const during = await state();
    const label = await pjs(`document.querySelector('[data-role="presenter-pause"]')?.textContent.trim() || ''`);
    check('presentations: the presenter\'s Pause holds the clock and the show goes on, on the slide it was on',
      paused === 'clicked' && held === later && during.running === true && during.index === 1 && during.pausedAt != null && label === 'Resume',
      JSON.stringify({ paused, held, later, running: during.running, index: during.index, label }));
    const reset = await press('presenter-reset');
    await until(async () => (await clock()) === '0:00', 'the clock at nought', 3000).catch(() => {});
    const afterReset = await state();
    const zero = await clock();
    await press('presenter-pause');
    await wait(2300);
    const going = await clock();
    const resumed = await state();
    check('presentations: the presenter\'s Reset puts the clock back to 0:00 and the slide stays; Resume takes the time up again',
      reset === 'clicked' && zero === '0:00' && afterReset.index === 1 && afterReset.running === true && going !== '0:00' && resumed.pausedAt == null && resumed.index === 1,
      JSON.stringify({ reset, zero, index: afterReset.index, going, pausedAt: resumed.pausedAt }));
  } catch (err) {
    check('presentations: the presenter clock checks ran', false, err?.message || String(err));
  } finally {
    await js?.(`window.rutbaOffice.present.set({ running: false })`).catch(() => {});
  }
}
