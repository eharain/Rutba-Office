// Presentations: the notes under the slide in Normal view are typed where
// they are shown. A slide with no notes says Click to add notes; what is
// typed there is kept once typing pauses, and kept for its own slide when
// another slide is chosen straight after; the next slide's box is its own.
// Run alone with RUTBA_VERIFY_ONLY=decknotes.

import fs from 'node:fs';
import path from 'node:path';
import { buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn, doc, sessionFor
 */
export async function verifyDeckNotes(h, { dir }) {
  const { open, check, until, wait, errorsIn, doc, sessionFor } = h;
  const file = path.join(dir, 'typed-notes.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Notes', slides: [{ layout: 'title', title: 'First' }, { layout: 'title', title: 'Second' }] }));
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const model = (slide) => doc.model({ id: sessionFor('deck').id, slide });
    const notesOf = (slide) => model(slide).slide?.notes || '';
    const box = () => js(`(() => { const b = document.querySelector('[data-role="slide-notes"]'); return b ? { slide: Number(b.dataset.slide), value: b.value, placeholder: b.placeholder, tag: b.tagName } : null; })()`);
    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2 && Boolean(document.querySelector('[data-role="slide-notes"]'))`), 'the notes box under the slide', 8000);
    const empty = await box();

    // Typed, then a pause: kept on the first slide.
    win.focus();
    await js(`document.querySelector('[data-role="slide-notes"]').focus(), 1`);
    wc.insertText('Welcome everyone.');
    const kept = await until(() => notesOf(0) === 'Welcome everyone.', 'the notes kept after a pause', 4000).then(() => true).catch(() => false);
    check('presentations: a slide with no notes says Click to add notes under it, and what is typed there is kept once typing pauses',
      empty?.tag === 'TEXTAREA' && empty.placeholder === 'Click to add notes' && empty.value === '' && kept,
      JSON.stringify({ empty, notes: notesOf(0) }));

    // More typed, and the second slide chosen at once: the first keeps it, the second's box is empty.
    wc.insertText(' Phones off.');
    await js(`(() => { const t = document.querySelectorAll('.sl-thumb')[1]; t.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); document.activeElement?.blur?.(); t.click(); return 1; })()`);
    await until(async () => (await box())?.slide === 1, 'the second slide', 4000).catch(() => {});
    await until(() => notesOf(0) === 'Welcome everyone. Phones off.', 'the first slide\'s notes', 4000).catch(() => {});
    const second = await box();
    check('presentations: notes typed just before another slide is chosen stay with their own slide, and the next slide\'s notes box is its own',
      notesOf(0) === 'Welcome everyone. Phones off.' && second?.slide === 1 && second.value === '' && notesOf(1) === '',
      JSON.stringify({ first: notesOf(0), second, secondNotes: notesOf(1) }));

    // Back on the first slide, its notes are in the box.
    await js(`document.querySelectorAll('.sl-thumb')[0].click(), 1`);
    await until(async () => (await box())?.slide === 0, 'the first slide again', 4000).catch(() => {});
    await wait(200);
    const back = await box();
    check('presentations: going back to a slide shows the notes typed on it', back?.value === 'Welcome everyone. Phones off.', JSON.stringify(back));

    const complaints = await errorsIn(win);
    check('presentations: typing notes under the slide reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('presentations: the notes typing checks ran', false, err?.message || String(err));
  }
}
