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
}
