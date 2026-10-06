// Presentations: View → Notes Master and Handout Master, pressed on the
// ribbon, on a deck that has neither — the notes page stands on the stage,
// portrait, with its six placeholders and the Notes Master tab; the Header
// box takes the header off; Notes Page Orientation turns the page; Close
// Master View goes back to the slides; the Handout Master shows where six
// slides go, then three beside lines for notes; and the saved deck keeps
// both masters as PowerPoint keeps them.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, doc, sessionFor, capture
 */
export async function verifyDeckNotesMaster(h, { dir }) {
  const { open, check, until, wait, doc, sessionFor, capture } = h;
  const file = path.join(dir, 'notes-master.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Notes', slides: [{ layout: 'title', title: 'Plan' }, { layout: 'obj', title: 'Steps', body: ['One', 'Two'] }] }));
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const tab = (label) => js(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(label)}); if (!t) return 'no tab'; t.click(); return 'tab'; })()`);
    const press = (label) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)});
      if (!b || b.disabled) return 'no button ' + ${JSON.stringify(label)};
      b.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 400, clientY: 110 }));
      return 'clicked';
    })()`);
    const pick = (label) => js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'no item'; b.click(); return 'picked'; })()`);
    const stage = () => js(`(() => { const s = document.querySelector('.sl-slide'); return s ? { w: parseFloat(s.style.width), h: parseFloat(s.style.height), svg: s.querySelector('.sl-svg')?.innerHTML || '' } : null; })()`);
    const tabs = () => js(`[...document.querySelectorAll('.rw-tab')].map((t) => t.textContent.trim())`);
    const model = (master) => doc.model({ id: sessionFor('deck').id, master });

    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slides', 8000);
    await tab('View');
    await wait(200);
    const opened = await press('Notes Master');
    await until(async () => (await tabs()).includes('Notes Master') && (await stage())?.svg.includes('Click to edit Master text styles'), 'the notes master on the stage', 8000).catch(() => {});
    const first = await stage();
    const boxes = await js(`[...document.querySelectorAll('.sl-nm-ph')].map((c) => c.checked)`);
    if (capture) await capture(win, 'deck-notes-master.png');
    check('presentations: View → Notes Master puts the notes page on the stage, portrait, with its placeholders and the Notes Master tab',
      opened === 'clicked' && Boolean(first) && first.h > first.w && (await tabs()).includes('Notes Master') && boxes.length === 6 && boxes.every(Boolean),
      JSON.stringify({ opened, size: first && [first.w, first.h], boxes, tabs: await tabs() }));

    // The Header box takes the header off the master.
    const part = model(null).slide ? (await js(`document.querySelector('.sl-mthumb.active')?.dataset.part || null`)) : null;
    await js(`(() => { const c = document.querySelector('.sl-nm-hdr'); if (c) c.click(); return Boolean(c); })()`);
    const off = await until(() => model(part).masterView?.placeholders?.hdr === false, 'the header off', 4000).then(() => true).catch(() => false);
    // Notes Page Orientation → Landscape.
    await press('Notes Page Orientation');
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === 'Landscape'))`), 'the menu', 3000).catch(() => {});
    await pick('Landscape');
    const turned = await until(async () => { const s = await stage(); return Boolean(s) && s.w > s.h; }, 'the page turned', 5000).then(() => true).catch(() => false);
    check('presentations: the Header box takes the header off the notes master, and Notes Page Orientation turns the page to landscape',
      Boolean(part) && off && turned, JSON.stringify({ part, off, turned, size: await stage().then((s) => s && [s.w, s.h]) }));

    // Close Master View, then the Handout Master: six slots, then three beside lines.
    await press('Close Master View');
    await until(async () => !(await tabs()).includes('Notes Master'), 'back to the slides', 4000).catch(() => {});
    await tab('View');
    await wait(200);
    await press('Handout Master');
    await until(async () => (await tabs()).includes('Handout Master'), 'the handout master', 8000).catch(() => {});
    const six = await until(() => js(`document.querySelectorAll('.sl-handout-slot').length === 6`), 'six slots', 4000).then(() => true).catch(() => false);
    await press('Slides Per Page');
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === '3 Slides'))`), 'the menu', 3000).catch(() => {});
    await pick('3 Slides');
    const three = await until(() => js(`document.querySelectorAll('.sl-handout-slot').length === 3 && document.querySelectorAll('.sl-handout-lines').length === 3`), 'three slots and lines', 4000).then(() => true).catch(() => false);
    const handoutBoxes = await js(`[...document.querySelectorAll('.sl-nm-ph')].length`);
    if (capture) await capture(win, 'deck-handout-master.png');
    check('presentations: View → Handout Master shows where six slides go on the page, Slides Per Page three beside lines for notes, with its four placeholders',
      six && three && handoutBoxes === 4, JSON.stringify({ six, three, handoutBoxes }));

    await press('Close Master View');
    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const deck = Deck.open(fs.readFileSync(file));
        const notes = deck.masterFor('notes');
        const handout = deck.masterFor('handout');
        return Boolean(notes && handout) && deck.masterPlaceholders(notes).hdr === false && deck.notesSize.cx > deck.notesSize.cy
          && /<p:notesMasterIdLst>[\s\S]*<p:handoutMasterIdLst>/.test(deck.pkg.text('ppt/presentation.xml'));
      } catch { return false; }
    }, 'the saved masters', 8000).then(() => true).catch(() => false);
    check('presentations: the saved deck keeps the notes master without its header, the handout master, and the turned notes page', saved, saved ? 'both masters listed in presentation.xml' : 'not in the file');
  } catch (err) {
    check('presentations: the Notes Master checks ran', false, err.message);
  }
}
