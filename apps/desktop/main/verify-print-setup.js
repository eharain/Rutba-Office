// File → Print starts from what the file keeps. A deck: Handout, three to a
// page, saved as a PDF, is kept in the deck where PowerPoint keeps it, and
// the next print opens on it. A document: its own paper shown, and
// Landscape chosen in the dialog turns the document's pages, as Word's Print
// does. Run alone with RUTBA_VERIFY_ONLY=printsetup.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { Deck, buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the files are written
 */
export async function verifyPrintSetup({ open, check, until, wait, press, errorsIn, doc, sessionFor }, { dir }) {
  const deckFile = path.join(dir, 'print-setup.pptx');
  const docFile = path.join(dir, 'print-setup.docx');
  // A select in the print dialog, by an option only it has, set as a person sets it.
  const choose = (js, option, value) => js(`(() => {
    const el = [...document.querySelectorAll('.rw-dialog select')].find((s) => [...s.options].some((o) => o.value === ${JSON.stringify(option)}));
    if (!el) return false;
    Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(el, ${JSON.stringify(value)});
    el.dispatchEvent(new Event('change', { bubbles: true }));
    return true;
  })()`);
  const valueOf = (js, option) => js(`[...document.querySelectorAll('.rw-dialog select')].find((s) => [...s.options].some((o) => o.value === ${JSON.stringify(option)}))?.value ?? null`);
  const button = (js, label) => js(`(() => { const b = [...document.querySelectorAll('.rw-dialog button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!b) return false; b.click(); return true; })()`);

  try {
    fs.writeFileSync(deckFile, buildPptx({ title: 'Print', slides: [{ layout: 'title', title: 'One' }, { layout: 'title', title: 'Two' }, { layout: 'title', title: 'Three' }] }));
    const win = await open('slides', deckFile);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 3`), 'the slides', 8000);
    await js(`document.querySelector('.sl-stage, .sl-canvas, body')?.focus(), 1`);
    await press(wc, 'p', { modifiers: ['control'] });
    await until(() => js(`[...document.querySelectorAll('.rw-dialog select option')].some((o) => o.value === 'handout')`), 'the print dialog', 5000);
    await choose(js, 'handout', 'handout');
    await until(() => js(`[...document.querySelectorAll('.rw-dialog select option')].some((o) => o.value === '9')`), 'slides to a page', 3000);
    await choose(js, '9', '3');
    await wait(200);
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: path.join(dir, 'print-setup.pdf') });
    await button(js, 'Save as PDF…');
    const kept = await until(() => {
      const s = doc.pageSetup({ id: sessionFor('deck').id });
      return s?.layout === 'handout' && s.perPage === 3;
    }, 'the print kept in the deck', 6000).then(() => true).catch(() => false);
    await until(() => js(`!document.querySelector('.rw-dialog select option[value="handout"]')`), 'the dialog to close', 4000).catch(() => {});
    await wait(800);
    // The next print opens on what the deck keeps.
    await press(wc, 'p', { modifiers: ['control'] });
    await until(() => js(`[...document.querySelectorAll('.rw-dialog select option')].some((o) => o.value === 'handout')`), 'the print dialog again', 5000);
    const reopened = await until(async () => (await valueOf(js, 'handout')) === 'handout' && (await valueOf(js, '9')) === '3', 'the kept choices', 4000).then(() => true).catch(() => false);
    await button(js, 'Cancel');
    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const inFile = await until(() => { try { return Deck.open(fs.readFileSync(deckFile)).printSettings()?.layout === 'handout'; } catch { return false; } }, 'the saved deck', 6000).then(() => true).catch(() => false);
    check('presentations: File → Print keeps a handout, three to a page, in the deck as PowerPoint does, and the next print opens on it',
      kept && reopened && inFile, JSON.stringify({ kept, reopened, inFile }));
    const complaints = await errorsIn(win);
    check('presentations: the print setup reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('presentations: the print setup checks ran', false, err?.message || String(err));
  }

  try {
    fs.writeFileSync(docFile, buildDocx({ styles: true, paragraphs: [{ text: 'A page to turn.' }, { text: 'Its second paragraph.' }] }));
    const win = await open('word', docFile);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="1"]'))`), 'the paragraphs', 8000);
    const shape = () => js(`(() => { const p = document.querySelector('.wd-page'); return p ? { w: p.offsetWidth, h: p.offsetHeight } : null; })()`);
    const before = await shape();
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(wc, 'p', { modifiers: ['control'] });
    await until(() => js(`[...document.querySelectorAll('.rw-dialog select option')].some((o) => o.value === 'landscape')`), 'the print dialog', 5000);
    const paper = await valueOf(js, 'Letter');
    await choose(js, 'landscape', 'landscape');
    const turned = await until(() => doc.pageSetup({ id: sessionFor('doc').id })?.orientation === 'landscape', 'the document turned', 5000).then(() => true).catch(() => false);
    await button(js, 'Cancel');
    const after = await until(async () => { const s = await shape(); return s && s.w > s.h ? s : null; }, 'the pages drawn landscape', 5000).catch(() => null);
    const pageAfter = await shape();
    check('documents: File → Print shows the document\'s own paper, and Landscape chosen there turns the document\'s pages, as Word\'s Print does',
      paper === 'A4' && turned && Boolean(after) && before && before.h > before.w,
      JSON.stringify({ paper, turned, before, after: pageAfter }));
    const complaints = await errorsIn(win);
    check('documents: the print setup reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('documents: the print setup checks ran', false, err?.message || String(err));
  }
}
