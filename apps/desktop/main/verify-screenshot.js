// Insert → Screenshot in Documents and Presentations, pressed on the ribbon.
// Each takes a picture of the other's window — never of a screen, so a
// check run puts nothing of what is on the desktop into a file — and the
// list a window shows leaves that window out, as Office's does.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, doc
 */
export async function verifyScreenshot(h, { dir }) {
  const { open, check, until, wait, doc } = h;
  const docFile = path.join(dir, 'shot-host.docx');
  const deckFile = path.join(dir, 'shot-target.pptx');
  try {
    fs.writeFileSync(docFile, buildDocx({ styles: true, paragraphs: [{ text: 'A screenshot goes here.' }] }));
    fs.writeFileSync(deckFile, buildPptx({ title: 'Target', slides: [{ layout: 'title', title: 'Target' }] }));
    const deckWin = await open('slides', deckFile);
    const docWin = await open('word', docFile);
    const deckSession = doc.sessions().filter((s) => s.kind === 'deck' && s.path === deckFile).pop();
    const docSession = doc.sessions().filter((s) => s.kind === 'doc' && s.path === docFile).pop();
    const press = (win) => win.webContents.executeJavaScript(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Screenshot');
      if (!b) return 'no button'; if (b.disabled) return 'disabled'; b.click(); return 'clicked';
    })()`);
    // Only the suite's own two windows are named in what the check reports:
    // the rest of the list is whatever else is open on this desktop.
    const ours = (titles) => titles.filter((t) => t.includes('shot-host') || t.includes('shot-target'));
    const listed = (win) => win.webContents.executeJavaScript(`[...document.querySelectorAll('.ss-source')].map((n) => n.title)`);
    const pick = (win, name) => win.webContents.executeJavaScript(`(() => { const b = [...document.querySelectorAll('.ss-source')].find((n) => n.title.includes(${JSON.stringify(name)})); if (!b) return 'no source'; b.click(); return 'picked'; })()`);

    // Documents takes the Presentations window.
    const pressedDoc = await press(docWin);
    await until(async () => (await listed(docWin)).some((t) => t.includes('shot-target')), 'the deck window in the list', 8000).catch(() => {});
    const fromDoc = await listed(docWin);
    const pickedDoc = await pick(docWin, 'shot-target');
    const images = () => (doc.model({ id: docSession.id }).blocks || []).reduce((n, b) => n + (b.images || []).length, 0);
    const inDoc = await until(() => images() > 0, 'the screenshot in the document', 8000).then(() => true).catch(() => false);
    check('documents: Insert → Screenshot lists the other windows, not this one, and puts the one picked in as a picture',
      pressedDoc === 'clicked' && fromDoc.some((t) => t.includes('shot-target')) && !fromDoc.some((t) => t.includes('shot-host')) && pickedDoc === 'picked' && inDoc,
      JSON.stringify({ pressedDoc, ours: ours(fromDoc), others: fromDoc.length, pickedDoc, images: images() }));

    // Presentations takes the Documents window.
    const pressedDeck = await press(deckWin);
    await until(async () => (await listed(deckWin)).some((t) => t.includes('shot-host')), 'the document window in the list', 8000).catch(() => {});
    const fromDeck = await listed(deckWin);
    const pickedDeck = await pick(deckWin, 'shot-host');
    const pictures = () => (doc.model({ id: deckSession.id, slide: 0 }).slide?.shapes || []).filter((s) => s.kind === 'picture');
    const onSlide = await until(() => pictures().length > 0, 'the screenshot on the slide', 8000).then(() => true).catch(() => false);
    const box = pictures()[0]?.geometry;
    const bounds = docWin.getContentBounds();
    const shaped = box ? Math.abs(box.w / box.h - bounds.width / bounds.height) < 0.15 : false;
    check('presentations: Insert → Screenshot lists the other windows, not this one, and puts the one picked on the slide in its own shape',
      pressedDeck === 'clicked' && !fromDeck.some((t) => t.includes('shot-target')) && pickedDeck === 'picked' && onSlide && shaped,
      JSON.stringify({ pressedDeck, ours: ours(fromDeck), others: fromDeck.length, pickedDeck, box, window: bounds }));
    await wait(100);
  } catch (err) {
    check('screenshot: the Screenshot checks ran', false, err.message);
  }
}
