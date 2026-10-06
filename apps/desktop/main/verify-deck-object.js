// Presentations: Insert → Object, pressed on the ribbon — Create from file
// embeds a Word document from this computer as its icon; a double-click on
// it opens the embedded document in a Documents window of its own; Create
// new embeds a blank Excel worksheet; and the saved deck keeps both as
// PowerPoint keeps embedded documents.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';
import { buildDocx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait, doc, sessionFor
 */
export async function verifyDeckObject(h, { dir }) {
  const { open, check, until, wait, doc, sessionFor } = h;
  const file = path.join(dir, 'object.pptx');
  const letter = path.join(dir, 'embedded-letter.docx');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Objects', slides: [{ layout: 'blank' }] }));
    fs.writeFileSync(letter, buildDocx({ styles: true, paragraphs: [{ text: 'A letter on a slide.' }] }));
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const objects = () => (doc.model({ id: sessionFor('deck').id }).slide?.shapes || []).filter((s) => s.object);
    const press = () => js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Object');
      if (!b || b.disabled) return 'no button'; b.click(); return 'clicked';
    })()`);

    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 1`), 'the slide', 8000);
    const pressed = await press();
    await until(() => js(`Boolean(document.querySelector('.obj-ok'))`), 'the Insert Object box', 4000).catch(() => {});
    await js(`document.querySelector('.obj-mode-file')?.click(), 1`);
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [letter] });
    await js(`document.querySelector('.obj-browse')?.click(), 1`);
    await until(() => js(`(document.querySelector('.obj-path')?.textContent || '').includes('embedded-letter')`), 'the file picked', 4000).catch(() => {});
    await js(`document.querySelector('.obj-ok')?.click(), 1`);
    const embedded = await until(() => objects().length === 1 && objects()[0].object.progId === 'Word.Document.12', 'the embedded letter', 8000).then(() => true).catch(() => false);
    check('presentations: Insert → Object → Create from file embeds a Word document on the slide as its icon',
      pressed === 'clicked' && embedded, JSON.stringify({ pressed, objects: objects().map((s) => s.object) }));

    // A double-click opens it in Documents.
    const before = new Set(doc.sessions().map((s) => s.id));
    const at = await js(`(() => { const id = ${JSON.stringify(String(objects()[0]?.id))}; const hit = document.querySelector('.sl-hit[data-shape="' + id + '"]'); if (!hit) return null; const r = hit.getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; })()`);
    if (at) {
      wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 2 });
      wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 2 });
    }
    const opened = await until(() => doc.sessions().some((s) => !before.has(s.id) && s.kind === 'doc' && /Microsoft_Word_Document\.docx$/.test(s.path || '')), 'the letter in Documents', 10000).then(() => true).catch(() => false);
    const words = (() => { const s = doc.sessions().find((x) => !before.has(x.id) && x.kind === 'doc'); return s ? (doc.model({ id: s.id }).blocks || []).map((b) => b.text).join(' ') : ''; })();
    check('presentations: a double-click on the embedded document opens it in a Documents window of its own', Boolean(at) && opened && /A letter on a slide/.test(words), JSON.stringify({ at, opened, words }));

    // Create new: a blank worksheet.
    await press();
    await until(() => js(`Boolean(document.querySelector('.obj-item[data-ext="xlsx"]'))`), 'the box again', 4000).catch(() => {});
    await js(`document.querySelector('.obj-item[data-ext="xlsx"]')?.click(), 1`);
    await js(`document.querySelector('.obj-ok')?.click(), 1`);
    const sheet = await until(() => objects().some((s) => s.object.progId === 'Excel.Sheet.12'), 'the embedded worksheet', 8000).then(() => true).catch(() => false);

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const deck = Deck.open(fs.readFileSync(file));
        const names = deck.pkg.partNames();
        return names.includes('ppt/embeddings/Microsoft_Word_Document1.docx') && names.includes('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx') && /progId="Excel\.Sheet\.12"/.test(deck.pkg.text('ppt/slides/slide1.xml'));
      } catch { return false; }
    }, 'the saved objects', 8000).then(() => true).catch(() => false);
    check('presentations: Insert → Object → Create new embeds a blank worksheet, and the saved deck keeps both documents as PowerPoint keeps them', sheet && saved, JSON.stringify({ sheet, saved }));
    await wait(50);
  } catch (err) {
    check('presentations: the Insert Object checks ran', false, err.message);
  }
}
