// Insert > QR Code in Documents, Worksheets and Presentations: the ribbon's
// button opens the dialog, a web address typed in at a chosen level of error
// correction and a colour goes in as a picture; the saved file's picture is
// read back module by module and decoded to the same words; a double-click
// on the picture opens the dialog again with its words, and changing them
// draws a new code in its place, which decodes to the new words. Run alone
// with RUTBA_VERIFY_ONLY=qrcode.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { SheetView } from '@rutba/sheet-view';
import { Deck, buildPptx } from '@rutba/presentation';
import { decodeQr, modulesFromPicture } from '@rutba/barcode/decode';
import { readPng, isDark } from '@rutba/barcode/png-read';

/** The text a PNG's QR Code holds, and what it was written with. */
function readCode(bytes) {
  const png = readPng(bytes);
  return decodeQr(modulesFromPicture({ width: png.width, height: png.height, dark: (x, y) => isDark(png.rgba(x, y)) }));
}

/** The words of the window's own gestures: a tab, a ribbon button, the dialog filled and confirmed. */
function gestures(win, { until, wait }) {
  const js = (code) => win.webContents.executeJavaScript(code);
  return {
    js,
    async ribbon(tab, tip) {
      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(tab)})?.click(), 1`);
      await wait(250);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(tip)}) && !n.disabled))`), `the ${tip} button`, 4000);
      return js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(tip)}) && !n.disabled); b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
    },
    async fill({ text, level, colour }) {
      await until(() => js(`Boolean(document.querySelector('.rw-dialog .qr-text'))`), 'the QR Code dialog', 5000);
      await js(`(() => {
        const ta = document.querySelector('.rw-dialog .qr-text');
        Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(ta, ${JSON.stringify(text)});
        ta.dispatchEvent(new Event('input', { bubbles: true }));
        const sel = document.querySelector('.rw-dialog .qr-level');
        Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(sel, ${JSON.stringify(level)});
        sel.dispatchEvent(new Event('change', { bubbles: true }));
        ${colour ? `document.querySelector('.rw-dialog .qr-swatch[data-colour="${colour}"]')?.click();` : ''}
        return 1;
      })()`);
      await until(() => js(`(() => { const ok = document.querySelector('.rw-dialog .qr-ok'); const about = document.querySelector('.rw-dialog .qr-about'); return Boolean(ok && !ok.disabled && about); })()`), 'the code drawn in the dialog', 4000);
      await wait(200);
    },
    async ok() {
      await js(`document.querySelector('.rw-dialog .qr-ok').click(), 1`);
      await until(() => js(`!document.querySelector('.rw-dialog .qr-text')`), 'the dialog to close', 5000).catch(() => {});
    },
    /** What the dialog shows when it opens again on a picture. */
    shown: () => js(`(() => ({ text: document.querySelector('.rw-dialog .qr-text')?.value ?? null, level: document.querySelector('.rw-dialog .qr-level')?.value ?? null, colour: document.querySelector('.rw-dialog .qr-swatch.on')?.dataset.colour ?? null, head: document.querySelector('.rw-dialog-head')?.textContent ?? null }))()`),
    dblclick: (selector) => js(`(() => { const n = document.querySelector(${JSON.stringify(selector)}); if (!n) return false; n.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true })); return true; })()`),
  };
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args
 */
export async function verifyQrCode(h, { dir }) {
  const { open, check, until, wait, errorsIn, doc, sessionFor } = h;
  const capture = async (win, name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE) return;
    win.webContents.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await win.webContents.capturePage()).toPNG());
  };
  const FIRST = { level: 'Q', colour: '#1F3864' };
  const SECOND = { level: 'H' };

  /* -- Documents ------------------------------------------------------------ */
  try {
    const file = path.join(dir, 'qr-words.docx');
    fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [{ text: 'Scan the code to read this online.' }] }));
    const win = await open('word', file);
    const g = gestures(win, { until, wait });
    await until(() => g.js(`Boolean(document.querySelector('.wd-page [data-block="0"]'))`), 'the page', 8000);
    const session = sessionFor('doc');
    const text = 'https://rutba.io/office?from=documents';
    const pressed = await g.ribbon('Insert', 'QR Code');
    await g.fill({ text, ...FIRST });
    await capture(win, 'qrcode-dialog.png');
    await g.ok();
    const qrImage = () => (doc.model({ id: session.id }).blocks || []).flatMap((b) => (b.images || []).map((img, i) => ({ ...img, block: b.index ?? null, i }))).find((img) => /^QR Code;/.test(img.title || ''));
    await until(() => Boolean(qrImage()), 'the QR Code picture in the document', 6000).catch(() => {});
    await until(() => g.js(`[...document.querySelectorAll('.wd-image')].some((n) => n.complete && n.naturalWidth > 50)`), 'the picture drawn on the page', 6000).catch(() => {});
    await capture(win, 'qrcode-word.png');
    const out = path.join(dir, 'qr-words-saved.docx');
    await doc.save({ id: session.id, path: out });
    const readBack = (p) => {
      const saved = openDocx(fs.readFileSync(p));
      const img = saved.blocks.flatMap((b) => b.images || []).find((x) => /^QR Code;/.test(x.title || ''));
      const pkg = OoxmlPackage.read(fs.readFileSync(p));
      return img ? { img, code: readCode(pkg.read(img.part)) } : null;
    };
    const first = readBack(out);
    check('qrcode: Documents\' Insert > QR Code puts in a picture whose saved modules decode to the words typed, at the level chosen, the words kept as its alt text',
      pressed === 'clicked' && first?.code.text === text && first.code.level === 'Q' && first.img.descr === text && first.img.title === 'QR Code; Q; #1F3864',
      JSON.stringify({ pressed, decoded: first?.code ?? null, descr: first?.img.descr, title: first?.img.title }));
    // A double-click opens it again with its words, level and colour.
    const clicked = await g.dblclick('.wd-image');
    await until(() => g.js(`Boolean(document.querySelector('.rw-dialog .qr-text'))`), 'the dialog again', 5000).catch(() => {});
    const again = await g.shown();
    const next = 'Changed in Documents';
    await g.fill({ text: next, ...SECOND });
    await g.ok();
    await until(() => qrImage()?.descr === next, 'the new words on the picture', 6000).catch(() => {});
    const out2 = path.join(dir, 'qr-words-saved-2.docx');
    await doc.save({ id: session.id, path: out2 });
    const second = readBack(out2);
    const count = openDocx(fs.readFileSync(out2)).blocks.flatMap((b) => b.images || []).length;
    check('qrcode: double-clicking the code in Documents opens the dialog with its words, level and colour, and OK draws the new code in its place',
      clicked && again.text === text && again.level === 'Q' && again.colour === '#1F3864' && second?.code.text === next && second.code.level === 'H' && count === 1,
      JSON.stringify({ again, decoded: second?.code ?? null, pictures: count }));
    const complaints = await errorsIn(win);
    check('qrcode: Documents reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('qrcode: the Documents checks ran', false, err?.message || String(err));
  }

  /* -- Worksheets ----------------------------------------------------------- */
  try {
    const file = path.join(dir, 'qr-sheet.xlsx');
    fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['Scan to order'], ['Price', 12]] }] }));
    const win = await open('sheets', file);
    const g = gestures(win, { until, wait });
    await until(() => g.js(`document.querySelectorAll('.sh-cell').length > 1`), 'the grid', 8000);
    const session = sessionFor('sheet');
    const text = 'https://rutba.io/office?from=worksheets';
    const pressed = await g.ribbon('Insert', 'QR Code');
    await g.fill({ text, ...FIRST });
    await g.ok();
    await until(() => g.js(`Boolean(document.querySelector('.sh-drawing[data-kind="image"]'))`), 'the picture on the sheet', 6000).catch(() => {});
    await capture(win, 'qrcode-sheets.png');
    const readBack = async (p) => {
      await doc.save({ id: session.id, path: p });
      const view = SheetView.open(fs.readFileSync(p));
      const d = (view.drawings.get('Sheet1') || []).find((x) => x.kind === 'image' && /^QR Code;/.test(x.title || ''));
      return d ? { d, code: readCode(view.pkg.read(d.descriptor.part)) } : null;
    };
    const first = await readBack(path.join(dir, 'qr-sheet-saved.xlsx'));
    check('qrcode: Worksheets\' Insert > QR Code puts in a picture whose saved modules decode to the words typed, the words kept as its alt text',
      pressed === 'clicked' && first?.code.text === text && first.code.level === 'Q' && first.d.descr === text,
      JSON.stringify({ pressed, decoded: first?.code ?? null, descr: first?.d.descr }));
    const clicked = await g.dblclick('.sh-drawing[data-kind="image"]');
    await until(() => g.js(`Boolean(document.querySelector('.rw-dialog .qr-text'))`), 'the dialog again', 5000).catch(() => {});
    const again = await g.shown();
    const next = 'Changed in Worksheets';
    await g.fill({ text: next, ...SECOND });
    await g.ok();
    await wait(600);
    const second = await readBack(path.join(dir, 'qr-sheet-saved-2.xlsx'));
    check('qrcode: double-clicking the code in Worksheets opens the dialog with its words, and OK draws the new code in its place',
      clicked && again.text === text && again.level === 'Q' && second?.code.text === next && second.code.level === 'H',
      JSON.stringify({ again, decoded: second?.code ?? null }));
    const complaints = await errorsIn(win);
    check('qrcode: Worksheets reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('qrcode: the Worksheets checks ran', false, err?.message || String(err));
  }

  /* -- Presentations -------------------------------------------------------- */
  try {
    const file = path.join(dir, 'qr-deck.pptx');
    fs.writeFileSync(file, buildPptx({ slides: [{ layout: 'obj', title: 'Find us online', body: ['Scan the code'] }] }));
    const win = await open('slides', file);
    const g = gestures(win, { until, wait });
    await until(() => g.js(`Boolean(document.querySelector('.sl-svg'))`), 'the slide', 8000);
    const session = sessionFor('deck');
    const text = 'https://rutba.io/office?from=presentations';
    const pressed = await g.ribbon('Insert', 'QR Code');
    await g.fill({ text, ...FIRST });
    await g.ok();
    const qrShape = () => (doc.model({ id: session.id, slide: 0 }).slide?.shapes || []).find((s) => /^QR Code;/.test(s.title || ''));
    await until(() => Boolean(qrShape()), 'the QR Code on the slide', 6000).catch(() => {});
    await capture(win, 'qrcode-slides.png');
    const readBack = async (p) => {
      await doc.save({ id: session.id, path: p });
      const deck = Deck.open(fs.readFileSync(p));
      const s = deck.slide(0).shapes.find((x) => /^QR Code;/.test(x.title || ''));
      return s ? { s, code: readCode(deck.pkg.read(s.source.part)) } : null;
    };
    const first = await readBack(path.join(dir, 'qr-deck-saved.pptx'));
    check('qrcode: Presentations\' Insert > QR Code puts in a picture whose saved modules decode to the words typed, the words kept as its alt text',
      pressed === 'clicked' && first?.code.text === text && first.code.level === 'Q' && first.s.descr === text,
      JSON.stringify({ pressed, decoded: first?.code ?? null, descr: first?.s.descr }));
    const id = qrShape()?.id;
    const clicked = await g.dblclick(`.sl-hit[data-shape="${id}"]`);
    await until(() => g.js(`Boolean(document.querySelector('.rw-dialog .qr-text'))`), 'the dialog again', 5000).catch(() => {});
    const again = await g.shown();
    const next = 'Changed in Presentations';
    await g.fill({ text: next, ...SECOND });
    await g.ok();
    await until(() => qrShape()?.descr === next, 'the new words on the picture', 6000).catch(() => {});
    const second = await readBack(path.join(dir, 'qr-deck-saved-2.pptx'));
    check('qrcode: double-clicking the code in Presentations opens the dialog with its words, and OK draws the new code in its place',
      clicked && again.text === text && again.level === 'Q' && second?.code.text === next && second.code.level === 'H',
      JSON.stringify({ again, decoded: second?.code ?? null }));
    const complaints = await errorsIn(win);
    check('qrcode: Presentations reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('qrcode: the Presentations checks ran', false, err?.message || String(err));
  }
}
