// Documents: Insert → Signature Line, pressed on the ribbon — Signature
// Setup filled in, the line drawn on the page with the signer's name under
// it, and kept in the saved file as Word keeps a signature line.

import fs from 'node:fs';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait
 */
export async function verifyWordSignature(h, { dir }) {
  const { open, check, until, wait } = h;
  const file = path.join(dir, 'signature.docx');
  try {
    fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [{ text: 'Agreed by:' }] }));
    const win = await open('word', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="0"]'))`), 'the paragraph', 8000);
    const pressed = await js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Signature Line');
      if (!b) return 'no button'; if (b.disabled) return 'disabled'; b.click(); return 'clicked';
    })()`);
    await until(() => js(`Boolean(document.querySelector('.wd-sig-signer'))`), 'Signature Setup', 4000).catch(() => {});
    await js(`(() => {
      const set = (sel, v) => { const el = document.querySelector(sel); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, v); el.dispatchEvent(new Event('input', { bubbles: true })); };
      set('.wd-sig-signer', 'Jo Bloggs');
      set('.wd-sig-title', 'Manager');
      set('.wd-sig-email', 'jo@example.com');
      return 1;
    })()`);
    await wait(100);
    await js(`document.querySelector('.wd-sig-ok')?.click(), 1`);
    const line = () => js(`(() => { const im = [...document.querySelectorAll('.wd-page img.wd-image')].find((n) => (n.title || '').startsWith('Signature line')); return im ? { title: im.title, loaded: im.complete && im.naturalWidth > 0, width: Math.round(im.getBoundingClientRect().width) } : null; })()`);
    await until(async () => (await line())?.loaded === true, 'the line on the page', 6000).catch(() => {});
    const drawn = await line();
    check('documents: Insert → Signature Line sets up the line in Signature Setup and draws it on the page for its signer',
      pressed === 'clicked' && drawn?.loaded === true && /Jo Bloggs/.test(drawn.title) && drawn.width > 100, JSON.stringify({ pressed, drawn }));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try { return /<o:signatureline\b[^>]*o:suggestedsigner="Jo Bloggs" o:suggestedsigner2="Manager" o:suggestedsigneremail="jo@example\.com"[^>]*issignatureline="t"/.test(OoxmlPackage.read(fs.readFileSync(file)).text('word/document.xml')); } catch { return false; }
    }, 'the saved signature line', 8000).then(() => true).catch(() => false);
    check('documents: the saved file carries the signature line as Word keeps one, for Word to sign', saved, saved ? 'o:signatureline in the document' : 'not in the file');
  } catch (err) {
    check('documents: the Signature Line checks ran', false, err.message);
  }
}
