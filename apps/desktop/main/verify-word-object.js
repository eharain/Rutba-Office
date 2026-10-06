// Documents: Insert → Object, pressed on the ribbon — Create new embeds a
// blank Excel worksheet as its icon in a paragraph after the caret's; a
// double-click on the icon opens the worksheet in a Worksheets window of
// its own; and the saved document keeps it as Word keeps an embedded one.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml';

/**
 * @param {object} h the harness: open, check, until, wait, doc
 */
export async function verifyWordObject(h, { dir }) {
  const { open, check, until, wait, doc } = h;
  const file = path.join(dir, 'word-object.docx');
  try {
    fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [{ text: 'The figures follow.' }, { text: 'The end.' }] }));
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'doc' && s.path === file).pop();
    const objects = () => (doc.model({ id: session.id }).blocks || []).flatMap((b) => b.images || []).filter((im) => im.kind === 'object');

    await until(() => js(`document.querySelectorAll('.wd-page [data-block]').length >= 2`), 'the page', 8000);
    const pressed = await js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Object');
      if (!b || b.disabled) return 'no button'; b.click(); return 'clicked';
    })()`);
    await until(() => js(`Boolean(document.querySelector('.obj-item[data-ext="xlsx"]'))`), 'the Insert Object box', 4000).catch(() => {});
    await js(`document.querySelector('.obj-item[data-ext="xlsx"]')?.click(), 1`);
    await js(`document.querySelector('.obj-ok')?.click(), 1`);
    const embedded = await until(() => objects().some((o) => o.object.progId === 'Excel.Sheet.12'), 'the embedded worksheet', 8000).then(() => true).catch(() => false);
    const drawn = await until(() => js(`Boolean(document.querySelector('.wd-page img.wd-drawing[data-kind="object"]'))`), 'its icon on the page', 4000).then(() => true).catch(() => false);
    check('documents: Insert → Object → Create new embeds a blank worksheet, drawn on the page as its icon',
      pressed === 'clicked' && embedded && drawn, JSON.stringify({ pressed, embedded, drawn, objects: objects().map((o) => o.object) }));

    const before = new Set(doc.sessions().map((s) => s.id));
    await js(`(() => { const im = document.querySelector('.wd-page img.wd-drawing[data-kind="object"]'); im?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })); return Boolean(im); })()`);
    const opened = await until(() => doc.sessions().some((s) => !before.has(s.id) && s.kind === 'sheet' && /Microsoft_Excel_Worksheet\.xlsx$/.test(s.path || '')), 'the worksheet in Worksheets', 10000).then(() => true).catch(() => false);
    check('documents: a double-click on the embedded worksheet opens it in a Worksheets window of its own', opened, JSON.stringify({ opened }));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const pkg = OoxmlPackage.read(fs.readFileSync(file));
        return pkg.has('word/embeddings/Microsoft_Excel_Worksheet1.xlsx') && /<o:OLEObject Type="Embed" ProgID="Excel\.Sheet\.12"/.test(pkg.text('word/document.xml'));
      } catch { return false; }
    }, 'the saved object', 8000).then(() => true).catch(() => false);
    check('documents: the saved document keeps the embedded worksheet as Word keeps one', saved, saved ? 'o:OLEObject and word/embeddings in the file' : 'not in the file');
    await wait(50);
  } catch (err) {
    check('documents: the Insert Object checks ran', false, err.message);
  }
}
