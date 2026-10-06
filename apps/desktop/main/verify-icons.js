// Insert → Icons in Documents, Worksheets and Presentations, pressed on the
// ribbon — the gallery found by its search, a colour picked, and the icon
// put in as a picture each app holds like any other.

import fs from 'node:fs';
import path from 'node:path';
import { nativeImage } from 'electron';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, doc
 */
export async function verifyIcons(h, { dir }) {
  const { open, check, until, wait, doc } = h;
  const files = {
    word: path.join(dir, 'icons.docx'),
    sheets: path.join(dir, 'icons.xlsx'),
    slides: path.join(dir, 'icons.pptx'),
  };
  fs.writeFileSync(files.word, buildDocx({ styles: true, paragraphs: [{ text: 'An icon:' }] }));
  fs.writeFileSync(files.sheets, buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['An icon']] }] }));
  fs.writeFileSync(files.slides, buildPptx({ title: 'Icons', slides: [{ layout: 'title', title: 'Icons' }] }));
  const kinds = { word: 'doc', sheets: 'sheet', slides: 'deck' };
  const counted = {
    word: (id) => (doc.model({ id }).blocks || []).reduce((n, b) => n + (b.images || []).length, 0),
    sheets: (id) => (doc.model({ id }).drawings || []).length,
    slides: (id) => (doc.model({ id, slide: 0 }).slide?.shapes || []).filter((s) => s.kind === 'picture').length,
  };
  const names = { word: 'documents', sheets: 'worksheets', slides: 'presentations' };

  for (const app of ['word', 'sheets', 'slides']) {
    try {
      const win = await open(app, files[app]);
      const js = (code) => win.webContents.executeJavaScript(code);
      const session = doc.sessions().filter((s) => s.kind === kinds[app] && s.path === files[app]).pop();
      await wait(400);
      const before = counted[app](session.id);
      const pressed = await js(`(async () => {
        [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click();
        await new Promise((r) => setTimeout(r, 200));
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Icons');
        if (!b) return 'no button'; if (b.disabled) return 'disabled'; b.click(); return 'clicked';
      })()`);
      await until(() => js(`Boolean(document.querySelector('.ic-find'))`), 'the Icons gallery', 4000).catch(() => {});
      const found = await js(`(async () => {
        const find = document.querySelector('.ic-find');
        Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(find, 'heart');
        find.dispatchEvent(new Event('input', { bubbles: true }));
        await new Promise((r) => setTimeout(r, 150));
        return [...document.querySelectorAll('.ic-cell')].map((n) => n.dataset.icon);
      })()`);
      await js(`(() => { document.querySelector('.ic-cell[data-icon="heart"]')?.click(); document.querySelector('.ic-swatch[data-colour="#C62828"]')?.click(); return 1; })()`);
      await wait(100);
      await js(`document.querySelector('.ic-insert')?.click(), 1`);
      const added = await until(() => counted[app](session.id) > before, 'the icon put in', 6000).then(() => true).catch(() => false);
      // In the deck, saved: the picture in the file is drawn in the red picked.
      let red = null;
      if (app === 'slides') {
        await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
        red = await until(() => {
          try {
            const pkg = OoxmlPackage.read(fs.readFileSync(files.slides));
            const media = pkg.partNames().filter((n) => n.startsWith('ppt/media/') && n.endsWith('.png'));
            if (!media.length) return false;
            const bitmap = nativeImage.createFromBuffer(pkg.read(media[media.length - 1])).toBitmap();
            // BGRA: a stroke pixel near #C62828.
            for (let i = 0; i < bitmap.length; i += 4) if (bitmap[i + 3] > 200 && bitmap[i + 2] > 170 && bitmap[i + 1] < 80 && bitmap[i] < 80) return true;
            return false;
          } catch {
            return false;
          }
        }, 'the red icon in the saved deck', 8000).then(() => true).catch(() => false);
      }
      check(`${names[app]}: Insert → Icons finds an icon by its name and puts it in, in the colour picked, as a picture`,
        pressed === 'clicked' && JSON.stringify(found) === '["heart"]' && added && red !== false, JSON.stringify({ pressed, found, before, after: counted[app](session.id), red }));
    } catch (err) {
      check(`${names[app]}: the Icons checks ran`, false, err.message);
    }
  }
}
