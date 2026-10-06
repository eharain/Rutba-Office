// Presentations: Insert → Photo Album, pressed on the ribbon — three
// pictures picked (the system's Open dialog answered as a person would),
// two to a slide with captions, and Create opens a new window on the album:
// a title slide and two slides of pictures, unsaved.

import fs from 'node:fs';
import path from 'node:path';
import { BrowserWindow } from 'electron';
import { buildPptx } from '@rutba/presentation';

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC', 'base64');

/**
 * @param {object} h the harness: open, check, until, wait, doc
 */
export async function verifyDeckAlbum(h, { dir }) {
  const { open, check, until, wait, doc } = h;
  const file = path.join(dir, 'album-host.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Host', slides: [{ layout: 'title', title: 'Host' }] }));
    const pictures = ['Beach.png', 'Hills.png', 'Harbour.png'].map((n) => { const f = path.join(dir, n); fs.writeFileSync(f, PNG); return f; });
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const pressed = await js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Photo Album');
      if (!b) return 'no button'; if (b.disabled) return 'disabled'; b.click(); return 'clicked';
    })()`);
    await until(() => js(`Boolean(document.querySelector('.sl-album'))`), 'the Photo Album dialog', 4000).catch(() => {});
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: pictures });
    await js(`document.querySelector('.sl-album-add')?.click(), 1`);
    await until(() => js(`document.querySelectorAll('.sl-album-row').length === 3`), 'the three pictures listed', 4000).catch(() => {});
    const listed = await js(`[...document.querySelectorAll('.sl-album-row')].map((n) => n.textContent.replace(/^\\d+/, ''))`);
    await js(`(() => {
      const layout = document.querySelector('.sl-album-layout');
      Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(layout, '2');
      layout.dispatchEvent(new Event('change', { bubbles: true }));
      document.querySelector('.sl-album-captions').click();
      const title = document.querySelector('.sl-album-title');
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(title, 'Trip');
      title.dispatchEvent(new Event('input', { bubbles: true }));
      return 1;
    })()`);
    await wait(100);
    const before = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
    await js(`document.querySelector('.sl-album-create')?.click(), 1`);
    const opened = await until(() => BrowserWindow.getAllWindows().some((w) => !before.has(w.id)), 'the album\'s window', 6000).then(() => true).catch(() => false);
    const album = BrowserWindow.getAllWindows().find((w) => !before.has(w.id));
    const session = doc.sessions().find((s) => s.kind === 'deck' && !s.path && s.id !== doc.sessions().find((x) => x.path === file)?.id && doc.model({ id: s.id }).count === 3);
    const drawn = album ? await until(() => album.webContents.executeJavaScript(`document.querySelectorAll('.sl-thumb, .sl-sorter button, [data-slide]').length >= 3 || document.title.includes('Photo Album')`), 'the album drawn', 8000).then(() => true).catch(() => false) : false;
    const shown = album ? await album.webContents.executeJavaScript('document.title') : null;
    const second = session ? doc.model({ id: session.id, slide: 1 }).slide : null;
    const pics = (second?.shapes || []).filter((s) => s.kind === 'picture').length;
    const captions = (second?.shapes || []).filter((s) => s.kind !== 'picture' && s.text).map((s) => s.text.paragraphs.map((p) => p.plain).join(''));
    check('presentations: Insert → Photo Album makes a new presentation of the picked pictures, two to a slide with captions, in a window of its own',
      pressed === 'clicked' && JSON.stringify(listed) === '["Beach.png","Hills.png","Harbour.png"]' && opened && drawn && Boolean(session) && pics === 2 && JSON.stringify(captions) === '["Beach","Hills"]',
      JSON.stringify({ pressed, listed, opened, drawn, shown, session: session?.id, pics, captions }));
  } catch (err) {
    check('presentations: the Photo Album checks ran', false, err.message);
  }
}
