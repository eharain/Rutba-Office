// Pictures: a tile for a picture the platform has no thumbnail for — an
// SVG on Windows, any large photo on Linux — is the picture itself, drawn
// by the window at the tile's size and handed to the platform to keep,
// rather than the kind's icon. And a tile's own menu acts on the tile
// right-clicked: with nothing on show, and with another picture on show
// beside the folder's grid, its Copy the path copies that tile's path.
// Run alone with RUTBA_VERIFY_ONLY=picturetiles.

import fs from 'node:fs';
import path from 'node:path';
import { clipboard } from '@rutba/office-shell/electron/clipboard';
import { gradientPng } from './sample-picture.js';

const DRAWING = '<svg xmlns="http://www.w3.org/2000/svg" width="400" height="300" viewBox="0 0 400 300">'
  + '<rect width="400" height="300" fill="#2f5597"/><circle cx="200" cy="150" r="90" fill="#f4b183"/></svg>';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the folder is made
 */
export async function verifyPictureTiles({ open, check, until, wait, errorsIn }, { dir }) {
  const root = path.join(dir, 'tiles');
  try {
    fs.mkdirSync(root, { recursive: true });
    fs.writeFileSync(path.join(root, 'a-photo.png'), gradientPng(320, 200, [40, 90, 160], [220, 200, 120]));
    fs.writeFileSync(path.join(root, 'b-drawing.svg'), DRAWING);
    const win = await open('pictures', path.join(root, 'a-photo.png'));
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.pv-image')) && document.querySelector('.pv-image').naturalWidth > 0`), 'the first picture', 8000);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith('Back to the folder')); b?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b?.click(); return 1; })()`);
    await until(() => js(`document.querySelectorAll('.pv-tile').length >= 2`), 'the tiles', 6000);
    const shown = await until(() => js(`Boolean(document.querySelector('.pv-tile[data-name="b-drawing.svg"] img.pv-thumb.loaded'))`), 'the drawing\'s tile', 15000).then(() => true).catch(() => false);
    const tile = await js(`(() => { const t = document.querySelector('.pv-tile[data-name="b-drawing.svg"]'); const img = t?.querySelector('img'); return { src: (img?.src || 'icon').split(':')[0], w: img?.naturalWidth || 0, h: img?.naturalHeight || 0, icon: Boolean(t?.querySelector('.pv-tile-icon')) }; })()`);
    // Drawn by the window: kept by the platform, which serves it from then on.
    const kept = tile.src === 'blob'
      ? await until(() => js(`fetch(${JSON.stringify(`rutba://thumb/${Buffer.from(path.join(root, 'b-drawing.svg')).toString('base64url')}?s=256`)}).then((r) => r.status === 200)`), 'the tile kept', 4000).then(() => true).catch(() => false)
      : true;
    check('pictures: a picture the platform has no tile for is drawn by the window at the tile\'s size, not shown as an icon, and kept for next time',
      shown && !tile.icon && tile.w > 0 && tile.w <= 256 && tile.h <= 256 && kept,
      `${tile.src === 'blob' ? 'drawn by the window' : tile.src === 'rutba' ? 'from the platform' : 'an icon'} (${tile.w}×${tile.h}); kept: ${kept}`);
    // The tile's menu, with nothing on show: Copy the path copies the
    // tile's own path, where it did nothing before.
    const clipWas = await clipboard().readText();
    const viaMenu = async (name, item) => {
      await js(`(() => { const t = document.querySelector('.pv-tile[data-name=${JSON.stringify(name)}]'); const r = t.getBoundingClientRect(); t.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: r.left + 10, clientY: r.top + 10 })); return 1; })()`);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith(${JSON.stringify(item)})))`), 'the tile\'s menu', 3000);
      const enabled = await js(`!([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith(${JSON.stringify(item)}))?.disabled)`);
      await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith(${JSON.stringify(item)})).click(); return 1; })()`);
      return enabled;
    };
    const norm = (p) => String(p || '').replace(/\\/g, '/').toLowerCase();
    const drawing = path.join(root, 'b-drawing.svg');
    await clipboard().writeText('');
    const enabledNone = await viaMenu('b-drawing.svg', 'Copy the path');
    const copiedNone = await until(async () => norm(await clipboard().readText()) === norm(drawing), 'the path copied', 3000).then(() => true).catch(() => false);
    // With the photo on show and the grid beside it: the drawing's tile.
    await js(`document.querySelector('.pv-tile[data-name="a-photo.png"]')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.pv-image'))`), 'the photo on show', 6000);
    await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'View')?.click(); return 1; })()`);
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-btn')].find((b) => b.textContent.trim() === 'Grid'))`), 'the Grid button', 3000);
    const gridWasOn = await js(`Boolean(document.querySelector('.pv-image') && document.querySelector('.pv-tile[data-name="b-drawing.svg"]'))`);
    if (!gridWasOn) await js(`[...document.querySelectorAll('.rw-btn')].find((b) => b.textContent.trim() === 'Grid').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.pv-tile[data-name="b-drawing.svg"]'))`), 'the grid beside the photo', 4000);
    await clipboard().writeText('');
    await viaMenu('b-drawing.svg', 'Copy the path');
    await wait(300);
    const copiedBeside = await clipboard().readText();
    const stillShown = await js(`Boolean(document.querySelector('.pv-image')) && document.querySelector('.pv-tile.active')?.dataset.name === 'a-photo.png'`);
    if (!gridWasOn) await js(`[...document.querySelectorAll('.rw-btn')].find((b) => b.textContent.trim() === 'Grid')?.click(), 1`);
    await clipboard().writeText(clipWas || '');
    check('pictures: a tile\'s menu acts on the tile right-clicked, with nothing on show and beside another picture',
      enabledNone && copiedNone && norm(copiedBeside) === norm(drawing) && stillShown,
      JSON.stringify({ enabledNone, copiedNone, copiedBeside: path.basename(copiedBeside || ''), stillShown }));

    const complaints = await errorsIn(win);
    check('pictures: the tiles report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    void wait;
  } catch (err) {
    check('pictures: the tile checks ran', false, err?.message || String(err));
  }
}
