// Pictures: a tile for a picture the platform has no thumbnail for — an
// SVG on Windows, any large photo on Linux — is the picture itself, drawn
// by the window at the tile's size and handed to the platform to keep,
// rather than the kind's icon. Run alone with RUTBA_VERIFY_ONLY=picturetiles.

import fs from 'node:fs';
import path from 'node:path';
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
    const complaints = await errorsIn(win);
    check('pictures: the tiles report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    void wait;
  } catch (err) {
    check('pictures: the tile checks ran', false, err?.message || String(err));
  }
}
