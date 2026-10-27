// Pictures: what a window is opened on. Edit this picture, on the picture
// Pictures was opened with, opens it in Image rather than bringing Pictures
// back to the front; a folder of folders only says how many folders it
// holds, with no "All 0" to choose; and a picture opened by a path written
// with forward slashes is found among its folder's pictures ("2 of 3").
// Run alone with RUTBA_VERIFY_ONLY=pictures-open.
import fs from 'node:fs';
import path from 'node:path';
import { BrowserWindow } from 'electron';
import { pngOf } from '../../../tests/fixtures/cube-glb.js';

export async function verifyPicturesOpen({ open, check, until, wait, errorsIn }, { dir, windows }) {
  const root = path.join(dir, 'pictures-open');
  const opened = [];
  try {
    const pics = path.join(root, 'Photos');
    const parent = path.join(root, 'Albums');
    fs.mkdirSync(pics, { recursive: true });
    for (const name of ['Spring', 'Summer']) fs.mkdirSync(path.join(parent, name), { recursive: true });
    const px = (c, n) => pngOf(Array.from({ length: n }, () => Array.from({ length: n }, () => c)));
    // Three pictures of their own sizes, so none is taken for another.
    fs.writeFileSync(path.join(pics, 'a.png'), px([200, 40, 40, 255], 8));
    fs.writeFileSync(path.join(pics, 'b.png'), px([40, 200, 40, 255], 10));
    fs.writeFileSync(path.join(pics, 'c.png'), px([40, 40, 200, 255], 12));
    const picture = path.join(pics, 'b.png');

    /** A window opened as Explorer or another window opens one: through windows.open. */
    const opening = async (app, file) => {
      const win = windows.open({ app, file });
      opened.push(win);
      if (win.webContents.isLoading()) await new Promise((resolve) => win.webContents.once('did-finish-load', resolve));
      await until(() => win.webContents.executeJavaScript(`Boolean(document.querySelector('.rw-app'))`), `the ${app} window to draw`, 12000).catch(() => {});
      return win;
    };
    const imageOn = (file) => BrowserWindow.getAllWindows().find((w) => {
      if (w.isDestroyed()) return false;
      const q = new URL(w.webContents.getURL()).searchParams;
      return q.get('app') === 'image' && q.get('file') && path.resolve(q.get('file')) === path.resolve(file);
    });
    const counter = (win) => win.webContents.executeJavaScript(`[...document.querySelectorAll('.chip')].map((c) => c.textContent.trim()).find((s) => / of \\d+$/.test(s)) || null`);

    // A folder holding only folders, opened first: a window stays open through the checks below,
    // which close theirs (the last window closed would quit the application).
    const albums = await open('pictures', parent);

    // Edit on the picture Pictures was opened with.
    const pv = await opening('pictures', picture);
    await until(async () => (await counter(pv)) === '2 of 3', 'the picture shown', 12000).catch(() => {});
    // Home > Edit on the ribbon (Details' Edit this picture runs the same command).
    const pressed = await pv.webContents.executeJavaScript(`(() => { const all = [...document.querySelectorAll('button:not([disabled])')]; const b = all.find((x) => x.textContent.trim() === 'Edit this picture') || all.find((x) => x.textContent.trim() === 'Edit'); b?.click(); return b ? b.textContent.trim() : null; })()`);
    let image = null;
    await until(() => (image = imageOn(picture)), 'an Image window on the picture', 12000).catch(() => {});
    if (image) opened.push(image);
    check('pictures: Edit this picture, on the picture Pictures was opened with, opens it in Image',
      pressed && Boolean(image) && image !== pv, JSON.stringify({ pressed, image: Boolean(image) }));

    // The same picture by a path written with forward slashes: found among its folder's.
    // Pictures on it already would only be brought to the front, so it goes first.
    const errors = [];
    for (const w of [pv, image]) if (w && !w.isDestroyed()) { errors.push(...(await errorsIn(w))); w.destroy(); }
    await wait(200);
    const slashed = await opening('pictures', picture.split(path.sep).join('/'));
    let shown = null;
    await until(async () => (shown = await counter(slashed)) === '2 of 3', 'the picture by its slashed path', 12000).catch(() => {});
    check('pictures: a picture opened by a path with forward slashes is the second of its folder\'s three, not "0 / 3"',
      shown === '2 of 3', JSON.stringify({ shown }));

    // The folder of folders.
    const read = () => albums.webContents.executeJavaScript(`(() => {
      const chips = [...document.querySelectorAll('.chip')].map((c) => c.textContent.trim());
      const kinds = document.querySelector('.pv-kinds');
      return { chips, kinds: kinds ? { hidden: kinds.hidden, drawn: kinds.getClientRects().length > 0, text: kinds.textContent.trim() } : null,
        folders: document.querySelectorAll('.pv-tile.folder').length };
    })()`);
    let seen = null;
    await until(async () => (seen = await read()).folders === 2, 'the two folders', 8000).catch(() => {});
    check('pictures: a folder of folders only says "2 folders", not "nothing here", and shows no "All 0"',
      seen?.folders === 2 && seen.chips.includes('2 folders') && !seen.chips.includes('nothing here') && (!seen.kinds || !seen.kinds.drawn), JSON.stringify(seen));

    for (const w of [...opened, albums]) if (!w.isDestroyed()) errors.push(...(await errorsIn(w)));
    check('pictures-open: those windows report nothing', errors.length === 0, errors.slice(0, 3).join(' | ') || 'nothing reported');
  } catch (e) {
    check('pictures-open: the checks ran', false, e.stack || e.message);
  } finally {
    // Only the windows opened here as Explorer opens them; the folder's is the checks' own, closed with the rest.
    for (const w of opened) if (w && !w.isDestroyed()) w.destroy();
    await wait(200);
  }
}
