// Pictures: ratings, tags, folder counts and Subfolders — a folder of three
// pictures and a folder inside it: the inner folder's tile says how much it
// holds; a picture opened and 4 pressed has four stars, in the details and
// on its tile; a tag added there is found by the filter; the rating filter
// keeps only the four-star picture; and Subfolders finds the picture inside
// the inner folder by its name. Run alone with RUTBA_VERIFY_ONLY=picturemarks.
import fs from 'node:fs';
import path from 'node:path';
import { pngOf } from '../../../tests/fixtures/cube-glb.js';

export async function verifyPicturesMarks({ open, check, until, wait, press, errorsIn }, { dir }) {
  try {
    const root = path.join(dir, 'marks-folder');
    const inner = path.join(root, 'Trips');
    fs.mkdirSync(inner, { recursive: true });
    const px = (c) => pngOf(Array.from({ length: 8 }, () => Array.from({ length: 8 }, () => c)));
    fs.writeFileSync(path.join(root, 'alpha.png'), px([200, 40, 40, 255]));
    fs.writeFileSync(path.join(root, 'bravo.png'), px([40, 200, 40, 255]));
    fs.writeFileSync(path.join(root, 'charlie.png'), px([40, 40, 200, 255]));
    fs.writeFileSync(path.join(inner, 'beach.png'), px([240, 220, 120, 255]));
    const win = await open('pictures', root);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const capture = async (name) => { if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG()); };
    const names = () => js(`[...document.querySelectorAll('.pv-tile:not(.folder)')].map((t) => t.dataset.name)`);
    await until(async () => (await names()).length === 3, 'the folder', 8000).catch(() => {});

    await until(() => js(`Boolean(document.querySelector('.pv-tile.folder .pv-folder-count'))`), 'the folder count', 4000).catch(() => {});
    const count = await js(`document.querySelector('.pv-tile.folder .pv-folder-count')?.textContent || ''`);
    check('pictures: a folder\'s tile says how much it holds before it is entered', count === '1 item', JSON.stringify(count));

    // Open bravo, press 4.
    await js(`[...document.querySelectorAll('.pv-tile')].find((t) => t.dataset.name === 'bravo.png')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.pv-stars'))`), 'the details', 5000).catch(() => {});
    win.focus();
    await press(wc, '4');
    await until(() => js(`document.querySelectorAll('.pv-star.on').length === 4`), 'four stars', 3000).catch(() => {});
    const stars = await js(`document.querySelectorAll('.pv-star.on').length`);
    // A tag, typed into the details.
    await js(`(() => { const i = document.querySelector('.pv-tag-input'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(i, 'holiday'); i.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await wait(100);
    await js(`document.querySelector('.pv-tag-input').dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })), 1`);
    await until(() => js(`Boolean(document.querySelector('.pv-tag[data-tag="holiday"]'))`), 'the tag', 3000).catch(() => {});
    await wait(300);
    await capture('pictures-marks-details.png');
    const tagged = await js(`Boolean(document.querySelector('.pv-tag[data-tag="holiday"]'))`);
    check('pictures: 4 gives the open picture four stars, and a tag typed into the details is kept', stars === 4 && tagged, JSON.stringify({ stars, tagged }));

    // Back to the folder: the stars on its tile; the filter finds the tag; the rating filter keeps it alone.
    await press(wc, 'Escape');
    await until(async () => (await names()).length === 3, 'the folder again', 4000).catch(() => {});
    const onTile = await js(`[...document.querySelectorAll('.pv-tile')].find((t) => t.dataset.name === 'bravo.png')?.querySelector('.pv-tile-stars')?.dataset.rating || null`);
    await capture('pictures-marks-grid.png');
    const setFilter = (q) => js(`(() => { const i = document.querySelector('.pv-tools .rw-search input'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(i, ${JSON.stringify(q)}); i.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await setFilter('holi');
    await until(async () => (await names()).length === 1, 'the tag filter', 3000).catch(() => {});
    const byTag = await names();
    await setFilter('');
    await until(async () => (await names()).length === 3, 'the filter cleared', 3000).catch(() => {});
    await js(`(() => { const s = document.querySelector('select[data-role="rated"]'); Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, '4'); s.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
    await until(async () => (await names()).length === 1, 'the rating filter', 3000).catch(() => {});
    const byRating = await names();
    await js(`(() => { const s = document.querySelector('select[data-role="rated"]'); Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, '0'); s.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
    check('pictures: the stars show on the tile, the filter finds the tag, and the rating filter keeps only the four-star picture',
      onTile === '4' && JSON.stringify(byTag) === '["bravo.png"]' && JSON.stringify(byRating) === '["bravo.png"]', JSON.stringify({ onTile, byTag, byRating }));

    // Subfolders: the picture inside Trips, by its name.
    await js(`document.querySelector('.pv-deep')?.click(), 1`);
    await setFilter('beach');
    await until(async () => JSON.stringify(await names()) === '["beach.png"]', 'the search through the folders inside', 6000).catch(() => {});
    const deep = await names();
    check('pictures: Subfolders finds a picture in the folders inside by its name', JSON.stringify(deep) === '["beach.png"]', JSON.stringify(deep));

    const complaints = await errorsIn(win);
    check('pictures: ratings, tags and Subfolders report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('pictures: the rating and tag checks ran', false, err?.message || JSON.stringify(err));
  }
}
