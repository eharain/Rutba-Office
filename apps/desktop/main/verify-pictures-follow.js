// Pictures: what a person set on their pictures follows the files, and the
// window follows the folder. A picture rated three stars and tagged keeps
// both when it is renamed on disk, and again when it is moved into another
// folder; a file written into the folder while it is open shows once the
// window is come back to; the slideshow fills the window, F11 takes it to
// the whole screen and Esc gives the window back as it was; and Batch edit
// opens with the choices it was last given. Run alone with
// RUTBA_VERIFY_ONLY=picturefollow.
import fs from 'node:fs';
import path from 'node:path';
import { pngOf } from '../../../tests/fixtures/cube-glb.js';

export async function verifyPicturesFollow({ open, check, until, wait, press, errorsIn }, { dir }) {
  try {
    const root = path.join(dir, 'follow-folder');
    const moved = path.join(root, 'Moved');
    fs.mkdirSync(moved, { recursive: true });
    const px = (c) => pngOf(Array.from({ length: 8 }, () => Array.from({ length: 8 }, () => c)));
    fs.writeFileSync(path.join(root, 'alpha.png'), px([200, 40, 40, 255]));
    fs.writeFileSync(path.join(root, 'bravo.png'), px([40, 200, 40, 255]));
    const win = await open('pictures', root);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const capture = async (name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      wc.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
    };
    const names = () => js(`[...document.querySelectorAll('.pv-tile:not(.folder)')].map((t) => t.dataset.name)`);
    const starsOn = (name) => js(`[...document.querySelectorAll('.pv-tile')].find((t) => t.dataset.name === ${JSON.stringify(name)})?.querySelector('.pv-tile-stars')?.dataset.rating || null`);
    const comeBack = () => js(`(() => { window.dispatchEvent(new Event('focus')); return 1; })()`);
    await until(async () => (await names()).length === 2, 'the folder', 8000).catch(() => {});

    // Bravo: three stars and a tag.
    await js(`[...document.querySelectorAll('.pv-tile')].find((t) => t.dataset.name === 'bravo.png')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.pv-stars'))`), 'the details', 5000).catch(() => {});
    win.focus();
    await press(wc, '3');
    await until(() => js(`document.querySelectorAll('.pv-star.on').length === 3`), 'three stars', 3000).catch(() => {});
    await js(`(() => { const i = document.querySelector('.pv-tag-input'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(i, 'garden'); i.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await wait(100);
    await js(`document.querySelector('.pv-tag-input').dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })), 1`);
    await until(() => js(`Boolean(document.querySelector('.pv-tag[data-tag="garden"]'))`), 'the tag', 3000).catch(() => {});
    // The fingerprint is read a moment after the mark is set.
    await wait(600);
    await press(wc, 'Escape');
    await until(async () => (await starsOn('bravo.png')) === '3', 'the stars on the tile', 4000).catch(() => {});

    // Renamed on disk, and a new file written beside it, while the window is open.
    fs.renameSync(path.join(root, 'bravo.png'), path.join(root, 'Garden view.png'));
    fs.writeFileSync(path.join(root, 'charlie.png'), px([40, 40, 200, 255]));
    await comeBack();
    await until(async () => (await names()).includes('Garden view.png') && (await starsOn('Garden view.png')) === '3', 'the renamed picture with its stars', 6000).catch(() => {});
    const afterRename = { names: await names(), stars: await starsOn('Garden view.png') };
    check('pictures: a file written into the open folder shows when the window is come back to, and a picture renamed on disk keeps its stars',
      afterRename.names.includes('charlie.png') && !afterRename.names.includes('bravo.png') && afterRename.stars === '3', JSON.stringify(afterRename));

    // Moved into the folder inside: its stars and its tag are there too.
    fs.renameSync(path.join(root, 'Garden view.png'), path.join(moved, 'Garden view.png'));
    await js(`[...document.querySelectorAll('.pv-tile.folder')].find((t) => t.querySelector('.pv-tile-name')?.textContent === 'Moved')?.click(), 1`);
    await until(async () => (await names()).includes('Garden view.png'), 'the folder inside', 6000).catch(() => {});
    await until(async () => (await starsOn('Garden view.png')) === '3', 'the moved picture with its stars', 6000).catch(() => {});
    const movedStars = await starsOn('Garden view.png');
    await js(`[...document.querySelectorAll('.pv-tile')].find((t) => t.dataset.name === 'Garden view.png')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.pv-tag[data-tag="garden"]'))`), 'the tag in the details', 5000).catch(() => {});
    const movedTag = await js(`Boolean(document.querySelector('.pv-tag[data-tag="garden"]'))`);
    await capture('pictures-follow-moved.png');
    check('pictures: a picture moved into another folder keeps its stars and its tag', movedStars === '3' && movedTag, JSON.stringify({ movedStars, movedTag }));

    // The slideshow fills the window; F11 takes it to the whole screen; Esc ends it and gives the window back.
    await press(wc, 'F5');
    await until(() => js(`Boolean(document.querySelector('.pv-show'))`), 'the slideshow', 4000).catch(() => {});
    const inWindow = await js(`Boolean(document.querySelector('.pv-show'))`) && !win.rutbaCheckFull && !win.isFullScreen();
    const hint = await js(`document.querySelector('[data-role="show-hint"]')?.textContent || ''`);
    await press(wc, 'F11');
    await until(() => Boolean(win.rutbaCheckFull || win.isFullScreen()), 'full screen', 3000).catch(() => {});
    const full = Boolean(win.rutbaCheckFull || win.isFullScreen());
    await press(wc, 'Escape');
    await until(() => js(`!document.querySelector('.pv-show')`), 'the show to end', 3000).catch(() => {});
    await until(() => !win.rutbaCheckFull && !win.isFullScreen(), 'the window back', 3000).catch(() => {});
    const backAgain = !win.rutbaCheckFull && !win.isFullScreen() && await js(`!document.querySelector('.pv-show')`);
    check('pictures: the slideshow fills the window and says so, F11 takes it to the whole screen, and Esc gives the window back',
      inWindow && /F11/.test(hint) && full && backAgain, JSON.stringify({ inWindow, hint, full, backAgain }));

    // Batch edit: a format and a pattern chosen, Cancel; opened again, they are there.
    await press(wc, 'Escape');
    await until(() => js(`document.querySelectorAll('.pv-tile').length >= 1 && !document.querySelector('.pv-stage.viewing')`), 'the folder again', 4000).catch(() => {});
    const openBatch = async () => {
      await js(`[...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Batch edit') && !n.disabled)?.click(), 1`);
      await until(() => js(`Boolean(document.querySelector('[data-role="batch-run"]'))`), 'the Batch edit dialog', 5000).catch(() => {});
    };
    const setValue = (selector, value) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return false;
      const proto = el.tagName === 'SELECT' ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(String(value))});
      el.dispatchEvent(new Event(el.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true }));
      return true;
    })()`);
    const cancel = async () => {
      await js(`[...document.querySelectorAll('.im-batch .rw-btn')].find((n) => n.textContent.trim() === 'Cancel')?.click(), 1`);
      await until(() => js(`!document.querySelector('.im-batch')`), 'the dialog closed', 4000).catch(() => {});
    };
    await openBatch();
    await setValue('[data-role="batch-format"]', 'png');
    await setValue('[data-role="batch-pattern"]', '{name} small');
    await wait(250);
    await cancel();
    await openBatch();
    await until(() => js(`document.querySelector('[data-role="batch-format"]')?.value === 'png'`), 'the format remembered', 3000).catch(() => {});
    const kept = await js(`({ format: document.querySelector('[data-role="batch-format"]')?.value, pattern: document.querySelector('[data-role="batch-pattern"]')?.value })`);
    await cancel();
    await js(`(() => { try { localStorage.removeItem('pictures.batch'); } catch {} return 1; })()`);
    check('pictures: Batch edit opens with the format and name pattern it was last given', kept.format === 'png' && kept.pattern === '{name} small', JSON.stringify(kept));

    const complaints = await errorsIn(win);
    check('pictures: following renamed files and the folder reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('pictures: the follow checks ran', false, err?.message || JSON.stringify(err));
  }
}
