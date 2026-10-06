// View → Window in three apps: Arrange All, Side by Side, Cascade, Switch
// Windows, Hide and Unhide, each pressed on the ribbon and read back from the
// windows' own bounds and visibility.

import fs from 'node:fs';
import path from 'node:path';

const overlap = (a, b) => a.x < b.x + b.width && b.x < a.x + a.width && a.y < b.y + b.height && b.y < a.y + a.height;

/**
 * @param {object} h the harness: open, check, until, wait
 */
export async function verifyWindows(h, { files }) {
  const { open, check, until, wait } = h;
  const copy = (src, name) => {
    const to = path.join(path.dirname(src), name);
    fs.copyFileSync(src, to);
    return to;
  };
  const press = (win, tab, label) => win.webContents.executeJavaScript(`(async () => {
    [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(tab)})?.click();
    await new Promise((r) => setTimeout(r, 200));
    const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)});
    if (!b) return 'no button';
    if (b.disabled) return 'disabled';
    b.click();
    return 'clicked';
  })()`);
  const menuItems = (win) => win.webContents.executeJavaScript(`[...document.querySelectorAll('.rw-menu button')].map((b) => b.textContent.trim())`);
  const pickItem = (win, label) => win.webContents.executeJavaScript(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'no item'; b.click(); return 'picked'; })()`);

  try {
    // Worksheets: three windows tiled, none over another.
    const sheets = [];
    for (const name of ['tile-a.xlsx', 'tile-b.xlsx', 'tile-c.xlsx']) sheets.push(await open('sheets', copy(files.xlsx, name)));
    const pressed = await press(sheets[0], 'View', 'Arrange All');
    const tiled = await until(() => {
      const b = sheets.map((w) => w.getBounds());
      return b.every((x, i) => b.every((y, j) => i === j || !overlap(x, y)));
    }, 'the three windows tiled', 4000).then(() => true).catch(() => false);
    check('windows: Worksheets → View → Arrange All tiles its three windows, none over another',
      pressed === 'clicked' && tiled === true, JSON.stringify({ pressed, bounds: sheets.map((w) => w.getBounds()) }));

    // Hide the second; Unhide from the first lists it and brings it back.
    const hid = await press(sheets[1], 'View', 'Hide');
    const hidden = await until(() => !sheets[1].isVisible(), 'the window to hide', 3000).then(() => true).catch(() => false);
    await sheets[0].webContents.executeJavaScript(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'View')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Unhide');
      b?.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 200, clientY: 120 }));
      return 1;
    })()`);
    await until(async () => (await menuItems(sheets[0])).length > 0, 'the Unhide list', 3000).catch(() => {});
    const offered = await menuItems(sheets[0]);
    const picked = await pickItem(sheets[0], 'tile-b.xlsx');
    const back = await until(() => sheets[1].isVisible(), 'the window to come back', 3000).then(() => true).catch(() => false);
    check('windows: Hide puts a Worksheets window away with its work kept, and Unhide lists only it and brings it back',
      hid === 'clicked' && hidden === true && JSON.stringify(offered) === '["tile-b.xlsx"]' && picked === 'picked' && back === true,
      JSON.stringify({ hid, hidden, offered, picked, back }));

    // Documents: Side by Side — this window and the one before it, half each.
    const docA = await open('word', copy(files.docx, 'side-a.docx'));
    const docB = await open('word', copy(files.docx, 'side-b.docx'));
    const sbs = await press(docB, 'View', 'Side by Side');
    const halves = await until(() => {
      const a = docA.getBounds();
      const b = docB.getBounds();
      return a.y === b.y && a.height === b.height && !overlap(a, b) && Math.abs(a.width - b.width) <= 1;
    }, 'two halves', 4000).then(() => true).catch(() => false);
    check('windows: Documents → View → Side by Side lays this document beside the other, half each, level',
      sbs === 'clicked' && halves === true, JSON.stringify({ sbs, a: docA.getBounds(), b: docB.getBounds() }));

    // Presentations: Cascade steps the windows; Switch Windows lists them.
    const deckA = await open('slides', copy(files.pptx, 'cascade-a.pptx'));
    const deckB = await open('slides', copy(files.pptx, 'cascade-b.pptx'));
    const cascaded = await press(deckA, 'View', 'Cascade');
    const stepped = await until(() => {
      const a = deckA.getBounds();
      const b = deckB.getBounds();
      return a.width === b.width && a.height === b.height && Math.abs(a.x - b.x) === 32 && Math.abs(a.y - b.y) === 32;
    }, 'a cascade', 4000).then(() => true).catch(() => false);
    await deckA.webContents.executeJavaScript(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Switch Windows');
      b?.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 300, clientY: 120 }));
      return 1;
    })()`);
    await until(async () => (await menuItems(deckA)).length > 0, 'the Switch Windows list', 3000).catch(() => {});
    const listed = await menuItems(deckA);
    await deckA.webContents.executeJavaScript(`document.querySelector('.rw-menu') && document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' })), 1`);
    check('windows: Presentations → View → Cascade steps its windows, and Switch Windows lists both decks',
      cascaded === 'clicked' && stepped === true && listed.includes('cascade-a.pptx') && listed.includes('cascade-b.pptx'),
      JSON.stringify({ cascaded, a: deckA.getBounds(), b: deckB.getBounds(), listed }));
    await wait(100);
  } catch (err) {
    check('windows: the window checks ran', false, err.message);
  }
}
