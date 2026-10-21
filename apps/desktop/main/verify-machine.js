// The computer's own clipboard and global shortcuts, checked on their own
// (the gate's machine lane, where nothing else runs). Everywhere else a
// check run uses stand-ins for both (machine-fakes.js), so check processes
// can run side by side; here the system adapters are proved: words, words
// with a table, and a picture written and read back through the system
// clipboard, the suite's clipboard service reaching it from a window, and a
// global shortcut held and let go. What was on the person's clipboard is put
// back at the end. Run alone with RUTBA_VERIFY_ONLY=machine; it runs only
// where the system adapters are in use.

import { clipboard as electronClipboard, globalShortcut, nativeImage } from 'electron';
import { clipboard as adapter, readHTML, readImage, writeParts, writeImage } from '@rutba/office-shell/electron/clipboard';
import { shortcuts } from '@rutba/office-shell/electron/shortcuts';

export async function verifyMachine({ open, check, until }) {
  const c = adapter();
  if (c.kind !== 'system' || shortcuts().kind !== 'system') {
    check('machine: the system adapters are in use in the machine lane', false, `clipboard ${c.kind}, shortcuts ${shortcuts().kind}`);
    return;
  }
  const kept = await electronClipboard.read().catch(() => []);
  try {
    // Words, through the adapter, are on the system clipboard, and back.
    await c.writeText('machine words');
    const sys = await electronClipboard.readText();
    check('machine: words written through the clipboard adapter are on the computer\'s clipboard, and read back', sys === 'machine words' && (await c.readText()) === 'machine words', JSON.stringify(sys));

    // Words and a table together, as a copy from Worksheets puts them.
    await writeParts({ text: 'a\tb', html: '<table><tr><td>a</td><td>b</td></tr></table>' });
    const html = await readHTML();
    check('machine: words and a table go on the computer\'s clipboard together and come back as they went', /<td>a<\/td>/.test(html) && (await electronClipboard.readText()) === 'a\tb', html.slice(0, 80));

    // A picture.
    const png = nativeImage.createFromBitmap(Buffer.alloc(2 * 2 * 4, 200), { width: 2, height: 2 });
    await writeImage(png);
    const back = await readImage();
    check('machine: a picture goes on the computer\'s clipboard and comes back at its size', !back.isEmpty() && back.getSize().width === 2 && back.getSize().height === 2, JSON.stringify(back.getSize()));

    // The suite's own clipboard service, from a window, reaching the system clipboard.
    const win = await open('word');
    await until(() => win.webContents.executeJavaScript('Boolean(window.rutbaOffice?.clipboard)'), 'the window', 10000).catch(() => {});
    await win.webContents.executeJavaScript(`window.rutbaOffice.clipboard.writeText({ text: 'from a window' })`);
    const fromWindow = await electronClipboard.readText();
    await electronClipboard.writeText('from the computer');
    const intoWindow = await win.webContents.executeJavaScript('window.rutbaOffice.clipboard.readText()');
    check('machine: the clipboard service a window uses writes to and reads from the computer\'s clipboard',
      fromWindow === 'from a window' && intoWindow === 'from the computer', JSON.stringify({ fromWindow, intoWindow }));
    win.close();

    // A global shortcut held through the adapter is held by the computer, and let go.
    const key = 'CommandOrControl+Alt+Shift+F12';
    const held = shortcuts().register(key, () => {});
    const registered = globalShortcut.isRegistered(key);
    shortcuts().unregister(key);
    check('machine: a global shortcut held through the adapter is held by the computer, and let go',
      held && registered && !globalShortcut.isRegistered(key), JSON.stringify({ held, registered }));
  } catch (err) {
    check('machine: the system adapter checks ran', false, err?.message || String(err));
  } finally {
    // What the person had on the clipboard, put back.
    try { if (kept.length) await electronClipboard.write(kept); else await electronClipboard.clear(); } catch { /* the clipboard is busy: nothing more to do */ }
  }
}
