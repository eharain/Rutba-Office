// Mail: a double-click on a message in the list opens it in a window of its
// own (owner, 2026-10-19: "the double click on mail in mail list should open
// a new window instead in the right pane").
//
// A message arrives and is double-clicked in the list, with a real mouse. A
// second Mail window opens on that message alone: its subject in the title,
// the message read, and no folders or list beside it. The arrow keys there
// stay on the message rather than moving to another. The first window still
// has its list. Deleted from its own window, the message goes to Deleted and
// the window closes. Run alone with RUTBA_VERIFY_ONLY=mail-window.

import { BrowserWindow } from 'electron';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 */
export async function verifyMailWindow(h) {
  const { open, check, until, wait, errorsIn } = h;
  const stamp = Date.now();
  const subject = `Window: the board pack ${stamp}`;
  const raw = [
    'From: Amina Yusuf <amina@work.example>', 'To: You <you@example.com>', `Subject: ${subject}`, `Date: ${new Date(stamp).toUTCString()}`,
    `Message-ID: <window-${stamp}@example.org>`, 'MIME-Version: 1.0', 'Content-Type: text/plain; charset=utf-8', '',
    'The board pack goes out on Friday. Anything you want changed should reach me by Thursday lunchtime.', '',
  ].join('\r\n');
  let win = null;
  let alone = null;
  try {
    win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the seeded account in the sidebar', 8000);
    const accountId = await js(`(async () => (await window.rutbaOffice.mail.accounts())[0]?.id || null)()`);
    const mail = (method, args) => js(`window.rutbaOffice.mail.${method}(${JSON.stringify(args)})`);
    await mail('deliverTest', { accountId, folder: 'Inbox', raw });
    const folders = await mail('folders', { accountId });
    const inbox = folders.find((f) => /^inbox$/i.test(f.name || f.path))?.path || 'Inbox';
    const openInbox = () => js(`(() => { const f = [...document.querySelectorAll('[title="${inbox}"]')].find((n) => n.textContent.trim().startsWith(${JSON.stringify(inbox)})); f?.click(); return 1; })()`);
    await openInbox();
    const rowAt = () => js(`(() => { const r = [...document.querySelectorAll('.ml-row')].find((row) => row.querySelector('.ml-subject')?.textContent.includes(${JSON.stringify(subject)})); if (!r) return null; r.scrollIntoView({ block: 'center' }); const b = r.getBoundingClientRect(); return { x: Math.round(b.left + b.width / 2), y: Math.round(b.top + b.height / 2) }; })()`);
    // The harness's until answers true, not the value: the row is read again after it.
    // The list is read again until the new message is in it.
    await until(async () => (await rowAt()) || (await openInbox(), await wait(400), null), 'the message in the list', 12000).catch(() => null);
    const at = await rowAt();

    // A real double-click: two presses, the second counted as the second click.
    const before = BrowserWindow.getAllWindows();
    if (at) {
      const wc = win.webContents;
      wc.sendInputEvent({ type: 'mouseMove', x: at.x, y: at.y });
      for (const clickCount of [1, 2]) {
        wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount });
        wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount });
      }
    }
    const opened = () => BrowserWindow.getAllWindows().find((w) => !before.includes(w) && !w.isDestroyed()) || null;
    await until(opened, 'a window of its own', 10000).catch(() => null);
    alone = opened();
    const ajs = (code) => alone.webContents.executeJavaScript(code);
    if (alone) {
      await new Promise((resolve) => (alone.webContents.isLoading() ? alone.webContents.once('did-finish-load', resolve) : resolve()));
      await until(() => ajs(`(document.querySelector('.ml-head h2')?.textContent || '').includes(${JSON.stringify(subject)})`), 'the message in its window', 10000).catch(() => {});
      await wait(500);
    }
    const shown = alone
      ? await ajs(`({ heading: document.querySelector('.ml-head h2')?.textContent || '', list: Boolean(document.querySelector('.ml-list')), folders: Boolean(document.querySelector('.ml-compose-cta')), title: document.title, url: location.search })`)
      : null;
    const first = await js(`({ list: document.querySelectorAll('.ml-list .ml-row').length, pane: document.querySelector('.ml-head h2')?.textContent || '' })`);
    check('mail: a double-click on a message in the list opens it in a window of its own, with its subject in the title, the message read and no folders or list beside it, while the first window keeps its list',
      Boolean(at) && Boolean(shown) && shown.heading.includes(subject) && !shown.list && !shown.folders && (shown.title.includes(subject) || alone.getTitle().includes(subject)) && first.list > 0,
      `${at ? 'clicked' : 'no row'}; ${JSON.stringify(shown)}; window title ${alone?.getTitle?.() || ''}; first window ${JSON.stringify(first)}`);

    // The arrow keys in that window scroll the message; they do not open another.
    if (alone) {
      alone.webContents.focus();
      for (const keyCode of ['Down', 'Down', 'Up', 'Down']) {
        alone.webContents.sendInputEvent({ type: 'keyDown', keyCode });
        alone.webContents.sendInputEvent({ type: 'keyUp', keyCode });
      }
      await wait(800);
    }
    const still = alone ? await ajs(`document.querySelector('.ml-head h2')?.textContent || ''`) : '';
    check('mail: in a message\'s own window the arrow keys stay on that message', still.includes(subject), still);

    // Deleted from its own window: the message goes to Deleted, and the window closes.
    if (alone) {
      await ajs(`(() => { const b = [...document.querySelectorAll('.rw-btn, button')].find((n) => (n.dataset.tip || n.title || n.getAttribute('aria-label') || '').trim().startsWith('Delete') && !n.disabled && n.offsetParent !== null); if (!b) return 'no button'; b.click(); return 'clicked'; })()`);
    }
    const closed = alone ? await until(() => alone.isDestroyed(), 'the window to close', 8000).then(() => true).catch(() => false) : false;
    const left = await mail('messages', { accountId, folder: inbox, query: '', limit: 500 });
    const inInbox = (left.rows || left || []).some((r) => (r.subject || '').includes(subject));
    check('mail: deleted from its own window, the message leaves the inbox and the window closes', closed && !inInbox, `closed ${closed}; still in the inbox ${inInbox}`);

    const complaints = [...(await errorsIn(win))];
    check('mail window: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('mail window: the check ran to the end', false, err?.message || String(err));
  } finally {
    if (alone && !alone.isDestroyed()) alone.destroy();
  }
}
