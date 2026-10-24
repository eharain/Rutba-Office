// Mail: a double-click on a message in the list opens it in a window of its
// own (owner, 2026-10-19: "the double click on mail in mail list should open
// a new window instead in the right pane").
//
// A message arrives and is double-clicked in the list, with a real mouse. A
// second Mail window opens on that message alone: its subject in the title,
// the message read, and no folders or list beside it. The arrow keys there
// stay on the message rather than moving to another. The first window still
// has its list. Deleted from its own window, the message goes to Deleted and
// the window closes. Deleted from the list, that row goes and the rest of the
// list stays as it was. Deleted from the middle of three, the one after it is
// selected and shown in the reading pane, as Outlook does, rather than the
// pane emptied and the list started again from the top. And the first message
// an account sends makes its Sent, which shows in the folders at once, not
// after the window is opened again. Import mail looks where the check points
// it, never at the computer's own mail, and names the files it found and the
// one being read without the folders they are in. Run alone with
// RUTBA_VERIFY_ONLY=mail-window.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
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

    // Deleted from the list (owner, 2026-10-23: "the email list entirely reloads when and email is
    // deleted rather than just single item removal"): that row goes, and the list stays as it was,
    // never emptied, no spinner over it, the other rows the same rows where they were.
    const keep = `Window: the minutes ${stamp}`;
    const drop = `Window: the agenda ${stamp}`;
    for (const [s, n] of [[keep, 'keep'], [drop, 'drop']]) {
      await mail('deliverTest', { accountId, folder: 'Inbox', raw: raw.replace(subject, s).replace(`<window-${stamp}@`, `<window-${n}-${stamp}@`) });
    }
    await openInbox();
    const rowOf = (s) => `[...document.querySelectorAll('.ml-row')].find((row) => row.querySelector('.ml-subject')?.textContent.includes(${JSON.stringify(s)}))`;
    await until(async () => (await js(`Boolean(${rowOf(keep)} && ${rowOf(drop)})`)) || (await openInbox(), await wait(400), false), 'both messages in the list', 12000).catch(() => {});
    const removed = await js(`(async () => {
      const kept = ${rowOf(keep)};
      const target = ${rowOf(drop)};
      if (!kept || !target) return { ready: false };
      kept.dataset.keptRow = '1';
      const top = kept.getBoundingClientRect().top;
      let emptied = false;
      let spun = false;
      const watch = new MutationObserver(() => {
        if (!document.querySelector('.ml-list')) emptied = true;
        // The list's spinner stands where the list was, a button's in the button.
        if ([...document.querySelectorAll('.rw-spinner')].some((n) => !n.closest('button'))) spun = true;
      });
      watch.observe(document.body, { childList: true, subtree: true });
      target.click();
      await new Promise((r) => setTimeout(r, 400));
      const del = [...document.querySelectorAll('.rw-btn, button')].find((n) => (n.dataset.tip || n.title || n.getAttribute('aria-label') || '').trim().startsWith('Delete') && !n.disabled && n.offsetParent !== null);
      if (!del) { watch.disconnect(); return { ready: true, button: false }; }
      del.click();
      for (let i = 0; i < 40 && ${rowOf(drop)}; i++) await new Promise((r) => setTimeout(r, 100));
      await new Promise((r) => setTimeout(r, 1500));
      watch.disconnect();
      const after = ${rowOf(keep)};
      return { ready: true, button: true, gone: !${rowOf(drop)}, sameRow: Boolean(after && after.dataset.keptRow === '1'), emptied, spun, moved: after ? Math.round(after.getBoundingClientRect().top - top) : null };
    })()`);
    check('mail: a message deleted from the list leaves it alone, the list never emptied nor a spinner shown over it, the other messages the same rows', removed.ready && removed.button && removed.gone && removed.sameRow && !removed.emptied && !removed.spun, JSON.stringify(removed));

    // Deleted from the middle of three (owner, 2026-10-24: "it should be
    // selecting the next rather than first"): the third is selected and read.
    const three = ['first', 'second', 'third'].map((n, i) => ({ n, s: `Window: the ${n} of three ${stamp}`, at: stamp + (3 - i) * 60000 }));
    for (const { n, s, at } of three) {
      await mail('deliverTest', { accountId, folder: 'Inbox', raw: raw.replace(subject, s).replace(`<window-${stamp}@`, `<window-three-${n}-${stamp}@`).replace(new Date(stamp).toUTCString(), new Date(at).toUTCString()) });
    }
    await openInbox();
    await until(async () => (await js(`Boolean(${three.map(({ s }) => rowOf(s)).join(' && ')})`)) || (await openInbox(), await wait(400), false), 'the three messages in the list', 12000).catch(() => {});
    const order = await js(`[...document.querySelectorAll('.ml-row .ml-subject')].map((n) => n.textContent).filter((s) => s.includes('of three')).map((s) => s.split(' ')[2])`);
    const next = await js(`(async () => {
      const middle = ${rowOf(three[1].s)};
      if (!middle) return { ready: false };
      middle.click();
      for (let i = 0; i < 40 && !(document.querySelector('.ml-reader h2')?.textContent || '').includes(${JSON.stringify(three[1].s)}); i++) await new Promise((r) => setTimeout(r, 100));
      const del = [...document.querySelectorAll('.rw-btn, button')].find((n) => (n.dataset.tip || n.title || n.getAttribute('aria-label') || '').trim().startsWith('Delete') && !n.disabled && n.offsetParent !== null);
      if (!del) return { ready: true, button: false };
      del.click();
      for (let i = 0; i < 40 && ${rowOf(three[1].s)}; i++) await new Promise((r) => setTimeout(r, 100));
      for (let i = 0; i < 40 && !(document.querySelector('.ml-reader h2')?.textContent || '').includes(${JSON.stringify(three[2].s)}); i++) await new Promise((r) => setTimeout(r, 100));
      return {
        ready: true,
        button: true,
        gone: !${rowOf(three[1].s)},
        selected: [...document.querySelectorAll('.ml-row.selected .ml-subject')].map((n) => n.textContent),
        shown: document.querySelector('.ml-reader h2')?.textContent || '',
      };
    })()`);
    check('mail: deleting the middle of three selects the third and shows it in the reading pane, not an empty pane',
      next.ready && next.button && next.gone && JSON.stringify(order) === '["first","second","third"]' && next.selected.length === 1 && next.selected[0].includes(three[2].s) && next.shown.includes(three[2].s),
      JSON.stringify({ order, ...next }));
    // The keys go on from there: Down moves to the message below the third, not the first.
    win.focus();
    win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Up' });
    win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Up' });
    const up = await until(() => js(`(document.querySelector('.ml-row.selected .ml-subject')?.textContent || '').includes(${JSON.stringify(three[0].s)})`), 'Up to the first of three', 3000).then(() => true).catch(() => false);
    check('mail: after the delete, Up goes on from the message selected, to the one above it', up, up ? 'the first of three' : await js(`document.querySelector('.ml-row.selected .ml-subject')?.textContent || 'nothing selected'`));

    // The first send makes Sent (A6.13): an account with no Sent yet, a
    // message sent from it, and Sent among its folders without a reload.
    const fresh = `sent.window.${stamp}@checks.example`;
    let freshId = null;
    try {
      await js(`window.rutbaOffice.mail.addAccounts({ text: ${JSON.stringify(`Email,Password,Name,IMAP server,IMAP port,SMTP server,SMTP port\r\n${fresh},pw,Sent Check,mail.checks.example,993,mail.checks.example,587\r\n`)} })`);
      freshId = (await js(`window.rutbaOffice.mail.accounts()`)).find((a) => a.email === fresh)?.id || null;
      // A message arrives, so the window lists the account and its Inbox.
      await mail('deliverTest', { accountId: freshId, folder: 'Inbox', raw: raw.replace(subject, `Window: hello ${stamp}`).replace(`<window-${stamp}@`, `<window-sent-${stamp}@`) });
      await until(() => js(`Boolean(document.querySelector('.ml-accounts [title=${JSON.stringify(fresh)}]'))`), 'the new account in the sidebar', 8000);
      await js(`document.querySelector('.ml-accounts [title=${JSON.stringify(fresh)}]').click(), 1`);
      await until(() => js(`Boolean(document.querySelector('[title="Inbox"]'))`), 'its Inbox', 6000);
      const sentBefore = await js(`Boolean(document.querySelector('[title="Sent"]'))`);
      const sentHad = (await mail('folders', { accountId: freshId })).some((f) => f.role === 'sent');
      await mail('queue', { accountId: freshId, draft: { to: 'amina@work.example', subject: `Window: the first sent ${stamp}`, text: 'Thank you.' }, holdSeconds: 0 });
      const sentNow = await until(() => js(`Boolean(document.querySelector('[title="Sent"]'))`), 'Sent among the folders', 8000).then(() => true).catch(() => false);
      check('mail: the first message an account sends makes Sent, which shows among its folders without a reload',
        !sentBefore && !sentHad && sentNow, JSON.stringify({ sentBefore, sentHad, sentNow }));
    } finally {
      if (freshId) await js(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(freshId)}, keepMail: false })`).catch(() => {});
    }

    // Import mail (A6.12): the scan pointed at a folder the check made, with
    // two Outlook data files named after an address, as Outlook names them.
    const was = process.env.RUTBA_MAIL_IMPORT_HOME;
    const own = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-import-home-'));
    try {
      const outlook = path.join(own, 'AppData', 'Local', 'Microsoft', 'Outlook');
      fs.mkdirSync(outlook, { recursive: true });
      for (const n of ['sam@example.com.pst', 'sam@example.com - old.pst']) fs.writeFileSync(path.join(outlook, n), Buffer.alloc(256, 7));
      process.env.RUTBA_MAIL_IMPORT_HOME = own;
      // Every line the window shows from here on, so a box shown for a moment is read too.
      await js(`(() => { window.__importSeen = new Set(); const note = () => { for (const n of document.querySelectorAll('.rw-dialog, [role="dialog"], .rw-toast')) window.__importSeen.add(n.textContent); }; window.__importWatch = new MutationObserver(note); window.__importWatch.observe(document.body, { childList: true, subtree: true, characterData: true }); return 1; })()`);
      await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Send / Receive')?.click(); return 1; })()`);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-btn')].find((b) => b.textContent.trim() === 'Import mail'))`), 'the Import mail button', 4000);
      await js(`[...document.querySelectorAll('.rw-btn')].find((b) => b.textContent.trim() === 'Import mail').click(), 1`);
      await until(() => js(`[...document.querySelectorAll('.ml-found-item')].some((n) => n.textContent.includes('sam@example.com - old.pst'))`), 'the files found', 8000).catch(() => {});
      const found = await js(`[...document.querySelectorAll('.ml-found-item')].map((n) => n.textContent)`);
      // The second file chosen: the Importing box names it while it is looked inside.
      await js(`(() => { [...document.querySelectorAll('.ml-found-item')].find((n) => n.textContent.includes('sam@example.com - old.pst'))?.click(); return 1; })()`);
      await wait(2500);
      const seen = await js(`(() => { window.__importWatch?.disconnect(); return [...(window.__importSeen || [])]; })()`);
      const named = (s) => s.includes('sam@example.com - old.pst');
      const leaks = [...found, ...seen].filter((s) => s.includes(own) || s.includes(outlook) || s.includes(os.homedir()));
      check('mail: Import mail looks only where it is pointed and names the files it found, and the one being read, without the folders they are in',
        found.some(named) && found.some((s) => s.includes('sam@example.com.pst')) && seen.some((s) => s.includes('Importing') && named(s)) && leaks.length === 0,
        JSON.stringify({ found: found.length, importing: seen.some((s) => s.includes('Importing') && named(s)), leaks: leaks.map((s) => s.slice(0, 120)) }));
      await js(`(() => { document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })); return 1; })()`);
      await wait(300);
    } finally {
      if (was === undefined) delete process.env.RUTBA_MAIL_IMPORT_HOME;
      else process.env.RUTBA_MAIL_IMPORT_HOME = was;
      try { fs.rmSync(own, { recursive: true, force: true, maxRetries: 10, retryDelay: 200 }); } catch { /* left in the temp folder */ }
    }

    const complaints = [...(await errorsIn(win))];
    check('mail window: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('mail window: the check ran to the end', false, err?.message || String(err));
  } finally {
    if (alone && !alone.isDestroyed()) alone.destroy();
  }
}
