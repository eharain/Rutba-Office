// Mail: messages as .eml files (Mail 2.25). Save as .eml on the row menu
// writes one message to the file chosen, and Ctrl+Shift+S on several writes
// each into the folder chosen; File, Open (Ctrl+O) brings .eml files into the
// folder open with their dates and read state; files dropped on a folder in
// the tree, or pasted with a folder holding the keys, come into that folder;
// and a .eml opened from Explorer shows alone, read only, with Add to a
// folder. The file dialogs are answered as a person would (ipc.js).
// Run alone with RUTBA_VERIFY_ONLY=mail-eml.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

export async function verifyMailEml({ open, check, until, wait, errorsIn }) {
  const stamp = Date.now();
  const tag = `Eml ${stamp}`;
  const subject = (w) => `${tag}: ${w}`;
  const raw = (w, n, extra = []) => [
    'From: Amina Yusuf <amina@work.example>', 'To: You <you@example.com>', `Subject: ${subject(w)}`,
    `Date: ${new Date(Date.UTC(2025, 4, 6, 10, n)).toUTCString()}`,
    `Message-ID: <eml-${w}-${stamp}@example.org>`, ...extra, 'MIME-Version: 1.0', 'Content-Type: text/plain; charset=utf-8', '',
    `The ${w} message.`, '',
  ].join('\r\n');
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-verify-eml-'));
  const answer = (kind, value) => globalThis.__rutbaCheckDialogAnswers?.push({ kind, answer: value });
  let win = null;
  let fileWin = null;
  let accountId = null;
  let made = null;
  try {
    win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the seeded account in the sidebar', 8000);
    accountId = await js(`(async () => (await window.rutbaOffice.mail.accounts())[0]?.id || null)()`);
    const mail = (method, args) => js(`window.rutbaOffice.mail.${method}(${JSON.stringify(args)})`);
    // A day ahead, so they are the top of the list.
    const ahead = (w, n) => raw(w, n).replace(/^Date: .*$/m, `Date: ${new Date(stamp + 86_400_000 + n * 60_000).toUTCString()}`);
    await mail('deliverTest', { accountId, folder: 'Inbox', raw: ahead('alpha', 1) });
    await mail('deliverTest', { accountId, folder: 'Inbox', raw: ahead('bravo', 2) });
    made = (await mail('createFolder', { accountId, name: `Files ${stamp}` })).path;
    win.webContents.reload();
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the window again', 10000);
    const inbox = (await mail('folders', { accountId })).find((f) => /^inbox$/i.test(f.name || f.path))?.path || 'Inbox';
    const openFolder = (p) => js(`(() => { const f = document.querySelector('.ml-drop[data-folder="${p}"] button'); f?.click(); return Boolean(f); })()`);
    const rowOf = (s) => `[...document.querySelectorAll('.ml-list > .ml-row')].find((row) => row.querySelector('.ml-subject')?.textContent.includes(${JSON.stringify(s)}))`;
    const rowsIn = async (folder) => (await mail('messages', { accountId, folder })).rows.filter((r) => String(r.subject).startsWith(tag));
    const key = (keyCode, modifiers) => {
      win.webContents.sendInputEvent({ type: 'keyDown', keyCode, modifiers });
      win.webContents.sendInputEvent({ type: 'keyUp', keyCode, modifiers });
    };
    const menuPick = async (w, label) => {
      const run = (code) => w.webContents.executeJavaScript(code);
      await until(() => run(`[...document.querySelectorAll('[role="menuitem"]')].some((n) => n.textContent.includes(${JSON.stringify(label)}))`), `the menu's ${label}`, 5000);
      return run(`(() => { const n = [...document.querySelectorAll('[role="menuitem"]')].reverse().find((x) => x.textContent.includes(${JSON.stringify(label)})); if (!n || n.disabled) return false; n.click(); return true; })()`);
    };

    await openFolder(inbox);
    await until(async () => (await js(`Boolean(${rowOf(subject('alpha'))} && ${rowOf(subject('bravo'))})`)) || (await openFolder(inbox), await wait(400), false), 'the two messages in the list', 12000).catch(() => {});

    // Save as .eml on the row menu: the one message, to the file chosen.
    const one = path.join(dir, 'chosen.eml');
    answer('save', one);
    await js(`${rowOf(subject('alpha'))}?.click(), 1`);
    await wait(300);
    await js(`(() => { const r = ${rowOf(subject('alpha'))}; const b = r.getBoundingClientRect(); r.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: b.left + 30, clientY: b.top + 8 })); return 1; })()`);
    const savePicked = await menuPick(win, 'Save as .eml');
    await until(() => fs.existsSync(one), 'the saved file', 6000).catch(() => {});
    const saved = fs.existsSync(one) ? fs.readFileSync(one, 'latin1') : '';
    check('mail: Save as .eml on the row menu writes the message to the file chosen',
      savePicked && saved.includes(`Subject: ${subject('alpha')}`) && saved.includes('The alpha message.'), JSON.stringify({ savePicked, size: saved.length }));

    // Ctrl+Shift+S on two: each its own file in the folder chosen, named by its subject.
    const several = path.join(dir, 'several');
    fs.mkdirSync(several);
    answer('open', [several]);
    // Had only one been chosen, the save dialog is answered too, and never left open on the screen.
    answer('save', path.join(dir, 'only-one.eml'));
    for (const w of ['alpha', 'bravo']) {
      await js(`(() => { const r = ${rowOf(subject(w))}; r?.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, ctrlKey: true })); return 1; })()`);
      await wait(200);
    }
    await js(`${rowOf(subject('bravo'))}?.focus(), 1`);
    key('S', ['control', 'shift']);
    await until(() => fs.readdirSync(several).length >= 2, 'the two files', 6000).catch(() => {});
    const queue = globalThis.__rutbaCheckDialogAnswers;
    if (Array.isArray(queue)) for (let i = queue.length - 1; i >= 0; i--) if (queue[i].kind === 'open' || queue[i].kind === 'save') queue.splice(i, 1);
    const names = fs.readdirSync(several).sort();
    check('mail: Save as .eml with two chosen writes each to a file of its own, named by its subject',
      names.length === 2 && names.every((n) => n.startsWith(`Eml ${stamp}`) && n.endsWith('.eml')), JSON.stringify(names));

    // File, Open (Ctrl+O): .eml files come into the folder open, their dates and read state kept.
    const read = path.join(dir, 'read.eml');
    const unread = path.join(dir, 'unread.eml');
    fs.writeFileSync(read, raw('charlie', 3));
    fs.writeFileSync(unread, raw('delta', 4, ['Status: O']));
    answer('open', [read, unread]);
    await js(`document.querySelector('.ml-list')?.focus(), document.activeElement?.blur?.(), 1`);
    key('O', ['control']);
    await until(async () => (await rowsIn(inbox)).some((r) => r.subject === subject('delta')), 'the opened files in the Inbox', 6000).catch(() => {});
    const byName = Object.fromEntries((await rowsIn(inbox)).map((r) => [r.subject, r]));
    const c = byName[subject('charlie')];
    const d = byName[subject('delta')];
    check('mail: File, Open brings .eml files into the folder open, each with the date it was written and whether it was read',
      Boolean(c && d) && c.unread === false && d.unread === true && new Date(c.date).toISOString() === '2025-05-06T10:03:00.000Z',
      JSON.stringify({ charlie: c && { unread: c.unread, date: c.date }, delta: d && { unread: d.unread } }));

    // Dropped on a folder in the tree: into that folder. The file here has no path, as a file from another program may not.
    const dropped = raw('echo', 5);
    const dropOk = await js(`(() => {
      const target = document.querySelector('.ml-drop[data-folder="${made}"]');
      if (!target) return false;
      const dt = new DataTransfer();
      dt.items.add(new File([${JSON.stringify(dropped)}], 'echo.eml', { type: 'message/rfc822' }));
      const over = new DragEvent('dragover', { bubbles: true, cancelable: true, dataTransfer: dt });
      target.dispatchEvent(over);
      target.dispatchEvent(new DragEvent('drop', { bubbles: true, cancelable: true, dataTransfer: dt }));
      return over.defaultPrevented;
    })()`);
    await until(async () => (await rowsIn(made)).some((r) => r.subject === subject('echo')), 'the dropped file in its folder', 6000).catch(() => {});
    check('mail: a .eml file dropped on a folder in the tree comes into that folder',
      dropOk && (await rowsIn(made)).some((r) => r.subject === subject('echo')), JSON.stringify({ dropOk, there: (await rowsIn(made)).map((r) => r.subject) }));

    // Pasted, as files copied in Explorer are, with that folder holding the keys.
    await js(`document.querySelector('.ml-drop[data-folder="${made}"] button')?.focus(), 1`);
    const pasteOk = await js(`(() => {
      const dt = new DataTransfer();
      dt.items.add(new File([${JSON.stringify(raw('foxtrot', 6))}], 'foxtrot.eml', { type: 'message/rfc822' }));
      const e = new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: dt });
      document.activeElement.dispatchEvent(e);
      return e.defaultPrevented;
    })()`);
    await until(async () => (await rowsIn(made)).some((r) => r.subject === subject('foxtrot')), 'the pasted file in the folder', 6000).catch(() => {});
    check('mail: .eml files pasted with a folder holding the keys come into that folder',
      pasteOk && (await rowsIn(made)).some((r) => r.subject === subject('foxtrot')), JSON.stringify({ pasteOk, there: (await rowsIn(made)).map((r) => r.subject) }));

    // Opened from Explorer: that message alone, read only, with Add to a folder.
    const outside = path.join(dir, 'from explorer.eml');
    fs.writeFileSync(outside, raw('golf', 7, ['Status: O']));
    fileWin = await open('mail', outside);
    const fjs = (code) => fileWin.webContents.executeJavaScript(code);
    await until(() => fjs(`Boolean(document.querySelector('.ml-opened-bar') && document.querySelector('.ml-reader'))`), 'the opened file, read', 12000).catch(() => {});
    const shown = await fjs(`(() => ({
      bar: document.querySelector('.ml-opened-bar')?.textContent || '',
      subject: document.querySelector('.ml-reader h2')?.textContent || '',
      deleteOff: [...document.querySelectorAll('.rw-ribbon button, button')].find((b) => b.textContent.trim() === 'Delete')?.disabled ?? null,
    }))()`);
    const notInInbox = !(await rowsIn(inbox)).some((r) => r.subject === subject('golf'));
    check('mail: a .eml opened from Explorer shows alone and read only, in no folder, and says where it came from',
      shown.bar.includes('from explorer.eml') && shown.subject.includes(subject('golf')) && shown.deleteOff !== false && notInInbox, JSON.stringify({ ...shown, notInInbox }));

    await wait(800); // the accounts' folders, for Add to a folder
    await fjs(`document.querySelector('.ml-opened-add')?.click(), 1`);
    const accountName = await fjs(`(async () => { const a = (await window.rutbaOffice.mail.accounts())[0]; return a?.name || a?.email || ''; })()`);
    const accountPicked = await menuPick(fileWin, accountName);
    await wait(300);
    const folderName = (await mail('folders', { accountId })).find((f) => f.path === inbox)?.name || inbox;
    const folderPicked = await fjs(`(() => { const n = [...document.querySelectorAll('.rw-submenu [role="menuitem"]')].find((x) => x.textContent.trim() === ${JSON.stringify(folderName)}); n?.click(); return Boolean(n); })()`);
    await until(async () => (await rowsIn(inbox)).some((r) => r.subject === subject('golf')), 'the opened file added to the Inbox', 6000).catch(() => {});
    const added = (await rowsIn(inbox)).find((r) => r.subject === subject('golf'));
    check('mail: Add to a folder brings the opened file into the folder chosen, unread as it was',
      accountPicked && folderPicked && Boolean(added) && added.unread === true, JSON.stringify({ accountPicked, folderPicked, added: added && { unread: added.unread } }));

    const complaints = [...(await errorsIn(win)), ...(await errorsIn(fileWin))];
    check('mail: messages as .eml files reported no error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('mail: the .eml checks ran to the end', false, err?.message || String(err));
  } finally {
    try {
      if (fileWin && !fileWin.isDestroyed()) fileWin.close();
    } catch { /* gone */ }
    try {
      const run = (code) => win?.webContents.executeJavaScript(code);
      if (accountId) {
        if (made) await run(`window.rutbaOffice.mail.deleteFolder({ accountId: ${JSON.stringify(accountId)}, folder: ${JSON.stringify(made)} })`).catch(() => {});
        for (const f of await run(`window.rutbaOffice.mail.folders({ accountId: ${JSON.stringify(accountId)} })`)) {
          const rows = (await run(`window.rutbaOffice.mail.messages({ accountId: ${JSON.stringify(accountId)}, folder: ${JSON.stringify(f.path)} })`)).rows.filter((r) => String(r.subject).startsWith(tag));
          if (rows.length) await run(`window.rutbaOffice.mail.delete({ accountId: ${JSON.stringify(accountId)}, folder: ${JSON.stringify(f.path)}, ids: ${JSON.stringify(rows.map((r) => r.id))} })`);
        }
      }
    } catch { /* the window is gone */ }
    fs.rmSync(dir, { recursive: true, force: true });
  }
}
