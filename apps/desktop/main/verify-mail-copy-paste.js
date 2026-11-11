// Mail: copy, cut and paste of messages between folders (Mail 2.25). Ctrl+C
// on a message in the list, then Ctrl+V on a folder in the tree, copies it
// there and leaves it where it was; Cut and Paste here on the right-click
// menus move one; a message copied in one account is pasted into a folder of
// another. Ctrl+C while typing in the search box copies words, not messages.
// Run alone with RUTBA_VERIFY_ONLY=mail-copy-paste.

export async function verifyMailCopyPaste({ open, check, until, wait, errorsIn }) {
  const stamp = Date.now();
  const subjects = { one: `Paste: one ${stamp}`, two: `Paste: two ${stamp}` };
  const raw = (subject, n) => [
    'From: Amina Yusuf <amina@work.example>', 'To: You <you@example.com>', `Subject: ${subject}`,
    `Date: ${new Date(stamp + 86_400_000 + n * 60_000).toUTCString()}`,
    `Message-ID: <paste-${n}-${stamp}@example.org>`, 'MIME-Version: 1.0', 'Content-Type: text/plain; charset=utf-8', '',
    `Message ${n}.`, '',
  ].join('\r\n');
  let win = null;
  let accountId = null;
  let other = null;
  let made = null;
  try {
    win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the seeded account in the sidebar', 8000);
    accountId = await js(`(async () => (await window.rutbaOffice.mail.accounts())[0]?.id || null)()`);
    const mail = (method, args) => js(`window.rutbaOffice.mail.${method}(${JSON.stringify(args)})`);
    await mail('deliverTest', { accountId, folder: 'Inbox', raw: raw(subjects.one, 1) });
    await mail('deliverTest', { accountId, folder: 'Inbox', raw: raw(subjects.two, 2) });
    made = (await mail('createFolder', { accountId, name: `Pasted ${stamp}` })).path;
    win.webContents.reload();
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the window again', 10000);
    const inbox = (await mail('folders', { accountId })).find((f) => /^inbox$/i.test(f.name || f.path))?.path || 'Inbox';
    const openFolder = (p) => js(`(() => { const f = document.querySelector('.ml-drop[data-folder="${p}"] .rw-item, .ml-drop[data-folder="${p}"] button') || [...document.querySelectorAll('[title="${p}"]')][0]; f?.click(); return Boolean(f); })()`);
    const rowOf = (s) => `[...document.querySelectorAll('.ml-list > .ml-row')].find((row) => row.querySelector('.ml-subject')?.textContent.includes(${JSON.stringify(s)}))`;
    const subjectsIn = async (folder, acc = accountId) => (await mail('messages', { accountId: acc, folder })).rows.map((r) => r.subject);
    const key = (keyCode, modifiers = ['control']) => {
      win.webContents.sendInputEvent({ type: 'keyDown', keyCode, modifiers });
      win.webContents.sendInputEvent({ type: 'keyUp', keyCode, modifiers });
    };
    const toastSaid = (re) => js(`[...document.querySelectorAll('.rw-toast')].map((n) => n.textContent).find((x) => ${re}.test(x)) || ''`);
    const menuPick = async (label) => {
      await until(() => js(`[...document.querySelectorAll('[role="menuitem"]')].some((n) => n.textContent.includes(${JSON.stringify(label)}))`), `the menu's ${label}`, 4000);
      return js(`(() => { const n = [...document.querySelectorAll('[role="menuitem"]')].find((x) => x.textContent.includes(${JSON.stringify(label)})); if (!n || n.getAttribute('aria-disabled') === 'true') return false; n.click(); return true; })()`);
    };
    const rightClick = (selector) => js(`(() => { const n = ${selector}; if (!n) return false; const b = n.getBoundingClientRect(); n.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: b.left + 20, clientY: b.top + 8 })); return true; })()`);

    await openFolder(inbox);
    await until(async () => (await js(`Boolean(${rowOf(subjects.one)} && ${rowOf(subjects.two)})`)) || (await openFolder(inbox), await wait(400), false), 'the two messages in the list', 12000).catch(() => {});

    // Ctrl+C on the first, Ctrl+V on the new folder in the tree.
    await js(`${rowOf(subjects.one)}?.click(), 1`);
    await wait(300);
    await js(`${rowOf(subjects.one)}?.focus(), 1`);
    key('C');
    await until(() => toastSaid('/copied/'), 'the copied note', 4000).catch(() => {});
    const copiedNote = await toastSaid('/copied/');
    await js(`document.querySelector('.ml-drop[data-folder="${made}"] button, .ml-drop[data-folder="${made}"] [tabindex]')?.focus(), 1`);
    key('V');
    await until(async () => (await subjectsIn(made)).includes(subjects.one), 'the copy in the new folder', 6000).catch(() => {});
    const afterCopy = { there: await subjectsIn(made), inbox: await subjectsIn(inbox) };
    check('mail: Ctrl+C on a message, then Ctrl+V on a folder in the tree, copies it there and leaves it where it was',
      /1 message copied/.test(copiedNote) && afterCopy.there.includes(subjects.one) && afterCopy.inbox.includes(subjects.one),
      JSON.stringify({ copiedNote, ...afterCopy }));

    // Cut on the list's menu, Paste here on the folder's.
    await openFolder(inbox);
    await until(() => js(`Boolean(${rowOf(subjects.two)})`), 'the second message', 6000).catch(() => {});
    await js(`${rowOf(subjects.two)}?.click(), 1`);
    await wait(300);
    await rightClick(rowOf(subjects.two));
    const cutPicked = await menuPick('Cut');
    await wait(300);
    await rightClick(`document.querySelector('.ml-drop[data-folder="${made}"] button, .ml-drop[data-folder="${made}"] .rw-item')`);
    const pastePicked = await menuPick('Paste here');
    await until(async () => !(await subjectsIn(inbox)).includes(subjects.two), 'the cut message gone from the Inbox', 6000).catch(() => {});
    const afterCut = { there: await subjectsIn(made), inbox: await subjectsIn(inbox) };
    check('mail: Cut on the list\'s menu, then Paste here on a folder\'s, moves the message there',
      cutPicked && pastePicked && afterCut.there.includes(subjects.two) && !afterCut.inbox.includes(subjects.two),
      JSON.stringify({ cutPicked, pastePicked, ...afterCut }));

    // While typing in the search box, Ctrl+C is the words' copy: no message is copied.
    await js(`(() => { const s = document.querySelector('input[type="search"], .rw-search input'); s?.focus(); return Boolean(s); })()`);
    const notesBefore = await js(`document.querySelectorAll('.rw-toast').length`);
    await wait(4500); // the earlier notes go
    key('C');
    await wait(600);
    const typingNote = await toastSaid('/copied/');
    check('mail: Ctrl+C while typing in the search box does not copy messages', !typingNote, JSON.stringify({ typingNote, notesBefore }));

    // Copied in one account, pasted into a folder of another.
    other = (await mail('addAccount', { account: { email: `paste-${stamp}@checks.example`, name: 'Paste Check', local: true, imap: { host: '', port: 993 }, smtp: { host: '', port: 465 } } })).id;
    await mail('createFolder', { accountId: other, name: 'Kept' });
    win.webContents.reload();
    await until(() => js(`[...document.querySelectorAll('.ml-accounts button')].some((b) => /Paste Check/.test(b.textContent))`), 'the second account', 10000);
    await js(`[...document.querySelectorAll('.ml-accounts button')].find((b) => !/Paste Check|All accounts/.test(b.textContent))?.click(), 1`);
    await wait(400);
    await openFolder(made);
    await until(() => js(`Boolean(${rowOf(subjects.one)})`), 'the copied message in its folder', 6000).catch(() => {});
    await js(`${rowOf(subjects.one)}?.click(), 1`);
    await wait(300);
    await js(`${rowOf(subjects.one)}?.focus(), 1`);
    key('C');
    await wait(400);
    await js(`[...document.querySelectorAll('.ml-accounts button')].find((b) => /Paste Check/.test(b.textContent))?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.ml-drop[data-folder="Kept"]'))`), "the other account's folder", 6000).catch(() => {});
    await rightClick(`document.querySelector('.ml-drop[data-folder="Kept"] button, .ml-drop[data-folder="Kept"] .rw-item')`);
    const acrossPicked = await menuPick('Paste here');
    await until(async () => (await subjectsIn('Kept', other)).includes(subjects.one), 'the message in the other account', 6000).catch(() => {});
    const across = { there: await subjectsIn('Kept', other), still: await subjectsIn(made) };
    check('mail: a message copied in one account is pasted into a folder of another, and stays where it was',
      acrossPicked && across.there.includes(subjects.one) && across.still.includes(subjects.one), JSON.stringify({ acrossPicked, ...across }));

    const complaints = await errorsIn(win);
    check('mail: copy and paste between folders reported no error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('mail: the copy and paste checks ran to the end', false, err?.message || String(err));
  } finally {
    const run = (code) => win?.webContents.executeJavaScript(code).catch(() => null);
    if (other) await run(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(other)} })`);
    if (accountId && made) await run(`window.rutbaOffice.mail.deleteFolder({ accountId: ${JSON.stringify(accountId)}, folder: ${JSON.stringify(made)} })`);
    if (accountId) {
      const rows = (await run(`window.rutbaOffice.mail.messages({ accountId: ${JSON.stringify(accountId)}, folder: 'Inbox' })`))?.rows || [];
      const ids = rows.filter((r) => Object.values(subjects).includes(r.subject)).map((r) => r.id);
      if (ids.length) await run(`window.rutbaOffice.mail.delete({ accountId: ${JSON.stringify(accountId)}, folder: 'Inbox', ids: ${JSON.stringify(ids)} })`);
    }
  }
}
