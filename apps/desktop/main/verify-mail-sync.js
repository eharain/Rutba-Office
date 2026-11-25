// Mail and its server, in the window (A6.1, A6.2, A6.4): against the test
// IMAP server, run in this process on 127.0.0.1, never a real account.
//
// The account's mail is fetched as the Mail window opens, with nobody
// pressing Get mail. A message opened is read at the server; marked unread
// and starred from the list, it is so there too. Deleted from the list it is
// in the server's Trash; moved with Move to, and dragged onto a folder, it is
// in that folder there. A folder made from the Folder tab is made at the
// server, and deleted from it is deleted there. Mail that reaches the server
// shows in the list by itself, as the server says it came (IDLE), and a mark
// made at the server elsewhere shows here. A draft written and saved is in
// the server's Drafts, one copy however often it is saved. On a second
// server that answers as Gmail does, Archive puts the message in All Mail
// and makes no Archive label. Run alone with RUTBA_VERIFY_ONLY=mail-sync.

import { dialog } from 'electron';
import { labelKeyed } from './verify-keys.js';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 */
export async function verifyMailSync(h) {
  const { open, check, until, wait, errorsIn } = h;
  // The test server lives with the mailbox package's tests: loaded only for this check.
  const { imapServer } = await import('../../../packages/mailbox/test/imap-server.js');
  const stamp = Date.now();
  const server = await imapServer({ users: { 'sync@checks.example': 'check' } });
  const subjects = { del: `Sync: delete me ${stamp}`, move: `Sync: move me ${stamp}`, mark: `Sync: mark me ${stamp}`, drag: `Sync: drag me ${stamp}`, elsewhere: `Sync: read elsewhere ${stamp}` };
  const raw = (subject, n) => [
    'From: Amina Yusuf <amina@work.example>', 'To: Sync <sync@checks.example>', `Subject: ${subject}`,
    `Date: ${new Date(stamp - n * 60000).toUTCString()}`, `Message-ID: <sync-${n}-${stamp}@work.example>`, '', `About ${subject}.`, '',
  ].join('\r\n');
  Object.values(subjects).forEach((s, n) => server.deliver('INBOX', raw(s, n + 1)));
  const at = (folder) => server.messages(folder).map((m) => m.subject);
  const flagsOf = (subject) => server.messages('INBOX').find((m) => m.subject === subject)?.flags || [];
  const realAsk = dialog.showMessageBox;
  const asked = [];
  let win = null;
  let accountId = null;
  let gmail = null;
  let gmailId = null;
  try {
    win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.ml-accounts'))`), 'the Mail window', 8000);
    const account = { ...server.account(), name: 'Sync Check' };
    accountId = (await js(`window.rutbaOffice.mail.addAccount(${JSON.stringify({ account, password: 'check' })})`))?.id || null;

    // The window opened again: it fetches by itself.
    await new Promise((resolve) => { win.webContents.once('did-finish-load', resolve); win.webContents.reload(); });
    await until(() => js(`Boolean(document.querySelector('.ml-accounts [title="sync@checks.example"]'))`), 'the account in the sidebar', 10000);
    await js(`document.querySelector('.ml-accounts [title="sync@checks.example"]').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.ml-drop[data-folder="INBOX"] button'))`), 'its Inbox', 10000).catch(() => {});
    await js(`document.querySelector('.ml-drop[data-folder="INBOX"] button')?.click(), 1`);
    const rowOf = (s) => `[...document.querySelectorAll('.ml-row')].find((row) => row.querySelector('.ml-subject')?.textContent.includes(${JSON.stringify(s)}))`;
    const listed = await until(() => js(`Boolean(${Object.values(subjects).map(rowOf).join(' && ')})`), 'the messages fetched', 15000).then(() => true).catch(() => false);
    check('mail sync: an account\'s mail is fetched as the Mail window opens, with nobody pressing Get mail', listed, listed ? 'all five listed' : await js(`[...document.querySelectorAll('.ml-row .ml-subject')].map((n) => n.textContent).join(' | ') || 'nothing listed'`));

    const button = (label) => `((list) => list.find((n) => ${labelKeyed(label, { tip: true, start: true })}(n) && !n.disabled && n.offsetParent !== null) || list.find((n) => (n.dataset.tip || n.title || n.getAttribute('aria-label') || n.textContent || '').trim().startsWith(${JSON.stringify(label)}) && !n.disabled && n.offsetParent !== null))([...document.querySelectorAll('.rw-btn, button')])`;
    const select = async (s) => {
      await js(`(${rowOf(s)})?.click(), 1`);
      await until(() => js(`(document.querySelector('.ml-reader h2')?.textContent || '').includes(${JSON.stringify(s)})`), `${s} open`, 6000).catch(() => {});
    };
    const menuPick = (label) => js(`(() => { const b = ((list) => list.find((n) => ${labelKeyed(label, { start: true })}(n) && !n.disabled) || list.find((n) => n.textContent.trim().startsWith(${JSON.stringify(label)}) && !n.disabled))([...document.querySelectorAll('.rw-menu [role^="menuitem"]')]); if (!b) return false; b.click(); return true; })()`);

    // Opened, then marked unread and starred from the list.
    await select(subjects.mark);
    const read = await until(() => flagsOf(subjects.mark).includes('\\Seen'), 'read at the server', 8000).then(() => true).catch(() => false);
    await js(`(() => { const r = ${rowOf(subjects.mark)}; const b = r.getBoundingClientRect(); r.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, clientX: b.left + 20, clientY: b.top + 10 })); return 1; })()`);
    await wait(300);
    const picked = await menuPick('Mark unread');
    const unread = picked && await until(() => !flagsOf(subjects.mark).includes('\\Seen'), 'unread at the server', 8000).then(() => true).catch(() => false);
    await js(`(${rowOf(subjects.mark)})?.querySelector('.ml-star')?.click(), 1`);
    const starred = await until(() => flagsOf(subjects.mark).includes('\\Flagged'), 'starred at the server', 8000).then(() => true).catch(() => false);
    check('mail sync: a message opened is read at the server, and marked unread and starred from the list it is so there too', read && unread && starred, JSON.stringify({ read, picked, unread, starred, flags: flagsOf(subjects.mark) }));

    // Deleted from the list: in the server's Trash.
    await select(subjects.del);
    const opened = await js(`(document.querySelector('.ml-reader h2')?.textContent || '').includes(${JSON.stringify(subjects.del)})`);
    const clicked = await until(() => js(`(() => { const b = ${button('Delete')}; if (!b) return false; b.click(); return true; })()`), 'the Delete button', 6000).then(() => true).catch(() => false);
    const trashed = await until(() => !at('INBOX').includes(subjects.del) && at('Trash').includes(subjects.del), 'in the Trash at the server', 8000).then(() => true).catch(() => false);
    check('mail sync: a message deleted from the list is in the server\'s Trash', trashed, JSON.stringify({ opened, clicked, inbox: at('INBOX').includes(subjects.del), trash: at('Trash').includes(subjects.del) }));

    // Moved with Move to.
    await select(subjects.move);
    await js(`(${button('Move to')})?.click(), 1`);
    await wait(300);
    const chose = await menuPick('Archive');
    const moved = chose && await until(() => at('Archive').includes(subjects.move) && !at('INBOX').includes(subjects.move), 'in Archive at the server', 8000).then(() => true).catch(() => false);
    check('mail sync: a message moved with Move to is in that folder at the server', moved, JSON.stringify({ chose, archive: at('Archive') }));

    // A folder made from the Folder tab, then a message dragged onto it.
    const made = `Sync folder ${stamp}`;
    await js(`(() => { ((list) => list.find((t) => ${labelKeyed("Folder")}(t)) || list.find((t) => t.textContent.trim() === 'Folder'))([...document.querySelectorAll('.rw-tab')])?.click(); return 1; })()`);
    await until(() => js(`Boolean(${button('New folder')})`), 'New folder', 4000).catch(() => {});
    await js(`(${button('New folder')})?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('input.ml-folder-name'))`), 'the name box', 4000).catch(() => {});
    win.webContents.focus();
    await js(`document.querySelector('input.ml-folder-name')?.focus(), 1`);
    await win.webContents.insertText(made);
    await js(`document.querySelector('button.ml-folder-name')?.click(), 1`);
    const atServer = await until(() => server.folders().includes(made), 'the folder at the server', 8000).then(() => true).catch(() => false);
    const shown = await until(() => js(`Boolean(document.querySelector('.ml-drop[data-folder=${JSON.stringify(made)}]'))`), 'the folder in the list', 6000).then(() => true).catch(() => false);
    check('mail sync: a folder made from the Folder tab is made at the server and shown at once', atServer && shown, JSON.stringify({ atServer, shown }));

    // Dragged: the events a mouse drag fires, from the row to the folder.
    const dropped = await js(`(() => {
      const row = ${rowOf(subjects.drag)};
      const target = document.querySelector('.ml-drop[data-folder=${JSON.stringify(made)}]');
      if (!row || !target) return { row: Boolean(row), target: Boolean(target) };
      const dt = new DataTransfer();
      const fire = (node, type) => node.dispatchEvent(new DragEvent(type, { bubbles: true, cancelable: true, dataTransfer: dt }));
      fire(row, 'dragstart');
      const accepted = !fire(target, 'dragover');
      const lit = target.classList.contains('over');
      fire(target, 'drop');
      fire(row, 'dragend');
      return { row: true, target: true, accepted, lit };
    })()`);
    const draggedThere = await until(() => at(made).includes(subjects.drag) && !at('INBOX').includes(subjects.drag), 'dragged at the server', 8000).then(() => true).catch(() => false);
    const leftList = await until(() => js(`!${rowOf(subjects.drag)}`), 'off the list', 4000).then(() => true).catch(() => false);
    check('mail sync: a message dragged onto a folder is moved there at the server and leaves the list', Boolean(dropped.accepted) && draggedThere && leftList, JSON.stringify({ ...dropped, draggedThere, leftList }));

    // Mail reaching the server shows by itself; a mark made there elsewhere shows here.
    const arriving = `Sync: arrived by itself ${stamp}`;
    server.deliver('INBOX', raw(arriving, 9));
    const elsewhere = server.messages('INBOX').find((m) => m.subject === subjects.elsewhere);
    if (elsewhere) server.setFlags('INBOX', elsewhere.uid, ['\\Seen']);
    const came = await until(() => js(`Boolean(${rowOf(arriving)})`), 'the new message in the list', 15000).then(() => true).catch(() => false);
    const readHere = await until(() => js(`(() => { const r = ${rowOf(subjects.elsewhere)}; return Boolean(r && !r.classList.contains('unread')); })()`), 'read here', 8000).then(() => true).catch(() => false);
    check('mail sync: mail that reaches the server shows in the list by itself, and a message read elsewhere shows read here', came && readHere, JSON.stringify({ came, readHere, idling: server.idling() }));

    // The folder deleted from the Folder tab, the question answered yes.
    dialog.showMessageBox = async (...args) => {
      const opts = args.find((a) => a && typeof a === 'object' && 'buttons' in a) || {};
      asked.push(opts.message || '');
      return { response: 0 };
    };
    await js(`document.querySelector('.ml-drop[data-folder=${JSON.stringify(made)}] button')?.click(), 1`);
    await wait(400);
    await js(`(() => { ((list) => list.find((t) => ${labelKeyed("Folder")}(t)) || list.find((t) => t.textContent.trim() === 'Folder'))([...document.querySelectorAll('.rw-tab')])?.click(); return 1; })()`);
    await js(`(${button('Delete folder')})?.click(), 1`);
    const deleted = await until(() => !server.folders().includes(made), 'the folder gone at the server', 8000).then(() => true).catch(() => false);
    const goneHere = await until(() => js(`!document.querySelector('.ml-drop[data-folder=${JSON.stringify(made)}]')`), 'the folder gone here', 6000).then(() => true).catch(() => false);
    check('mail sync: a folder deleted from the Folder tab, after it asks, is deleted at the server and here', deleted && goneHere && asked.some((q) => q.includes(made)), JSON.stringify({ deleted, goneHere, asked }));

    // A draft written here and saved: in the server's Drafts, marked a draft there. Opened from Drafts and
    // saved again with other words, the server has one copy, the new one.
    const setSubject = (s) => js(`(() => { const n = [...document.querySelectorAll('.ml-compose .rw-field')].find((f) => f.querySelector('label')?.textContent === 'Subject')?.querySelector('input'); if (!n) return 0; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(s)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    const saveDraft = () => js(`(() => { const b = ((list) => list.find((n) => ${labelKeyed('Save draft')}(n)) || list.find((n) => n.textContent.trim() === 'Save draft'))([...document.querySelectorAll('.ml-compose .rw-btn, .ml-compose button')]); if (!b) return false; b.click(); return true; })()`);
    const draftsAt = () => server.messages('Drafts').map((m) => ({ subject: m.subject, draft: (m.flags || []).includes('\\Draft') }));
    const firstWords = `Sync: a draft ${stamp}`;
    const laterWords = `Sync: the draft written on ${stamp}`;
    await js(`document.querySelector('.ml-accounts [title="sync@checks.example"]')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.ml-drop[data-folder="INBOX"] button'))`), 'its Inbox again', 6000).catch(() => {});
    await js(`document.querySelector('.ml-drop[data-folder="INBOX"] button')?.click(), 1`);
    await js(`document.querySelector('.ml-compose-cta button')?.click(), 1`);
    const writing = await until(() => js(`Boolean(document.querySelector('.ml-compose .ml-form'))`), 'a new message', 6000).then(() => true).catch(() => false);
    const typed = Boolean(await setSubject(firstWords));
    await wait(300);
    const saved = await saveDraft();
    const reached = await until(() => draftsAt().some((d) => d.subject === firstWords && d.draft), 'the draft in the server\'s Drafts', 10000).then(() => true).catch(() => false);
    await until(() => js(`!document.querySelector('.ml-compose .ml-form')`), 'Compose closed', 6000).catch(() => {});
    await js(`document.querySelector('.ml-drop[data-folder="Drafts"] button')?.click(), 1`);
    const inList = await until(() => js(`Boolean(${rowOf(firstWords)})`), 'the draft in the list', 8000).then(() => true).catch(() => false);
    await js(`(${rowOf(firstWords)})?.click(), 1`);
    const openedAgain = await until(() => js(`Boolean(document.querySelector('.ml-compose .ml-form'))`), 'the draft in Compose', 8000).then(() => true).catch(() => false);
    const typedAgain = Boolean(await setSubject(laterWords));
    await wait(300);
    const savedAgain = await saveDraft();
    const oneCopy = await until(() => { const d = draftsAt(); return d.length === 1 && d[0].subject === laterWords && d[0].draft; }, 'one copy in the server\'s Drafts, the new one', 10000).then(() => true).catch(() => false);
    const oneRow = await until(() => js(`[...document.querySelectorAll('.ml-row .ml-subject')].filter((n) => n.textContent.includes(${JSON.stringify(String(stamp))}) && n.textContent.includes('draft')).length === 1`), 'one draft in the list', 8000).then(() => true).catch(() => false);
    check('mail sync: a draft written and saved here is in the server\'s Drafts, and saved again the server has one copy, the new one', reached && oneCopy && oneRow, JSON.stringify({ writing, typed, saved, reached, inList, openedAgain, typedAgain, savedAgain, oneCopy, oneRow, atServer: draftsAt() }));

    const complaints = [...(await errorsIn(win))];

    // Archive on a Gmail account: the message goes to All Mail, and no Archive label is made there.
    const ALL = '[Gmail]/All Mail';
    gmail = await imapServer({ users: { 'gsync@checks.example': 'check' }, capabilities: { gmail: true }, folders: [['INBOX', null], [ALL, '\\All'], ['[Gmail]/Sent Mail', '\\Sent'], ['[Gmail]/Drafts', '\\Drafts'], ['[Gmail]/Bin', '\\Trash'], ['[Gmail]/Spam', '\\Junk']] });
    const filed = `Sync: archive me ${stamp}`;
    gmail.deliver('INBOX', raw(filed, 11));
    const gAt = (folder) => gmail.messages(folder).map((m) => m.subject);
    gmailId = (await js(`window.rutbaOffice.mail.addAccount(${JSON.stringify({ account: { ...gmail.account(), name: 'Gmail Check' }, password: 'check' })})`))?.id || null;
    await new Promise((resolve) => { win.webContents.once('did-finish-load', resolve); win.webContents.reload(); });
    await until(() => js(`Boolean(document.querySelector('.ml-accounts [title="gsync@checks.example"]'))`), 'the Gmail account in the sidebar', 10000).catch(() => {});
    await js(`document.querySelector('.ml-accounts [title="gsync@checks.example"]')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.ml-drop[data-folder="INBOX"] button'))`), 'its Inbox', 10000).catch(() => {});
    await js(`document.querySelector('.ml-drop[data-folder="INBOX"] button')?.click(), 1`);
    const fetched = await until(() => js(`Boolean(${rowOf(filed)})`), 'the message fetched', 15000).then(() => true).catch(() => false);
    await select(filed);
    const allListed = await js(`Boolean(document.querySelector('.ml-drop[data-folder=${JSON.stringify(ALL)}]'))`);
    const pressed = await until(() => js(`(() => { const b = ${button('Archive')}; if (!b) return false; b.click(); return true; })()`), 'the Archive button', 6000).then(() => true).catch(() => false);
    const inAllMail = await until(() => gAt(ALL).includes(filed) && !gAt('INBOX').includes(filed), 'in All Mail at the server', 10000).then(() => true).catch(() => false);
    const offList = await until(() => js(`!${rowOf(filed)}`), 'off the list', 6000).then(() => true).catch(() => false);
    const noLabelMade = !gmail.folders().includes('Archive');
    check('mail sync: Archive on a Gmail account puts the message in All Mail at the server, and makes no Archive label there', inAllMail && offList && noLabelMade && allListed, JSON.stringify({ fetched, allListed, pressed, inAllMail, offList, noLabelMade, folders: gmail.folders() }));

    for (const c of await errorsIn(win)) if (!complaints.includes(c)) complaints.push(c);
    check('mail sync: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('mail sync: the check ran to the end', false, err?.message || String(err));
  } finally {
    dialog.showMessageBox = realAsk;
    if (accountId && win && !win.isDestroyed()) await win.webContents.executeJavaScript(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(accountId)}, keepMail: false })`).catch(() => {});
    if (gmailId && win && !win.isDestroyed()) await win.webContents.executeJavaScript(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(gmailId)}, keepMail: false })`).catch(() => {});
    await server.close().catch(() => {});
    if (gmail) await gmail.close().catch(() => {});
  }
}
