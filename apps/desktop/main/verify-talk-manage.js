// Talk: contacts removed and blocked, invitations withdrawn and sent again
// (owner, 2026-10-22, D-2026-10-22-09: "on talk ability to reinvite after
// couple of days also remove invites and friends"), on the stand-in
// introducer: a contact's menu removes it (asked first) and blocks another;
// an invitation asked shows when it was sent, Invite again sends it again once
// the introducer allows it (the stand-in here at once; Rutba's after two days)
// by Rutba's email as the first was, and Withdraw takes it back. Run alone
// with RUTBA_VERIFY_ONLY=talk-manage.

import { BrowserWindow } from 'electron';
import { standInIntroducer } from './verify-talk.js';

const PAGE = (url) => `<!doctype html><meta charset="utf-8"><body><script>
const log = []; window.peer = { log };
const ws = new WebSocket(${JSON.stringify(url)});
ws.onopen = () => ws.send(JSON.stringify({ type: 'hello', token: 'ben' }));
ws.onmessage = (e) => log.push(JSON.parse(e.data));
window.peer.send = (m) => ws.send(JSON.stringify(m));
</script></body>`;

export async function verifyTalkManage({ open, check, until, wait, errorsIn }) {
  const ok = (p) => p.then(() => true, () => false);
  const intro = standInIntroducer({
    page: PAGE,
    people: { grace: { id: 'u-grace', name: 'Grace Okafor', email: 'grace@example.com' } },
    contacts: [['u-amira', 'u-ben'], ['u-amira', 'u-grace']],
    remindAfterMs: 0,
  });
  let ben = null;
  try {
    const port = await intro.listen();
    const win = await open('talk');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.tk-signin'))`), 'the sign-in', 8000);
    await js(`window.rutbaOffice.talk.signIn({ token: 'amira', server: ${JSON.stringify(`ws://127.0.0.1:${port}/ws`)} }).then(() => 1)`);
    await until(() => js(`document.querySelectorAll('.tk-contact').length === 2`), 'the two contacts', 8000);
    ben = new BrowserWindow({ show: false, webPreferences: { contextIsolation: true, nodeIntegration: false, partition: 'talk-manage-ben' } });
    await ben.loadURL(`http://127.0.0.1:${port}/ben`);
    const bjs = (code) => ben.webContents.executeJavaScript(code);
    await until(() => bjs(`window.peer.log.some((m) => m.type === 'welcome')`), 'Ben signed in', 6000);

    const menuItem = async (sel, label) => {
      await js(`document.querySelector(${JSON.stringify(sel)}).dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 120, clientY: 160 })), 1`);
      await until(() => js(`[...document.querySelectorAll('[role="menuitem"]')].some((n) => n.textContent.includes(${JSON.stringify(label)}))`), label, 4000);
      await js(`[...document.querySelectorAll('[role="menuitem"]')].find((n) => n.textContent.includes(${JSON.stringify(label)})).click(), 1`);
    };

    // Ben removed from the contacts, asked first; Ben is told.
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'message', answer: { response: 0 } });
    await menuItem('.tk-contact[data-contact="u-ben"]', 'Remove from contacts');
    const removed = await ok(until(() => js(`!document.querySelector('.tk-contact[data-contact="u-ben"]')`), 'Ben gone', 5000));
    const benTold = await ok(until(() => bjs(`window.peer.log.some((m) => m.type === 'contact.removed' && m.id === 'u-amira')`), 'Ben told', 5000));
    // Grace blocked.
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'message', answer: { response: 0 } });
    await menuItem('.tk-contact[data-contact="u-grace"]', 'Block');
    const blocked = await ok(until(() => js(`!document.querySelector('.tk-contact[data-contact="u-grace"]')`), 'Grace gone', 5000));
    check('talk: a contact\'s menu removes it, asked first and the other side told, and blocks another',
      removed && benTold && blocked, JSON.stringify({ removed, benTold, blocked }));

    // An invitation asked by Rutba's email: when it was sent, Invite again, Withdraw.
    const asked = await js(`window.rutbaOffice.talk.invite({ email: 'dana@example.com', how: 'rutba' })`);
    await until(() => js(`Boolean(document.querySelector('[data-asked="dana@example.com"] .tk-asked-again:not([disabled])'))`), 'Invite again ready', 5000).catch(() => {});
    const row = await js(`({ words: document.querySelector('[data-asked="dana@example.com"]')?.textContent || '', again: !document.querySelector('[data-asked="dana@example.com"] .tk-asked-again')?.disabled })`);
    const mailedBefore = intro.mailed().length;
    await js(`document.querySelector('[data-asked="dana@example.com"] .tk-asked-again').click(), 1`);
    const again = await ok(until(() => Promise.resolve(intro.mailed().length === mailedBefore + 1), 'sent again', 6000));
    await until(() => js(`/sent/.test(document.querySelector('[data-asked="dana@example.com"]')?.textContent || '')`), 'the row', 3000).catch(() => {});
    await js(`document.querySelector('[data-asked="dana@example.com"] .tk-asked-withdraw').click(), 1`);
    const withdrawn = await ok(until(() => js(`!document.querySelector('[data-asked="dana@example.com"]')`), 'withdrawn', 5000));
    check('talk: an invitation asked shows when it was sent; Invite again sends it again by Rutba\'s email as the first was, once allowed; Withdraw takes it back',
      Boolean(asked.request) && /sent/.test(row.words) && /Emailed by Rutba/.test(row.words) && row.again && again && intro.mailed().at(-1).again === true && withdrawn,
      JSON.stringify({ asked: Boolean(asked.request), row, again, withdrawn }));

    const complaints = await errorsIn(win);
    check('talk: removing, blocking, withdrawing and inviting again reported no error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    await js(`window.rutbaOffice.talk.signOut().then(() => 1)`).catch(() => {});
  } catch (err) {
    check('talk: the contact and invitation checks ran to the end', false, err?.message || String(err));
  } finally {
    try { ben?.destroy(); } catch { /* gone */ }
    await intro.close().catch(() => {});
  }
}
