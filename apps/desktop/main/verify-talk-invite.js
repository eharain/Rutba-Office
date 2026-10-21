// Talk invitations (owner, 2026-10-21, D-2026-10-21-27 and -28), checked on
// the stand-in introducer: asking an address emails it from the person's own
// Mail account, the copy in Sent; Rutba's email is the other way, until its
// limit, and the dialog then says so; each request says how it went; a link
// Ben made opens Talk on his invitation and Accept makes him a contact; a
// link Amira makes is copied, sent by a chat app or read off a phone code,
// and withdrawn. Run alone with RUTBA_VERIFY_ONLY=talk-invite.

import { BrowserWindow } from 'electron';
import { standInIntroducer } from './verify-talk.js';

const BEN_PAGE = (url) => `<!doctype html><meta charset="utf-8"><body><script>
const log = []; window.ben = { log };
const ws = new WebSocket(${JSON.stringify(url)});
ws.onopen = () => ws.send(JSON.stringify({ type: 'hello', token: 'ben' }));
ws.onmessage = (e) => log.push(JSON.parse(e.data));
window.ben.send = (m) => ws.send(JSON.stringify(m));
</script></body>`;

export async function verifyTalkInvite({ open, check, until, wait, errorsIn }) {
  const ok = (p) => p.then(() => true).catch(() => false);
  // One email a day from Rutba, so its limit is reached within the check.
  const intro = standInIntroducer({ page: BEN_PAGE, mailPerDay: 1 });
  let ben = null;
  try {
    const port = await intro.listen();
    const win = await open('talk');
    const js = (code) => win.webContents.executeJavaScript(code);
    const set = (sel, value) => js(`(() => { const n = document.querySelector(${JSON.stringify(sel)}); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(value)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.tk-signin'))`), 'the sign-in', 8000);
    await js(`window.rutbaOffice.talk.signIn({ token: 'amira', server: ${JSON.stringify(`ws://127.0.0.1:${port}/ws`)} }).then(() => 1)`);
    await until(() => js(`Boolean(document.querySelector('.tk-add'))`), 'signed in', 8000);
    // An account of her own to send from (a check run's mail goes nowhere, its copy into Sent).
    const account = await js(`window.rutbaOffice.mail.addAccount({ account: { email: 'amira@checks.example', name: 'Amira Khan', imap: { host: 'mail.checks.example', port: 993, secure: true }, smtp: { host: 'mail.checks.example', port: 587 } }, password: 'check' })`);

    // From her own address: the dialog offers it first, and the email is in Sent with the invitation's link.
    await js(`document.querySelector('.tk-add').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.tk-how [data-how^="own:"] input:checked'))`), 'her own address chosen', 8000);
    const offered = await js(`[...document.querySelectorAll('.tk-how [data-how]')].map((n) => n.dataset.how.split(':')[0])`);
    await set('.tk-email', 'chen@example.com');
    await set('.tk-note', 'From the supplier day');
    await js(`document.querySelector('.tk-ask').click(), 1`);
    const ownTold = await ok(until(() => js(`document.querySelector('[data-asked="chen@example.com"]')?.dataset.told === 'own:sent'`), 'told from her own address', 8000));
    const ownWords = await js(`document.querySelector('[data-asked="chen@example.com"]')?.textContent || ''`);
    const sent = await js(`(async () => {
      for (const a of [${JSON.stringify(account)}]) {
        const { rows } = await window.rutbaOffice.mail.messages({ accountId: a.id, folder: 'Sent', limit: 20 });
        const row = rows.find((r) => /Rutba Talk/.test(r.subject || '') && JSON.stringify(r.to || '').includes('chen@example.com'));
        if (!row) continue;
        const m = await window.rutbaOffice.mail.message({ accountId: a.id, folder: 'Sent', id: row.id ?? row.uid });
        return { subject: row.subject, link: /\\/i\\/[\\w-]{16,}/.test(m?.text || ''), join: /office\\.rutba\\.io\\/download\\?invite=/.test(m?.text || '') };
      }
      return null;
    })()`);
    check('talk: asking an address offers the person\'s own Mail address first, emails the invitation from it with the link to accept and the way to join, keeps the copy in Sent, and the request says it was emailed from that address',
      offered[0] === 'own' && !ownWords.includes('.mbox') && offered.includes('rutba') && offered.includes('none') && ownTold && /chen@example.com/.test(ownWords) && /Emailed from/.test(ownWords) && sent?.link && sent?.join,
      JSON.stringify({ offered, ownTold, ownWords, sent }));

    // From Rutba, once: then its limit for the day is reached, and the dialog says so.
    await js(`document.querySelector('.tk-add').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.tk-how [data-how="rutba"] input'))`), 'the dialog again', 8000);
    await js(`document.querySelector('.tk-how [data-how="rutba"] input').click(), 1`);
    await set('.tk-email', 'dana@example.com');
    await js(`document.querySelector('.tk-ask').click(), 1`);
    const rutbaTold = await ok(until(() => js(`document.querySelector('[data-asked="dana@example.com"]')?.dataset.told === 'rutba:sent'`), 'told by Rutba', 8000));
    const mailed = intro.mailed();
    await js(`document.querySelector('.tk-add').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.tk-how [data-how="rutba"] input'))`), 'the dialog a third time', 8000);
    await wait(400);
    const closed = await js(`(() => { const n = document.querySelector('.tk-how [data-how="rutba"]'); return { disabled: n.querySelector('input').disabled, words: n.textContent }; })()`);
    await js(`[...document.querySelectorAll('.rw-dialog button')].find((b) => b.textContent.trim() === 'Cancel')?.click(), 1`);
    check('talk: Rutba emails an invitation in the person\'s name within its daily limit, the request says so, and once the limit is reached the dialog closes that way and says why',
      rutbaTold && mailed.length === 1 && mailed[0].to === 'dana@example.com' && /\/i\//.test(mailed[0].link) && closed.disabled && /limit/.test(closed.words),
      JSON.stringify({ rutbaTold, mailed: mailed.map((m) => m.to), closed }));

    // Ben's link opens Talk on his invitation; Accept makes him a contact.
    ben = new BrowserWindow({ show: false, webPreferences: { contextIsolation: true, nodeIntegration: false, partition: 'talk-invite-ben' } });
    await ben.loadURL(`http://127.0.0.1:${port}/ben`);
    const bjs = (code) => ben.webContents.executeJavaScript(code);
    await until(() => bjs(`window.ben.log.some((m) => m.type === 'welcome')`), 'Ben signed in', 6000);
    await bjs(`window.ben.send({ type: 'invite.create', note: 'Lunch on Friday?' }), 1`);
    await until(() => bjs(`window.ben.log.some((m) => m.type === 'invite.created')`), 'Ben\'s link', 6000);
    const benLink = await bjs(`window.ben.log.find((m) => m.type === 'invite.created').invite`);
    const opened = await open('talk', null, { invite: benLink.code });
    const ojs = (code) => opened.webContents.executeJavaScript(code);
    const card = await ok(until(() => ojs(`/Ben Ortiz invites you/.test(document.querySelector('.tk-invitation')?.textContent || '')`), 'the invitation', 8000));
    const cardNote = await ojs(`document.querySelector('.tk-invitation')?.textContent || ''`);
    await ojs(`document.querySelector('.tk-invite-accept').click(), 1`);
    const benAdded = await ok(until(() => ojs(`Boolean(document.querySelector('.tk-contact[data-contact="u-ben"]')) && !document.querySelector('.tk-invitation')`), 'Ben a contact', 8000));
    const benTold = await ok(until(() => bjs(`window.ben.log.some((m) => m.type === 'contact.added' && m.contact.id === 'u-amira')`), 'Ben told', 5000));
    check('talk: an invitation link opens Talk on who it is from and their note; Accept makes them a contact on both sides',
      card && /Lunch on Friday/.test(cardNote) && benAdded && benTold, JSON.stringify({ card, cardNote, benAdded, benTold }));
    opened.close();

    // Amira's own link: copied, sent by WhatsApp, read off a phone code, withdrawn.
    await js(`document.querySelector('.tk-links-btn').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.tk-link-new'))`), 'the links dialog', 6000);
    await set('.tk-link-note', 'Come and talk');
    await js(`document.querySelector('.tk-link-new').click(), 1`);
    const made = await ok(until(() => js(`/\\/i\\/[\\w-]{16,}$/.test(document.querySelector('.tk-link-url')?.value || '')`), 'the link', 6000));
    const url = await js(`document.querySelector('.tk-link-url')?.value || ''`);
    globalThis.__rutbaCheckOpened = [];
    await js(`document.querySelector('.tk-share-whatsapp').click(), 1`);
    await js(`document.querySelector('.tk-share-email-app').click(), 1`);
    await wait(300);
    const sentTo = (globalThis.__rutbaCheckOpened || []).slice();
    await js(`document.querySelector('.tk-share-qr').click(), 1`);
    const qr = await ok(until(() => js(`(document.querySelector('.tk-qr path')?.getAttribute('d') || '').length > 200`), 'the phone code', 4000));
    await js(`document.querySelector('.tk-link-copy').click(), 1`);
    await wait(200);
    const copied = await js(`window.rutbaOffice.clipboard.readText({}).then((r) => (typeof r === 'string' ? r : r?.text || ''))`).catch(() => '');
    await js(`document.querySelector('.tk-link-revoke').click(), 1`);
    const withdrawn = await ok(until(() => js(`!document.querySelector('.tk-link')`), 'the link withdrawn', 6000));
    check('talk: a link of one\'s own is made, copied, sent through a chat app or the person\'s own email app with the link in the words, shown as a code a phone reads, and withdrawn',
      made && sentTo.some((u) => u.startsWith('https://wa.me/?text=') && decodeURIComponent(u).includes(url)) && sentTo.some((u) => u.startsWith('mailto:') && decodeURIComponent(u).includes(url)) && qr && copied === url && withdrawn,
      JSON.stringify({ made, url, sentTo, qr, copied, withdrawn }));
    await js(`[...document.querySelectorAll('.rw-dialog button')].find((b) => b.textContent.trim() === 'Done')?.click(), 1`);

    const complaints = await errorsIn(win);
    check('talk: none of the invitations reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    await js(`window.rutbaOffice.talk.signOut().then(() => 1)`).catch(() => {});
    await js(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(account.id)} }).then(() => 1)`).catch(() => {});
  } catch (err) {
    check('talk: the invitation checks ran to the end', false, err?.message || String(err));
  } finally {
    try { ben?.destroy(); } catch { /* gone */ }
    await intro.close().catch(() => {});
  }
}
