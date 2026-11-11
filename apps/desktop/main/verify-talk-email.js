// Talk, Invite by email (Office 2.25, D-2026-11-11-01 and -02), on the
// stand-in introducer: with no Mail account to send from, Rutba emails the
// invitation and the dialog says so, with what it still allows; with one,
// people picked from Contacts and typed in are each asked, and each gets the
// invitation from that account as the person's own note, their name in it
// and their own link below, the copy in Sent; Rutba sends nothing for them.
// Run alone with RUTBA_VERIFY_ONLY=talk-email.

import { standInIntroducer } from './verify-talk.js';

export async function verifyTalkEmail({ open, check, until, errorsIn }) {
  const ok = (p) => p.then(() => true, () => false);
  const intro = standInIntroducer({ mailPerDay: 5 });
  let win = null;
  let account = null;
  const cards = [];
  try {
    const port = await intro.listen();
    win = await open('talk');
    const js = (code) => win.webContents.executeJavaScript(code);
    // A field's value set as typing sets it, so React hears it.
    const set = (sel, value) => js(`(() => { const n = document.querySelector(${JSON.stringify(sel)}); const proto = n instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : n instanceof HTMLSelectElement ? HTMLSelectElement.prototype : HTMLInputElement.prototype; Object.getOwnPropertyDescriptor(proto, 'value').set.call(n, ${JSON.stringify(value)}); n.dispatchEvent(new Event(n instanceof HTMLSelectElement ? 'change' : 'input', { bubbles: true })); return 1; })()`);
    const key = (sel, k) => js(`document.querySelector(${JSON.stringify(sel)}).dispatchEvent(new KeyboardEvent('keydown', { key: ${JSON.stringify(k)}, bubbles: true, cancelable: true })), 1`);
    await until(() => js(`Boolean(document.querySelector('.tk-signin'))`), 'the sign-in', 8000);
    await js(`window.rutbaOffice.talk.signIn({ token: 'amira', server: ${JSON.stringify(`ws://127.0.0.1:${port}/ws`)} }).then(() => 1)`);
    await until(() => js(`Boolean(document.querySelector('.tk-ie-btn'))`), 'signed in', 8000);
    for (const [full, email] of [['Chen Wei', 'chen@example.com'], ['Dana Ray', 'dana.ray@example.com']]) {
      cards.push(await js(`window.rutbaOffice.contacts.save({ contact: { name: { full: ${JSON.stringify(full)} }, emails: [{ value: ${JSON.stringify(email)}, type: 'work' }] }, book: 'local' })`));
    }

    // No account of her own: Rutba sends it, and the dialog says so and what is left.
    const before = await js(`window.rutbaOffice.talk.sender()`);
    await js(`document.querySelector('.tk-ie-btn').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.tk-ie-from'))`), 'the dialog', 8000);
    if (before.account) await set('.tk-ie-from', 'rutba');
    await until(() => js(`Boolean(document.querySelector('.tk-ie-rutba'))`), 'Rutba\'s way', 4000);
    const rutbaSays = await js(`(() => { const n = document.querySelector('.tk-ie-rutba'); return { words: n.textContent, left: n.dataset.left, note: document.querySelector('.tk-ie-note').value, subject: Boolean(document.querySelector('.tk-ie-subject')) }; })()`);
    await set('.tk-ie-typed', 'gus@example.com');
    await key('.tk-ie-typed', 'Enter');
    const chip = await ok(until(() => js(`Boolean(document.querySelector('.tk-ie-chip[data-email="gus@example.com"]'))`), 'the address taken', 3000));
    await js(`document.querySelector('.tk-ie-send').click(), 1`);
    const rutbaTold = await ok(until(() => js(`!document.querySelector('.tk-ie') && document.querySelector('[data-asked="gus@example.com"]')?.dataset.told === 'rutba:sent'`), 'Rutba emailed it', 8000));
    const mailed = intro.mailed().map((m) => m.to);
    check('talk: Invite by email with no Mail account of one\'s own lets Rutba send it, says so with what Rutba still allows today and this week, and the request says Rutba emailed it',
      chip && rutbaTold && mailed.includes('gus@example.com') && /no-reply@rutba\.io/.test(rutbaSays.words) && /left today/.test(rutbaSays.words) && rutbaSays.left === '5/30' && !rutbaSays.subject && rutbaSays.note.length > 0
        && (before.account || /no Mail account/.test(rutbaSays.words)),
      JSON.stringify({ before: before.account, chip, rutbaTold, mailed, rutbaSays }));

    // An account of her own: the invitation goes from it, as her note.
    account = await js(`window.rutbaOffice.mail.addAccount({ account: { email: 'amira@email-checks.example', name: 'Amira Khan', imap: { host: 'mail.email-checks.example', port: 993, secure: true }, smtp: { host: 'mail.email-checks.example', port: 587 } }, password: 'check' })`);
    const sender = await js(`window.rutbaOffice.talk.sender()`);
    await js(`document.querySelector('.tk-ie-btn').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.tk-ie-pick[data-email="chen@example.com"]'))`), 'Contacts in the dialog', 8000);
    await set('.tk-ie-from', String(account.id));
    await until(() => js(`Boolean(document.querySelector('.tk-ie-subject'))`), 'her own account chosen', 4000);
    await set('.tk-ie-search', 'chen');
    await until(() => js(`!document.querySelector('.tk-ie-pick[data-email="dana.ray@example.com"]') && Boolean(document.querySelector('.tk-ie-pick[data-email="chen@example.com"]'))`), 'Contacts searched', 4000);
    await js(`document.querySelector('.tk-ie-pick[data-email="chen@example.com"] input').click(), 1`);
    await set('.tk-ie-typed', 'Eve Stone <eve@example.com>');
    const label = await js(`document.querySelector('.tk-ie-send').textContent`);
    await set('.tk-ie-note', 'Hello {name},\n\nShall we move our calls to Talk? From the supplier day.\n\nAmira');
    await set('.tk-ie-subject', 'Our calls, on Talk');
    await js(`document.querySelector('.tk-ie-send').click(), 1`);
    const ownTold = await ok(until(() => js(`!document.querySelector('.tk-ie') && ['chen@example.com', 'eve@example.com'].every((e) => document.querySelector('[data-asked="' + e + '"]')?.dataset.told === 'own:sent')`), 'emailed from her account', 10000));
    const sent = await js(`(async () => {
      const { rows } = await window.rutbaOffice.mail.messages({ accountId: ${JSON.stringify(account.id)}, folder: 'Sent', limit: 20 });
      const out = {};
      for (const r of rows.filter((x) => x.subject === 'Our calls, on Talk')) {
        const m = await window.rutbaOffice.mail.message({ accountId: ${JSON.stringify(account.id)}, folder: 'Sent', id: r.id ?? r.uid });
        out[JSON.stringify(r.to || '').includes('chen@') ? 'chen' : 'eve'] = { hello: /^Hello (Chen Wei|Eve Stone),/.exec(m?.text || '')?.[1] || '', note: /supplier day/.test(m?.text || ''), link: /\\/i\\/[\\w-]{16,}/.exec(m?.text || '')?.[0] || '', join: /office\\.rutba\\.io\\/download\\?invite=/.test(m?.text || '') };
      }
      return out;
    })()`);
    check('talk: Invite by email from one\'s own sending account takes people from Contacts and typed in, sends each the person\'s own note from that account with their name and their own link, keeps the copies in Sent, and Rutba sends none of them',
      sender.account?.address === 'amira@email-checks.example' && /2/.test(label) && ownTold && sent.chen?.hello === 'Chen Wei' && sent.eve?.hello === 'Eve Stone' && sent.chen.note && sent.chen.join && sent.chen.link && sent.eve.link && sent.chen.link !== sent.eve.link
        && !intro.mailed().some((m) => /chen@|eve@/.test(m.to)),
      JSON.stringify({ sender: sender.account, label, ownTold, sent, mailed: intro.mailed().map((m) => m.to) }));

    const complaints = await errorsIn(win);
    check('talk: Invite by email reported no error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('talk: the Invite by email checks ran to the end', false, err?.message || String(err));
  } finally {
    const js = (code) => (win ? win.webContents.executeJavaScript(code).catch(() => {}) : Promise.resolve());
    await js(`window.rutbaOffice.talk.signOut().then(() => 1)`);
    if (account) await js(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(account.id)} }).then(() => 1)`);
    if (cards.length) await js(`window.rutbaOffice.contacts.remove({ ids: ${JSON.stringify(cards.map((c) => c?.id).filter(Boolean))} }).then(() => 1)`);
    await intro.close().catch(() => {});
  }
}
