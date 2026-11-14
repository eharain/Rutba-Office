// Mail: the sending account (D-2026-11-11-02). Accounts has "Send new mail
// from", listing only the accounts that can send; the one chosen is kept,
// new mail written in a mailbox brought in from a file goes from it, and
// mail.sendFrom, which Calendar and Talk call, sends from it to the server
// (a test SMTP server on this computer), a 'text/calendar; method=REQUEST'
// attachment going as the invitation, with the message kept in Sent. Run
// alone with RUTBA_VERIFY_ONLY=mail-default-account.

import { startTestSmtp } from './check-smtp.js';
import { labelKeyed } from './verify-keys.js';

const INVITE = [
  'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Rutba//Calendar//EN', 'METHOD:REQUEST',
  'BEGIN:VEVENT', 'UID:check-planning@rutba', 'DTSTAMP:20261111T090000Z', 'DTSTART:20261112T100000Z', 'DTEND:20261112T110000Z',
  'SUMMARY:Planning', 'ORGANIZER:mailto:raj@checks.example', 'ATTENDEE;RSVP=TRUE:mailto:ann@checks.example', 'END:VEVENT', 'END:VCALENDAR', '',
].join('\r\n');

export async function verifyMailDefaultAccount({ open, check, until, wait, errorsIn }) {
  let win = null;
  let smtp = null;
  const made = [];
  try {
    win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the accounts', 8000);
    const seeded = await js(`(async () => (await window.rutbaOffice.mail.accounts()).find((a) => a.local)?.id || null)()`);
    for (const [email, name] of [['ana@checks.example', 'Ana Silva'], ['raj@checks.example', 'Raj Patel']]) {
      made.push((await js(`window.rutbaOffice.mail.addAccount({ account: { email: ${JSON.stringify(email)}, name: ${JSON.stringify(name)}, imap: { host: 'mail.checks.example', port: 993, secure: true }, smtp: { host: 'mail.checks.example', port: 587 } }, password: 'check' })`)).id);
    }
    win.webContents.reload();
    await until(() => js(`[...document.querySelectorAll('.ml-accounts button')].some((b) => /Raj Patel/.test(b.textContent))`), 'the two accounts', 10000);
    const first = await js(`window.rutbaOffice.mail.defaultAccount()`);

    // Accounts: Send new mail from, the two that can send and not the mailbox brought in from a file.
    const openAccounts = async () => {
      await js(`[...document.querySelectorAll('.ml-add-account button')].find((b) => /Manage accounts/.test(b.textContent)).click(), 1`);
      await until(() => js(`Boolean(document.querySelector('select.ml-default-account'))`), 'Send new mail from', 6000);
    };
    const closeDialog = () => js(`[...document.querySelectorAll('.rw-dialog button')].find((b) => /^(Done|Close)$/.test(b.textContent.trim()))?.click(), 1`);
    await openAccounts();
    const offered = await js(`(() => { const s = document.querySelector('select.ml-default-account'); return { value: s.value, options: [...s.options].map((o) => o.value), label: s.closest('.rw-field')?.querySelector('label')?.textContent || '' }; })()`);
    await js(`(() => { const s = document.querySelector('select.ml-default-account'); Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, ${JSON.stringify(made[1])}); s.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
    await until(() => js(`(async () => (await window.rutbaOffice.mail.defaultAccount())?.id === ${JSON.stringify(made[1])})()`), 'the chosen account kept', 5000).catch(() => {});
    const chosen = await js(`window.rutbaOffice.mail.defaultAccount()`);
    await closeDialog();
    await wait(300);
    await openAccounts();
    const shownAgain = await js(`document.querySelector('select.ml-default-account').value`);
    await closeDialog();
    check('mail: Accounts has Send new mail from with only the accounts that can send, the first of them until one is chosen, and the one chosen is kept',
      first?.id === made[0] && offered.label === 'Send new mail from' && offered.value === made[0] && offered.options.length === 2 && made.every((id) => offered.options.includes(id)) && !offered.options.includes(seeded)
        && chosen?.id === made[1] && chosen.address === 'raj@checks.example' && chosen.name === 'Raj Patel' && shownAgain === made[1],
      JSON.stringify({ first, offered, chosen, shownAgain, seeded }));

    // New mail written in the mailbox brought in from a file goes from the chosen account.
    if (seeded) {
      const seededEmail = await js(`(async () => (await window.rutbaOffice.mail.accounts()).find((a) => a.id === ${JSON.stringify(seeded)})?.email || '')()`);
      await js(`[...document.querySelectorAll('.ml-accounts button')].find((b) => b.title === ${JSON.stringify(seededEmail)} || b.closest('[title]')?.title === ${JSON.stringify(seededEmail)})?.click(), 1`);
      await wait(400);
      await js(`document.querySelector('.ml-compose-cta button')?.click(), 1`);
      await until(() => js(`Boolean(document.querySelector('.ml-form'))`), 'the new message', 5000);
      const from = await js(`(() => { const f = [...document.querySelectorAll('.ml-form .rw-field')].find((x) => x.querySelector('label')?.textContent === 'From'); return f?.querySelector('select')?.value || null; })()`);
      check('mail: a new message written in a mailbox brought in from a file goes from the account chosen in Send new mail from', from === made[1], JSON.stringify({ from, seededEmail }));
      await js(`((list) => list.find((b) => ${labelKeyed("Discard")}(b)) || list.find((b) => b.textContent.trim() === 'Discard'))([...document.querySelectorAll('.ml-compose .rw-btn, .ml-compose button')])?.click() || document.querySelector('.ml-compose-close')?.click(), 1`);
      await wait(300);
    }

    // mail.sendFrom, as Calendar and Talk call it: from the default, the invitation as its calendar part, kept in Sent.
    smtp = await startTestSmtp();
    globalThis.__rutbaCheckSmtp = { port: smtp.port };
    const done = await js(`window.rutbaOffice.mail.sendFrom({
      to: [{ name: 'Ann Lee', address: 'ann@checks.example' }],
      subject: 'Invitation: Planning',
      text: 'Planning, Thursday 12 November, 10:00.',
      attachments: [{ filename: 'invite.ics', content: ${JSON.stringify(INVITE)}, contentType: 'text/calendar; method=REQUEST' }],
    })`);
    await until(() => Promise.resolve(smtp.received.length > 0), 'the message at the test server', 6000).catch(() => {});
    const got = smtp.received[0] || { from: null, to: [], data: '' };
    // The calendar part's words, however the message encoded them (as they are, quoted-printable or base64).
    const carriesInvite = /BEGIN:VCALENDAR/.test(got.data.replace(/=\r\n/g, '')) || got.data.includes(Buffer.from('BEGIN:VCALENDAR').toString('base64').slice(0, 16));
    const sentFolder = await js(`(async () => (await window.rutbaOffice.mail.folders({ accountId: ${JSON.stringify(made[1])} })).find((f) => f.role === 'sent')?.path || null)()`);
    const kept = sentFolder ? await js(`(async () => (await window.rutbaOffice.mail.messages({ accountId: ${JSON.stringify(made[1])}, folder: ${JSON.stringify(sentFolder)} })).rows.map((r) => r.subject))()`) : [];
    check('mail: sendFrom sends from the default account to its server, a text/calendar; method=REQUEST attachment going as the invitation, and keeps it in Sent',
      done?.accountId === made[1] && done.from === 'raj@checks.example' && got.from === 'raj@checks.example' && got.to.includes('ann@checks.example')
        && /text\/calendar;[^\r\n]*method=REQUEST/i.test(got.data) && carriesInvite,
      JSON.stringify({ done, from: got.from, to: got.to, calendar: /text\/calendar[^\r\n]*/i.exec(got.data)?.[0] || null }));
    check('mail: what sendFrom sent is in that account\'s Sent', kept.includes('Invitation: Planning'), JSON.stringify({ sentFolder, kept }));

    const refused = await js(`window.rutbaOffice.mail.sendFrom({ account: ${JSON.stringify(seeded || 'none')}, to: 'ann@checks.example', subject: 'x' }).then(() => 'sent', (e) => e.message)`);
    check('mail: sendFrom from a mailbox brought in from a file is refused in words', /cannot send|not set up/.test(String(refused)), String(refused));

    const complaints = await errorsIn(win);
    check('mail: the sending account reported no error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('mail: the sending account checks ran to the end', false, err?.message || String(err));
  } finally {
    delete globalThis.__rutbaCheckSmtp;
    await smtp?.close().catch(() => {});
    for (const id of made) await win?.webContents.executeJavaScript(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(id)} }).then(() => 1, () => 1)`).catch(() => {});
  }
}
