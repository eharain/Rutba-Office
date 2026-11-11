// Calendar: meeting invitations by email (2.25, iTIP over iMIP). A person
// picked from Contacts for the Invite line; Send mails the REQUEST from the
// account Mail sends new mail from, to the test SMTP server on this computer
// (check-smtp.js), the organizer and the people on it, the time zone with
// it; an answer that arrives puts "accepted" on that person's line; a new
// time is sent as an update with SEQUENCE 1 that asks again; and deleting
// the meeting, asked first, sends the CANCEL. Run alone with
// RUTBA_VERIFY_ONLY=calendar-invites.

import { parseMessage } from '@rutba/mailbox/mime';
import { readCalendar, writeReply } from '@rutba/calendar';
import { startTestSmtp } from './check-smtp.js';

/** A message the test server took: who it went to, its subject, and its calendar part read. */
function received(m) {
  const parsed = parseMessage(m.data);
  const part = (parsed.attachments || []).find((a) => /calendar/i.test(a.type || '') || /\.ics$/i.test(a.filename || ''));
  const ics = part?.bytes ? Buffer.from(part.bytes).toString('utf8') : '';
  return { to: m.to, subject: parsed.subject || '', text: parsed.text || '', ics, cal: ics ? readCalendar(ics) : null };
}

export async function verifyCalendarInvites({ open, check, until, wait, errorsIn }) {
  let smtp = null;
  let accountId = null;
  let contactId = null;
  let mail = null;
  let win = null;
  try {
    smtp = await startTestSmtp();
    globalThis.__rutbaCheckSmtp = { port: smtp.port };
    mail = await open('mail');
    const mjs = (code) => mail.webContents.executeJavaScript(code);
    // An account on a made-up server, the only one with an address: what it sends goes to the test server and nowhere else.
    accountId = (await mjs(`window.rutbaOffice.mail.addAccount({ account: { email: 'planner@checks.example', name: 'Pat Planner', imap: { host: 'mail.checks.example', port: 993, secure: true }, smtp: { host: 'mail.checks.example', port: 587 } }, password: 'check' })`)).id;
    contactId = (await mjs(`window.rutbaOffice.contacts.save({ contact: { name: { full: 'Kim Lee', given: 'Kim', family: 'Lee', middle: '', prefix: '', suffix: '' }, emails: [{ value: 'kim@checks.example', type: 'work', pref: true, label: null }], phones: [], addresses: [], urls: [], categories: [], kind: 'individual', extra: [] } })`)).id;

    win = await open('calendar');
    const js = (code) => win.webContents.executeJavaScript(code);
    win.setContentSize(1456, 910);
    const click = (sel, text) => js(`(() => { const b = [...document.querySelectorAll(${JSON.stringify(sel)})].find((x) => x.textContent.trim() === ${JSON.stringify(text)}); b?.click(); return Boolean(b); })()`);
    const type = (sel, value) => js(`(() => { const i = document.querySelector(${JSON.stringify(sel)}); if (!i) return false; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(i, ${JSON.stringify(value)}); i.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
    const value = (sel) => js(`document.querySelector(${JSON.stringify(sel)})?.value ?? null`);
    const sendButton = () => js(`(() => { const b = document.querySelector('.rw-dialog .cal-send'); return b ? { label: b.textContent.trim(), title: b.dataset.tip || b.title || '' } : null; })()`);
    const settle = (n, what) => until(() => smtp.received.length >= n || null, what, 15000).catch(() => {});

    // A new meeting: Kim picked from Contacts, Sam typed.
    await click('.rw-ribbon .rw-btn', 'New event');
    await until(() => js(`Boolean(document.querySelector('.cal-invitees'))`), 'the editor', 5000).catch(() => {});
    await type('.cal-form .cal-title', 'Budget review');
    await type('.cal-invitees', 'Kim');
    await until(() => js(`Boolean(document.querySelector('.cal-suggest-item'))`), 'Kim from Contacts', 5000).catch(() => {});
    const offered = await js(`[...document.querySelectorAll('.cal-suggest-item')].map((b) => b.textContent)`);
    await js(`[...document.querySelectorAll('.cal-suggest-item')].find((b) => /kim@checks\\.example/.test(b.textContent))?.click(), 1`);
    await wait(150);
    const picked = await value('.cal-invitees');
    check('calendar invites: typing a name on the Invite line offers the person from Contacts, and picking them puts their address there',
      offered.some((o) => /Kim Lee/.test(o) && /kim@checks\.example/.test(o)) && /kim@checks\.example, $/.test(picked || ''), JSON.stringify({ offered, picked }));
    await type('.cal-invitees', `${picked}sam@checks.example`);
    await wait(150);
    const button = await sendButton();
    check('calendar invites: with people invited and an account to send from, the editor offers Send, from that account',
      button?.label === 'Send' && /planner@checks\.example/.test(button.title || ''), JSON.stringify(button));

    // Send: the REQUEST, from the account, to both, with the organizer and the zone.
    await js(`document.querySelector('.rw-dialog .cal-send')?.click(), 1`);
    await settle(1, 'the invitation to reach the test server');
    const first = smtp.received[0] ? received(smtp.received[0]) : null;
    const req = first?.cal?.events?.[0];
    check('calendar invites: Send mails the invitation from the account to the people invited: METHOD:REQUEST, the organizer, each asked to answer, the time zone with it',
      first && first.subject === 'Invitation: Budget review' && first.cal.method === 'REQUEST'
        && ['kim@checks.example', 'sam@checks.example'].every((a) => first.to.includes(a))
        && req?.organizer?.email === 'planner@checks.example'
        && req.attendees.some((a) => a.email === 'kim@checks.example' && a.name === 'Kim Lee' && a.partstat === 'NEEDS-ACTION' && a.rsvp)
        && (!req.start.tzid || /BEGIN:VTIMEZONE/.test(first.ics)),
      JSON.stringify({ subject: first?.subject, to: first?.to, method: first?.cal?.method, organizer: req?.organizer, attendees: req?.attendees, tz: req?.start?.tzid }));

    // Kim's answer arrives in the Inbox.
    const stored = await js(`(async () => { const all = await window.rutbaOffice.calendar.events({ from: Date.now() - 86400000, to: Date.now() + 40 * 86400000 }); const s = all.find((e) => e.summary === 'Budget review'); return s ? window.rutbaOffice.calendar.get({ id: s.id }) : null; })()`);
    const reply = writeReply(stored, { email: 'kim@checks.example', name: 'Kim Lee', partstat: 'ACCEPTED' });
    const raw = ['From: Kim Lee <kim@checks.example>', 'To: Pat Planner <planner@checks.example>', 'Subject: Accepted: Budget review', `Date: ${new Date().toUTCString()}`, 'Message-ID: <reply-1@checks.example>', 'MIME-Version: 1.0',
      'Content-Type: multipart/mixed; boundary="b1"', '', '--b1', 'Content-Type: text/plain; charset=utf-8', '', 'Kim Lee accepted.', '--b1',
      'Content-Type: text/calendar; method=REPLY; charset=utf-8', 'Content-Disposition: attachment; filename="reply.ics"', '', reply, '--b1--', ''].join('\r\n');
    await mjs(`window.rutbaOffice.mail.deliverTest({ accountId: ${JSON.stringify(accountId)}, folder: 'Inbox', raw: ${JSON.stringify(raw)} })`);
    const taken = await js(`window.rutbaOffice.calendar.checkMail()`);
    await click('.cal-view', 'Agenda');
    await until(() => js(`[...document.querySelectorAll('.cal-agenda-item')].some((b) => /Budget review/.test(b.textContent))`), 'the meeting in the agenda', 6000).catch(() => {});
    await js(`[...document.querySelectorAll('.cal-agenda-item')].find((b) => /Budget review/.test(b.textContent))?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.cal-detail'))`), 'the meeting', 4000).catch(() => {});
    const people = await js(`document.querySelector('.cal-detail')?.textContent || ''`);
    check('calendar invites: an answer that arrives in the Inbox puts "accepted" on that person\'s line',
      taken.some((r) => r.kind === 'reply' && r.answers?.[0]?.partstat === 'ACCEPTED') && /Kim Lee \(accepted\)/.test(people), JSON.stringify({ taken, people: people.slice(0, 200) }));

    // Sam's answer, with Calendar following Mail's arrivals: on the line with no read of the Inbox, where Mail
    // tells of them (2.25); where it cannot, the timed read still takes it.
    const follow = await js(`window.rutbaOffice.calendar.followMail()`);
    const samReply = writeReply(stored, { email: 'sam@checks.example', partstat: 'TENTATIVE' });
    const samRaw = ['From: sam@checks.example', 'To: Pat Planner <planner@checks.example>', 'Subject: Tentative: Budget review', `Date: ${new Date().toUTCString()}`, 'Message-ID: <reply-2@checks.example>', 'MIME-Version: 1.0',
      'Content-Type: multipart/mixed; boundary="b2"', '', '--b2', 'Content-Type: text/plain; charset=utf-8', '', 'Maybe.', '--b2',
      'Content-Type: text/calendar; method=REPLY; charset=utf-8', 'Content-Disposition: attachment; filename="reply.ics"', '', samReply, '--b2--', ''].join('\r\n');
    await mjs(`window.rutbaOffice.mail.deliverTest({ accountId: ${JSON.stringify(accountId)}, folder: 'Inbox', raw: ${JSON.stringify(samRaw)} })`);
    const samLine = () => js(`(async () => (await window.rutbaOffice.calendar.get({ id: ${JSON.stringify(stored.id)} }))?.attendees?.find((a) => a.email === 'sam@checks.example')?.partstat || null)()`);
    if (follow.following) await until(async () => (await samLine()) === 'TENTATIVE', "Sam's answer heard arriving", 6000).catch(() => {});
    const heardLine = await samLine();
    const after = await js(`window.rutbaOffice.calendar.checkMail()`);
    const samAfter = await samLine();
    await js(`window.rutbaOffice.calendar.followMail({ on: false })`);
    check('calendar invites: with Mail telling of arrivals, an answer goes onto its line as it comes and the timed read passes it by; with a Mail that cannot, the timed read takes it',
      follow.following ? heardLine === 'TENTATIVE' && !after.some((r) => r.kind === 'reply') && samAfter === 'TENTATIVE'
        : after.some((r) => r.kind === 'reply' && r.answers?.[0]?.partstat === 'TENTATIVE') && samAfter === 'TENTATIVE',
      JSON.stringify({ follow, heardLine, after, samAfter }));

    // An hour later: sent as an update, SEQUENCE 1, asking again.
    await click('.rw-dialog .rw-btn', 'Edit');
    await until(() => js(`Boolean(document.querySelector('.cal-invitees'))`), 'the editor', 4000).catch(() => {});
    const time = await value('.cal-form input[type=time]');
    const [H, M] = String(time || '10:00').split(':').map(Number);
    const later = `${String((H + 1) % 24).padStart(2, '0')}:${String(M).padStart(2, '0')}`;
    await type('.cal-form input[type=time]', later);
    await wait(150);
    const updateButton = await sendButton();
    await js(`document.querySelector('.rw-dialog .cal-send')?.click(), 1`);
    await settle(2, 'the update to reach the test server');
    const second = smtp.received[1] ? received(smtp.received[1]) : null;
    const upd = second?.cal?.events?.[0];
    check('calendar invites: a meeting sent before offers Send update; a new time goes as the update, SEQUENCE 1, everyone asked again',
      updateButton?.label === 'Send update' && second?.subject === 'Updated invitation: Budget review' && upd?.sequence === 1
        && upd.attendees.find((a) => a.email === 'kim@checks.example')?.partstat === 'NEEDS-ACTION',
      JSON.stringify({ updateButton, subject: second?.subject, sequence: upd?.sequence, attendees: upd?.attendees }));

    // Deleted, the people told: asked first (Send the cancellation), then the CANCEL.
    await until(() => js(`[...document.querySelectorAll('.cal-agenda-item')].some((b) => /Budget review/.test(b.textContent))`), 'the meeting again', 6000).catch(() => {});
    await js(`[...document.querySelectorAll('.cal-agenda-item')].find((b) => /Budget review/.test(b.textContent))?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.cal-detail'))`), 'the meeting', 4000).catch(() => {});
    globalThis.__rutbaCheckDialogAnswers.push({ kind: 'message', answer: { response: 0 } });
    await click('.rw-dialog .rw-btn', 'Delete');
    await settle(3, 'the cancellation to reach the test server');
    const third = smtp.received[2] ? received(smtp.received[2]) : null;
    await wait(300);
    const left = await js(`(async () => (await window.rutbaOffice.calendar.events({ from: Date.now() - 86400000, to: Date.now() + 40 * 86400000 })).filter((e) => e.summary === 'Budget review').length)()`);
    check('calendar invites: deleting a meeting that was sent asks first, then mails the CANCEL to the people invited and takes it off',
      third?.cal?.method === 'CANCEL' && third.cal.events[0]?.status === 'CANCELLED' && third.subject === 'Cancelled: Budget review' && left === 0,
      JSON.stringify({ method: third?.cal?.method, subject: third?.subject, left }));

    const complaints = await errorsIn(win);
    check('calendar invites: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('calendar invites: the checks ran', false, err.stack || err.message);
  } finally {
    delete globalThis.__rutbaCheckSmtp;
    await smtp?.close().catch(() => {});
    if (mail && !mail.isDestroyed()) {
      if (accountId) await mail.webContents.executeJavaScript(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(accountId)} }).then(() => 1, () => 1)`).catch(() => {});
      if (contactId) await mail.webContents.executeJavaScript(`window.rutbaOffice.contacts.remove({ id: ${JSON.stringify(contactId)} }).then(() => 1, () => 1)`).catch(() => {});
      mail.destroy();
    }
    if (win && !win.isDestroyed()) win.destroy();
  }
}
