// Calendar: a room's mailbox run by Office (2.25). View → Room mailboxes
// lists Mail's accounts and says plainly that it works only while this copy
// of Office is open; ticking one makes it a room's, with a room card in
// Contacts for "Add a room". Then, from the room's Inbox: an invitation that
// clashes with nothing is booked and accepted from the room's address, to
// the test SMTP server on this computer (check-smtp.js); one that clashes is
// declined with the times free that day; and a plain "free on ...?" gets the
// free times. Run alone with RUTBA_VERIFY_ONLY=calendar-desk.

import { parseMessage } from '@rutba/mailbox/mime';
import { readCalendar, writeRequest, dateTimeAt, localZone } from '@rutba/calendar';
import { startTestSmtp } from './check-smtp.js';

function received(m) {
  const parsed = parseMessage(m.data);
  const part = (parsed.attachments || []).find((a) => /calendar/i.test(a.type || '') || /\.ics$/i.test(a.filename || ''));
  const ics = part?.bytes ? Buffer.from(part.bytes).toString('utf8') : '';
  return { to: m.to, from: m.from, subject: parsed.subject || '', text: parsed.text || '', cal: ics ? readCalendar(ics) : null, ics };
}

const H = 3600000;

export async function verifyCalendarDesk({ open, check, until, wait, errorsIn }) {
  let smtp = null;
  let accountId = null;
  let cardId = null;
  let mail = null;
  let win = null;
  try {
    smtp = await startTestSmtp();
    globalThis.__rutbaCheckSmtp = { port: smtp.port };
    mail = await open('mail');
    const mjs = (code) => mail.webContents.executeJavaScript(code);
    accountId = (await mjs(`window.rutbaOffice.mail.addAccount({ account: { email: 'room-check@rooms.example', name: 'Check room', imap: { host: 'mail.rooms.example', port: 993, secure: true }, smtp: { host: 'mail.rooms.example', port: 587 } }, password: 'check' })`)).id;

    win = await open('calendar');
    const js = (code) => win.webContents.executeJavaScript(code);
    win.setContentSize(1456, 910);
    const click = (sel, text) => js(`(() => { const b = [...document.querySelectorAll(${JSON.stringify(sel)})].find((x) => x.textContent.trim() === ${JSON.stringify(text)}); b?.click(); return Boolean(b); })()`);

    // View → Room mailboxes: the accounts, the plain warning, and the room ticked.
    await click('.rw-tab', 'View');
    await until(() => js(`Boolean(document.querySelector('.cal-desk-open'))`), 'View → Room mailboxes', 3000).catch(() => {});
    await js(`document.querySelector('.cal-desk-open')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.cal-room-desk .cal-desk-row[data-account=${JSON.stringify(accountId)}]'))`), 'the room mailboxes', 5000).catch(() => {});
    const dialog = await js(`({ warn: document.querySelector('.cal-room-desk .cal-desk-warn')?.textContent || '', row: document.querySelector('.cal-room-desk .cal-desk-row[data-account=${JSON.stringify(accountId)}]')?.textContent || '' })`);
    await js(`document.querySelector('.cal-room-desk .cal-desk-row[data-account=${JSON.stringify(accountId)}] input')?.click(), 1`);
    await until(() => js(`(async () => (await window.rutbaOffice.calendar.roomMailboxes()).some((r) => r.id === ${JSON.stringify(accountId)} && r.room))()`), 'the account marked as a room', 5000).catch(() => {});
    await wait(300);
    const marked = await js(`(async () => ({ rows: await window.rutbaOffice.calendar.roomMailboxes(), cards: await window.rutbaOffice.contacts.bookable(), state: document.querySelector('.cal-room-desk .cal-desk-row[data-account=${JSON.stringify(accountId)}] .cal-desk-state')?.textContent || '' }))()`);
    const card = marked.cards.find((c) => c.email === 'room-check@rooms.example');
    cardId = card?.id || null;
    check('calendar desk: Room mailboxes lists Mail\'s accounts, says it works only while this copy of Office is open, and ticking one makes it a room\'s, with a room card in Contacts',
      /only while this copy of Office is open/.test(dialog.warn) && /room-check@rooms\.example/.test(dialog.row)
        && marked.rows.find((r) => r.id === accountId)?.room === true && card?.kind === 'room' && /A room: 0 bookings/.test(marked.state),
      JSON.stringify({ dialog, rows: marked.rows, card, state: marked.state }));
    await click('.cal-room-desk .rw-btn', 'Done');

    // Three days ahead, ten to eleven on this computer's clock.
    const day = new Date();
    day.setHours(0, 0, 0, 0);
    day.setDate(day.getDate() + 3);
    const at = (h) => day.getTime() + h * H;
    const zone = localZone();
    const invite = (uid, from, to) => writeRequest([{ uid, sequence: 0, summary: `Desk check ${uid}`, start: dateTimeAt(at(from), zone), end: dateTimeAt(at(to), zone),
      organizer: { email: 'kim@checks.example', name: 'Kim Lee' }, attendees: [{ email: 'room-check@rooms.example', name: 'Check room', type: 'ROOM', role: 'NON-PARTICIPANT' }] }], {});
    let n = 0;
    const deliver = (subject, parts) => {
      n += 1;
      const head = ['From: Kim Lee <kim@checks.example>', 'To: Check room <room-check@rooms.example>', `Subject: ${subject}`, `Date: ${new Date().toUTCString()}`, `Message-ID: <desk-${n}-${Date.now()}@checks.example>`, 'MIME-Version: 1.0'];
      const raw = parts.ics
        ? [...head, 'Content-Type: multipart/mixed; boundary="d1"', '', '--d1', 'Content-Type: text/plain; charset=utf-8', '', parts.text || '', '--d1',
          'Content-Type: text/calendar; method=REQUEST; charset=utf-8', 'Content-Disposition: attachment; filename="invite.ics"', '', parts.ics, '--d1--', ''].join('\r\n')
        : [...head, 'Content-Type: text/plain; charset=utf-8', '', parts.text, ''].join('\r\n');
      return mjs(`window.rutbaOffice.mail.deliverTest({ accountId: ${JSON.stringify(accountId)}, folder: 'Inbox', raw: ${JSON.stringify(raw)} })`);
    };
    const settle = (k, what) => until(() => smtp.received.length >= k || null, what, 15000).catch(() => {});

    // An invitation that clashes with nothing: booked, and accepted from the room.
    await deliver('Invitation: Desk check one', { ics: invite('desk-one@checks.example', 10, 11), text: 'Please come.' });
    const booked = await js(`window.rutbaOffice.calendar.checkMail()`);
    await settle(1, 'the room\'s acceptance');
    const yes = smtp.received[0] ? received(smtp.received[0]) : null;
    const yesLine = yes?.cal?.events?.[0]?.attendees?.[0];
    check('calendar desk: an invitation that clashes with nothing is booked, and accepted from the room\'s own address',
      booked.some((r) => r.kind === 'room-booked') && yes?.cal?.method === 'REPLY' && yes.to.includes('kim@checks.example') && /room-check@rooms\.example/.test(yes.from || '')
        && yesLine?.email === 'room-check@rooms.example' && yesLine.partstat === 'ACCEPTED' && /is booked for/.test(yes.text),
      JSON.stringify({ booked, to: yes?.to, from: yes?.from, subject: yes?.subject, line: yesLine }));

    // Half an hour into it: declined, with what is free that day.
    await deliver('Invitation: Desk check two', { ics: invite('desk-two@checks.example', 10.5, 11.5), text: 'And this?' });
    const declined = await js(`window.rutbaOffice.calendar.checkMail()`);
    await settle(2, 'the room\'s refusal');
    const no = smtp.received[1] ? received(smtp.received[1]) : null;
    const noLine = no?.cal?.events?.[0]?.attendees?.[0];
    const free = declined.find((r) => r.kind === 'room-declined')?.free || [];
    check('calendar desk: an invitation that clashes is declined, the reply saying the times free that day',
      noLine?.partstat === 'DECLINED' && /is not free for/.test(no.text) && /Free that day: /.test(no.text) && free.length >= 1 && free.every((f) => no.text.includes(f)),
      JSON.stringify({ declined, subject: no?.subject, text: no?.text?.slice(0, 300), line: noLine }));

    // A plain question: the free times that day.
    const asked = day.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
    await deliver('Room?', { text: `Hello, is the room free on ${asked}?` });
    const answered = await js(`window.rutbaOffice.calendar.checkMail()`);
    await settle(3, 'the answer to the enquiry');
    const words = smtp.received[2] ? received(smtp.received[2]) : null;
    check('calendar desk: a plain "free on ...?" to the room gets the free times that day',
      answered.some((r) => r.kind === 'room-enquiry') && words?.subject === 'Free times: Check room' && /Check room, .*: free /.test(words.text) && words.to.includes('kim@checks.example'),
      JSON.stringify({ answered, subject: words?.subject, text: words?.text?.slice(0, 300) }));

    // Add a room knows the room's whole day from its mailbox.
    const rooms = await js(`window.rutbaOffice.calendar.rooms({ start: ${at(10)}, end: ${at(11)} })`);
    const later = await js(`window.rutbaOffice.calendar.rooms({ start: ${at(14)}, end: ${at(15)} })`);
    const mine = (list) => list.find((r) => r.email === 'room-check@rooms.example');
    check('calendar desk: Add a room shows the room Busy at its booking and Free beside it, from what its mailbox knows',
      mine(rooms)?.free === false && mine(later)?.free === true && mine(later).sources.includes('room-mailbox'), JSON.stringify({ at10: mine(rooms), at14: mine(later) }));

    const complaints = await errorsIn(win);
    check('calendar desk: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('calendar desk: the checks ran', false, err.stack || err.message);
  } finally {
    delete globalThis.__rutbaCheckSmtp;
    await smtp?.close().catch(() => {});
    const api = win && !win.isDestroyed() ? win : mail && !mail.isDestroyed() ? mail : null;
    if (api && accountId) await api.webContents.executeJavaScript(`window.rutbaOffice.calendar.setRoomMailbox({ accountId: ${JSON.stringify(accountId)}, room: false }).then(() => 1, () => 1)`).catch(() => {});
    if (api && cardId) await api.webContents.executeJavaScript(`window.rutbaOffice.contacts.remove({ id: ${JSON.stringify(cardId)} }).then(() => 1, () => 1)`).catch(() => {});
    if (mail && !mail.isDestroyed()) {
      if (accountId) await mail.webContents.executeJavaScript(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(accountId)} }).then(() => 1, () => 1)`).catch(() => {});
      mail.destroy();
    }
    if (win && !win.isDestroyed()) win.destroy();
  }
}
