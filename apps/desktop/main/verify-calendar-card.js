// Calendar: its card in Mail's reading pane (calendar-invite-card.js, 2.25).
// An invitation shows Accept, Tentative and Decline and what else is on at
// that time; Accept puts it in the calendar and mails the REPLY from the
// address invited (to the test SMTP server on this computer, check-smtp.js);
// the organizer's cancellation, opened, marks it cancelled and offers Remove
// from calendar; and an answer to this person's own meeting goes onto that
// person's line as the card opens. Every message here goes to an account of
// the check's own, on a made-up server. Run alone with
// RUTBA_VERIFY_ONLY=calendar-card.

import { parseMessage } from '@rutba/mailbox/mime';
import { readCalendar, writeReply, localZone } from '@rutba/calendar';
import { startTestSmtp } from './check-smtp.js';
import { labelKeyed } from './verify-keys.js';

const pad = (n) => String(n).padStart(2, '0');
const utc = (ms) => { const d = new Date(ms); return `${d.getUTCFullYear()}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}T${pad(d.getUTCHours())}${pad(d.getUTCMinutes())}00Z`; };

/** A message with a calendar part, as a calendar program mails it. */
function scheduling({ from, to, subject, method, ics, id }) {
  return [`From: ${from}`, `To: ${to}`, `Subject: ${subject}`, `Date: ${new Date().toUTCString()}`, `Message-ID: <${id}>`, 'MIME-Version: 1.0',
    'Content-Type: multipart/mixed; boundary="b1"', '', '--b1', 'Content-Type: text/plain; charset=utf-8', '', subject, '--b1',
    `Content-Type: text/calendar; method=${method}; charset=utf-8`, `Content-Disposition: attachment; filename="${method === 'REPLY' ? 'reply' : 'invite'}.ics"`, '', ics, '--b1--', ''].join('\r\n');
}

const invite = ({ start, end, method = 'REQUEST', sequence = 0 }) => ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Northwind//Calendar//EN', `METHOD:${method}`, 'BEGIN:VEVENT',
  'UID:card-1@northwind.example', `DTSTAMP:${utc(Date.now())}`, `DTSTART:${utc(start)}`, `DTEND:${utc(end)}`, 'SUMMARY:Card check', 'LOCATION:Room 2', `SEQUENCE:${sequence}`,
  `STATUS:${method === 'CANCEL' ? 'CANCELLED' : 'CONFIRMED'}`, 'ORGANIZER;CN=Amina Yusuf:mailto:amina@northwind.example',
  'ATTENDEE;CN=You;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE:mailto:you.here@checks.example', 'END:VEVENT', 'END:VCALENDAR', ''].join('\r\n');

export async function verifyCalendarCard({ open, check, until, wait, errorsIn }) {
  let smtp = null;
  let accountId = null;
  let mail = null;
  const made = [];
  try {
    smtp = await startTestSmtp();
    globalThis.__rutbaCheckSmtp = { port: smtp.port };
    mail = await open('mail');
    const js = (code) => mail.webContents.executeJavaScript(code);
    accountId = (await js(`window.rutbaOffice.mail.addAccount({ account: { email: 'you.here@checks.example', name: 'Card Check', imap: { host: 'mail.checks.example', port: 993, secure: true }, smtp: { host: 'mail.checks.example', port: 587 } }, password: 'check' })`)).id;
    const deliver = (raw) => js(`window.rutbaOffice.mail.deliverTest({ accountId: ${JSON.stringify(accountId)}, folder: 'Inbox', raw: ${JSON.stringify(raw)} })`);

    // Tomorrow at 15:00 here, with a dentist's appointment already in the calendar then.
    const day = new Date();
    day.setDate(day.getDate() + 1);
    day.setHours(15, 0, 0, 0);
    const start = day.getTime();
    const end = start + 3600000;
    const local = localZone();
    const at = (ms) => { const x = new Date(ms); return { date: `${x.getFullYear()}-${pad(x.getMonth() + 1)}-${pad(x.getDate())}`, time: `${pad(x.getHours())}:${pad(x.getMinutes())}:00`, tzid: local, utc: false, allDay: false, at: ms }; };
    made.push((await js(`(async () => { const c = await window.rutbaOffice.calendar.calendars({}); return (await window.rutbaOffice.calendar.save({ calendarId: c[0].id, event: { summary: 'Dentist', start: ${JSON.stringify(at(start + 1800000))}, end: ${JSON.stringify(at(start + 5400000))} } })).id; })()`)));

    await deliver(scheduling({ from: 'Amina Yusuf <amina@northwind.example>', to: 'you.here@checks.example', subject: 'Invitation: Card check', method: 'REQUEST', ics: invite({ start, end }), id: 'card-1@northwind.example' }));
    mail.webContents.reload();
    await until(() => js(`[...document.querySelectorAll('.ml-accounts button')].some((b) => /Card Check/.test(b.textContent))`), 'the check\'s account', 10000).catch(() => {});
    await js(`[...document.querySelectorAll('.ml-accounts button')].find((b) => /Card Check/.test(b.textContent))?.click(), 1`);
    const openRow = async (subject) => {
      await until(() => js(`[...document.querySelectorAll('.ml-row')].some((r) => r.textContent.includes(${JSON.stringify(subject)}))`), subject, 8000).catch(() => {});
      await js(`[...document.querySelectorAll('.ml-row')].find((r) => r.textContent.includes(${JSON.stringify(subject)}))?.click(), 1`);
    };
    const cardNow = () => js(`(() => { const c = document.querySelector('.cal-card'); return c ? { method: c.dataset.method, title: c.querySelector('.ml-invite-title')?.textContent || '', who: c.querySelector('.ml-invite-who')?.textContent || '', clash: c.querySelector('.cal-card-clash')?.textContent || '', note: c.querySelector('.cal-card-note')?.textContent || '', chip: c.querySelector('.chip')?.textContent || '', buttons: [...c.querySelectorAll('.rw-btn')].map((b) => b.textContent.trim()) } : null; })()`);

    // The invitation: its card, with what else is on then.
    await openRow('Invitation: Card check');
    await until(() => js(`Boolean(document.querySelector('.cal-card .rw-btn'))`), 'the card', 8000).catch(() => {});
    const asked = await cardNow();
    check('calendar card: an invitation in Mail shows the calendar\'s card, with Accept, Tentative and Decline, and what else is on at that time',
      asked?.method === 'REQUEST' && asked.title === 'Card check' && /Amina Yusuf/.test(asked.who) && asked.buttons.join(',') === 'Accept,Tentative,Decline' && /Dentist/.test(asked.clash),
      JSON.stringify(asked));

    // Accept: in the calendar, and the answer mailed from the address invited.
    await js(`((list) => list.find((b) => ${labelKeyed("Accept")}(b)) || list.find((b) => b.textContent.trim() === 'Accept'))([...document.querySelectorAll('.cal-card .rw-btn')])?.click(), 1`);
    await until(() => smtp.received.length >= 1 || null, 'the answer to reach the test server', 15000).catch(() => {});
    const sent = smtp.received[0] ? parseMessage(smtp.received[0].data) : null;
    const part = (sent?.attachments || []).find((a) => /calendar/i.test(a.type || '') || /\.ics$/i.test(a.filename || ''));
    const reply = part?.bytes ? readCalendar(Buffer.from(part.bytes).toString('utf8')) : null;
    await until(async () => /you answered accepted/.test((await cardNow())?.who || '') || null, 'the answer on the card', 4000).catch(() => {});
    const answered = await cardNow();
    const kept = await js(`(async () => { const all = await window.rutbaOffice.calendar.events({ from: ${start - 1}, to: ${end + 1} }); return all.filter((e) => e.summary === 'Card check').map((e) => e.id); })()`);
    check('calendar card: Accept keeps the meeting and mails the REPLY to the organizer from the address invited',
      kept.length === 1 && smtp.received[0]?.to.includes('amina@northwind.example') && /^Accepted: Card check/.test(sent?.subject || '')
        && reply?.method === 'REPLY' && reply.events[0]?.attendees?.[0]?.email === 'you.here@checks.example' && reply.events[0].attendees[0].partstat === 'ACCEPTED' && /you answered accepted/.test(answered?.who || ''),
      JSON.stringify({ kept, to: smtp.received[0]?.to, subject: sent?.subject, reply: reply?.events?.[0]?.attendees, who: answered?.who }));

    // The organizer cancels: the card marks it cancelled, and Remove from calendar takes it off.
    await deliver(scheduling({ from: 'Amina Yusuf <amina@northwind.example>', to: 'you.here@checks.example', subject: 'Cancelled: Card check', method: 'CANCEL', ics: invite({ start, end, method: 'CANCEL', sequence: 1 }), id: 'card-2@northwind.example' }));
    await openRow('Cancelled: Card check');
    await until(async () => (await cardNow())?.method === 'CANCEL' || null, 'the cancellation\'s card', 8000).catch(() => {});
    const off = await cardNow();
    const marked = await js(`(async () => (await window.rutbaOffice.calendar.get({ id: ${JSON.stringify(kept[0] || '')} }))?.status || null)()`);
    await js(`((list) => list.find((b) => ${labelKeyed("Remove from calendar")}(b)) || list.find((b) => b.textContent.trim() === 'Remove from calendar'))([...document.querySelectorAll('.cal-card .rw-btn')])?.click(), 1`);
    await until(async () => /Taken off/.test((await cardNow())?.note || '') || null, 'the meeting taken off', 4000).catch(() => {});
    const gone = await js(`(async () => (await window.rutbaOffice.calendar.events({ from: ${start - 1}, to: ${end + 1} })).filter((e) => e.summary === 'Card check').length)()`);
    check('calendar card: the organizer\'s cancellation, opened, marks the meeting cancelled, and Remove from calendar takes it off',
      off?.title === 'Cancelled: Card check' && /Marked cancelled/.test(off.note) && off.buttons.includes('Remove from calendar') && marked === 'CANCELLED' && gone === 0,
      JSON.stringify({ off, marked, gone }));

    // An answer to this person's own meeting: onto that person's line as the card opens.
    const own = await js(`(async () => { const c = await window.rutbaOffice.calendar.calendars({}); return window.rutbaOffice.calendar.save({ calendarId: c[0].id, event: { summary: 'Own meeting', start: ${JSON.stringify(at(end + 3600000))}, end: ${JSON.stringify(at(end + 7200000))}, organizer: { email: 'you.here@checks.example', name: 'Card Check' }, attendees: [{ email: 'lee@checks.example', name: 'Lee', role: 'REQ-PARTICIPANT', partstat: 'NEEDS-ACTION', rsvp: true }] } }); })()`);
    made.push(own.id);
    await deliver(scheduling({ from: 'Lee <lee@checks.example>', to: 'you.here@checks.example', subject: 'Accepted: Own meeting', method: 'REPLY', ics: writeReply(own, { email: 'lee@checks.example', name: 'Lee', partstat: 'ACCEPTED' }), id: 'card-3@checks.example' }));
    await openRow('Accepted: Own meeting');
    await until(async () => (await cardNow())?.method === 'REPLY' || null, 'the answer\'s card', 8000).catch(() => {});
    const theirs = await cardNow();
    const line = await js(`(async () => (await window.rutbaOffice.calendar.get({ id: ${JSON.stringify(own.id)} }))?.attendees?.find((a) => a.email === 'lee@checks.example')?.partstat || null)()`);
    check('calendar card: an answer to one\'s own meeting shows who answered what, and is on their line in the calendar as the card opens',
      /Lee accepted/.test(theirs?.chip || '') && /on the meeting in your calendar/.test(theirs?.note || '') && line === 'ACCEPTED', JSON.stringify({ theirs, line }));

    const complaints = await errorsIn(mail);
    check('calendar card: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('calendar card: the checks ran', false, err.stack || err.message);
  } finally {
    delete globalThis.__rutbaCheckSmtp;
    await smtp?.close().catch(() => {});
    if (mail && !mail.isDestroyed()) {
      const js = (code) => mail.webContents.executeJavaScript(code).catch(() => {});
      for (const id of made) await js(`window.rutbaOffice.calendar.remove({ id: ${JSON.stringify(id)} }).then(() => 1, () => 1)`);
      if (accountId) await js(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(accountId)} }).then(() => 1, () => 1)`);
      mail.destroy();
    }
  }
}
