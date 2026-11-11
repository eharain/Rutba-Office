// Meeting invitations by email (iTIP over iMIP, RFC 5546 and 6047): the
// REQUEST a Send mails, with its time zone; an edit sent again as an update
// that asks again only when the time or place moved; the CANCEL to someone
// taken off the list and to everyone when it is deleted; the REPLYs that put
// each answer on its line; and an answer sent from the address invited.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { readCalendar, unfold, writeRequest, writeCancel, applyReply, significantChange, vtimezone, readScheduling, writeReply, localToUtc } from '../packages/calendar/src/index.js';
import { createCalendarService } from '../apps/desktop/main/calendar.js';
import { createMailLink } from '../apps/desktop/main/calendar-mail.js';
import { ENGLISH, inviteText, wordsFrom } from '../apps/desktop/main/calendar-text.js';

const at = (y, m, d, H, M, tz) => ({ date: `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`, time: `${String(H).padStart(2, '0')}:${String(M).padStart(2, '0')}:00`, tzid: tz, utc: false, allDay: false, at: localToUtc({ y, m, d, H, M }, tz) });

const meeting = (over = {}) => ({
  uid: 'plan@rutba.io', sequence: 0, summary: 'Planning', location: 'Room 4',
  start: at(2026, 11, 17, 14, 0, 'Europe/London'), end: at(2026, 11, 17, 15, 0, 'Europe/London'),
  organizer: { email: 'me@example.com', name: 'Me' },
  attendees: [{ email: 'kim@example.org', name: 'Kim', role: 'REQ-PARTICIPANT' }, { email: 'sam@example.org', name: null }],
  ...over,
});

test('a REQUEST names the organizer, asks each person to answer, and carries the time zone it uses', () => {
  const text = writeRequest([meeting()], {});
  assert.match(text, /METHOD:REQUEST/);
  assert.match(text, /ORGANIZER;CN=Me:mailto:me@example\.com/);
  assert.match(unfold(text).join('\n'), /ATTENDEE;CN=Kim;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE:mailto:kim@example\.org/);
  assert.match(text, /BEGIN:VTIMEZONE\r\nTZID:Europe\/London/);
  assert.ok(text.indexOf('BEGIN:VTIMEZONE') < text.indexOf('BEGIN:VEVENT'), 'the zone before the event that names it');
  const read = readCalendar(text);
  assert.equal(read.method, 'REQUEST');
  assert.equal(read.events[0].start.at, meeting().start.at);
  assert.equal(read.events[0].attendees.length, 2);
});

test('a VTIMEZONE has the year\'s changes of offset, each at its wall-clock time', () => {
  const lines = vtimezone('Europe/London', Date.UTC(2026, 5, 1), Date.UTC(2026, 6, 1));
  const text = lines.join('\n');
  assert.match(text, /BEGIN:DAYLIGHT\nDTSTART:20260329T010000\nTZOFFSETFROM:\+0000\nTZOFFSETTO:\+0100/);
  assert.match(text, /BEGIN:STANDARD\nDTSTART:20261025T020000\nTZOFFSETFROM:\+0100\nTZOFFSETTO:\+0000/);
  assert.equal(vtimezone('UTC', 0, 1), null);
  const kolkata = vtimezone('Asia/Kolkata', Date.UTC(2026, 0, 1), Date.UTC(2026, 1, 1)).join('\n');
  assert.match(kolkata, /TZOFFSETTO:\+0530/);
  assert.doesNotMatch(kolkata, /DAYLIGHT/);
});

test('a change asks again only when the time, the place or the rule moved', () => {
  assert.equal(significantChange(meeting(), meeting({ summary: 'Planning, with numbers' })), false);
  assert.equal(significantChange(meeting(), meeting({ location: 'Room 5' })), true);
  assert.equal(significantChange(meeting(), meeting({ start: at(2026, 11, 17, 15, 0, 'Europe/London') })), true);
});

test('a CANCEL says CANCELLED to the people it names', () => {
  const text = writeCancel([meeting({ sequence: 2 })], { attendees: [{ email: 'sam@example.org' }] });
  const read = readCalendar(text);
  assert.equal(read.method, 'CANCEL');
  assert.equal(read.events[0].status, 'CANCELLED');
  assert.equal(read.events[0].sequence, 2);
  assert.deepEqual(read.events[0].attendees.map((a) => a.email), ['sam@example.org']);
});

test('a REPLY puts the answer on its line; an older one does not undo a newer', () => {
  const reply = readCalendar(writeReply(meeting(), { email: 'Kim@Example.org', partstat: 'accepted' })).events[0];
  const r = applyReply(meeting(), reply);
  assert.equal(r.changed, true);
  assert.equal(r.attendees.find((a) => a.email === 'kim@example.org').partstat, 'ACCEPTED');
  const older = { ...reply, dtstamp: { at: 1 }, attendees: [{ email: 'kim@example.org', partstat: 'DECLINED' }] };
  assert.equal(applyReply({ ...meeting(), attendees: r.attendees }, older).changed, false);
  const crasher = applyReply(meeting(), { ...reply, attendees: [{ email: 'lee@example.net', partstat: 'ACCEPTED' }] });
  assert.equal(crasher.attendees.find((a) => a.email === 'lee@example.net').uninvited, true);
});

test('the words of a message come from the window, English for any it left out', () => {
  const words = wordsFrom({ language: 'en', inviteSubject: 'Einladung: {summary}', nonsense: 'x' });
  assert.equal(words.inviteSubject, 'Einladung: {summary}');
  assert.equal(words.when, ENGLISH.when);
  assert.equal(words.nonsense, undefined);
  const text = inviteText(meeting(), wordsFrom(null));
  assert.match(text, /^Planning\n\nWhen: .*Europe\/London/);
  assert.match(text, /Where: Room 4/);
  assert.match(text, /Invited: Kim, sam@example\.org/);
});

/* ── the service, with a Mail that records what it is asked to send ──── */

function fakeMail({ withSendFrom = true, email = 'me@example.com' } = {}) {
  const sent = [];
  const inbox = [];
  const service = {
    accounts: () => [{ id: 'acc-1', email, name: 'Me' }],
    folders: () => [{ path: 'INBOX', role: 'inbox' }],
    messages: ({ limit }) => ({ rows: inbox.slice(-limit).reverse() }),
    message: ({ id }) => inbox.find((m) => m.id === id)?.message || null,
    attachmentText: ({ id }) => inbox.find((m) => m.id === id)?.part || null,
    send: async ({ accountId, draft }) => { sent.push({ via: 'send', accountId, ...draft }); return { messageId: `m${sent.length}`, refused: [] }; },
  };
  if (withSendFrom) service.sendFrom = async (msg) => { sent.push({ via: 'sendFrom', ...msg }); return { messageId: `m${sent.length}`, refused: [] }; };
  const deliver = (from, text) => inbox.push({ id: `in-${inbox.length + 1}`, from: { address: from }, hasAttachments: true, attachmentCount: 1, message: { attachments: [{ type: 'text/calendar', filename: 'reply.ics' }] }, part: { name: 'reply.ics', type: 'text/calendar', text } });
  return { service, sent, deliver };
}

function setup(opts) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-itip-'));
  const mail = fakeMail(opts);
  const calendar = createCalendarService({ stores: { dir }, broadcast: null, mail: createMailLink(() => mail.service) });
  return { dir, calendar, mail, done: () => fs.rmSync(dir, { recursive: true, force: true }) };
}

const icsOf = (m) => fs.readFileSync(m.attachments[0].path, 'utf8');

test('Send mails the REQUEST from the account; an edit sends the update, asking again only when it moved; those taken off get a CANCEL', async () => {
  const { calendar, mail, done } = setup();
  try {
    const cal = calendar.calendars()[0];
    const saved = calendar.save({ calendarId: cal.id, event: { ...meeting(), organizer: null, uid: null } });
    const first = await calendar.sendInvitation({ id: saved.id });
    assert.deepEqual(first.to, ['kim@example.org', 'sam@example.org']);
    assert.equal(first.from, 'me@example.com');
    assert.equal(first.update, false);
    const m1 = mail.sent[0];
    assert.equal(m1.via, 'sendFrom');
    assert.equal(m1.account, 'acc-1');
    assert.equal(m1.subject, 'Invitation: Planning');
    assert.equal(m1.attachments[0].contentType, 'text/calendar; method=REQUEST; charset=utf-8');
    assert.match(m1.attachments[0].content, /METHOD:REQUEST/);
    assert.match(m1.attachments[0].content, /ORGANIZER;CN=Me:mailto:me@example\.com/);

    // Kim accepts; a new title is sent without asking her again.
    calendar.takeMail({ from: 'kim@example.org', text: writeReply(readCalendar(m1.attachments[0].content).events[0], { email: 'kim@example.org', partstat: 'ACCEPTED' }) });
    let kept = calendar.get({ id: saved.id });
    calendar.save({ calendarId: cal.id, event: { ...kept, summary: 'Planning 2027' } });
    const second = await calendar.sendInvitation({ id: saved.id });
    assert.equal(second.update, true);
    assert.equal(second.significant, false);
    assert.equal(mail.sent[1].subject, 'Updated invitation: Planning 2027');
    const update = readCalendar(mail.sent[1].attachments[0].content).events[0];
    assert.equal(update.sequence, 1, 'SEQUENCE raised by the edit');
    assert.equal(update.attendees.find((a) => a.email === 'kim@example.org').partstat, 'ACCEPTED');

    // A new time, and Sam taken off: Kim is asked again, Sam is told it is off for him.
    kept = calendar.get({ id: saved.id });
    calendar.save({ calendarId: cal.id, event: { ...kept, start: at(2026, 11, 18, 10, 0, 'Europe/London'), end: at(2026, 11, 18, 11, 0, 'Europe/London'), attendees: kept.attendees.filter((a) => a.email !== 'sam@example.org') } });
    const third = await calendar.sendInvitation({ id: saved.id });
    assert.equal(third.significant, true);
    assert.deepEqual(third.removed, ['sam@example.org']);
    const moved = readCalendar(mail.sent[2].attachments[0].content).events[0];
    assert.equal(moved.attendees.find((a) => a.email === 'kim@example.org').partstat, 'NEEDS-ACTION');
    const offFor = mail.sent[3];
    assert.deepEqual(offFor.to, ['sam@example.org']);
    assert.equal(readCalendar(offFor.attachments[0].content).method, 'CANCEL');
  } finally {
    done();
  }
});

test('an answer that arrives goes onto its line, from the Inbox as well as from a message opened; a stranger\'s line for someone else is not taken', async () => {
  const { calendar, mail, done } = setup();
  try {
    const cal = calendar.calendars()[0];
    const saved = calendar.save({ calendarId: cal.id, event: { ...meeting(), organizer: null } });
    await calendar.sendInvitation({ id: saved.id });
    const request = readCalendar(mail.sent[0].attachments[0].content).events[0];
    // Someone else answering for Kim is not Kim.
    assert.equal(calendar.takeMail({ from: 'mallory@example.net', text: writeReply(request, { email: 'lee@example.net', partstat: 'ACCEPTED' }) }).answers.length, 0);
    mail.deliver('sam@example.org', writeReply(request, { email: 'sam@example.org', partstat: 'TENTATIVE' }));
    const seen = await calendar.checkMail();
    assert.equal(seen.length, 1);
    assert.equal(seen[0].kind, 'reply');
    assert.equal(seen[0].answers[0].partstat, 'TENTATIVE');
    assert.equal(calendar.get({ id: saved.id }).attendees.find((a) => a.email === 'sam@example.org').partstat, 'TENTATIVE');
    assert.equal((await calendar.checkMail()).length, 0, 'a message read once is not read again');
  } finally {
    done();
  }
});

test('deleting a meeting sent before tells everyone; one occurrence names its RECURRENCE-ID', async () => {
  const { calendar, mail, done } = setup();
  try {
    const cal = calendar.calendars()[0];
    const weekly = { freq: 'WEEKLY', interval: 1, count: null, until: null, byDay: null, byMonthDay: null, byMonth: null, bySetPos: null, wkst: 'MO' };
    const saved = calendar.save({ calendarId: cal.id, event: { ...meeting(), organizer: null, rrule: weekly } });
    assert.equal((await calendar.sendCancel({ id: saved.id })).sent, false, 'never sent: nobody to tell');
    await calendar.sendInvitation({ id: saved.id });
    const second = saved.start.at + 7 * 86400000;
    const one = await calendar.sendCancel({ id: saved.id, scope: 'this', original: second });
    assert.equal(one.sent, true);
    const cancel = readCalendar(mail.sent[1].attachments[0].content).events[0];
    assert.equal(cancel.recurrenceId.at, second);
    assert.equal(cancel.status, 'CANCELLED');
    const all = await calendar.sendCancel({ id: saved.id });
    assert.deepEqual(all.to, ['kim@example.org', 'sam@example.org']);
    assert.equal(mail.sent[2].subject, 'Cancelled: Planning');
  } finally {
    done();
  }
});

const invitation = ({ sequence = 0, start = '20261028T120000Z', status = 'CONFIRMED', method = 'REQUEST' } = {}) => [
  'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Other//EN', `METHOD:${method}`, 'BEGIN:VEVENT', 'UID:lunch@other.example', `SEQUENCE:${sequence}`,
  'DTSTAMP:20261020T090000Z', `DTSTART:${start}`, 'DTEND:20261028T130000Z', 'SUMMARY:Lunch', `STATUS:${status}`,
  'ORGANIZER;CN=Sam:mailto:sam@other.example',
  'ATTENDEE;CN=Me;PARTSTAT=NEEDS-ACTION;RSVP=TRUE:mailto:me@example.com',
  'END:VEVENT', 'END:VCALENDAR', '',
].join('\r\n');

test('Accept sends the REPLY to the organizer from the address invited, and the organizer\'s change and cancellation reach the calendar', async () => {
  const { calendar, mail, done } = setup();
  try {
    const r = await calendar.answer({ text: invitation(), partstat: 'accepted', comment: 'See you there.' });
    assert.equal(r.sent, true);
    assert.equal(r.from, 'me@example.com');
    const m = mail.sent[0];
    assert.deepEqual(m.to, ['sam@other.example']);
    assert.equal(m.subject, 'Accepted: Lunch');
    assert.match(m.text, /^Me accepted\.\n\nSee you there\./);
    const reply = readCalendar(m.attachments[0].content);
    assert.equal(reply.method, 'REPLY');
    assert.equal(reply.events[0].attendees[0].partstat, 'ACCEPTED');
    assert.match(m.attachments[0].content, /COMMENT:See you there\./);

    // Moved by the organizer: the event moves, and asks again.
    assert.equal(calendar.takeMail({ from: 'someone@else.example', text: invitation({ sequence: 1, start: '20261028T113000Z' }) }), null, 'only the organizer moves it');
    const moved = calendar.takeMail({ from: 'sam@other.example', text: invitation({ sequence: 1, start: '20261028T113000Z' }) });
    assert.equal(moved.kind, 'update');
    const kept = calendar.get({ id: moved.eventId });
    assert.equal(kept.start.at, Date.UTC(2026, 9, 28, 11, 30));
    assert.equal(kept.attendees[0].partstat, 'NEEDS-ACTION');
    assert.equal(calendar.takeMail({ from: 'sam@other.example', text: invitation({ sequence: 1, start: '20261028T113000Z' }) }), null, 'the same change twice is one change');

    const off = calendar.takeMail({ from: 'sam@other.example', text: invitation({ sequence: 2, method: 'CANCEL', status: 'CANCELLED' }) });
    assert.equal(off.kind, 'cancel');
    assert.equal(calendar.get({ id: off.eventId }).status, 'CANCELLED');
  } finally {
    done();
  }
});

test('with a Mail that has no sendFrom yet, the invitation goes through its send, the file on disk', async () => {
  const { calendar, mail, done } = setup({ withSendFrom: false });
  try {
    const cal = calendar.calendars()[0];
    const saved = calendar.save({ calendarId: cal.id, event: { ...meeting(), organizer: null } });
    await calendar.sendInvitation({ id: saved.id });
    const m = mail.sent[0];
    assert.equal(m.via, 'send');
    assert.equal(m.accountId, 'acc-1');
    assert.equal(m.to, 'kim@example.org, sam@example.org');
    assert.match(icsOf(m), /METHOD:REQUEST/);
    assert.equal(readScheduling(icsOf(m)).events[0].uid, saved.uid);
  } finally {
    done();
  }
});

test('an invitation someone else organizes is not sent from here, and with no account there is nothing to send from', async () => {
  const { calendar, done } = setup();
  try {
    const cal = calendar.calendars()[0];
    const theirs = calendar.save({ calendarId: cal.id, event: { ...meeting(), organizer: { email: 'sam@other.example' } } });
    await assert.rejects(calendar.sendInvitation({ id: theirs.id }), /Only the organizer/);
  } finally {
    done();
  }
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-itip-'));
  try {
    const calendar = createCalendarService({ stores: { dir }, broadcast: null, mail: null });
    const saved = calendar.save({ calendarId: calendar.calendars()[0].id, event: { ...meeting(), organizer: null } });
    assert.equal(calendar.sendingAccount(), null);
    await assert.rejects(calendar.sendInvitation({ id: saved.id }), /no Mail account/);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
