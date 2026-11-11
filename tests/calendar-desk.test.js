// A room's mailbox run by Office: an invitation with no clash is booked and
// accepted from the room's address; one that clashes is declined with the
// times free that day; a cancellation frees the time; a VFREEBUSY REQUEST
// gets the busy times; a plain "free on ...?" gets the free times; and what
// arrived before the mailbox was marked, or from a machine, is left alone.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { writeRequest, writeCancel, writeFreeBusyRequest, readScheduling, unfold, dateTimeAt, localZone } from '../packages/calendar/src/index.js';
import { createCalendarService } from '../apps/desktop/main/calendar.js';
import { createMailLink } from '../apps/desktop/main/calendar-mail.js';
import { askedDay } from '../apps/desktop/main/calendar-desk.js';

const H = 3600000;
const zone = localZone();
/** Three days ahead, at an hour on this computer's clock. */
const day = (() => {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + 3);
  return d.getTime();
})();
const at = (h) => day + h * H;

function fakeMail() {
  const sent = [];
  const boxes = { me: [], room: [] };
  let n = 0;
  const service = {
    accounts: () => [{ id: 'me', email: 'me@example.com', name: 'Me' }, { id: 'room', email: 'board@rooms.example', name: 'Board room' }],
    folders: () => [{ path: 'INBOX', role: 'inbox' }],
    messages: ({ accountId, limit }) => ({ rows: boxes[accountId].slice(-limit).reverse() }),
    message: ({ accountId, id }) => boxes[accountId].find((m) => m.id === id)?.message || null,
    attachmentText: ({ accountId, id }) => boxes[accountId].find((m) => m.id === id)?.part || null,
    send: async () => ({ refused: [] }),
    sendFrom: async (msg) => { sent.push(msg); return { messageId: `s${sent.length}`, refused: [] }; },
  };
  const date = (when) => new Date(when ?? Date.now()).toISOString();
  const deliver = (box, from, text, { when = null } = {}) => {
    n += 1;
    boxes[box].push({ id: `in-${n}`, from: { address: from }, date: date(when), hasAttachments: true, attachmentCount: 1, message: { attachments: [{ type: 'text/calendar', filename: 'invite.ics' }] }, part: { name: 'invite.ics', type: 'text/calendar', text } });
  };
  const write = (box, from, subject, text, { when = null } = {}) => {
    n += 1;
    boxes[box].push({ id: `in-${n}`, from: { address: from }, subject, date: date(when), hasAttachments: false, attachmentCount: 0, message: { subject, text, attachments: [] } });
  };
  return { service, sent, deliver, write };
}

function setup() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-desk-'));
  const mail = fakeMail();
  const contacts = { bookable: () => [{ id: 'c1', kind: 'room', name: 'Board room', email: 'board@rooms.example', capacity: 10, where: null, equipment: [] }] };
  const calendar = createCalendarService({ stores: { dir }, broadcast: null, mail: createMailLink(() => mail.service), contacts: () => contacts });
  return { calendar, mail, done: () => fs.rmSync(dir, { recursive: true, force: true }) };
}

const meeting = (uid, from, to, over = {}) => ({
  uid, sequence: 0, summary: `Meeting ${uid}`,
  start: dateTimeAt(at(from), zone), end: dateTimeAt(at(to), zone),
  organizer: { email: 'kim@example.org', name: 'Kim' },
  attendees: [{ email: 'board@rooms.example', name: 'Board room', type: 'ROOM', role: 'NON-PARTICIPANT' }, { email: 'sam@example.org' }],
  ...over,
});
const replyOf = (m) => readScheduling(m.attachments[0].content);

test('the day an enquiry asks about', () => {
  const now = new Date(2026, 10, 11, 9).getTime(); // a Wednesday
  const local = (y, m, d) => new Date(y, m - 1, d).getTime();
  assert.equal(askedDay('Is the board room free on 17 November?', now), local(2026, 11, 17));
  assert.equal(askedDay('free on Nov 17th, 2027?', now), local(2027, 11, 17));
  assert.equal(askedDay('available 2026-12-01', now), local(2026, 12, 1));
  assert.equal(askedDay('free on 3 March?', now), local(2027, 3, 3), 'a date passed this year is next year');
  assert.equal(askedDay('free 5/12?', now), local(2026, 12, 5), 'day first');
  assert.equal(askedDay('free tomorrow?', now), local(2026, 11, 12));
  assert.equal(askedDay('is it free on Friday', now), local(2026, 11, 13));
  assert.equal(askedDay('Lunch on 17 November?', now), null, 'not an enquiry');
  assert.equal(askedDay('free soon?', now), null, 'no day');
});

test('an invitation with no clash is booked and accepted from the room; one that clashes is declined with the times free', async () => {
  const { calendar, mail, done } = setup();
  try {
    assert.equal(calendar.roomMailboxes().find((a) => a.id === 'room').room, false);
    calendar.setRoomMailbox({ accountId: 'room', room: true });
    mail.deliver('room', 'kim@example.org', writeRequest([meeting('one@x', 10, 11)], {}));
    const first = await calendar.checkMail();
    assert.equal(first[0]?.kind, 'room-booked');
    const yes = mail.sent[0];
    assert.equal(yes.account, 'room');
    assert.deepEqual(yes.to, ['kim@example.org']);
    assert.equal(yes.subject, 'Accepted: Meeting one@x');
    assert.match(yes.text, /Board room is booked for/);
    assert.match(yes.text, /while it is open/);
    const back = replyOf(yes);
    assert.equal(back.method, 'REPLY');
    assert.equal(back.events[0].attendees[0].partstat, 'ACCEPTED');
    assert.match(unfold(yes.attachments[0].content).join('\n'), /CUTYPE=ROOM/);
    assert.equal(calendar.roomMailboxes().find((a) => a.id === 'room').bookings, 1);
    // Not this person's own meeting: nothing new in their calendar.
    assert.equal(calendar.events({ from: day, to: day + 24 * H }).length, 0);

    // Half an hour into it: declined, with what is free that day.
    mail.deliver('room', 'kim@example.org', writeRequest([meeting('two@x', 10.5, 11.5)], {}));
    const second = await calendar.checkMail();
    assert.equal(second[0]?.kind, 'room-declined');
    assert.equal(second[0].free.length, 2, 'before ten and after eleven');
    const no = mail.sent[1];
    assert.equal(no.subject, 'Declined: Meeting two@x');
    assert.match(no.text, /Board room is not free for/);
    assert.match(no.text, /Free that day: /);
    assert.ok(no.text.includes(second[0].free[0]));
    const declined = replyOf(no).events[0];
    assert.equal(declined.attendees[0].partstat, 'DECLINED');
    assert.match(unfold(no.attachments[0].content).join('\n'), /COMMENT:.*Free that day/);

    // The editor's Add a room knows the room's whole day now.
    const busy = calendar.rooms({ start: at(10), end: at(11) });
    const free = calendar.rooms({ start: at(14), end: at(15) });
    assert.equal(busy[0].free, false);
    assert.equal(free[0].free, true);
    assert.ok(free[0].sources.includes('room-mailbox'));

    // Cancelled by Kim: the time is free again, and the next ask is booked.
    mail.deliver('room', 'kim@example.org', writeCancel([meeting('one@x', 10, 11, { sequence: 1 })], {}));
    const third = await calendar.checkMail();
    assert.equal(third[0]?.kind, 'room-cancelled');
    assert.equal(calendar.roomMailboxes().find((a) => a.id === 'room').bookings, 0);
    mail.deliver('room', 'kim@example.org', writeRequest([meeting('three@x', 10.5, 11.5)], {}));
    assert.equal((await calendar.checkMail())[0]?.kind, 'room-booked');
  } finally {
    done();
  }
});

test('an enquiry is answered: VFREEBUSY with the busy times, a plain "free on" with the free times', async () => {
  const { calendar, mail, done } = setup();
  try {
    calendar.setRoomMailbox({ accountId: 'room', room: true });
    mail.deliver('room', 'kim@example.org', writeRequest([meeting('one@x', 10, 11)], {}));
    await calendar.checkMail();

    mail.deliver('room', 'sam@example.org', writeFreeBusyRequest({ uid: 'fb-1', organizer: { email: 'sam@example.org' }, attendees: [{ email: 'board@rooms.example' }], start: day, end: day + 24 * H }));
    const fb = await calendar.checkMail();
    assert.equal(fb[0]?.kind, 'room-freebusy');
    const answer = mail.sent.at(-1);
    assert.deepEqual(answer.to, ['sam@example.org']);
    const read = readScheduling(answer.attachments[0].content);
    assert.equal(read.method, 'REPLY');
    assert.deepEqual(read.freebusy[0].busy.map((b) => [b.start, b.end]), [[at(10), at(11)]]);
    assert.doesNotMatch(answer.attachments[0].content, /Meeting one/, 'busy times only, never what fills them');

    const asked = new Date(day).toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
    mail.write('room', 'sam@example.org', 'Room?', `Hello, is the board room free on ${asked}?`);
    const plain = await calendar.checkMail();
    assert.equal(plain[0]?.kind, 'room-enquiry');
    assert.equal(plain[0].day, day);
    assert.equal(plain[0].free.length, 2);
    const words = mail.sent.at(-1);
    assert.equal(words.subject, 'Free times: Board room');
    assert.match(words.text, /Board room, .*: free /);
    assert.ok(words.text.includes(plain[0].free[1]));
    assert.equal(words.attachments.length, 0);

    // Mail with no question in it, from a machine, or from before the mailbox was marked: left alone.
    mail.write('room', 'sam@example.org', 'Thanks', 'Thank you for the room.');
    mail.write('room', 'no-reply@example.org', 'Free on 17 November?', 'free?');
    mail.deliver('room', 'kim@example.org', writeRequest([meeting('old@x', 14, 15)], {}), { when: Date.now() - 48 * H });
    const before = mail.sent.length;
    assert.deepEqual(await calendar.checkMail(), []);
    assert.equal(mail.sent.length, before);
  } finally {
    done();
  }
});

test("a room's Inbox is not this person's: its address is not theirs, and its invitations stay out of their calendar", async () => {
  const { calendar, mail, done } = setup();
  try {
    calendar.setRoomMailbox({ accountId: 'room', room: true });
    calendar.setRoomMailbox({ accountId: 'room', room: false });
    // Unmarked, the room's Inbox is an Inbox like any other again; marked, it is the room's.
    calendar.setRoomMailbox({ accountId: 'room', room: true });
    const cal = calendar.calendars()[0];
    const mine = calendar.save({ calendarId: cal.id, event: meeting('mine@x', 9, 10, { organizer: null, attendees: [{ email: 'board@rooms.example', name: 'Board room', type: 'ROOM' }] }) });
    const sent = await calendar.sendInvitation({ id: mine.id });
    assert.equal(sent.from, 'me@example.com', 'sent from the person, not the room');
    // The room takes the invitation from this person, as from anyone.
    mail.deliver('room', 'me@example.com', mail.sent[0].attachments[0].content);
    const r = await calendar.checkMail();
    assert.equal(r.find((x) => x.accountId === 'room')?.kind, 'room-booked');
    assert.deepEqual(mail.sent[1].to, ['me@example.com']);
  } finally {
    done();
  }
});
