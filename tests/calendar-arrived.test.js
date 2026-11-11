// Mail heard arriving (Mail's onArrived, 2.25) instead of the Inboxes read
// every 90 seconds: an answer goes onto its line and a room's mailbox is
// answered as the mail comes, from the calendar text the arrival carries;
// the timed read passes by what an arrival took, and the two never take a
// message twice. A Mail without onArrived (older than 2.25) keeps the timer.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { readCalendar, writeRequest, writeReply, readScheduling, dateTimeAt } from '../packages/calendar/src/index.js';
import { createCalendarService } from '../apps/desktop/main/calendar.js';
import { createMailLink } from '../apps/desktop/main/calendar-mail.js';
import { register } from '../apps/desktop/main/calendar-pack.js';

const H = 3600 * 1000;
const zone = 'Europe/London';
const day = (() => {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + 3);
  return d.getTime();
})();

/** A Mail with two accounts, a person's and a room's, that tells of what arrives as Mail 2.25 does. */
function fakeMail({ tells = true } = {}) {
  const sent = [];
  const boxes = { me: [], room: [] };
  const listeners = new Set();
  const reads = { messages: 0, message: 0, attachmentText: 0 };
  let n = 0;
  const service = {
    accounts: () => [{ id: 'me', email: 'me@example.com', name: 'Me' }, { id: 'room', email: 'board@rooms.example', name: 'Board room' }],
    folders: () => [{ path: 'INBOX', role: 'inbox' }, { path: 'Archive', role: null }],
    messages: ({ accountId, limit }) => { reads.messages += 1; return { rows: boxes[accountId].slice(-limit).reverse() }; },
    message: ({ accountId, id }) => { reads.message += 1; return boxes[accountId].find((m) => m.id === id)?.message || null; },
    attachmentText: ({ accountId, id }) => { reads.attachmentText += 1; return boxes[accountId].find((m) => m.id === id)?.part || null; },
    send: async () => ({ refused: [] }),
    sendFrom: async (msg) => { sent.push(msg); return { messageId: `s${sent.length}`, refused: [] }; },
  };
  if (tells) service.onArrived = (listener) => { listeners.add(listener); return () => listeners.delete(listener); };
  /** Mail arrived: kept in the box, and told as Mail tells it (`carry` false: the calendar text too large to carry). */
  const arrive = (box, from, { text = null, subject = '', body = '', folder = 'INBOX', carry = true, tell = true } = {}) => {
    n += 1;
    const id = `in-${n}`;
    const date = new Date().toISOString();
    boxes[box].push(text
      ? { id, from: { address: from }, subject, date, hasAttachments: true, attachmentCount: 1, message: { subject, attachments: [{ type: 'text/calendar', filename: 'invite.ics' }] }, part: { name: 'invite.ics', type: 'text/calendar', text } }
      : { id, from: { address: from }, subject, date, hasAttachments: false, attachmentCount: 0, message: { subject, text: body, from: [{ name: '', address: from }], date, attachments: [] } });
    const event = {
      accountId: box, folder, ids: [id], count: 1, at: date, filedByRules: false,
      invitations: text ? [{ id, index: 0, messageId: `<${id}@x>`, subject, from: { name: '', address: from }, method: /METHOD:(\w+)/.exec(text)?.[1] || null, text: carry ? text : null }] : [],
    };
    if (tell) for (const l of [...listeners]) l(event);
    return { id, event };
  };
  return { service, sent, arrive, reads, listeners };
}

function setup(t, opts) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-arrived-'));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  const mail = fakeMail(opts);
  const contacts = { bookable: () => [{ id: 'c1', kind: 'room', name: 'Board room', email: 'board@rooms.example', capacity: 10, where: null, equipment: [] }] };
  const calendar = createCalendarService({ stores: { dir }, broadcast: null, mail: createMailLink(() => mail.service), contacts: () => contacts });
  return { calendar, mail, dir };
}

const settle = () => new Promise((r) => setImmediate(r));
const until = async (fn, what) => {
  for (let i = 0; i < 100; i++) {
    if (fn()) return;
    await new Promise((r) => setTimeout(r, 5));
  }
  assert.fail(`waited for ${what}`);
};

async function invited(calendar, mail) {
  const cal = calendar.calendars()[0];
  const saved = calendar.save({ calendarId: cal.id, event: { summary: 'Plan', start: dateTimeAt(day + 10 * H, zone), end: dateTimeAt(day + 11 * H, zone), attendees: [{ email: 'kim@example.org' }, { email: 'sam@example.org' }] } });
  await calendar.sendInvitation({ id: saved.id });
  return { saved, request: readCalendar(mail.sent.at(-1).attachments[0].content).events[0] };
}
const lineOf = (calendar, id, email) => calendar.get({ id }).attendees.find((a) => a.email === email)?.partstat;

test("the adapter hears Mail's arrivals as Mail tells them, and says when this Mail cannot", () => {
  const mail = fakeMail();
  const link = createMailLink(() => mail.service);
  const heard = [];
  const stop = link.onArrived((e) => heard.push(e));
  assert.equal(typeof stop, 'function');
  mail.arrive('me', 'kim@example.org', { text: 'BEGIN:VCALENDAR\r\nMETHOD:reply\r\nEND:VCALENDAR\r\n', subject: 'Accepted: Plan' });
  assert.equal(heard.length, 1);
  assert.deepEqual(heard[0].invitations[0], { id: 'in-1', subject: 'Accepted: Plan', from: { name: '', address: 'kim@example.org' }, method: 'REPLY', text: 'BEGIN:VCALENDAR\r\nMETHOD:reply\r\nEND:VCALENDAR\r\n' });
  assert.deepEqual(heard[0].ids, ['in-1']);
  mail.arrive('me', 'kim@example.org', { text: 'BEGIN:VCALENDAR\r\nEND:VCALENDAR\r\n', carry: false });
  assert.equal(heard[1].invitations[0].text, null, 'a part too large to carry is read from Mail');
  stop();
  mail.arrive('me', 'kim@example.org', { body: 'hello' });
  assert.equal(heard.length, 2, 'stopped');
  assert.equal(createMailLink(() => fakeMail({ tells: false }).service).onArrived(() => {}), null, 'a Mail older than 2.25');
  assert.equal(createMailLink(() => null).onArrived(() => {}), null, 'no Mail');
  assert.equal(createMailLink(() => ({ onArrived: () => { throw new Error('x'); } })).onArrived(() => {}), null);
});

test('an answer heard arriving goes onto its line at once, and the timed read passes it by', async (t) => {
  const { calendar, mail } = setup(t);
  const { saved, request } = await invited(calendar, mail);
  assert.deepEqual(calendar.followMail(), { following: true });
  assert.deepEqual(calendar.followMail(), { following: true }, 'followed once');
  assert.equal(mail.listeners.size, 1);
  const before = mail.reads.attachmentText;
  mail.arrive('me', 'sam@example.org', { text: writeReply(request, { email: 'sam@example.org', partstat: 'TENTATIVE' }) });
  await until(() => lineOf(calendar, saved.id, 'sam@example.org') === 'TENTATIVE', "Sam's answer");
  assert.equal(mail.reads.attachmentText, before, 'taken from the text the arrival carried');
  assert.deepEqual(await calendar.checkMail(), [], 'the timed read does not take it again');

  // A part too large to carry is read from Mail.
  mail.arrive('me', 'kim@example.org', { text: writeReply(request, { email: 'kim@example.org', partstat: 'ACCEPTED' }), carry: false });
  await until(() => lineOf(calendar, saved.id, 'kim@example.org') === 'ACCEPTED', "Kim's answer");
  assert.equal(mail.reads.attachmentText, before + 1);

  // A plain message to the person: nothing to read; a folder other than the Inbox: left to the Inbox's own arrival.
  const reads = { ...mail.reads };
  mail.arrive('me', 'lee@example.net', { body: 'Lunch?' });
  mail.arrive('me', 'sam@example.org', { text: writeReply(request, { email: 'sam@example.org', partstat: 'DECLINED' }), folder: 'Archive' });
  await settle();
  assert.equal(mail.reads.message, reads.message, "a plain message to the person is not read");
  assert.equal(lineOf(calendar, saved.id, 'sam@example.org') === 'TENTATIVE', true);

  // Followed no longer: an arrival changes nothing until the timed read.
  assert.deepEqual(calendar.followMail({ on: false }), { following: false });
  assert.equal(mail.listeners.size, 0);
});

test("a room's mailbox is answered as the mail comes: an invitation booked, a plain enquiry read and answered", async (t) => {
  const { calendar, mail } = setup(t);
  calendar.setRoomMailbox({ accountId: 'room', room: true });
  calendar.followMail();
  const meeting = { uid: 'one@x', sequence: 0, summary: 'Meeting one', start: dateTimeAt(day + 10 * H, zone), end: dateTimeAt(day + 11 * H, zone), organizer: { email: 'kim@example.org', name: 'Kim' }, attendees: [{ email: 'board@rooms.example', name: 'Board room', type: 'ROOM', role: 'NON-PARTICIPANT' }] };
  mail.arrive('room', 'kim@example.org', { text: writeRequest([meeting], {}), subject: 'Invitation: Meeting one' });
  await until(() => mail.sent.length === 1, 'the room to accept');
  assert.equal(mail.sent[0].account, 'room');
  assert.equal(readScheduling(mail.sent[0].attachments[0].content).events[0].attendees[0].partstat, 'ACCEPTED');
  assert.equal(calendar.roomMailboxes().find((a) => a.id === 'room').bookings, 1);
  assert.equal(calendar.events({ from: day, to: day + 24 * H }).length, 0, "not taken into the person's own calendar");

  const when = new Date(day);
  mail.arrive('room', 'sam@example.org', { subject: 'Board room', body: `Is the board room free on ${when.getFullYear()}-${String(when.getMonth() + 1).padStart(2, '0')}-${String(when.getDate()).padStart(2, '0')}?` });
  await until(() => mail.sent.length === 2, 'the enquiry answered');
  assert.deepEqual(mail.sent[1].to, ['sam@example.org']);
  assert.deepEqual(await calendar.checkMail(), [], 'nothing left for the timed read');
});

test('an arrival and the timed read at the same moment take a message once', async (t) => {
  const { calendar, mail } = setup(t);
  const { saved, request } = await invited(calendar, mail);
  calendar.followMail();
  // In the box before the read begins, and told while it runs.
  const { event } = mail.arrive('me', 'sam@example.org', { text: writeReply(request, { email: 'sam@example.org', partstat: 'ACCEPTED' }), tell: false });
  const [read, heard] = await Promise.all([calendar.checkMail(), calendar.arrived(event)]);
  assert.equal(read.length + heard.length, 1);
  assert.equal(lineOf(calendar, saved.id, 'sam@example.org'), 'ACCEPTED');
});

test('the pack follows Mail when it can and reads the Inboxes on the timer only when it cannot', async (t) => {
  const was = { verify: process.env.RUTBA_OFFICE_VERIFY_APPS, smoke: process.env.RUTBA_OFFICE_SMOKE };
  delete process.env.RUTBA_OFFICE_VERIFY_APPS;
  delete process.env.RUTBA_OFFICE_SMOKE;
  t.after(() => {
    for (const [k, v] of [['RUTBA_OFFICE_VERIFY_APPS', was.verify], ['RUTBA_OFFICE_SMOKE', was.smoke]]) if (v === undefined) delete process.env[k]; else process.env[k] = v;
  });
  t.mock.timers.enable({ apis: ['setTimeout', 'setInterval'] });
  for (const tells of [true, false]) {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-arrived-pack-'));
    t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
    const mail = fakeMail({ tells });
    register({ stores: { dir }, broadcast: null, packs: { service: (id) => (id === 'mail' ? mail.service : null) } });
    t.mock.timers.tick(20 * 1000);
    await settle();
    assert.equal(mail.listeners.size, tells ? 1 : 0);
    const first = mail.reads.messages;
    assert.ok(first > 0, 'one read for what came before');
    for (let i = 0; i < 3; i++) {
      t.mock.timers.tick(90 * 1000);
      await settle();
    }
    assert.equal(mail.reads.messages - first, tells ? 0 : 3 * 2, tells ? 'no timed reads while Mail tells' : 'both Inboxes read every 90 seconds');
  }
});
