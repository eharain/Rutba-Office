// Rooms and resources as contacts, booked by inviting their mailbox: a room's
// card keeps its kind, seats, place and equipment through a vCard; Contacts
// lists the bookable ones; the editor's "Add a room" sorts them by whether they
// are free then, and then by whether they seat the people invited; and a room
// invited goes out with CUTYPE=ROOM.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { readVCards, writeVCard, isBookable } from '../packages/contacts/src/index.js';
import { writeRequest, unfold, localToUtc } from '../packages/calendar/src/index.js';
import { createContactsService } from '../apps/desktop/main/contacts.js';
import { createCalendarService } from '../apps/desktop/main/calendar.js';
import { roomsFor, busyInCalendars } from '../apps/desktop/main/calendar-rooms.js';

const scratch = (t) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-rooms-'));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  return dir;
};

const room = { name: { given: 'Board room' }, kind: 'room', emails: [{ value: 'board@rooms.example' }], room: { capacity: 12, where: 'Floor 3', equipment: ['Screen', 'Phone; speaker'] } };

test('a room card keeps its kind, seats, place and equipment through a vCard, 3.0 and 4.0', () => {
  for (const version of ['3.0', '4.0']) {
    const text = writeVCard(room, { version });
    if (version === '4.0') assert.match(text, /KIND:location/);
    assert.match(text, /X-RUTBA-KIND:room/);
    const [back] = readVCards(text);
    assert.equal(back.kind, 'room');
    assert.deepEqual(back.room, { capacity: 12, where: 'Floor 3', equipment: ['Screen', 'Phone', 'speaker'] });
    assert.ok(isBookable(back));
  }
  const [resource] = readVCards(['BEGIN:VCARD', 'VERSION:4.0', 'FN:Projector', 'KIND:x-resource', 'EMAIL:projector@rooms.example', 'END:VCARD', ''].join('\r\n'));
  assert.equal(resource.kind, 'resource');
  assert.ok(isBookable(resource));
  const [person] = readVCards(['BEGIN:VCARD', 'VERSION:3.0', 'FN:Ann', 'EMAIL:ann@example.com', 'END:VCARD', ''].join('\r\n'));
  assert.equal(person.kind, 'individual');
  assert.equal(person.room, null);
  assert.ok(!isBookable(person));
  // A room with no mailbox cannot be booked: nothing to invite.
  assert.ok(!isBookable({ ...room, emails: [] }));
});

test('Contacts lists the bookable cards with their mailbox, and the kind filter finds them', (t) => {
  const contacts = createContactsService({ stores: { dir: scratch(t) }, broadcast: null });
  contacts.save({ contact: room });
  contacts.save({ contact: { name: { given: 'Projector' }, kind: 'resource', emails: [{ value: 'projector@rooms.example' }], room: { capacity: null, where: null, equipment: [] } } });
  contacts.save({ contact: { name: { given: 'Ann', family: 'Lee' }, emails: [{ value: 'ann@example.com' }] } });
  const listed = contacts.bookable();
  assert.deepEqual(listed.map((r) => [r.kind, r.name, r.email, r.capacity]), [['room', 'Board room', 'board@rooms.example', 12], ['resource', 'Projector', 'projector@rooms.example', null]]);
  assert.equal(contacts.list({ kind: 'bookable' }).length, 2);
  assert.equal(contacts.list({ kind: 'room' }).length, 1);
  assert.equal(contacts.list({ query: 'floor 3' }).length, 1, 'found by where it is');
  assert.equal(contacts.list({ query: 'screen' }).length, 1, 'found by what it has');
  const offered = contacts.suggest({ query: 'board' });
  assert.equal(offered[0]?.kind, 'room', 'a suggestion says it is a room');
});

const H = 3600000;
const base = Date.UTC(2026, 10, 17, 9, 0);

test('the rooms free then come first, then those that seat everyone, smallest first', () => {
  const bookable = [
    { id: 'a', kind: 'room', name: 'Atrium', email: 'atrium@r', capacity: 40 },
    { id: 'b', kind: 'room', name: 'Booth', email: 'booth@r', capacity: 2 },
    { id: 'c', kind: 'room', name: 'Corner', email: 'corner@r', capacity: 6 },
    { id: 'd', kind: 'resource', name: 'Display', email: 'display@r', capacity: null },
    { id: 'e', kind: 'room', name: 'East', email: 'east@r', capacity: 8 },
  ];
  // A source with the whole day of every room but East, which is busy at the time.
  const sources = [{ busy: (email) => (email === 'east@r' ? { busy: [{ start: base, end: base + H }], whole: false, source: 'calendar' } : email === 'atrium@r' ? null : { busy: [], whole: true, source: 'server' }) }];
  const rows = roomsFor({ bookable, sources, start: base, end: base + H, people: 5 });
  assert.deepEqual(rows.map((r) => [r.name, r.free]), [
    ['Corner', true], // free, seats five, the smallest that does
    ['Display', true], // free, seats not known: counted as fitting, after those with a number
    ['Booth', true], // free but too small
    ['Atrium', null], // nothing knows its day
    ['East', false], // busy then
  ]);
  assert.deepEqual(rows[4].sources, ['calendar']);
});

test("the meetings here that have a room on are its busy time, leaving out the one being edited and those it declined", () => {
  const ev = (uid, h, partstat = 'ACCEPTED', over = {}) => ({ uid, summary: uid, start: { at: base + h * H, allDay: false }, end: { at: base + (h + 1) * H, allDay: false }, attendees: [{ email: 'Board@Rooms.example', partstat }], ...over });
  const calendars = [{ events: [ev('one', 0), ev('two', 1), ev('three', 3, 'DECLINED'), ev('four', 4, 'ACCEPTED', { status: 'CANCELLED' }), ev('five', 5, 'ACCEPTED', { attendees: [] })] }];
  const spans = (except) => busyInCalendars(calendars, 'board@rooms.example', base - H, base + 10 * H, except).map((b) => [b.start, b.end]);
  assert.deepEqual(spans(), [[base, base + 2 * H]]);
  assert.deepEqual(spans('two'), [[base, base + H]]);
});

test('calendar.rooms reads Contacts and the calendars, and a room invited goes out with CUTYPE=ROOM', (t) => {
  const dir = scratch(t);
  const contacts = createContactsService({ stores: { dir }, broadcast: null });
  contacts.save({ contact: room });
  contacts.save({ contact: { name: { given: 'Small room' }, kind: 'room', emails: [{ value: 'small@rooms.example' }], room: { capacity: 4, where: null, equipment: [] } } });
  const calendar = createCalendarService({ stores: { dir }, broadcast: null, contacts: () => contacts });
  const cal = calendar.calendars()[0];
  const zone = 'Europe/London';
  const at = (H, M = 0) => ({ date: '2026-11-17', time: `${String(H).padStart(2, '0')}:${String(M).padStart(2, '0')}:00`, tzid: zone, utc: false, allDay: false, at: localToUtc({ y: 2026, m: 11, d: 17, H, M }, zone) });
  calendar.save({ calendarId: cal.id, event: { summary: 'Board', start: at(10), end: at(11), attendees: [{ email: 'board@rooms.example', name: 'Board room', type: 'ROOM', partstat: 'ACCEPTED' }] } });
  const busyThen = calendar.rooms({ start: at(10, 30).at, end: at(11, 30).at, people: 3 });
  assert.deepEqual(busyThen.map((r) => [r.name, r.free]), [['Small room', null], ['Board room', false]]);
  const later = calendar.rooms({ start: at(14).at, end: at(15).at, people: 6 });
  assert.deepEqual(later.map((r) => r.name), ['Board room', 'Small room'], 'the one that seats six first');

  const text = writeRequest([{ uid: 'r@x', sequence: 0, summary: 'Review', start: at(14), end: at(15), organizer: { email: 'me@example.com' }, attendees: [{ email: 'board@rooms.example', name: 'Board room', type: 'ROOM', role: 'NON-PARTICIPANT' }] }], {});
  assert.match(unfold(text).join('\n'), /ATTENDEE;CN=Board room;ROLE=NON-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CUTYPE=ROOM:mailto:board@rooms\.example/);
});
