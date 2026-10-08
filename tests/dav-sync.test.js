/**
 * Calendar and contacts accounts: a CalDAV and CardDAV server (the small
 * one in fixtures) kept in step with the calendars and cards here — an
 * account added brings its calendars and cards, a change made here goes to
 * the server, a change there comes back, deletions go both ways, and when
 * both sides changed one event the server's copy wins.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { dateTimeAt } from '@rutba/calendar';
import { createCalendarService } from '../apps/desktop/main/calendar.js';
import { createContactsService } from '../apps/desktop/main/contacts.js';
import { createDavService } from '../apps/desktop/main/dav-sync.js';
import { startDavServer } from './fixtures/dav-server.js';

function memory() {
  const map = new Map();
  return { get: (k) => map.get(k), set: (k, v) => { map.set(k, v); }, delete: (k) => { map.delete(k); } };
}

function setup() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-dav-'));
  const stores = { dir, settings: memory(), secrets: memory() };
  const calendar = createCalendarService({ stores, broadcast: null });
  const contacts = createContactsService({ stores, broadcast: null });
  const dav = createDavService({ stores, calendar, contacts, timers: false });
  return { stores, calendar, contacts, dav };
}

const ics = (uid, summary, start = Date.UTC(2026, 9, 12, 9)) => [
  'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Other device//EN', 'BEGIN:VEVENT', `UID:${uid}`,
  `DTSTAMP:20261001T000000Z`, `DTSTART:${new Date(start).toISOString().replace(/[-:]/g, '').slice(0, 15)}Z`,
  `DTEND:${new Date(start + 3600000).toISOString().replace(/[-:]/g, '').slice(0, 15)}Z`, `SUMMARY:${summary}`, 'END:VEVENT', 'END:VCALENDAR', '',
].join('\r\n');
const vcf = (uid, name, email) => ['BEGIN:VCARD', 'VERSION:3.0', `UID:${uid}`, `FN:${name}`, `N:${name.split(' ').reverse().join(';')};;;`, `EMAIL:${email}`, 'END:VCARD', ''].join('\r\n');
const OCT = { from: Date.UTC(2026, 9, 1), to: Date.UTC(2026, 10, 1) };

test('An account added brings its calendars and cards; a change made here goes to the server, and one there comes back', async () => {
  const server = await startDavServer();
  const { calendar, contacts, dav } = setup();
  try {
    const work = server.calendarPath('work');
    server.put(`${work}standup.ics`, ics('standup@other', 'Stand-up'));
    server.put(`${server.bookPath('friends')}bo.vcf`, vcf('bo@other', 'Bo Lund', 'bo@example.com'));
    const added = await dav.add({ url: server.url, user: 'ann', password: 'secret' });
    assert.equal(added.calendars, 1);
    assert.equal(added.books, 1);
    const cal = calendar.calendars().find((c) => c.name === 'Work');
    assert.ok(cal && cal.account === added.id, 'the server\'s calendar is here, marked as the account\'s');
    assert.deepEqual(calendar.events(OCT).map((e) => e.summary), ['Stand-up']);
    assert.deepEqual(contacts.list().map((c) => c.display), ['Bo Lund']);

    // Made here: sent on the next sync, under its UID.
    const made = calendar.save({ calendarId: cal.id, event: { summary: 'Review', start: dateTimeAt(Date.UTC(2026, 9, 14, 13)), end: dateTimeAt(Date.UTC(2026, 9, 14, 14)) } });
    let [r] = await dav.sync();
    assert.equal(r.sent, 1);
    const sent = [...server.collections.get(work).items.values()].map((i) => i.data).find((d) => d.includes('SUMMARY:Review'));
    assert.ok(sent && sent.includes(`UID:${made.uid}`));

    // Changed there: fetched on the next sync.
    server.put(`${work}standup.ics`, ics('standup@other', 'Stand-up, moved'));
    [r] = await dav.sync();
    assert.equal(r.received, 1);
    assert.deepEqual(calendar.events(OCT).map((e) => e.summary).sort(), ['Review', 'Stand-up, moved']);

    // Removed there: gone here.
    server.remove(`${work}standup.ics`);
    [r] = await dav.sync();
    assert.equal(r.removed, 1);
    assert.deepEqual(calendar.events(OCT).map((e) => e.summary), ['Review']);

    // Removed here: gone there.
    calendar.remove({ id: made.id });
    await dav.sync();
    assert.equal(server.collections.get(work).items.size, 0);

    // Nothing changed: one question each and no listing.
    const before = server.log.length;
    await dav.sync();
    assert.ok(!server.log.slice(before).some((l) => l.method === 'REPORT'), 'an unchanged collection is not listed again');
  } finally {
    await server.close();
  }
});

test('When an event changed here and on the server both, the server\'s copy wins and comes back', async () => {
  const server = await startDavServer();
  const { calendar, dav } = setup();
  try {
    const work = server.calendarPath('work');
    server.put(`${work}plan.ics`, ics('plan@other', 'Plan'));
    await dav.add({ url: server.url, user: 'ann', password: 'secret' });
    const [event] = calendar.events(OCT);
    const full = calendar.get({ id: event.id });
    calendar.save({ calendarId: full.calendarId, event: { ...full, summary: 'Plan (mine)' } });
    server.put(`${work}plan.ics`, ics('plan@other', 'Plan (theirs)'));
    const [r] = await dav.sync();
    assert.equal(r.conflicts, 1);
    assert.deepEqual(calendar.events(OCT).map((e) => e.summary), ['Plan (theirs)']);
    assert.match([...server.collections.get(work).items.values()][0].data, /SUMMARY:Plan \(theirs\)/, 'the server\'s copy was not overwritten');
  } finally {
    await server.close();
  }
});

test('Cards: a new card goes to the address book chosen for new cards, a change and a removal follow it there', async () => {
  const server = await startDavServer();
  const { contacts, dav } = setup();
  try {
    const book = server.bookPath('friends');
    await dav.add({ url: server.url, user: 'ann', password: 'secret' });
    const local = contacts.save({ contact: { name: { full: 'Kept Here', given: 'Kept', family: 'Here' }, emails: [{ value: 'kept@example.com' }] } });
    const [{ id }] = dav.books().books;
    dav.setDefaultBook({ id });
    const card = contacts.save({ contact: { name: { full: 'Ann Novak', given: 'Ann', family: 'Novak' }, emails: [{ value: 'ann@example.com' }], phones: [], addresses: [] } });
    await dav.sync();
    const items = () => [...server.collections.get(book).items.values()].map((i) => i.data);
    assert.equal(items().length, 1, 'the card made before a book was chosen stays on this computer');
    assert.match(items()[0], /FN:Ann Novak/);
    contacts.save({ contact: { ...contacts.get({ id: card.id }), title: 'Editor' } });
    await dav.sync();
    assert.match(items()[0], /TITLE:Editor/);
    contacts.remove({ id: card.id });
    await dav.sync();
    assert.equal(items().length, 0);
    assert.ok(contacts.get({ id: local.id }), 'the local card is untouched');
  } finally {
    await server.close();
  }
});

test('A wrong password is refused with why; removing an account takes away the calendars and cards it brought', async () => {
  const server = await startDavServer();
  const { calendar, contacts, dav } = setup();
  try {
    await assert.rejects(dav.add({ url: server.url, user: 'ann', password: 'wrong' }), /did not accept the user name and password/);
    assert.deepEqual(dav.accounts(), []);
    server.put(`${server.bookPath('friends')}bo.vcf`, vcf('bo@other', 'Bo Lund', 'bo@example.com'));
    const added = await dav.add({ url: server.url, user: 'ann', password: 'secret' });
    assert.equal(calendar.calendars().length, 2);
    assert.throws(() => calendar.removeCalendar({ id: calendar.calendars().find((c) => c.account).id }), /remove the account/);
    dav.remove({ id: added.id });
    assert.deepEqual(calendar.calendars().map((c) => c.name), ['My calendar']);
    assert.deepEqual(contacts.list(), []);
    await assert.rejects(setup().dav.add({ url: 'http://example.com', user: 'a', password: 'b' }), /HTTPS/, 'plain HTTP only for this computer');
  } finally {
    await server.close();
  }
});
