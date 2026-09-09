// iCalendar, read and written — the files Google, Outlook and Apple export,
// the invitations mail carries, and the recurrence rules people write.

import test from 'node:test';
import assert from 'node:assert/strict';
import {
  readCalendar, writeCalendar, writeReply, isInvitation, occurrences, describeRule,
  parseDuration, formatDuration, fold, unfold, parseLine, localToUtc, utcToLocal, ianaZone,
} from '../packages/calendar/src/index.js';

const CRLF = (s) => s.replace(/\n/g, '\r\n');

const GOOGLE_INVITE = CRLF(`BEGIN:VCALENDAR
PRODID:-//Google Inc//Google Calendar 70.9054//EN
VERSION:2.0
CALSCALE:GREGORIAN
METHOD:REQUEST
BEGIN:VTIMEZONE
TZID:Europe/London
BEGIN:DAYLIGHT
TZOFFSETFROM:+0000
TZOFFSETTO:+0100
TZNAME:BST
DTSTART:19700329T010000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:+0100
TZOFFSETTO:+0000
TZNAME:GMT
DTSTART:19701025T020000
RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU
END:STANDARD
END:VTIMEZONE
BEGIN:VEVENT
DTSTART;TZID=Europe/London:20260615T090000
DTEND;TZID=Europe/London:20260615T093000
DTSTAMP:20260601T101500Z
ORGANIZER;CN=Sam Patel:mailto:sam@example.org
UID:abc123@google.com
ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=
 TRUE;CN=Kim Lee;X-NUM-GUESTS=0:mailto:kim@example.com
ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=ACCEPTED;CN=Sam Pate
 l:mailto:sam@example.org
CREATED:20260601T100000Z
DESCRIPTION:Weekly sync\\, with the numbers.\\nBring the report.
LAST-MODIFIED:20260601T101500Z
LOCATION:Room 4\\, second floor
SEQUENCE:0
STATUS:CONFIRMED
SUMMARY:Pricing review
TRANSP:OPAQUE
BEGIN:VALARM
ACTION:DISPLAY
DESCRIPTION:This is an event reminder
TRIGGER:-P0DT0H15M0S
END:VALARM
END:VEVENT
END:VCALENDAR
`);

test('a Google invitation reads whole: zone, people, folded lines, escapes, the alarm', () => {
  const cal = readCalendar(GOOGLE_INVITE);
  assert.equal(cal.method, 'REQUEST');
  assert.ok(isInvitation(cal));
  assert.equal(cal.timezones[0], 'Europe/London');
  const e = cal.events[0];
  assert.equal(e.uid, 'abc123@google.com');
  assert.equal(e.summary, 'Pricing review');
  assert.equal(e.description, 'Weekly sync, with the numbers.\nBring the report.');
  assert.equal(e.location, 'Room 4, second floor');
  assert.equal(e.start.tzid, 'Europe/London');
  assert.equal(e.start.time, '09:00:00');
  // 9am in London in June is 8am UTC.
  assert.equal(e.start.at, Date.UTC(2026, 5, 15, 8, 0, 0));
  assert.equal(e.end.at - e.start.at, 30 * 60000);
  assert.equal(e.organizer.email, 'sam@example.org');
  assert.equal(e.organizer.name, 'Sam Patel');
  assert.equal(e.attendees.length, 2);
  assert.equal(e.attendees[0].email, 'kim@example.com');
  assert.equal(e.attendees[0].name, 'Kim Lee', 'a line folded inside a parameter value');
  assert.equal(e.attendees[0].partstat, 'NEEDS-ACTION');
  assert.equal(e.attendees[0].rsvp, true);
  assert.equal(e.attendees[1].name, 'Sam Patel', 'a line folded inside a name');
  assert.equal(e.alarms[0].before, 15);
  assert.equal(e.status, 'CONFIRMED');
});

test('an Outlook event names its zone the Windows way, and the instant is still right', () => {
  const cal = readCalendar(CRLF(`BEGIN:VCALENDAR
PRODID:-//Microsoft Corporation//Outlook 16.0 MIMEDIR//EN
VERSION:2.0
METHOD:PUBLISH
BEGIN:VEVENT
UID:040000008200E00074C5B7101A82E008
DTSTART;TZID="Pacific Standard Time":20260120T140000
DTEND;TZID="Pacific Standard Time":20260120T150000
SUMMARY:Vendor call
DTSTAMP:20260110T000000Z
END:VEVENT
END:VCALENDAR
`));
  const e = cal.events[0];
  assert.equal(e.start.tzid, 'Pacific Standard Time');
  assert.equal(ianaZone(e.start.tzid), 'America/Los_Angeles');
  // 2pm Pacific in January is 10pm UTC.
  assert.equal(e.start.at, Date.UTC(2026, 0, 20, 22, 0, 0));
});

test('an all-day event ends the next day, and a DURATION stands in for a missing DTEND', () => {
  const cal = readCalendar(CRLF(`BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Apple Inc.//macOS 14.0//EN
BEGIN:VEVENT
UID:a1
DTSTART;VALUE=DATE:20260301
SUMMARY:Bank holiday
END:VEVENT
BEGIN:VEVENT
UID:a2
DTSTART:20260301T100000Z
DURATION:PT1H30M
SUMMARY:Call
END:VEVENT
END:VCALENDAR
`));
  const [day, call] = cal.events;
  assert.equal(day.allDay, true);
  assert.equal(day.start.date, '2026-03-01');
  assert.equal(day.end.date, '2026-03-02');
  assert.equal(call.allDay, false);
  assert.equal(call.end.at - call.start.at, 90 * 60000);
  assert.equal(parseDuration('P1DT2H30M'), (26 * 3600 + 1800) * 1000);
  assert.equal(parseDuration('-PT15M'), -15 * 60000);
  assert.equal(formatDuration(-15 * 60000), '-PT15M');
  assert.equal(formatDuration(90 * 60000), 'PT1H30M');
  assert.equal(formatDuration(2 * 86400000), 'P2D');
});

test('a weekly rule on two days with a count, an exception on one date, and a moved instance', () => {
  const cal = readCalendar(CRLF(`BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//x//x//EN
BEGIN:VEVENT
UID:w1
DTSTART;TZID=Europe/London:20260601T090000
DTEND;TZID=Europe/London:20260601T100000
RRULE:FREQ=WEEKLY;BYDAY=MO,WE;COUNT=6
EXDATE;TZID=Europe/London:20260608T090000
SUMMARY:Stand-up
END:VEVENT
BEGIN:VEVENT
UID:w1
RECURRENCE-ID;TZID=Europe/London:20260610T090000
DTSTART;TZID=Europe/London:20260610T150000
DTEND;TZID=Europe/London:20260610T160000
SUMMARY:Stand-up (moved)
END:VEVENT
END:VCALENDAR
`));
  const [rule, moved] = cal.events;
  const got = occurrences(rule, { from: Date.UTC(2026, 5, 1), to: Date.UTC(2026, 6, 1), siblings: [moved] });
  const days = got.map((o) => new Date(o.start).toISOString().slice(0, 16));
  // 1 Jun is a Monday. Six by count: 1, 3, 8 (removed), 10 (moved to 3pm), 15, 17.
  assert.deepEqual(days, ['2026-06-01T08:00', '2026-06-03T08:00', '2026-06-10T14:00', '2026-06-15T08:00', '2026-06-17T08:00']);
  assert.equal(got[2].event.summary, 'Stand-up (moved)');
  assert.equal(got[2].original, Date.UTC(2026, 5, 10, 8));
  assert.equal(describeRule(rule.rrule, rule.start), 'Every week on Monday, Wednesday, 6 times');
});

test('monthly on the last Friday, yearly on a date, and daily until a date across a clock change', () => {
  const lastFriday = readCalendar(CRLF(`BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:m1
DTSTART;TZID=Europe/London:20260130T170000
DTEND;TZID=Europe/London:20260130T173000
RRULE:FREQ=MONTHLY;BYDAY=-1FR
SUMMARY:Month end
END:VEVENT
END:VCALENDAR
`)).events[0];
  const fridays = occurrences(lastFriday, { from: Date.UTC(2026, 0, 1), to: Date.UTC(2026, 4, 1) }).map((o) => new Date(o.start).toISOString().slice(0, 10));
  assert.deepEqual(fridays, ['2026-01-30', '2026-02-27', '2026-03-27', '2026-04-24']);
  assert.equal(describeRule(lastFriday.rrule), 'Every month on the last Friday');

  const birthday = readCalendar(CRLF(`BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:y1
DTSTART;VALUE=DATE:20200229
RRULE:FREQ=YEARLY
SUMMARY:Leap birthday
END:VEVENT
END:VCALENDAR
`)).events[0];
  const years = occurrences(birthday, { from: Date.UTC(2020, 0, 1), to: Date.UTC(2029, 0, 1) }).map((o) => new Date(o.start).toISOString().slice(0, 10));
  assert.deepEqual(years, ['2020-02-29', '2024-02-29', '2028-02-29'], 'a 29 February birthday falls in leap years only');

  const daily = readCalendar(CRLF(`BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:d1
DTSTART;TZID=Europe/London:20260327T090000
DTEND;TZID=Europe/London:20260327T091500
RRULE:FREQ=DAILY;UNTIL=20260331T235959Z
SUMMARY:Nine o'clock
END:VEVENT
END:VCALENDAR
`)).events[0];
  const nine = occurrences(daily, {}).map((o) => new Date(o.start).toISOString().slice(0, 16));
  // The clocks go forward on 29 March 2026; 9am London is 9am UTC before and 8am UTC after.
  assert.deepEqual(nine, ['2026-03-27T09:00', '2026-03-28T09:00', '2026-03-29T08:00', '2026-03-30T08:00', '2026-03-31T08:00']);
});

test('written and read back: escapes, folding, zones, people, rules and alarms survive', () => {
  const cal = readCalendar(GOOGLE_INVITE);
  const e = { ...cal.events[0], rrule: { freq: 'WEEKLY', interval: 2, count: null, until: null, byDay: [{ n: 0, day: 1 }], byMonthDay: null, byMonth: null, bySetPos: null, wkst: 'MO' } };
  const text = writeCalendar({ events: [e], name: 'Work' });
  assert.ok(text.split('\r\n').every((l) => Buffer.byteLength(l, 'utf8') <= 75), 'every line is folded within 75 octets');
  const back = readCalendar(text);
  assert.equal(back.name, 'Work');
  const r = back.events[0];
  assert.equal(r.summary, e.summary);
  assert.equal(r.description, e.description);
  assert.equal(r.location, e.location);
  assert.equal(r.start.at, e.start.at);
  assert.equal(r.start.tzid, 'Europe/London');
  assert.equal(r.attendees.length, 2);
  assert.equal(r.attendees[0].name, 'Kim Lee');
  assert.equal(r.organizer.name, 'Sam Patel');
  assert.equal(r.alarms[0].before, 15);
  assert.equal(r.rrule.freq, 'WEEKLY');
  assert.equal(r.rrule.interval, 2);
  assert.deepEqual(r.rrule.byDay, [{ n: 0, day: 1 }]);
  assert.ok(text.includes('X-NUM-GUESTS') === false, 'unknown attendee parameters are not invented back');
});

test('a reply carries the answer and nothing the organizer does not read', () => {
  const cal = readCalendar(GOOGLE_INVITE);
  const reply = writeReply(cal.events[0], { email: 'kim@example.com', name: 'Kim Lee', partstat: 'accepted' });
  assert.ok(reply.includes('METHOD:REPLY'));
  assert.ok(reply.includes('UID:abc123@google.com'));
  assert.ok(reply.includes('ATTENDEE;CN=Kim Lee;PARTSTAT=ACCEPTED:mailto:kim@example.com'));
  assert.ok(reply.includes('ORGANIZER;CN=Sam Patel:mailto:sam@example.org'));
  assert.ok(!reply.includes('kim@example.com:'), 'only the one attendee, once');
  const back = readCalendar(reply);
  assert.equal(back.method, 'REPLY');
  assert.equal(back.events[0].attendees[0].partstat, 'ACCEPTED');
});

test('lines: folding never splits a character, and a quoted parameter keeps its colon', () => {
  const long = 'DESCRIPTION:' + 'é'.repeat(100);
  const folded = fold(long);
  assert.ok(folded.split('\r\n').every((l) => Buffer.byteLength(l, 'utf8') <= 75));
  assert.equal(unfold(folded)[0], long);
  const p = parseLine('ATTENDEE;CN="Lee, Kim: Ops";RSVP=TRUE:mailto:kim@example.com');
  assert.equal(p.params.CN[0], 'Lee, Kim: Ops');
  assert.equal(p.value, 'mailto:kim@example.com');
});

test('zones: a wall-clock time becomes the right instant either side of a clock change', () => {
  assert.equal(localToUtc({ y: 2026, m: 7, d: 1, H: 12 }, 'Europe/London'), Date.UTC(2026, 6, 1, 11));
  assert.equal(localToUtc({ y: 2026, m: 1, d: 1, H: 12 }, 'Europe/London'), Date.UTC(2026, 0, 1, 12));
  assert.equal(localToUtc({ y: 2026, m: 1, d: 1, H: 9 }, 'Asia/Karachi'), Date.UTC(2026, 0, 1, 4));
  assert.deepEqual(utcToLocal(Date.UTC(2026, 6, 1, 11), 'Europe/London'), { y: 2026, m: 7, d: 1, H: 12, M: 0, S: 0 });
  assert.equal(ianaZone('GMT Standard Time'), 'Europe/London');
  assert.equal(ianaZone('Nowhere/Land'), null);
});
