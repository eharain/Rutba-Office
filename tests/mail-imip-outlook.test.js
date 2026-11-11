// An invitation as Outlook sends it (iMIP): the calendar is a third
// alternative beside the plain and HTML text, with no file name and no
// Content-Disposition. It is listed with the attachments all the same, so the
// reader shows its invitation card (Mail 2.25, the calendar session's ask).
import test from 'node:test';
import assert from 'node:assert/strict';
import { parseMessage } from '@rutba/mailbox/mime';

const ics = [
  'BEGIN:VCALENDAR', 'METHOD:REQUEST', 'PRODID:Microsoft Exchange Server 2010', 'VERSION:2.0',
  'BEGIN:VEVENT', 'ORGANIZER;CN=Amina Yusuf:mailto:amina@work.example',
  'ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=Sam:mailto:sam@checks.example',
  'SUMMARY;LANGUAGE=en-GB:Budget review', 'DTSTART:20261120T100000Z', 'DTEND:20261120T110000Z',
  'UID:040000008200E00074C5B7101A82E00800000000outlook', 'SEQUENCE:0', 'DTSTAMP:20261111T090000Z', 'END:VEVENT', 'END:VCALENDAR', '',
].join('\r\n');

const outlook = [
  'From: Amina Yusuf <amina@work.example>', 'To: Sam <sam@checks.example>', 'Subject: Budget review',
  'Date: Wed, 11 Nov 2026 09:00:00 +0000', 'Message-ID: <outlook-invite@work.example>', 'MIME-Version: 1.0',
  'Content-Type: multipart/alternative; boundary="_000_alt_"', '',
  '--_000_alt_', 'Content-Type: text/plain; charset="us-ascii"', '', 'Shall we go through the numbers?', '',
  '--_000_alt_', 'Content-Type: text/html; charset="us-ascii"', '', '<p>Shall we go through the numbers?</p>', '',
  '--_000_alt_', 'Content-Type: text/calendar; charset="utf-8"; method=REQUEST', 'Content-Transfer-Encoding: base64', '',
  Buffer.from(ics).toString('base64').replace(/.{76}/g, '$&\r\n'), '',
  '--_000_alt_--', '',
].join('\r\n');

test('an Outlook invitation: the calendar alternative with no file name is listed with the attachments', () => {
  const m = parseMessage(outlook);
  const cal = m.attachments.filter((a) => a.type === 'text/calendar');
  assert.equal(cal.length, 1, JSON.stringify(m.attachments.map((a) => ({ type: a.type, filename: a.filename }))));
  assert.match(cal[0].filename, /\.ics$/);
  assert.equal(cal[0].inline, false, 'not an inline picture: it is offered');
  assert.match(Buffer.from(cal[0].bytes).toString('utf8'), /METHOD:REQUEST[\s\S]*SUMMARY;LANGUAGE=en-GB:Budget review/);
  assert.equal(m.text, 'Shall we go through the numbers?', 'the words are still the plain alternative');
  assert.match(m.html, /go through the numbers/);
  assert.equal(m.hasAttachments, true);
});

test('a message whose whole body is the calendar is listed too', () => {
  const m = parseMessage([
    'From: a@work.example', 'To: b@checks.example', 'Subject: Invitation', 'MIME-Version: 1.0',
    'Content-Type: text/calendar; method=REQUEST; charset=utf-8', '', ics,
  ].join('\r\n'));
  assert.equal(m.attachments.filter((a) => a.type === 'text/calendar').length, 1);
});

test('ordinary text alternatives are still not attachments', () => {
  const m = parseMessage(outlook.replace(/--_000_alt_\r\nContent-Type: text\/calendar[\s\S]*?(?=--_000_alt_--)/, ''));
  assert.deepEqual(m.attachments, []);
});
