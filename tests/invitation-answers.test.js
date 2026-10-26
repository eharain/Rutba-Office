// An invitation answered is remembered: opened again (Mail's strip reads the
// file, which still says NEEDS-ACTION), it says what was answered, until a
// later version of it asks again. The reply and invitation files come with
// their sizes for the message they go on; and the unsubscribe question names
// a list by its name, not its raw List-Id.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createCalendarService } from '../apps/desktop/main/calendar.js';
import { unsubscribeOffer, listName } from '../apps/desktop/main/mail-insight.js';

const invitation = ({ sequence = 0, uid = 'lunch@other.example' } = {}) => [
  'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Other//EN', 'METHOD:REQUEST', 'BEGIN:VEVENT', `UID:${uid}`, `SEQUENCE:${sequence}`,
  'DTSTAMP:20261020T090000Z', 'DTSTART:20261028T120000Z', 'DTEND:20261028T130000Z', 'SUMMARY:Lunch',
  'ORGANIZER;CN=Sam:mailto:sam@other.example',
  'ATTENDEE;CN=Me;PARTSTAT=NEEDS-ACTION;RSVP=TRUE:mailto:me@example.com',
  'END:VEVENT', 'END:VCALENDAR', '',
].join('\r\n');

function setup() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-invite-'));
  const calendar = createCalendarService({ stores: { dir }, broadcast: null, mail: { accounts: () => [{ email: 'me@example.com' }] } });
  return { dir, calendar };
}

test('an answer is remembered when the invitation is opened again, and a later version asks again', async () => {
  const { dir, calendar } = setup();
  try {
    assert.equal(calendar.openFile({ text: invitation() }).invitation.me.partstat, 'NEEDS-ACTION');
    const r = calendar.respond({ text: invitation(), partstat: 'TENTATIVE' });
    assert.equal(calendar.openFile({ text: invitation() }).invitation.me.partstat, 'TENTATIVE');
    assert.equal(r.replySize, fs.statSync(r.replyPath).size, 'the reply file\'s size, for the message it goes on');
    assert.ok(r.replySize > 0);
    assert.equal(calendar.openFile({ text: invitation({ sequence: 1 }) }).invitation.me.partstat, 'NEEDS-ACTION', 'a changed meeting asks again');

    // Declined: the event goes, the answer stays.
    calendar.respond({ text: invitation({ uid: 'review@other.example' }), partstat: 'DECLINED' });
    assert.equal(calendar.openFile({ text: invitation({ uid: 'review@other.example' }) }).invitation.me.partstat, 'DECLINED');

    // Kept by the next run.
    await new Promise((resolve) => setTimeout(resolve, 600));
    const again = createCalendarService({ stores: { dir }, broadcast: null, mail: { accounts: () => [{ email: 'me@example.com' }] } });
    assert.equal(again.openFile({ text: invitation() }).invitation.me.partstat, 'TENTATIVE');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('an answer given before answers were remembered is read from the event kept', () => {
  const { dir, calendar } = setup();
  try {
    calendar.respond({ text: invitation(), partstat: 'ACCEPTED' });
    delete calendar._dav.state().answers;
    assert.equal(calendar.openFile({ text: invitation() }).invitation.me.partstat, 'ACCEPTED');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('with no mail account among the people invited, the answer given here is still the one shown', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-invite-'));
  try {
    const calendar = createCalendarService({ stores: { dir }, broadcast: null, mail: { accounts: () => [] } });
    assert.equal(calendar.openFile({ text: invitation() }).invitation.me, null);
    calendar.respond({ text: invitation(), partstat: 'ACCEPTED' });
    assert.deepEqual(calendar.openFile({ text: invitation() }).invitation.me, { email: 'me@example.com', partstat: 'ACCEPTED' }, 'the line it was answered for');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('an invitation written for people invited comes with its size', () => {
  const { dir, calendar } = setup();
  try {
    const at = (ms) => ({ date: new Date(ms).toISOString().slice(0, 10), time: new Date(ms).toISOString().slice(11, 19), tzid: null, utc: true, allDay: false, at: ms });
    const saved = calendar.save({ event: { summary: 'Plan', start: at(Date.UTC(2026, 10, 2, 9)), end: at(Date.UTC(2026, 10, 2, 10)), attendees: [{ email: 'kim@example.com' }] } });
    const file = calendar.invitationFile({ id: saved.id });
    assert.equal(file.size, fs.statSync(file.path).size);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('the unsubscribe question names the list by its name', () => {
  assert.equal(listName('"Weekly News" <news.example.com>'), 'Weekly News');
  assert.equal(listName('Weekly News <news.example.com>'), 'Weekly News');
  assert.equal(listName('<news.example.com>'), 'news.example.com');
  assert.equal(listName(''), null);
  const offer = unsubscribeOffer({ headers: [{ key: 'List-Unsubscribe', value: '<https://news.example.com/u?x=1>' }, { key: 'List-Id', value: '"Weekly News" <news.example.com>' }] });
  assert.equal(offer.list, 'Weekly News');
});
