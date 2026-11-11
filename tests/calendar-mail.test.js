// The calendar's Mail adapter against Mail's final sendFrom (2.25): the
// arguments pass Mail's own check (an attachment is content only, never a
// path; an account is an id or an address), the calendar file goes typed
// with its method for Mail to send as the invitation part, an answer goes
// from the address invited, and Mail's 'no-account' and 'cannot-send' come
// back said plainly. Mail's older send, which takes an id and a file, still
// works with an address.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { readCalendar, writeRequest, localToUtc } from '../packages/calendar/src/index.js';
import { createCalendarService } from '../apps/desktop/main/calendar.js';
import { createMailLink } from '../apps/desktop/main/calendar-mail.js';

const isCalendarPart = (a) => /^text\/calendar/i.test(String(a?.contentType || '')) || /\.ics$/i.test(String(a?.filename || ''));

// Mail's argument check for an attachment, as office-mail ceabfdb2 has it (main/mail.js, handedAttachment).
function handedAttachment(a, n) {
  if (!a || typeof a !== 'object') throw new Error('An attachment is a name, its content and its type.');
  if (a.path || a.href) throw new Error('An attachment is sent from its content, not from a file or a web address.');
  if (a.content === undefined || a.content === null) throw new Error(`Attachment ${n + 1} has no content.`);
  const content = typeof a.content === 'string' ? a.content : Buffer.from(a.content);
  return { filename: String(a.filename || (isCalendarPart(a) ? 'invite.ics' : `attachment-${n + 1}`)), content, ...(a.contentType ? { contentType: String(a.contentType) } : {}), ...(a.method ? { method: String(a.method) } : {}) };
}

function mailWithSendFrom({ accounts, fail = null }) {
  const sent = [];
  const service = {
    accounts: () => accounts,
    defaultAccount: () => {
      const a = accounts.find((x) => x.isDefault) || accounts[0];
      return a ? { id: a.id, address: a.email, name: a.name } : null;
    },
    folders: () => [{ path: 'INBOX', role: 'inbox' }],
    messages: () => ({ rows: [] }),
    // As Mail's: the account by id or address, or the default; its checks, then the message.
    sendFrom: async ({ account = null, to = '', cc = '', subject = '', text = '', attachments = [], headers = null } = {}) => {
      if (fail) throw Object.assign(new Error(fail.message), { code: fail.code });
      const a = account ? accounts.find((x) => x.id === account || x.email.toLowerCase() === String(account).toLowerCase()) : service.defaultAccount() && accounts.find((x) => x.id === service.defaultAccount().id);
      if (!a) throw Object.assign(new Error('That account is not set up.'), { code: 'no-account' });
      const files = (Array.isArray(attachments) ? attachments : []).map(handedAttachment);
      if (!to && !cc) throw new Error('A message needs someone to go to.');
      sent.push({ accountId: a.id, from: a.email, to, cc, subject, text, attachments: files, headers });
      return { messageId: `<m${sent.length}@test>`, refused: [], accountId: a.id, from: a.email };
    },
  };
  return { service, sent };
}

const zone = 'Europe/London';
const at = (H) => ({ date: '2026-11-24', time: `${String(H).padStart(2, '0')}:00:00`, tzid: zone, utc: false, allDay: false, at: localToUtc({ y: 2026, m: 11, d: 24, H, M: 0 }, zone) });

function setup(t, mail) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-calmail-'));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  return createCalendarService({ stores: { dir }, broadcast: null, mail: createMailLink(() => mail.service) });
}

test("an invitation passes Mail's own argument check: content only, typed with its method, from the organizer's account", async (t) => {
  const mail = mailWithSendFrom({ accounts: [{ id: 'a1', email: 'me@example.com', name: 'Me', isDefault: true, canSend: true }] });
  const calendar = setup(t, mail);
  const cal = calendar.calendars()[0];
  const saved = calendar.save({ calendarId: cal.id, event: { summary: 'Plan', start: at(10), end: at(11), attendees: [{ email: 'kim@example.org' }] } });
  await calendar.sendInvitation({ id: saved.id });
  const [m] = mail.sent;
  assert.equal(m.accountId, 'a1');
  assert.equal(m.attachments.length, 1, 'one calendar file: Mail makes the invitation part from it');
  const [file] = m.attachments;
  assert.equal(file.contentType, 'text/calendar; method=REQUEST; charset=utf-8');
  assert.equal(typeof file.content, 'string');
  assert.equal(readCalendar(file.content).method, 'REQUEST');
  assert.ok(!('path' in file));
});

test('an answer goes from the address invited, by address, and passes the same check', async (t) => {
  const mail = mailWithSendFrom({ accounts: [{ id: 'a1', email: 'me@example.com', name: 'Me', isDefault: true, canSend: true }, { id: 'a2', email: 'work@example.com', name: 'Me at work', canSend: true }] });
  const calendar = setup(t, mail);
  const invite = writeRequest([{ uid: 'inv@x', sequence: 0, summary: 'Review', start: at(14), end: at(15), organizer: { email: 'kim@example.org', name: 'Kim' }, attendees: [{ email: 'work@example.com' }] }], {});
  const r = await calendar.answer({ text: invite, partstat: 'ACCEPTED' });
  assert.equal(r.sent, true);
  const [m] = mail.sent;
  assert.equal(m.accountId, 'a2');
  assert.equal(m.from, 'work@example.com');
  assert.equal(m.attachments[0].contentType, 'text/calendar; method=REPLY; charset=utf-8');
  assert.equal(readCalendar(m.attachments[0].content).method, 'REPLY');
});

test("Mail's no-account and cannot-send come back said plainly, the code kept", async (t) => {
  for (const [code, said] of [['no-account', /no Mail account to send it from/], ['cannot-send', /cannot send from me@example\.com, as that account has no outgoing server/], [null, /^Not sent: the server hung up$/]]) {
    const mail = mailWithSendFrom({ accounts: [{ id: 'a1', email: 'me@example.com', name: 'Me', isDefault: true, canSend: code !== 'cannot-send' }], fail: { code, message: 'the server hung up' } });
    const calendar = setup(t, mail);
    const cal = calendar.calendars()[0];
    const saved = calendar.save({ calendarId: cal.id, event: { summary: 'Plan', start: at(10), end: at(11), attendees: [{ email: 'kim@example.org' }] } });
    await assert.rejects(calendar.sendInvitation({ id: saved.id }), (err) => said.test(err.message) && err.code === code);
  }
});

test("Mail's older send still takes the answer from the address invited, as that account's id with the file", async (t) => {
  const sent = [];
  const accounts = [{ id: 'a1', email: 'me@example.com', name: 'Me' }, { id: 'a2', email: 'work@example.com', name: 'Work' }];
  const service = {
    accounts: () => accounts,
    folders: () => [], messages: () => ({ rows: [] }),
    send: async ({ accountId, draft }) => { sent.push({ accountId, ...draft }); return { refused: [] }; },
  };
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-calmail-'));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  const calendar = createCalendarService({ stores: { dir }, broadcast: null, mail: createMailLink(() => service) });
  const invite = writeRequest([{ uid: 'old@x', sequence: 0, summary: 'Old', start: at(9), end: at(10), organizer: { email: 'kim@example.org' }, attendees: [{ email: 'work@example.com' }] }], {});
  await calendar.answer({ text: invite, partstat: 'DECLINED' });
  assert.equal(sent[0].accountId, 'a2');
  assert.ok(fs.existsSync(sent[0].attachments[0].path));
});
