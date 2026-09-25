// Send later and the Outbox it shares with the undo window.
//
// `createMailService` takes its clock and its SMTP transport as arguments —
// see `now` and `smtp` in apps/desktop/main/mail.js — so every test here
// drives a scheduler that never waits on a real timer and never opens a
// socket. `scheduleChoices` and `parseCustomSchedule` (packages/mailbox/src/
// schedule.js) are the pure half: what Send later's own buttons compute,
// with no store or service involved at all.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createMailService } from '../apps/desktop/main/mail.js';
import { scheduleChoices, parseCustomSchedule } from '../packages/mailbox/src/schedule.js';

/** An in-memory stand-in for the settings/secrets files the real shell keeps. */
function fakeStores() {
  const settings = new Map();
  const secrets = new Map();
  return {
    settings: {
      get: (k, fallback) => (settings.has(k) ? settings.get(k) : fallback),
      set: (k, v) => settings.set(k, v),
      delete: (k) => settings.delete(k),
      all: () => Object.fromEntries(settings),
    },
    secrets: {
      available: () => true,
      get: (k) => secrets.get(k) ?? null,
      set: (k, v) => secrets.set(k, v),
      delete: (k) => secrets.delete(k),
      keys: () => [...secrets.keys()],
    },
  };
}

function fixture({ fail = () => false } = {}) {
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mail-sendlater-'));
  const stores = fakeStores();
  let clock = Date.parse('2026-03-02T09:00:00Z'); // a Monday
  const sent = [];
  const smtp = {
    async send({ draft }) {
      if (fail(draft)) throw new Error('the server refused the connection');
      sent.push(draft);
      return { messageId: `m-${sent.length}` };
    },
  };
  const service = createMailService({
    stores,
    holdBlob: () => ({ url: '' }),
    broadcast: () => {},
    userData,
    now: () => clock,
    smtp,
  });
  const account = service.addAccount({
    account: { email: 'me@example.com', name: 'Me', imap: { host: 'imap.example.com', port: 993 }, smtp: { host: 'smtp.example.com', port: 465 } },
    password: 'secret',
  });
  return {
    service,
    account,
    stores,
    sent,
    setClock: (ms) => { clock = ms; },
    clock: () => clock,
  };
}

test('scheduling a message stores it in the Outbox with the chosen time', () => {
  const { service, account } = fixture();
  const at = '2026-03-05T08:00:00.000Z';
  const item = service.queue({ accountId: account.id, draft: { to: 'a@b.com', subject: 'Later' }, at });
  assert.equal(item.at, at);
  const outbox = service.outbox();
  assert.equal(outbox.length, 1);
  assert.equal(outbox[0].id, item.id);
  assert.equal(outbox[0].draft.subject, 'Later');
});

test('the presets land where a person reading them would expect, including across a weekend', () => {
  // A Saturday: "Monday morning" must be the day after tomorrow, not today
  // and not a week further off, and "Later today" must not appear once 9pm
  // has passed.
  const saturdayEvening = new Date('2026-03-07T22:00:00'); // Saturday, past 9pm local
  const choices = scheduleChoices(saturdayEvening);
  assert.ok(!choices.some((c) => c.id === 'tonight'), 'no "later today" once 9pm has passed');
  const monday = choices.find((c) => c.id === 'monday').at;
  assert.equal(monday.getDay(), 1);
  assert.equal(monday.getDate(), 9); // the 7th is a Saturday; the next Monday is the 9th
  assert.equal(monday.getHours(), 8);

  // A Monday morning itself: "Monday morning" must mean *next* Monday, a
  // full week on, never today.
  const mondayMorning = new Date('2026-03-02T07:00:00');
  const again = scheduleChoices(mondayMorning).find((c) => c.id === 'monday').at;
  assert.equal(again.getDate(), 9);

  // A person's own date and time, read as local time; an empty or
  // unparseable field is `null` rather than "now" or a thrown error.
  assert.equal(parseCustomSchedule(''), null);
  assert.equal(parseCustomSchedule('not a date'), null);
  const custom = parseCustomSchedule('2026-03-05T08:00');
  assert.equal(custom.getFullYear(), 2026);
  assert.equal(custom.getMonth(), 2);
  assert.equal(custom.getDate(), 5);
  assert.equal(custom.getHours(), 8);
});

test('the scheduler sends only what is due, in order, and leaves the rest waiting', async () => {
  const { service, account, sent, setClock, clock } = fixture();
  const early = service.queue({ accountId: account.id, draft: { to: 'first@example.com', subject: 'First' }, at: new Date(clock() + 60_000).toISOString() });
  const later = service.queue({ accountId: account.id, draft: { to: 'second@example.com', subject: 'Second' }, at: new Date(clock() + 120_000).toISOString() });
  assert.equal(service.outbox().length, 2);

  setClock(clock() + 60_000); // only `early` is due
  await service.pumpOutbox();
  assert.deepEqual(sent.map((d) => d.subject), ['First']);
  assert.deepEqual(service.outbox().map((o) => o.id), [later.id]);

  setClock(clock() + 60_000); // now `later` is due too
  await service.pumpOutbox();
  assert.deepEqual(sent.map((d) => d.subject), ['First', 'Second']);
  assert.equal(service.outbox().length, 0);
});

test('a failed send stays in the Outbox with the error on it, and is retried', async () => {
  let failNext = true;
  const { service, account, sent, setClock, clock } = fixture({ fail: () => failNext });
  service.queue({ accountId: account.id, draft: { to: 'a@b.com', subject: 'Try again' }, at: new Date(clock()).toISOString() });

  await service.pumpOutbox();
  assert.equal(sent.length, 0, 'nothing sent while it is failing');
  let [item] = service.outbox();
  assert.match(item.error, /refused/);
  assert.equal(item.attempts, 1);
  assert.ok(new Date(item.at).getTime() > clock(), 'the retry is scheduled for later, not dropped');

  // Too soon: the retry is not due yet, so nothing changes.
  await service.pumpOutbox();
  assert.equal(service.outbox().length, 1);
  assert.equal(sent.length, 0);

  // Once the clock reaches the retry time, and the server stops refusing it,
  // the same message goes out — the Outbox never lost it.
  failNext = false;
  setClock(new Date(item.at).getTime());
  await service.pumpOutbox();
  assert.deepEqual(sent.map((d) => d.subject), ['Try again']);
  assert.equal(service.outbox().length, 0);
});

test('taking a scheduled message back removes it from the Outbox, unsent', () => {
  const { service, account, sent } = fixture();
  const item = service.queue({ accountId: account.id, draft: { to: 'a@b.com', subject: 'Oops' }, at: new Date(Date.now() + 3600_000).toISOString() });
  const taken = service.unsend({ id: item.id });
  assert.equal(taken.draft.subject, 'Oops');
  assert.equal(service.outbox().length, 0);
  assert.equal(sent.length, 0);
});

test('taking a message back and not re-sending it leaves nothing scheduled; queuing it again keeps the new time', () => {
  const { service, account } = fixture();
  const first = service.queue({ accountId: account.id, draft: { to: 'a@b.com', subject: 'Edit me' }, at: '2026-03-05T08:00:00.000Z' });
  const taken = service.unsend({ id: first.id });
  assert.equal(service.outbox().length, 0, 'edited out of the Outbox until it is put back');
  const requeued = service.queue({ accountId: taken.accountId, draft: taken.draft, at: '2026-03-06T09:00:00.000Z' });
  assert.equal(service.outbox().length, 1);
  assert.equal(service.outbox()[0].at, '2026-03-06T09:00:00.000Z');
  assert.equal(service.outbox()[0].draft.subject, 'Edit me');
  assert.notEqual(requeued.id, first.id);
});

test('the Outbox survives a restart because it lives in settings, not in a timer', async () => {
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mail-sendlater-restart-'));
  const stores = fakeStores();
  const sent = [];
  const smtp = { async send({ draft }) { sent.push(draft); return { messageId: 'm' }; } };
  const first = createMailService({ stores, holdBlob: () => ({ url: '' }), broadcast: () => {}, userData, now: () => Date.parse('2026-03-02T09:00:00Z'), smtp });
  const account = first.addAccount({ account: { email: 'me@example.com', name: 'Me', imap: { host: 'h', port: 993 }, smtp: { host: 'h', port: 465 } }, password: 'secret' });
  first.queue({ accountId: account.id, draft: { to: 'a@b.com', subject: 'Still here' }, at: '2026-03-09T08:00:00.000Z' });

  // A fresh service over the same `stores` — the "restart" — reads the same
  // outbox back, not an empty one.
  const restarted = createMailService({ stores, holdBlob: () => ({ url: '' }), broadcast: () => {}, userData, now: () => Date.parse('2026-03-02T09:00:01Z'), smtp });
  assert.equal(restarted.outbox().length, 1);
  assert.equal(restarted.outbox()[0].draft.subject, 'Still here');
  assert.equal(sent.length, 0, 'not due yet, so nothing has gone out just from starting up');
});

test('a message that fell due while the application was closed goes out at start-up', async () => {
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mail-sendlater-overdue-'));
  const stores = fakeStores();
  const sent = [];
  const smtp = { async send({ draft }) { sent.push(draft); return { messageId: 'm' }; } };
  const before = createMailService({ stores, holdBlob: () => ({ url: '' }), broadcast: () => {}, userData, now: () => Date.parse('2026-03-02T09:00:00Z'), smtp });
  const account = before.addAccount({ account: { email: 'me@example.com', name: 'Me', imap: { host: 'h', port: 993 }, smtp: { host: 'h', port: 465 } }, password: 'secret' });
  before.queue({ accountId: account.id, draft: { to: 'a@b.com', subject: 'Overdue' }, at: '2026-03-02T09:05:00.000Z' });

  // The application was closed until well after that, then opened again —
  // constructing the service is what a start-up does, including its own
  // "anything left over from the last run goes out now".
  const after = createMailService({ stores, holdBlob: () => ({ url: '' }), broadcast: () => {}, userData, now: () => Date.parse('2026-03-03T10:00:00Z'), smtp });
  await after.pumpOutbox();
  assert.deepEqual(sent.map((d) => d.subject), ['Overdue']);
  assert.equal(after.outbox().length, 0);
});
