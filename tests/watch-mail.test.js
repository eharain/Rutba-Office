// Mail, as the notifier hears of it (apps/desktop/main/watch-mail.js,
// D-2026-11-22-23): new email. Mail's arrivals are handed in here as
// main/mail.js tells them (its onArrived); nothing is fetched and nothing is
// shown on this computer.

import test from 'node:test';
import assert from 'node:assert/strict';
import { watchMail, newsOf, QUIET_ROLES, FRESH_MS, NAMES_MOST } from '../apps/desktop/main/watch-mail.js';

const NOW = new Date(2026, 10, 23, 10, 0).getTime();
const lately = new Date(NOW - 60_000).toISOString();

const message = (id, name, subject, more = {}) => ({ id, subject, from: { name, address: `${name.toLowerCase().replace(/\s+/g, '.')}@example.org` }, unread: true, date: lately, ...more });
const arrival = (messages, more = {}) => ({ accountId: 'a1', folder: 'INBOX', role: 'inbox', ids: messages.map((m) => m.id), count: messages.length, at: new Date(NOW).toISOString(), invitations: [], filedByRules: false, messages, ...more });

function setup({ accounts = [{ id: 'a1', email: 'sam@example.org', name: 'Work' }], mail = undefined } = {}) {
  const reports = [];
  const notifier = { report: (note) => { reports.push(note); return { shown: true }; } };
  const listeners = new Set();
  const service = mail !== undefined ? mail : {
    onArrived: (fn) => { listeners.add(fn); return () => listeners.delete(fn); },
    accounts: () => accounts,
  };
  const ctx = { front: false };
  const watch = watchMail({ notifier, mail: service, front: (app) => ctx.front && app === 'mail', now: () => NOW });
  const arrive = (a) => { for (const fn of [...listeners]) fn(a); };
  return { watch, reports, arrive, listeners, ctx, notifier };
}

test('one new email is reported with who it is from and its subject, and opens that message in a window of its own', () => {
  const { reports, arrive } = setup();
  arrive(arrival([message(41, 'Amina Yusuf', 'Quarterly numbers')]));
  assert.equal(reports.length, 1);
  const r = reports[0];
  assert.equal(r.kind, 'mail');
  assert.equal(r.id, 'mail:a1:41');
  assert.equal(r.from, 'Amina Yusuf');
  assert.equal(r.subject, 'Quarterly numbers');
  assert.equal(r.account, 'sam@example.org');
  assert.equal(r.count, 1);
  assert.equal(r.app, 'mail');
  assert.equal(r.front, false);
  assert.deepEqual(r.open, { app: 'mail', fresh: true, query: { message: JSON.stringify({ accountId: 'a1', folder: 'INBOX', id: 41 }) } });
});

test('several at once are one report: how many, who from (each named once, a few at most), opened on the account', () => {
  const { reports, arrive } = setup();
  arrive(arrival([message(1, 'Amina Yusuf', 'One'), message(2, 'Bilal Khan', 'Two'), message(3, 'Amina Yusuf', 'Three')]));
  assert.equal(reports[0].count, 3);
  assert.equal(reports[0].from, 'Amina Yusuf, Bilal Khan');
  assert.equal(reports[0].subject, '');
  assert.deepEqual(reports[0].open, { app: 'mail', query: { account: 'a1' } }, 'the Inbox is where Mail opens an account anyway');
  const many = ['A', 'B', 'C', 'D', 'E'].map((n, i) => message(10 + i, `Person ${n}`, n));
  arrive(arrival(many, { folder: 'Projects', role: 'folder' }));
  assert.equal(reports[1].count, 5);
  assert.equal(reports[1].from.split(', ').length, NAMES_MOST);
  assert.ok(reports[1].from.endsWith('…'));
  assert.deepEqual(reports[1].open, { app: 'mail', query: { account: 'a1', folder: 'Projects' } });
});

test('a sender with no name is named by the address', () => {
  const { reports, arrive } = setup();
  arrive(arrival([{ id: 5, subject: 'Hello', from: { name: '', address: 'bilal@example.org' }, unread: true, date: lately }]));
  assert.equal(reports[0].from, 'bilal@example.org');
  arrive(arrival([{ id: 6, subject: 'No sender', from: null, unread: true, date: null }]));
  assert.equal(reports[1].from, '', 'the notifier says "New email" then');
  assert.equal(reports[1].count, 1, 'a message with no date of its own is taken as new');
});

test('nothing is told of what arrives in Sent, Drafts, Trash, Junk or Archive', () => {
  const { reports, arrive } = setup();
  for (const role of QUIET_ROLES) arrive(arrival([message(7, 'Amina Yusuf', 'Kept')], { role, folder: role }));
  assert.deepEqual(reports, []);
});

test('mail already read elsewhere and old mail a first fetch brings are not news; the rest is counted alone', () => {
  const { reports, arrive } = setup();
  const old = new Date(NOW - FRESH_MS - 3600_000).toISOString();
  arrive(arrival([message(1, 'Amina Yusuf', 'Read', { unread: false }), message(2, 'Bilal Khan', 'Old', { date: old })]));
  assert.deepEqual(reports, [], 'none of it is news');
  arrive(arrival([message(3, 'Amina Yusuf', 'Read', { unread: false }), message(4, 'Bilal Khan', 'New'), message(5, 'Chen Li', 'Old', { date: old })]));
  assert.equal(reports.length, 1);
  assert.equal(reports[0].count, 1);
  assert.equal(reports[0].from, 'Bilal Khan');
  assert.equal(reports[0].id, 'mail:a1:4');
  assert.equal(JSON.parse(reports[0].open.query.message).id, 4);
  assert.deepEqual(newsOf(arrival([message(8, 'Amina Yusuf', 'x', { date: 'not a date' })]), NOW).map((m) => m.id), [8]);
});

test('mail a rule filed is no longer where it arrived: Mail is opened on the account', () => {
  const { reports, arrive } = setup();
  arrive(arrival([message(9, 'Amina Yusuf', 'Filed')], { filedByRules: true }));
  assert.equal(reports[0].count, 1);
  assert.equal(reports[0].subject, 'Filed');
  assert.deepEqual(reports[0].open, { app: 'mail', query: { account: 'a1' } });
  arrive(arrival([message(10, 'A B', 'x'), message(11, 'C D', 'y')], { filedByRules: true, folder: 'Projects', role: 'folder' }));
  assert.deepEqual(reports[1].open, { app: 'mail', query: { account: 'a1' } });
});

test('a Mail window in front is said so, and the notifier then shows nothing', () => {
  const { reports, arrive, ctx } = setup();
  ctx.front = true;
  arrive(arrival([message(12, 'Amina Yusuf', 'Seen')]));
  assert.equal(reports[0].front, true);
});

test('an arrival that does not say what came is told by its count; an account not known has no address', () => {
  const { reports, arrive } = setup({ accounts: [] });
  assert.equal(newsOf({ count: 2 }), null);
  arrive({ accountId: 'a9', folder: 'INBOX', ids: [31, 32], count: 2, at: 'then' });
  assert.equal(reports[0].count, 2);
  assert.equal(reports[0].id, 'mail:a9:31');
  assert.equal(reports[0].from, '');
  assert.equal(reports[0].account, '');
  assert.deepEqual(reports[0].open, { app: 'mail', query: { account: 'a9', folder: 'INBOX' } });
  arrive({ accountId: 'a9', folder: 'INBOX', ids: [], count: 0 });
  arrive(null);
  assert.equal(reports.length, 1);
});

test('stopped, nothing more is heard; a fault here never reaches Mail; a Mail that cannot tell of arrivals is not watched', () => {
  const a = setup();
  a.watch.stop();
  assert.equal(a.listeners.size, 0);
  a.watch.stop();

  const b = setup();
  b.notifier.report = () => { throw new Error('a notifier with a fault'); };
  assert.doesNotThrow(() => b.arrive(arrival([message(13, 'Amina Yusuf', 'x')])));

  assert.equal(setup({ mail: {} }).watch, null);
  assert.equal(setup({ mail: null }).watch, null);
  assert.equal(setup({ mail: { onArrived: () => { throw new Error('refused'); } } }).watch, null);
});
