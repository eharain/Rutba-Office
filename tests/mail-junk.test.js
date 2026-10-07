// Mail's junk filter: the lists, the server's verdict, and a filter that
// learns from what is marked Junk and Not junk — the pure half first, then
// through the service, where arriving mail is filed and moving a message
// teaches it.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { defaultJunk, junkSettings, learn, junkScore, junkVerdict, listSender, listEntry, onList, tokensOf, serverSaysJunk, MIN_LEARNED } from '../apps/desktop/main/mail-junk.js';

// deliverTest, the stand-in for an IMAP fetch, runs only in a check run.
process.env.RUTBA_OFFICE_VERIFY_APPS = '1';
const { createMailService } = await import('../apps/desktop/main/mail.js');

const JUNK = [
  ['winner@prize-claims.example', 'You have WON a cash prize', 'Claim your lottery prize now, send your bank details to receive the transfer of one million dollars.'],
  ['offers@cheap-meds.example', 'Cheap pills online', 'Buy cheap pills online with no prescription, discount offer, click here http://cheap-meds.example/buy'],
  ['agent@prize-claims.example', 'Urgent transfer pending', 'Urgent: your transfer of funds is pending, send bank details and a processing fee.'],
  ['deals@spam-shop.example', 'Limited offer just for you', 'Limited time offer, click here to claim your free gift card, the discount ends today.'],
  ['coins@moon.example', 'Double your money', 'Invest now and double your money, guaranteed returns, click here http://moon.example'],
  ['claims@lotto.example', 'Lottery winner notification', 'Congratulations winner, your address won the international lottery, claim your prize with your bank details.'],
];
const GOOD = [
  ['sam@work.example', 'Minutes from Tuesday', 'Here are the minutes from the planning meeting on Tuesday. The next review is on the 14th.'],
  ['ana@work.example', 'Draft report attached', 'I attached the draft of the quarterly report, could you read section two before Friday?'],
  ['sam@work.example', 'Lunch on Thursday?', 'Are you free for lunch on Thursday? The usual place near the office.'],
  ['it@work.example', 'Laptop upgrade schedule', 'Your laptop upgrade is scheduled for Monday morning, please save your work beforehand.'],
  ['mum@family.example', 'Sunday dinner', 'Coming for dinner on Sunday? Bring the photos from the trip.'],
  ['ana@work.example', 'Re: the budget review', 'Thanks for the numbers, the budget review meeting moves to Wednesday afternoon.'],
];
const message = ([address, subject, text], id = subject) => ({ from: [{ address, name: '' }], subject, text, messageId: `${id}@test` });
const taught = () => {
  let s = junkSettings(defaultJunk());
  JUNK.forEach((m) => { s = learn(s, message(m), true); });
  GOOD.forEach((m) => { s = learn(s, message(m), false); });
  return s;
};

test('list entries are addresses or domains, and a domain covers the ones under it', () => {
  assert.equal(listEntry(' Friend@Example.COM '), 'friend@example.com');
  assert.equal(listEntry('example.org'), '@example.org');
  assert.equal(listEntry('@example.org'), '@example.org');
  assert.equal(listEntry('not an address'), null);
  assert.equal(onList(['@example.org'], 'a@mail.example.org'), true, 'a subdomain');
  assert.equal(onList(['@example.org'], 'a@notexample.org'), false, 'not a suffix of the name');
  assert.equal(onList(['friend@example.com'], 'FRIEND@example.com'), true);
});

test('a message\'s words: its subject, its sender, its body and where its links go', () => {
  const t = tokensOf({ from: [{ address: 'deals@shop.example', name: 'Deals' }], subject: 'Free gift', html: '<p>Click <a href="https://track.example/x">here</a></p>' });
  for (const want of ['subject:free', 'free', 'from:deals@shop.example', 'domain:shop.example', 'name:deals', 'click', 'url:track.example', 'meta:markup-only']) assert.ok(t.includes(want), want);
});

test('the filter says nothing until it has learned enough of each kind, then tells junk from good', () => {
  let s = junkSettings(defaultJunk());
  for (const m of JUNK.slice(0, MIN_LEARNED)) s = learn(s, message(m), true);
  assert.equal(junkScore(s.model, message(JUNK[0])), null, 'no good mail learned yet');
  s = taught();
  const junkLike = message(['new@prize-claims.example', 'Claim your prize', 'You are a winner, send your bank details to claim the cash prize transfer.'], 'n1');
  const goodLike = message(['newperson@work.example', 'Report question', 'A question about section two of the quarterly report draft before the meeting.'], 'n2');
  assert.ok(junkScore(s.model, junkLike) >= 0.9, `junk scores high: ${junkScore(s.model, junkLike)}`);
  assert.ok(junkScore(s.model, goodLike) <= 0.2, `good scores low: ${junkScore(s.model, goodLike)}`);
  assert.deepEqual(junkVerdict(junkLike, s), { junk: true, why: 'filter', score: junkScore(s.model, junkLike) });
  assert.equal(junkVerdict(goodLike, s).junk, false);
});

test('learning a message twice counts it once, and marking it the other way moves it across', () => {
  const s = taught();
  assert.equal(learn(s, message(JUNK[0]), true), s, 'the same way again: nothing changes');
  const swapped = learn(s, message(JUNK[0]), false);
  assert.deepEqual([swapped.model.junk, swapped.model.good], [JUNK.length - 1, GOOD.length + 1]);
  assert.deepEqual([s.model.junk, s.model.good], [JUNK.length, GOOD.length], 'the settings it was given are not touched');
});

test('the lists, contacts, the server\'s verdict and the level, in Outlook\'s order', () => {
  const s = taught();
  const junkLike = message(JUNK[0], 'x');
  assert.equal(junkVerdict(junkLike, listSender(s, 'winner@prize-claims.example', 'safe')).why, 'safe', 'a Safe Sender beats the filter');
  assert.equal(junkVerdict(message(GOOD[0], 'y'), listSender(s, '@work.example', 'blocked')).why, 'blocked', 'a Blocked Sender beats it too');
  assert.equal(junkVerdict(junkLike, s, { isContact: () => true }).why, 'contact');
  assert.equal(junkVerdict(junkLike, { ...s, trustContacts: false }, { isContact: () => true }).why, 'filter', 'contacts trusted only while that is on');
  assert.equal(junkVerdict(junkLike, { ...s, level: 'off' }).junk, false, 'off: the filter is not asked');
  const flagged = { ...message(GOOD[0], 'z'), headers: [{ name: 'X-Spam-Flag', key: 'x-spam-flag', value: 'YES' }] };
  assert.equal(serverSaysJunk(flagged), true);
  assert.equal(junkVerdict(flagged, s).why, 'server');
  assert.equal(junkVerdict(message(GOOD[0], 'w'), { ...s, level: 'safeOnly' }).why, 'notSafe', 'Safe Lists Only: a stranger is junk');
  // Moving a sender from one list takes them off the other.
  const both = listSender(listSender(s, 'a@b.example', 'safe'), 'a@b.example', 'blocked');
  assert.deepEqual([both.safe, both.blocked], [[], ['a@b.example']]);
  assert.throws(() => listSender(s, 'nonsense', 'safe'), /not an address or a domain/);
});

/* ── through the service ─────────────────────────────────────────────────── */

function service() {
  const settings = new Map();
  const stores = {
    settings: { get: (k, f) => (settings.has(k) ? settings.get(k) : f), set: (k, v) => settings.set(k, v), delete: (k) => settings.delete(k), all: () => Object.fromEntries(settings) },
    secrets: { available: () => true, get: () => null, set: () => {}, delete: () => {}, keys: () => [] },
  };
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-junk-'));
  const mail = createMailService({ stores, holdBlob: () => ({ url: '' }), broadcast: () => {}, userData, smtp: { send: async () => ({ messageId: 'x' }) } });
  const account = mail.addAccount({ account: { email: 'me@example.com', name: 'Me', imap: { host: 'h', port: 993 }, smtp: { host: 'h', port: 465 } }, password: 'secret' });
  return { mail, accountId: account.id };
}
const raw = ([from, subject, body], n) => [`From: <${from}>`, 'To: <me@example.com>', `Subject: ${subject}`, `Date: ${new Date(Date.UTC(2026, 9, 1) + n * 60000).toUTCString()}`, `Message-ID: <t${n}-${subject.length}@test>`, 'MIME-Version: 1.0', 'Content-Type: text/plain; charset=utf-8', '', body, ''].join('\r\n');
const subjects = (mail, accountId, folder) => mail.messages({ accountId, folder, limit: 100 }).rows.map((r) => r.subject);

test('moving a message into Junk teaches the filter junk, and Not junk teaches it good', async () => {
  const { mail, accountId } = service();
  for (const [i, m] of [...JUNK, ...GOOD].entries()) await mail.deliverTest({ accountId, raw: raw(m, i) });
  assert.equal(subjects(mail, accountId, 'Inbox').length, JUNK.length + GOOD.length, 'nothing filed before it has learned');
  const rows = mail.messages({ accountId, folder: 'Inbox', limit: 100 }).rows;
  const junkIds = rows.filter((r) => JUNK.some(([, s]) => s === r.subject)).map((r) => r.id);
  mail.move({ accountId, folder: 'Inbox', ids: junkIds, to: 'Junk' });
  assert.deepEqual(mail.junk().learned, { junk: JUNK.length, good: 0 });
  // The Inbox learned as good from the folders, and the filter is ready.
  const learned = mail.learnJunk({ accountId });
  assert.deepEqual([learned.added, learned.ready], [{ junk: 0, good: GOOD.length }, true]);

  // Now it files what arrives — and says why on the message.
  const r1 = await mail.deliverTest({ accountId, raw: raw(['new@prize-claims.example', 'Claim your prize', 'You are a winner, send your bank details to claim the cash prize transfer.'], 100) });
  const r2 = await mail.deliverTest({ accountId, raw: raw(['newperson@work.example', 'Report question', 'A question about section two of the quarterly report draft before the review meeting.'], 101) });
  assert.deepEqual([r1.junked, r2.junked], [1, 0]);
  assert.ok(subjects(mail, accountId, 'Junk').includes('Claim your prize'));
  assert.ok(subjects(mail, accountId, 'Inbox').includes('Report question'));
  const filed = mail.messages({ accountId, folder: 'Junk', limit: 100 }).rows.find((r) => r.subject === 'Claim your prize');
  const full = mail.message({ accountId, folder: 'Junk', id: filed.id });
  assert.equal(full.junk.why, 'filter');
  assert.match(full.junk.reason, /junk filter judged it junk/);

  // Not junk: back in the Inbox, its note gone, learned as good.
  mail.notJunk({ accountId, folder: 'Junk', ids: [filed.id] });
  const back = mail.messages({ accountId, folder: 'Inbox', limit: 100 }).rows.find((r) => r.subject === 'Claim your prize');
  assert.ok(back, 'back in the Inbox');
  assert.equal(mail.message({ accountId, folder: 'Inbox', id: back.id }).junk, undefined);
  assert.equal(mail.junk().learned.good, GOOD.length + 1);
  // Deleting from Junk teaches nothing.
  const before = mail.junk().learned;
  const another = mail.messages({ accountId, folder: 'Junk', limit: 100 }).rows[0];
  mail.move({ accountId, folder: 'Junk', ids: [another.id], to: 'Trash' });
  assert.deepEqual(mail.junk().learned, before);

  // A message that comes late, carrying a date months before everything
  // else here, is still the one that arrived — filed, not passed over for
  // the newest.
  const late = await mail.deliverTest({ accountId, raw: raw(['sales@cheap-meds.example', 'Cheap pills, discount offer', 'Buy cheap pills online with a discount offer, click here http://cheap-meds.example/buy'], -200000) });
  assert.equal(late.junked, 1);
  assert.ok(subjects(mail, accountId, 'Junk').includes('Cheap pills, discount offer'));
});

test('a blocked sender goes to Junk at once, whatever the filter knows; options refuse a bad entry by name', async () => {
  const { mail, accountId } = service();
  mail.setJunk({ patch: { blocked: ['offers.example'] } });
  assert.deepEqual(mail.junk().blocked, ['@offers.example']);
  const r = await mail.deliverTest({ accountId, raw: raw(['news@mail.offers.example', 'Hello', 'Hello.'], 1) });
  assert.equal(r.junked, 1);
  assert.equal(mail.message({ accountId, folder: 'Junk', id: mail.messages({ accountId, folder: 'Junk', limit: 5 }).rows[0].id }).junk.why, 'blocked');
  assert.throws(() => mail.setJunk({ patch: { safe: ['friend@example.com', 'not an address'] } }), /"not an address" is not an address or a domain/);
  assert.throws(() => mail.setJunk({ patch: { level: 'extreme' } }), /not a junk filter level/);
  mail.listSender({ address: 'news@mail.offers.example', list: 'safe' });
  assert.deepEqual([mail.junk().safe, mail.junk().blocked], [['news@mail.offers.example'], ['@offers.example']], 'an address on the Safe list, its domain still blocked');
  const r2 = await mail.deliverTest({ accountId, raw: raw(['news@mail.offers.example', 'Hello again', 'Hello.'], 2) });
  assert.equal(r2.junked, 1, 'Blocked comes first, as in Outlook');
});

/* ── per account, Safe Recipients, and the International lists ───────────── */

test('Safe Recipients keep mail to a group or a mailing list, and the International lists block by country and encoding', async () => {
  const { tldEntry, groupAddressOf, charsetsOf, listRecipient } = await import('../apps/desktop/main/mail-junk.js');
  assert.deepEqual(['ru', '.CN', '*.br', 'xn--p1ai', 'r', 'co.uk', '1x'].map(tldEntry), ['ru', 'cn', 'br', 'xn--p1ai', null, null, null]);
  const list = { from: [{ address: 'someone@example.org' }], to: [{ address: 'dev@lists.example.org' }], subject: 'Build broken', text: 'x', headers: [{ key: 'list-post', value: '<mailto:dev@lists.example.org>' }] };
  assert.equal(groupAddressOf(list, ['me@example.com']), 'dev@lists.example.org');
  const toGroup = { from: [{ address: 'boss@work.example' }], to: [{ address: 'everyone@work.example' }] };
  assert.equal(groupAddressOf(toGroup, ['me@work.example']), 'everyone@work.example', 'the one address it went to, not mine');
  assert.equal(groupAddressOf({ from: [{ address: 'a@b.example' }], to: [{ address: 'me@example.com' }] }, ['me@example.com']), null, 'sent to me: no group');
  assert.deepEqual(charsetsOf({ charsets: ['koi8-r'], headers: [{ key: 'subject', value: '=?windows-1251?B?0eru7Ozl?=' }] }).sort(), ['koi8-r', 'windows-1251']);

  let s = junkSettings({ ...defaultJunk(), level: 'safeOnly' });
  assert.equal(junkVerdict(list, s).why, 'notSafe', 'Safe Lists Only: a list not yet kept goes to Junk');
  s = listRecipient(s, 'dev@lists.example.org');
  assert.deepEqual(junkVerdict(list, s), { junk: false, why: 'safeRecipient', score: null });

  s = junkSettings({ ...defaultJunk(), level: 'off', blockedTlds: ['ru'], blockedEncodings: ['cyrillic', 'nonsense'] });
  assert.deepEqual(s.blockedEncodings, ['cyrillic'], 'an encoding it does not know is dropped');
  assert.equal(junkVerdict({ from: [{ address: 'x@mail.ru' }], subject: 'Hi', text: 'hi' }, s).why, 'blockedTld', 'whatever the level, as Blocked Senders');
  assert.equal(junkVerdict({ from: [{ address: 'x@example.com' }], subject: 'Hi', text: 'hi', charsets: ['koi8-r'] }, s).why, 'blockedEncoding');
  assert.equal(junkVerdict({ from: [{ address: 'x@example.ru.example' }], subject: 'Hi', text: 'hi', charsets: ['utf-8'] }, s).junk, false, 'only the last part of the domain is its country');
});

test('each account keeps its own junk options once it has them; what the filter learned is shared', async () => {
  const { mail, accountId } = service();
  const other = mail.addAccount({ account: { email: 'me@other.example', name: 'Me', imap: { host: 'h2', port: 993 }, smtp: { host: 'h2', port: 465 } }, password: 'secret' }).id;
  assert.equal(mail.junk({ accountId }).own, false, 'until it has its own, an account uses the shared ones');
  mail.setJunk({ accountId, patch: { blocked: ['@noisy.example'] } });
  assert.deepEqual([mail.junk({ accountId }).blocked, mail.junk({ accountId }).own], [['@noisy.example'], true]);
  assert.deepEqual([mail.junk({ accountId: other }).blocked, mail.junk().blocked], [[], []], 'the other account, and the shared ones, untouched');
  const noisy = (n) => raw(['news@noisy.example', `News ${n}`, 'News.'], n).replace('To: <me@example.com>', `To: <${n % 2 ? 'me@other.example' : 'me@example.com'}>`);
  const a = await mail.deliverTest({ accountId, raw: noisy(2) });
  const b = await mail.deliverTest({ accountId: other, raw: noisy(3) });
  assert.deepEqual([a.junked, b.junked], [1, 0], 'blocked for the one account only');
  // Block Sender on a message of the other account lists it there.
  mail.listSender({ accountId: other, address: 'pest@pests.example', list: 'blocked' });
  assert.deepEqual([mail.junk({ accountId: other }).blocked, mail.junk({ accountId }).blocked], [['pest@pests.example'], ['@noisy.example']]);
  // Learning is one: a message marked junk in either account teaches both.
  const row = mail.messages({ accountId: other, folder: 'Inbox', limit: 5 }).rows[0];
  mail.move({ accountId: other, folder: 'Inbox', ids: [row.id], to: 'Junk' });
  assert.equal(mail.junk({ accountId }).learned.junk, 1);
  assert.equal(mail.junk({ accountId }).blocked[0], '@noisy.example', 'learning leaves the options as they were');
});

test('Never block this group or mailing list keeps the list safe for the account; mail sent to a person is refused by saying so', async () => {
  const { mail, accountId } = service();
  mail.setJunk({ accountId, patch: { level: 'safeOnly' } });
  const post = (n, subject) => raw(['someone@example.org', subject, 'A post to the list.'], n)
    .replace('To: <me@example.com>', 'To: <dev@lists.example.org>\r\nList-Id: Developers <dev.lists.example.org>\r\nList-Post: <mailto:dev@lists.example.org>');
  const first = await mail.deliverTest({ accountId, raw: post(1, 'Build broken') });
  assert.equal(first.junked, 1, 'Safe Lists Only files it, until the list is kept');
  const row = mail.messages({ accountId, folder: 'Junk', limit: 5 }).rows[0];
  const view = mail.listRecipient({ accountId, folder: 'Junk', ids: [row.id] });
  assert.deepEqual([view.added, view.safeRecipients], [['dev@lists.example.org'], ['dev@lists.example.org']]);
  const second = await mail.deliverTest({ accountId, raw: post(2, 'Build fixed') });
  assert.equal(second.junked, 0, 'mail to the list now stays in the Inbox');
  const direct = await mail.deliverTest({ accountId, raw: raw(['someone@example.org', 'Just to you', 'Hello.'], 3) });
  assert.equal(direct.junked, 1);
  const mine = mail.messages({ accountId, folder: 'Junk', limit: 5 }).rows.find((r) => r.subject === 'Just to you');
  assert.throws(() => mail.listRecipient({ accountId, folder: 'Junk', ids: [mine.id] }), /sent to you, not to a group or a mailing list/);
  assert.throws(() => mail.setJunk({ accountId, patch: { blockedTlds: ['ru', 'russia!'] } }), /"russia!" is not a top-level domain/);
  assert.throws(() => mail.setJunk({ accountId, patch: { blockedEncodings: ['klingon'] } }), /"klingon" is not an encoding/);
});

test('mail written in a blocked encoding goes to Junk, a multipart one by its text part', async () => {
  const { mail, accountId } = service();
  mail.setJunk({ accountId, patch: { level: 'low', blockedEncodings: ['cyrillic'] } });
  const plain = raw(['x@example.com', 'Privet', 'Privet.'], 1).replace('charset=utf-8', 'charset=koi8-r');
  const multi = [
    'From: <y@example.com>', 'To: <me@example.com>', 'Subject: Hello', `Date: ${new Date(Date.UTC(2026, 9, 1)).toUTCString()}`, 'Message-ID: <m2@test>', 'MIME-Version: 1.0',
    'Content-Type: multipart/alternative; boundary="b1"', '', '--b1', 'Content-Type: text/plain; charset="windows-1251"', '', 'Hello.', '--b1--', '',
  ].join('\r\n');
  const ok = raw(['z@example.com', 'Fine', 'Fine.'], 3);
  const results = [await mail.deliverTest({ accountId, raw: plain }), await mail.deliverTest({ accountId, raw: multi }), await mail.deliverTest({ accountId, raw: ok })];
  assert.deepEqual(results.map((r) => r.junked), [1, 1, 0]);
  const filed = mail.messages({ accountId, folder: 'Junk', limit: 5 }).rows[0];
  const full = mail.message({ accountId, folder: 'Junk', id: filed.id });
  assert.equal(full.junk.why, 'blockedEncoding');
  assert.match(full.junk.reason, /Blocked Encodings/);
});
