// Invite by email (Office 2.25): several people at once, each asked in Talk
// and emailed from the person's own sending account in Mail as a note they
// wrote, or, with no account to send from, by Rutba within its limits.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createTalkService, fillLetter, DOWNLOAD_URL, INVITE_BATCH } from '../apps/desktop/main/talk.js';
import { talkMail } from '../apps/desktop/main/talk-mail.js';
import { parseAddresses, INVITE_BATCH as DIALOG_BATCH } from '../apps/desktop/renderer/apps/talk/addresses.js';

const CODE = 'AbCdEfGh12345678_-xy';
const folders = [];
test.after(() => { for (const f of folders) fs.rmSync(f, { recursive: true, force: true }); });

/** The Talk service on a fake introducer that answers as `answer(msg)` says. */
function service({ answer, mail = null, allowance = { day: { used: 0, limit: 10 }, week: { used: 0, limit: 30 } } }) {
  const map = (m) => ({ get: (k, d) => (m.has(k) ? m.get(k) : d), set: (k, v) => m.set(k, v), delete: (k) => m.delete(k) });
  const settings = map(new Map());
  const secrets = map(new Map([['talk.token', 'amira']]));
  const out = [];
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'talk-invite-email-'));
  folders.push(dataDir);
  let ws = null;
  const svc = createTalkService({
    stores: { settings, secrets },
    dataDir,
    connect: () => {
      ws = { readyState: 1, send: (text) => {
        const msg = JSON.parse(text);
        out.push(msg);
        if (msg.type === 'hello') return void setImmediate(() => ws.onmessage({ data: JSON.stringify({ type: 'welcome', me: { id: 'u-a', name: 'Amira Khan', email: 'amira@example.com' }, contacts: [{ id: 'u-ben', name: 'Ben Ortiz', email: 'ben@example.com' }], requests: { incoming: [], outgoing: [] }, mailAllowance: allowance }) }));
        for (const reply of [].concat(answer(msg) || [])) setImmediate(() => ws.onmessage({ data: JSON.stringify(reply) }));
      }, close: () => {} };
      setImmediate(() => ws.onopen());
      return ws;
    },
    mail,
  });
  return { svc, out, settings };
}
const online = async (svc) => {
  svc.resume();
  for (let i = 0; i < 50 && svc.state().phase !== 'online'; i++) await new Promise((r) => setTimeout(r, 5));
  assert.equal(svc.state().phase, 'online');
};
let n = 0;
const asked = (msg, extra = {}) => ({ type: 'contact.request.sent', ref: msg.ref, request: { id: `req-${++n}`, email: msg.email, note: msg.note, at: 1, link: `https://introducer.rutba.io/i/${CODE}${n}` }, ...extra });
/** An introducer that knows Ben as a contact already and asks everyone else. */
const introducer = (extra = () => ({})) => (m) => {
  if (m.type !== 'contact.request') return null;
  if (m.email === 'ben@example.com') return { type: 'contact.request.sent', ref: m.ref, request: null, contact: { id: 'u-ben', name: 'Ben Ortiz', email: 'ben@example.com' } };
  return asked(m, extra(m));
};

test('the letter\'s places are filled: the links, and each person\'s name or none', () => {
  const letter = 'Hello {name},\n\nAccept: {accept}\n\nGet it: {join}';
  assert.equal(fillLetter(letter, { accept: 'A', join: 'J', name: 'Chen Wei' }), 'Hello Chen Wei,\n\nAccept: A\n\nGet it: J');
  assert.equal(fillLetter(letter, { accept: 'A', join: 'J' }), 'Hello,\n\nAccept: A\n\nGet it: J');
  // A name is never read as a place to fill.
  assert.equal(fillLetter('Hi {name}: {accept}', { accept: 'A', name: '{join}' }), 'Hi {join}: A');
});

test('addresses typed or pasted are split, named where they carry a name, and each kept once', () => {
  assert.deepEqual(parseAddresses('chen@example.com, Dana Ray <dana@example.com>; "Eve Stone" <eve@example.com>\nfred@example.com CHEN@example.com not-an-address mailto:gus@example.com'), [
    { email: 'dana@example.com', name: 'Dana Ray' },
    { email: 'eve@example.com', name: 'Eve Stone' },
    { email: 'chen@example.com', name: '' },
    { email: 'fred@example.com', name: '' },
    { email: 'gus@example.com', name: '' },
  ]);
  assert.deepEqual(parseAddresses(''), []);
  assert.equal(DIALOG_BATCH, INVITE_BATCH);
});

test('Mail\'s sending account and its send are used when Mail has them', async () => {
  const sent = [];
  const mail = {
    accounts: () => [{ id: 'a1', email: 'amira@shop.example', name: 'Amira' }, { id: 'a2', email: 'amira@home.example' }, { id: 'arc', email: 'old.mbox', local: true }],
    defaultAccount: async () => ({ id: 'a2', address: 'amira@home.example', name: '' }),
    sendFrom: async (args) => { sent.push(['sendFrom', args]); },
    send: async (args) => { sent.push(['send', args]); },
  };
  const post = talkMail(mail);
  assert.deepEqual(await post.accounts(), [{ id: 'a1', address: 'amira@shop.example', name: 'Amira' }, { id: 'a2', address: 'amira@home.example', name: '' }]);
  assert.deepEqual(await post.defaultAccount(), { id: 'a2', address: 'amira@home.example', name: '' });
  assert.deepEqual(await post.sendFrom({ to: 'x@example.com', subject: 's', text: 't', html: 'h' }), { from: { id: 'a2', address: 'amira@home.example', name: '' } });
  await post.sendFrom({ account: 'a1', to: 'y@example.com', subject: 's', text: 't', html: 'h' });
  assert.deepEqual(sent, [
    ['sendFrom', { account: 'a2', to: 'x@example.com', subject: 's', text: 't', html: 'h' }],
    ['sendFrom', { account: 'a1', to: 'y@example.com', subject: 's', text: 't', html: 'h' }],
  ]);
  // An archive kept on this computer has no server: it is never sent from.
  await assert.rejects(post.sendFrom({ account: 'arc', to: 'z@example.com', subject: 's', text: 't' }), (err) => err.code === 'no-account');
});

test('a Mail from before the sending account: the first account that can send is the default, and send() sends', async () => {
  const sent = [];
  const post = talkMail({ accounts: () => [{ id: 'arc', email: 'old.mbox', local: true }, { id: 'a1', email: 'amira@shop.example' }], send: async (args) => { sent.push(args); } });
  assert.deepEqual(await post.defaultAccount(), { id: 'a1', address: 'amira@shop.example', name: '' });
  await post.sendFrom({ to: 'x@example.com', subject: 's', text: 't', html: 'h' });
  assert.deepEqual(sent, [{ accountId: 'a1', draft: { to: 'x@example.com', subject: 's', text: 't', html: 'h' } }]);
  assert.equal(await talkMail({ accounts: () => [] }).defaultAccount(), null);
  assert.equal(await talkMail(null).defaultAccount(), null);
  await assert.rejects(talkMail(null).sendFrom({ to: 'x@example.com' }), (err) => err.code === 'no-account');
});

test('invited by email from the sending account: each asked in Talk with no Rutba email, each sent a note with their name and their own link, a contact already left as one', async () => {
  const sent = [];
  const mail = { accounts: () => [{ id: 'a1', email: 'amira@shop.example', name: 'Amira Khan' }], defaultAccount: () => ({ id: 'a1', address: 'amira@shop.example', name: 'Amira Khan' }), sendFrom: async (args) => { sent.push(args); } };
  const { svc, out } = service({ mail, answer: introducer() });
  await online(svc);
  assert.deepEqual((await svc.sender()).account, { id: 'a1', address: 'amira@shop.example', name: 'Amira Khan' });
  const r = await svc.inviteByEmail({
    people: [{ email: 'chen@example.com', name: 'Chen Wei' }, { email: 'dana@example.com' }, { email: 'CHEN@example.com' }, { email: 'ben@example.com', name: 'Ben' }, { email: 'nobody' }],
    letter: { subject: 'Talk with Amira', text: 'Hello {name},\n\nCome and talk.\n\nAccept: {accept}\n\nGet Office: {join}' },
  });
  assert.equal(r.via, 'own');
  assert.equal(r.from, 'amira@shop.example');
  const requests = out.filter((m) => m.type === 'contact.request');
  assert.deepEqual(requests.map((m) => [m.email, m.notify]), [['chen@example.com', 'none'], ['dana@example.com', 'none'], ['ben@example.com', 'none']]);
  assert.deepEqual(r.results.map((x) => [x.email, x.contact ? 'contact' : x.told?.by, x.told?.sent ?? null]), [['chen@example.com', 'own', true], ['dana@example.com', 'own', true], ['ben@example.com', 'contact', null]]);
  assert.equal(sent.length, 2);
  assert.deepEqual(sent.map((s) => [s.account, s.to, s.subject]), [['a1', 'chen@example.com', 'Talk with Amira'], ['a1', 'dana@example.com', 'Talk with Amira']]);
  const chenLink = r.results[0].request.link;
  assert.equal(sent[0].text, `Hello Chen Wei,\n\nCome and talk.\n\nAccept: ${chenLink}\n\nGet Office: ${DOWNLOAD_URL}?invite=${chenLink.split('/i/')[1]}`);
  assert.match(sent[1].text, /^Hello,\n\nCome and talk\./);
  assert.notEqual(r.results[1].request.link, chenLink, 'each has their own link');
  assert.match(sent[0].html, /<a href="https:\/\/introducer\.rutba\.io\/i\//);
  // Each request shows how its person was told.
  assert.deepEqual(svc.state().requests.outgoing.map((o) => [o.email, o.told?.by, o.told?.sent]), [['chen@example.com', 'own', true], ['dana@example.com', 'own', true]]);
});

test('a note that lost its link still carries it: the email can always be accepted', async () => {
  const sent = [];
  const mail = { accounts: () => [{ id: 'a1', email: 'amira@shop.example' }], sendFrom: async (args) => { sent.push(args); } };
  const { svc } = service({ mail, answer: introducer() });
  await online(svc);
  const r = await svc.inviteByEmail({ people: [{ email: 'chen@example.com' }], letter: { subject: 'Hi', text: 'Just words.' } });
  assert.equal(r.results[0].told.sent, true);
  assert.equal(sent[0].text, `Just words.\n\n${r.results[0].request.link}\n`);
});

test('with no Mail account Rutba emails them, within its limits, the note in its email; asking for Rutba works with an account too', async () => {
  const { svc, out } = service({ answer: introducer((m) => ({ mail: m.email === 'fred@example.com' ? { sent: false, reason: 'limit' } : { sent: true, allowance: { day: { used: 1, limit: 10 }, week: { used: 1, limit: 30 } } } })) });
  await online(svc);
  assert.equal((await svc.sender()).account, null);
  const r = await svc.inviteByEmail({ people: [{ email: 'chen@example.com' }, { email: 'fred@example.com' }], note: 'x'.repeat(400) });
  assert.equal(r.via, 'rutba');
  assert.equal(r.from, null);
  const requests = out.filter((m) => m.type === 'contact.request');
  assert.deepEqual(requests.map((m) => m.notify), ['rutba', 'rutba']);
  assert.equal(requests[0].note.length, 300, 'the note as the introducer keeps it');
  assert.deepEqual(r.results.map((x) => x.told), [{ by: 'rutba', sent: true, reason: null }, { by: 'rutba', sent: false, reason: 'limit' }]);

  const sent = [];
  const withAccount = service({ mail: { accounts: () => [{ id: 'a1', email: 'amira@shop.example' }], sendFrom: async (a) => { sent.push(a); } }, answer: introducer(() => ({ mail: { sent: true } })) });
  await online(withAccount.svc);
  const viaRutba = await withAccount.svc.inviteByEmail({ people: [{ email: 'gus@example.com' }], via: 'rutba', note: 'Hi' });
  assert.equal(viaRutba.via, 'rutba');
  assert.equal(sent.length, 0);
  assert.equal(withAccount.out.find((m) => m.type === 'contact.request').notify, 'rutba');
  // Asked for one's own account where there is none to send from.
  assert.deepEqual(await svc.inviteByEmail({ people: [{ email: 'gus@example.com' }], via: 'own' }), { error: 'no-account' });
});

test('once the introducer takes no more today the rest are not asked; too many at once, or none, is refused before anything is asked', async () => {
  const mail = { accounts: () => [{ id: 'a1', email: 'amira@shop.example' }], sendFrom: async () => {} };
  const { svc, out } = service({ mail, answer: (m) => (m.type !== 'contact.request' ? null : m.email === 'dana@example.com' ? { type: 'error', for: 'contact.request', code: 'too-many' } : asked(m)) });
  await online(svc);
  const r = await svc.inviteByEmail({ people: ['chen', 'dana', 'eve', 'fred'].map((p) => ({ email: `${p}@example.com` })), letter: { subject: 's', text: '{accept}' } });
  assert.deepEqual(r.results.map((x) => x.error || 'asked'), ['asked', 'too-many', 'too-many', 'too-many']);
  assert.equal(out.filter((m) => m.type === 'contact.request').length, 2);
  const many = Array.from({ length: INVITE_BATCH + 1 }, (_, i) => ({ email: `p${i}@example.com` }));
  assert.deepEqual(await svc.inviteByEmail({ people: many }), { error: 'too-many-at-once', max: INVITE_BATCH });
  assert.deepEqual(await svc.inviteByEmail({ people: [{ email: 'not an address' }] }), { error: 'address' });
  assert.equal(out.filter((m) => m.type === 'contact.request').length, 2);
});

test('a send that fails says so for that person, and the others still go', async () => {
  const mail = { accounts: () => [{ id: 'a1', email: 'amira@shop.example' }], sendFrom: async ({ to }) => { if (to === 'chen@example.com') throw new Error('The server refused it.'); } };
  const { svc } = service({ mail, answer: introducer() });
  await online(svc);
  const r = await svc.inviteByEmail({ people: [{ email: 'chen@example.com' }, { email: 'dana@example.com' }], letter: { subject: 's', text: '{accept}' } });
  assert.deepEqual(r.results.map((x) => [x.told.sent, x.told.reason || null]), [[false, 'failed'], [true, null]]);
  assert.equal(r.results[0].told.detail, 'The server refused it.');
});
