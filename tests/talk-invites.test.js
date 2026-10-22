// Talk invitations (D-2026-10-21-27 and -28): the links Office takes, the
// codes it reads from what someone pastes, the letter it emails, and the
// Talk service asking the introducer and telling the person how it went.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { parseLink, LINK_SCHEME } from '@rutba/office-shell/links';
import { createTalkService, inviteCode, letterHtml, DOWNLOAD_URL } from '../apps/desktop/main/talk.js';

const CODE = 'AbCdEfGh12345678_-xy';

test('only a Talk invitation is a link Office acts on', () => {
  assert.equal(LINK_SCHEME, 'rutba-office');
  assert.deepEqual(parseLink(`rutba-office://talk/invite/${CODE}`), { app: 'talk', query: { invite: CODE } });
  assert.deepEqual(parseLink(`  rutba-office://talk/invite/${CODE}/ `), { app: 'talk', query: { invite: CODE } });
  for (const bad of ['rutba-office://word/open/C:/x.docx', `rutba-office://talk/invite/${CODE}/../x`, 'rutba-office://talk/invite/short', `https://introducer.rutba.io/i/${CODE}`, '', null]) {
    assert.equal(parseLink(bad), null, String(bad));
  }
});

test('an invitation code is read from the page, the Office link or the code itself', () => {
  assert.equal(inviteCode(`https://introducer.rutba.io/i/${CODE}`), CODE);
  assert.equal(inviteCode(`http://127.0.0.1:5000/i/${CODE}?from=mail#top`), CODE);
  assert.equal(inviteCode(`rutba-office://talk/invite/${CODE}`), CODE);
  assert.equal(inviteCode(` ${CODE} `), CODE);
  for (const bad of ['https://introducer.rutba.io/i/short', 'https://x.example/other/' + CODE, 'two words', '']) assert.equal(inviteCode(bad), null, bad);
});

test('the letter goes as HTML too, its addresses links and its words escaped', () => {
  const html = letterHtml('Hello <b>you</b>,\n\nAccept: https://introducer.rutba.io/i/abc\nthen talk');
  assert.match(html, /<p>Hello &lt;b&gt;you&lt;\/b&gt;,<\/p>/);
  assert.match(html, /<a href="https:\/\/introducer\.rutba\.io\/i\/abc">https:\/\/introducer\.rutba\.io\/i\/abc<\/a><br>then talk/);
});

/** The Talk service on a fake introducer that answers as `answer(msg)` says. */
function service({ answer, accounts = [], sent = [] }) {
  const map = (m) => ({ get: (k, d) => (m.has(k) ? m.get(k) : d), set: (k, v) => m.set(k, v), delete: (k) => m.delete(k) });
  const settings = map(new Map());
  const secrets = map(new Map([['talk.token', 'amira']]));
  const out = [];
  let ws = null;
  const svc = createTalkService({
    stores: { settings, secrets },
    dataDir: fs.mkdtempSync(path.join(os.tmpdir(), 'talk-invites-')),
    connect: () => {
      ws = { readyState: 1, send: (text) => {
        const msg = JSON.parse(text);
        out.push(msg);
        if (msg.type === 'hello') return void setImmediate(() => ws.onmessage({ data: JSON.stringify({ type: 'welcome', me: { id: 'u-a', name: 'Amira Khan', email: 'amira@example.com' }, contacts: [], requests: { incoming: [], outgoing: [] }, mailAllowance: { day: { used: 0, limit: 10 }, week: { used: 0, limit: 30 } } }) }));
        for (const reply of [].concat(answer(msg) || [])) setImmediate(() => ws.onmessage({ data: JSON.stringify(reply) }));
      }, close: () => {} };
      setImmediate(() => ws.onopen());
      return ws;
    },
    mail: { accounts: () => accounts, send: async (args) => { sent.push(args); return { messageId: 'x' }; } },
  });
  return { svc, out, settings };
}
const online = async (svc) => {
  svc.resume();
  for (let i = 0; i < 50 && svc.state().phase !== 'online'; i++) await new Promise((r) => setTimeout(r, 5));
  assert.equal(svc.state().phase, 'online');
};
const request = (msg, extra = {}) => ({ type: 'contact.request.sent', ref: msg.ref, request: { id: 'req-1', email: msg.email, note: msg.note, at: 1, link: `https://introducer.rutba.io/i/${CODE}` }, ...extra });

test('from the person\'s own address: asked without Rutba\'s email, the letter sent from that account with both links filled, and the request says so', async () => {
  const sent = [];
  const { svc, out } = service({ sent, accounts: [{ id: 'acc1', email: 'amira@shop.example' }, { id: 'old', email: 'archive.mbox', local: true }], answer: (m) => (m.type === 'contact.request' ? request(m) : null) });
  await online(svc);
  const r = await svc.invite({ email: 'chen@example.com', note: 'Hi', how: 'own', accountId: 'acc1', letter: { subject: 'Talk?', text: 'Accept: {accept}\n\nJoin: {join}' } });
  assert.equal(out.find((m) => m.type === 'contact.request').notify, 'none');
  // Kept with the account it went from, so Invite again sends from the same one.
  assert.deepEqual(r.told, { by: 'own', from: 'amira@shop.example', accountId: 'acc1', sent: true });
  assert.equal(sent.length, 1);
  assert.equal(sent[0].accountId, 'acc1');
  assert.equal(sent[0].draft.to, 'chen@example.com');
  assert.equal(sent[0].draft.text, `Accept: https://introducer.rutba.io/i/${CODE}\n\nJoin: ${DOWNLOAD_URL}?invite=${CODE}`);
  assert.match(sent[0].draft.html, /<a href=/);
  // Kept beside the request, so every window shows how it went.
  assert.deepEqual(svc.state().requests.outgoing[0].told.by, 'own');
  // An imported archive cannot send.
  const archived = await svc.invite({ email: 'dee@example.com', how: 'own', accountId: 'old' });
  assert.equal(archived.told.reason, 'no-account');
});

test('from Rutba: the introducer is asked to email, and its answer, limit and all, is what the person sees', async () => {
  const allowance = { day: { used: 10, limit: 10 }, week: { used: 12, limit: 30 } };
  const { svc, out } = service({ answer: (m) => (m.type === 'contact.request' ? request(m, { mail: { by: 'rutba', sent: false, reason: 'limit', allowance } }) : null) });
  await online(svc);
  const r = await svc.invite({ email: 'dana@example.com', how: 'rutba' });
  assert.equal(out.find((m) => m.type === 'contact.request').notify, 'rutba');
  assert.deepEqual(r.told, { by: 'rutba', sent: false, reason: 'limit' });
  assert.deepEqual(svc.state().mailAllowance, allowance);
});

test('a refusal comes back as its code, and an address that is a contact already as the contact', async () => {
  const { svc } = service({ answer: (m) => (m.email === 'me@example.com' ? { type: 'error', for: 'contact.request', code: 'yourself' } : { type: 'contact.request.sent', ref: m.ref, request: null, contact: { id: 'u-b', name: 'Ben' } }) });
  await online(svc);
  assert.deepEqual(await svc.invite({ email: 'me@example.com', how: 'none' }), { error: 'yourself' });
  assert.deepEqual(await svc.invite({ email: 'ben@example.com', how: 'none' }), { contact: { id: 'u-b', name: 'Ben' } });
});

test('links: made, looked at and accepted by their code, withdrawn', async () => {
  const invite = { code: CODE, url: `https://introducer.rutba.io/i/${CODE}`, note: 'Lunch', uses: 0, maxUses: 20 };
  const { svc, out } = service({ answer: (m) => ({
    'invite.create': { type: 'invite.created', ref: m.ref, invite },
    'invite.peek': { type: 'invite.peeked', ref: m.ref, code: m.code, from: { id: 'u-b', name: 'Ben Ortiz' }, note: 'Lunch', kind: 'link' },
    'invite.accept': { type: 'contact.added', contact: { id: 'u-b', name: 'Ben Ortiz' } },
    'invite.revoke': { type: 'invite.revoked', ref: m.ref, code: m.code },
  })[m.type] });
  await online(svc);
  assert.deepEqual(await svc.inviteLink({ note: 'Lunch' }), { invite });
  assert.equal(svc.state().invites.length, 1);
  const peeked = await svc.peekInvite({ link: `rutba-office://talk/invite/${CODE}` });
  assert.equal(peeked.from.name, 'Ben Ortiz');
  assert.equal(out.find((m) => m.type === 'invite.peek').code, CODE);
  assert.equal((await svc.acceptInvite({ link: invite.url })).contact.id, 'u-b');
  assert.deepEqual(await svc.revokeInvite({ code: CODE }), { revoked: true });
  assert.equal(svc.state().invites.length, 0);
  assert.deepEqual(await svc.peekInvite({ link: 'not a link' }), { error: 'no-invite' });
});
