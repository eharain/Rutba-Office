// Connect (Office 2.25): someone who already uses Talk is asked at once, in
// Talk, with no email. Where a person stands comes from what this copy has
// seen, never from asking the introducer whether an address has an account.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createTalkService } from '../apps/desktop/main/talk.js';

const folders = [];
test.after(() => { for (const f of folders) fs.rmSync(f, { recursive: true, force: true }); });

const WELCOME = {
  type: 'welcome',
  me: { id: 'u-a', name: 'Amira Khan', email: 'amira@example.com' },
  contacts: [{ id: 'u-ben', name: 'Ben Ortiz', email: 'Ben@Example.com', presence: 'online' }, { id: 'u-grace', name: 'Grace Okafor', email: 'grace@example.com' }],
  requests: { incoming: [{ id: 'in-1', from: { id: 'u-hana', name: 'Hana Ito', email: 'hana@example.com' }, note: '', at: 1 }], outgoing: [{ id: 'out-1', email: 'ivan@example.com', at: 1 }] },
};

function service({ answer, settings: kept = new Map() }) {
  const map = (m) => ({ get: (k, d) => (m.has(k) ? m.get(k) : d), set: (k, v) => m.set(k, v), delete: (k) => m.delete(k) });
  const settings = map(kept);
  const secrets = map(new Map([['talk.token', 'amira']]));
  const out = [];
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'talk-connect-'));
  folders.push(dataDir);
  let ws = null;
  const say = (m) => setImmediate(() => ws.onmessage({ data: JSON.stringify(m) }));
  const svc = createTalkService({
    stores: { settings, secrets },
    dataDir,
    connect: () => {
      ws = { readyState: 1, send: (text) => {
        const msg = JSON.parse(text);
        out.push(msg);
        if (msg.type === 'hello') return void say(WELCOME);
        for (const reply of [].concat(answer(msg) || [])) say(reply);
      }, close: () => {} };
      setImmediate(() => ws.onopen());
      return ws;
    },
  });
  return { svc, out, say, settings };
}
const online = async (svc) => {
  svc.resume();
  for (let i = 0; i < 50 && svc.state().phase !== 'online'; i++) await new Promise((r) => setTimeout(r, 5));
  assert.equal(svc.state().phase, 'online');
};
const tick = () => new Promise((r) => setTimeout(r, 10));

test('where each person stands: a contact, someone who asked, someone asked, someone seen using Talk, or nobody known', async () => {
  const { svc, say } = service({ answer: (m) => (m.type === 'invite.peek' ? { type: 'invite.peeked', ref: m.ref, code: m.code, from: { id: 'u-jo', name: 'Jo', email: 'jo@example.com' }, note: '', kind: 'link' } : null) });
  await online(svc);
  // Someone who became a contact and was later taken off is still known to use Talk.
  say({ type: 'contact.added', contact: { id: 'u-kai', name: 'Kai', email: 'kai@example.com' } });
  await tick();
  say({ type: 'contact.removed', id: 'u-kai' });
  await svc.peekInvite({ link: 'AbCdEfGh12345678_-xy' });
  await tick();
  assert.deepEqual(svc.reach({ emails: ['ben@example.com', ' HANA@example.com', 'ivan@example.com', 'kai@example.com', 'jo@example.com', 'lee@example.com', ''] }), {
    'ben@example.com': 'contact',
    'hana@example.com': 'asked-you',
    'ivan@example.com': 'asked',
    'kai@example.com': 'uses-talk',
    'jo@example.com': 'uses-talk',
    'lee@example.com': null,
  });
});

test('what was seen is kept on this computer, so it is known after a restart', async () => {
  const kept = new Map();
  const first = service({ answer: () => null, settings: kept });
  await online(first.svc);
  assert.ok((kept.get('talk.known') || []).includes('ben@example.com'));
  assert.ok(kept.get('talk.known').includes('hana@example.com'));
  assert.ok(!kept.get('talk.known').includes('ivan@example.com'), 'someone asked is not known to use Talk');
});

test('Connect asks at once, in Talk, with no email, and the request says so; a contact already comes back as one', async () => {
  const { svc, out } = service({ answer: (m) => {
    if (m.type !== 'contact.request') return null;
    if (m.email === 'hana@example.com') return { type: 'contact.request.sent', ref: m.ref, request: null, contact: { id: 'u-hana', name: 'Hana Ito', email: 'hana@example.com' } };
    return { type: 'contact.request.sent', ref: m.ref, request: { id: 'req-9', email: m.email, note: m.note, at: 2, link: 'https://introducer.rutba.io/i/AbCdEfGh12345678_-xy' } };
  } });
  await online(svc);
  const r = await svc.connect({ email: 'kai@example.com' });
  const sent = out.find((m) => m.type === 'contact.request');
  assert.equal(sent.notify, 'none');
  assert.equal(sent.note, '');
  assert.deepEqual(r.told, { by: 'talk', sent: true });
  assert.deepEqual(svc.state().requests.outgoing.find((o) => o.id === 'req-9').told.by, 'talk');
  // Hana had asked first: the introducer makes them contacts.
  assert.deepEqual(await svc.connect({ email: 'hana@example.com' }), { contact: { id: 'u-hana', name: 'Hana Ito', email: 'hana@example.com' } });
  assert.deepEqual(await svc.connect({ email: 'not an address' }), { error: 'address' });
});

test('a refusal comes back as its code', async () => {
  const { svc } = service({ answer: (m) => (m.type === 'contact.request' ? { type: 'error', for: 'contact.request', code: 'too-many' } : null) });
  await online(svc);
  assert.deepEqual(await svc.connect({ email: 'kai@example.com' }), { error: 'too-many' });
});

test('connections: the people one is in contact with, for Calendar and others to pick from', async () => {
  const { svc } = service({ answer: () => null });
  await online(svc);
  assert.deepEqual(svc.connections(), [
    { id: 'u-ben', name: 'Ben Ortiz', email: 'Ben@Example.com', presence: 'online' },
    { id: 'u-grace', name: 'Grace Okafor', email: 'grace@example.com', presence: 'offline' },
  ]);
});
