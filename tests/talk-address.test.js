// My address in Talk (Office 2.25): which of the person's proved addresses
// Talk shows and finds them by. The live introducer keeps only the sign-in
// address; a server that keeps more says so with me.addresses, and the
// client speaks to it in address.add, .confirm, .set and .remove.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createTalkService } from '../apps/desktop/main/talk.js';

const folders = [];
test.after(() => { for (const f of folders) fs.rmSync(f, { recursive: true, force: true }); });

const SIGN_IN = { address: 'amira@example.com', signIn: true, findable: true, shown: true, verified: 1 };

/** A fake introducer; with `keeps`, one that keeps several addresses, its codes in `codes`. */
function service({ keeps = false } = {}) {
  const map = (m) => ({ get: (k, d) => (m.has(k) ? m.get(k) : d), set: (k, v) => m.set(k, v), delete: (k) => m.delete(k) });
  const out = [];
  const codes = new Map();
  let list = [{ ...SIGN_IN }];
  const me = () => ({ id: 'u-a', name: 'Amira Khan', email: list.find((a) => a.shown).address, ...(keeps ? { addresses: list.map((a) => ({ ...a })) } : {}) });
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'talk-address-'));
  folders.push(dataDir);
  let ws = null;
  const answer = (m) => {
    const fail = (code, extra = {}) => ({ type: 'error', for: m.type, code, ref: m.ref, ...extra });
    if (!keeps && m.type.startsWith('address.')) return fail('unknown');
    if (m.type === 'address.add') {
      if (list.some((a) => a.address === m.address)) return fail('taken');
      codes.set(m.address, { code: '482913', tries: 3 });
      return { type: 'address.code.sent', ref: m.ref, address: m.address, expires: 900000 };
    }
    if (m.type === 'address.confirm') {
      const c = codes.get(m.address);
      if (!c) return fail('expired');
      if (c.code !== m.code) { c.tries -= 1; return fail('bad-code', { left: c.tries }); }
      codes.delete(m.address);
      list.push({ address: m.address, signIn: false, findable: true, shown: false, verified: 2 });
      return { type: 'me', ref: m.ref, me: me() };
    }
    if (m.type === 'address.set') {
      if (m.shown) list = list.map((a) => ({ ...a, shown: a.address === m.address }));
      if (typeof m.findable === 'boolean') list = list.map((a) => (a.address === m.address ? { ...a, findable: m.findable } : a));
      return { type: 'me', ref: m.ref, me: me() };
    }
    if (m.type === 'address.remove') {
      if (list.find((a) => a.address === m.address)?.signIn) return fail('sign-in');
      const wasShown = list.find((a) => a.address === m.address)?.shown;
      list = list.filter((a) => a.address !== m.address);
      if (wasShown) list = list.map((a) => ({ ...a, shown: a.signIn }));
      return { type: 'me', ref: m.ref, me: me() };
    }
    return null;
  };
  const svc = createTalkService({
    stores: { settings: map(new Map()), secrets: map(new Map([['talk.token', 'amira']])) },
    dataDir,
    connect: () => {
      ws = { readyState: 1, send: (text) => {
        const msg = JSON.parse(text);
        out.push(msg);
        if (msg.type === 'hello') return void setImmediate(() => ws.onmessage({ data: JSON.stringify({ type: 'welcome', me: me(), contacts: [], requests: { incoming: [], outgoing: [] } }) }));
        const reply = answer(msg);
        if (reply) setImmediate(() => ws.onmessage({ data: JSON.stringify(reply) }));
      }, close: () => {} };
      setImmediate(() => ws.onopen());
      return ws;
    },
  });
  return { svc, out, codes };
}
const online = async (svc) => {
  svc.resume();
  for (let i = 0; i < 50 && svc.state().phase !== 'online'; i++) await new Promise((r) => setTimeout(r, 5));
  assert.equal(svc.state().phase, 'online');
};

test('an introducer that keeps only the sign-in address: that one is shown and finds you, and nothing else is asked of it', async () => {
  const { svc, out } = service();
  await online(svc);
  assert.deepEqual(svc.addresses(), { supported: false, list: [{ address: 'amira@example.com', signIn: true, findable: true, shown: true }] });
  for (const call of [() => svc.addAddress({ address: 'amira@shop.example' }), () => svc.confirmAddress({ address: 'amira@shop.example', code: '1' }), () => svc.setAddress({ address: 'amira@example.com', shown: true }), () => svc.removeAddress({ address: 'amira@shop.example' })]) {
    assert.deepEqual(await call(), { error: 'unsupported' });
  }
  assert.equal(out.filter((m) => m.type.startsWith('address.')).length, 0);
});

test('signed out, there is no address to show', () => {
  const { svc } = service();
  assert.deepEqual(svc.addresses(), { supported: false, list: [] });
});

test('another address is proved with the code sent to it, then shown, kept from finding you, and taken off', async () => {
  const { svc, out, codes } = service({ keeps: true });
  await online(svc);
  assert.equal(svc.addresses().supported, true);
  assert.deepEqual(await svc.addAddress({ address: 'not an address' }), { error: 'address' });
  assert.deepEqual(await svc.addAddress({ address: ' Amira@Shop.example ' }), { sent: true, address: 'amira@shop.example', expires: 900000 });
  assert.equal(codes.get('amira@shop.example').code, '482913');
  assert.deepEqual(await svc.confirmAddress({ address: 'amira@shop.example', code: '000000' }), { error: 'bad-code', left: 2 });
  const proved = await svc.confirmAddress({ address: 'amira@shop.example', code: ' 482913 ' });
  assert.deepEqual(proved.addresses.map((a) => [a.address, a.signIn, a.findable, a.shown]), [['amira@example.com', true, true, true], ['amira@shop.example', false, true, false]]);

  // Shown: the people she talks with see that one; Talk's own state says so.
  await svc.setAddress({ address: 'amira@shop.example', shown: true });
  assert.equal(svc.state().me.email, 'amira@shop.example');
  assert.deepEqual(svc.addresses().list.filter((a) => a.shown).map((a) => a.address), ['amira@shop.example']);
  // The sign-in address kept from finding her.
  await svc.setAddress({ address: 'amira@example.com', findable: false });
  assert.equal(svc.addresses().list.find((a) => a.signIn).findable, false);
  assert.deepEqual(out.filter((m) => m.type === 'address.set').map(({ type, ref, ...rest }) => rest), [{ address: 'amira@shop.example', shown: true }, { address: 'amira@example.com', findable: false }]);

  // The sign-in address stays; the other goes, and the sign-in address is shown again.
  assert.deepEqual(await svc.removeAddress({ address: 'amira@example.com' }), { error: 'sign-in' });
  const after = await svc.removeAddress({ address: 'amira@shop.example' });
  assert.deepEqual(after.addresses.map((a) => a.address), ['amira@example.com']);
  assert.equal(svc.state().me.email, 'amira@example.com');
});

test('an address already proved is refused as the introducer says', async () => {
  const { svc } = service({ keeps: true });
  await online(svc);
  assert.deepEqual(await svc.addAddress({ address: 'amira@example.com' }), { error: 'taken' });
  assert.deepEqual(await svc.confirmAddress({ address: 'never@example.com', code: '1' }), { error: 'expired' });
});
