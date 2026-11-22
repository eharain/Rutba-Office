// Office has no place to type a server (the owner, 22 Nov 2026: the apps are
// given their servers by Rutba's introducer, not told them by hand). Up to
// 2.26.1 Talk's sign-in had a box for a server address and a token in every
// copy, and the service took whatever was typed. The box and the typed
// address are for the window checks and the screens build only; a copy in
// use goes to Rutba's introducer, and drops an address an earlier copy kept.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createTalkService, SERVER_KEY, DEFAULT_SERVER } from '../apps/desktop/main/talk.js';

const folders = [];
test.after(() => { for (const f of folders) fs.rmSync(f, { recursive: true, force: true }); });

const TYPED = 'wss://someone-elses.example/ws';

function service({ testServers, settings: kept = new Map(), secrets: held = new Map() }) {
  const map = (m) => ({ get: (k, d) => (m.has(k) ? m.get(k) : d), set: (k, v) => m.set(k, v), delete: (k) => m.delete(k) });
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'talk-typed-'));
  folders.push(dataDir);
  const asked = [];
  const sent = [];
  const svc = createTalkService({
    stores: { settings: map(kept), secrets: map(held) },
    dataDir,
    ...(testServers === undefined ? {} : { testServers }),
    connect: (url) => {
      asked.push(url);
      const ws = { readyState: 1, send: (text) => {
        const msg = JSON.parse(text);
        sent.push(msg);
        if (msg.type === 'hello') setImmediate(() => ws.onmessage({ data: JSON.stringify({ type: 'welcome', me: { id: 'u-a', name: 'Amira Khan', email: 'amira@example.com' }, contacts: [], requests: { incoming: [], outgoing: [] } }) }));
      }, close: () => {} };
      setImmediate(() => ws.onopen());
      return ws;
    },
  });
  return { svc, asked, sent, kept, held };
}
const settle = async (svc, phase) => { for (let i = 0; i < 50 && svc.state().phase !== phase; i++) await new Promise((r) => setTimeout(r, 5)); };

test('a copy in use takes no typed server and no token given outright, and says to sign in with Rutba', () => {
  for (const testServers of [undefined, false]) {
    const { svc, asked, kept, held } = service({ testServers });
    assert.equal(svc.state().testServers, false);
    assert.throws(() => svc.signIn({ token: 'amira', server: TYPED }), /Sign in with Rutba/);
    assert.throws(() => svc.signIn({ token: 'amira' }), /Sign in with Rutba/);
    assert.deepEqual(asked, [], 'nothing was connected to');
    assert.equal(kept.has(SERVER_KEY), false, 'the address was not kept');
    assert.equal(held.size, 0, 'the token was not kept');
    assert.equal(svc.state().server, DEFAULT_SERVER);
    assert.equal(svc.state().phase, 'signed-out');
  }
});

test('a copy in use drops an address an earlier copy kept, with the token that went with it, and connects to nothing', async () => {
  const { svc, asked, sent, kept, held } = service({ testServers: false, settings: new Map([[SERVER_KEY, TYPED]]), secrets: new Map([['talk.token', 'typed-token'], ['talk.refresh', 'typed-refresh']]) });
  assert.equal(kept.get(SERVER_KEY), DEFAULT_SERVER);
  assert.equal(held.size, 0);
  svc.resume();
  await new Promise((r) => setTimeout(r, 30));
  assert.deepEqual(asked, [], 'the typed server is not connected to, and Rutba\'s introducer is not shown the old token');
  assert.deepEqual(sent, []);
  assert.equal(svc.state().phase, 'signed-out');
  assert.equal(svc.state().server, DEFAULT_SERVER);
});

test('a copy in use that signed in with Rutba keeps its sign-in', async () => {
  const { svc, asked, held } = service({ testServers: false, settings: new Map([[SERVER_KEY, DEFAULT_SERVER]]), secrets: new Map([['talk.token', 'amira']]) });
  assert.equal(held.get('talk.token'), 'amira');
  svc.resume();
  await settle(svc, 'online');
  assert.equal(svc.state().phase, 'online');
  assert.deepEqual(asked, [DEFAULT_SERVER]);
});

test('a check run may give a server and a token, and an address it kept stays', async () => {
  const { svc, asked, kept } = service({ testServers: true });
  assert.equal(svc.state().testServers, true);
  assert.throws(() => svc.signIn({ token: 'amira', server: 'https://not-a-socket.example' }), /starts with wss/);
  svc.signIn({ token: 'amira', server: 'ws://127.0.0.1:9/ws' });
  await settle(svc, 'online');
  assert.equal(svc.state().phase, 'online');
  assert.deepEqual(asked, ['ws://127.0.0.1:9/ws']);
  assert.equal(kept.get(SERVER_KEY), 'ws://127.0.0.1:9/ws');
  // Signed out, the screen still knows it may draw the box.
  assert.equal(svc.signOut().testServers, true);

  const again = service({ testServers: true, settings: new Map([[SERVER_KEY, 'ws://127.0.0.1:9/ws']]), secrets: new Map([['talk.token', 'amira']]) });
  assert.equal(again.kept.get(SERVER_KEY), 'ws://127.0.0.1:9/ws');
  assert.equal(again.held.get('talk.token'), 'amira');
});
