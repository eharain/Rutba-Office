// AI Tools over the network (apps/desktop/main/mcp/remote.js): answered only
// with the key, never to a web page, only while the tools are on and the
// network chosen, a session to each AI app that connects, and nothing
// listening once it is stopped.
import test from 'node:test';
import assert from 'node:assert/strict';
import { createRemote, newKey } from '../apps/desktop/main/mcp/remote.js';

const tools = {
  list: () => [{ name: 'list_folders', description: 'The folders', inputSchema: { type: 'object' } }],
  call: async (name, args, ctx) => ({ content: [{ type: 'text', text: `${name} from ${ctx?.where}` }] }),
};

test('the network: the key, no web pages, a session, the switch, and nothing left listening', async () => {
  let key = newKey();
  let open = true;
  const seen = new Map();
  const remote = createRemote({
    tools,
    version: '9.9.9',
    instructions: 'Test',
    allowed: () => (open ? { open: true } : { open: false, reason: 'The tools are off.' }),
    key: () => key,
    onSession: (id, entry) => (entry ? seen.set(id, entry) : seen.delete(id)),
  });
  const { port } = await remote.start(0);
  const url = `http://127.0.0.1:${port}/mcp`;
  const post = (body, headers = {}, at = url) => fetch(at, { method: 'POST', headers: { 'content-type': 'application/json', ...headers }, body: JSON.stringify(body) })
    .then(async (r) => ({ status: r.status, session: r.headers.get('mcp-session-id'), body: await r.json().catch(() => null) }));
  const init = { jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-06-18', clientInfo: { name: 'Far app', version: '1' } } };
  const auth = { authorization: `Bearer ${key}` };
  try {
    assert.equal((await post(init)).status, 401, 'no key');
    assert.equal((await post(init, { authorization: 'Bearer wrong' })).status, 401, 'a wrong key');
    assert.equal((await post(init, { ...auth, origin: 'https://example.com' })).status, 403, 'a web page');
    assert.equal((await post(init, auth, `http://127.0.0.1:${port}/other`)).status, 404, 'another path');
    assert.equal((await fetch(url, { headers: auth })).status, 405, 'no stream to open');

    const hello = await post(init, auth);
    assert.equal(hello.status, 200);
    assert.equal(hello.body.result.serverInfo.version, '9.9.9');
    assert.ok(hello.session);
    assert.equal([...seen.values()][0].client, 'Far app');
    assert.equal([...seen.values()][0].where, '127.0.0.1');

    const call = await post({ jsonrpc: '2.0', id: 2, method: 'tools/call', params: { name: 'list_folders', arguments: {} } }, { ...auth, 'mcp-session-id': hello.session });
    assert.equal(call.body.result.content[0].text, 'list_folders from 127.0.0.1');
    assert.equal((await post({ jsonrpc: '2.0', method: 'notifications/initialized' }, { ...auth, 'mcp-session-id': hello.session })).status, 202, 'a notification');
    assert.equal((await post({ jsonrpc: '2.0', id: 3, method: 'ping' }, { ...auth, 'mcp-session-id': 'gone' })).status, 404, 'a session that has ended');

    open = false;
    const off = await post({ jsonrpc: '2.0', id: 4, method: 'ping' }, { ...auth, 'mcp-session-id': hello.session });
    assert.equal(off.status, 403);
    assert.match(off.body.error.message, /off/);
    open = true;

    key = newKey();
    assert.equal((await post(init, auth)).status, 401, 'the old key, once renewed');

    // Ten wrong keys in a minute and the address is turned away.
    for (let i = 0; i < 10; i++) await post(init, { authorization: 'Bearer nope' });
    assert.equal((await post(init, { authorization: `Bearer ${key}` })).status, 429);
  } finally {
    await remote.stop();
  }
  assert.equal(seen.size, 0, 'the sessions go when it stops');
  await assert.rejects(fetch(url, { method: 'POST', body: '{}' }), 'nothing listens once stopped');
});
