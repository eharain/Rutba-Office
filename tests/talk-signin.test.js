// Talk's Rutba sign-in where it goes wrong or is given up
// (main/talk-signin.js, and the service that runs it): no connection is said
// in words and never as the network's own error, and the wait for the browser
// can be cancelled, or the browser opened again. Run on a stand-in for Rutba
// sign-in; the only thing reached is the listener on this computer.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import http from 'node:http';
import net from 'node:net';
import { discover, signInWithBrowser, refreshTokens, UNREACHED, CLIENT_ID } from '../apps/desktop/main/talk-signin.js';
import { createTalkService } from '../apps/desktop/main/talk.js';
import { METHODS } from '@rutba/office-shell/contract';

const folders = [];
test.after(() => { for (const f of folders) fs.rmSync(f, { recursive: true, force: true }); });

const ISSUER = 'https://auth.example';
const WELL_KNOWN = '/.well-known/openid-configuration';
const json = (body) => ({ ok: true, status: 200, json: async () => body });

/** Rutba sign-in standing in: its discovery document for whoever asks, and tokens for a code. */
function provider() {
  const asked = [];
  const fetchImpl = async (url, init) => {
    const at = String(url);
    asked.push({ url: at, init });
    if (at.endsWith(WELL_KNOWN)) {
      const issuer = at.slice(0, -WELL_KNOWN.length);
      return json({ issuer, authorization_endpoint: `${issuer}/authorize`, token_endpoint: `${issuer}/token` });
    }
    if (at.endsWith('/token')) return json({ access_token: 'a.b.c', refresh_token: 'r1', expires_in: 600 });
    throw new Error('not asked for');
  };
  return { fetchImpl, asked };
}
/** No connection at all, as the network says it. */
const offline = async () => { throw new TypeError('fetch failed'); };
/** The browser coming back to the listener Office holds. */
const get = (url) => new Promise((resolve, reject) => { http.get(url, (res) => { res.resume(); res.on('end', () => resolve(res.statusCode)); }).on('error', reject); });
/** Whether nothing listens on a port of this computer any more. */
const free = (port) => new Promise((resolve) => {
  const s = net.createServer();
  s.once('error', () => resolve(false));
  s.listen(port, '127.0.0.1', () => s.close(() => resolve(true)));
});

/** Talk's service with nobody signed in, never connected. */
function talk({ fetchImpl, openBrowser }) {
  const map = (m) => ({ get: (k, d) => (m.has(k) ? m.get(k) : d), set: (k, v) => m.set(k, v), delete: (k) => m.delete(k) });
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'talk-signin-'));
  folders.push(dir);
  const told = [];
  const svc = createTalkService({
    stores: { settings: map(new Map()), secrets: map(new Map()) },
    dataDir: dir,
    broadcast: (event, payload) => told.push({ event, payload }),
    connect: () => { throw new Error('nothing connects in this check'); },
    openBrowser,
    fetchImpl,
  });
  return { svc, told };
}

test('with no connection, sign-in says so in words and never as the network error itself', async () => {
  const runs = [
    () => discover(ISSUER, { fetchImpl: offline }),
    () => signInWithBrowser({ issuer: ISSUER, open: () => {}, fetchImpl: offline }),
    () => refreshTokens({ issuer: ISSUER, refreshToken: 'r', fetchImpl: offline }),
  ];
  for (const run of runs) {
    await assert.rejects(run, (err) => {
      assert.equal(err.message, UNREACHED);
      assert.equal(err.code, 'unreachable');
      assert.ok(!/fetch failed/i.test(err.message));
      assert.ok(err.cause instanceof TypeError, 'what the network said is kept for the log');
      return true;
    });
  }
  assert.ok(UNREACHED.includes('online') && UNREACHED.includes('Sign in with Rutba'), 'what to check, and what to press');
  // Something else answering in its place (a network's own sign-in page) is not Rutba sign-in reached.
  const page = async () => ({ ok: true, status: 200, json: async () => { throw new SyntaxError('Unexpected token <'); } });
  await assert.rejects(() => discover(ISSUER, { fetchImpl: page }), { message: UNREACHED, code: 'unreachable' });
  // The connection going after the browser came back is said the same way.
  let n = 0;
  const cut = async (url, init) => (n++ === 0 ? provider().fetchImpl(url, init) : offline());
  await assert.rejects(() => refreshTokens({ issuer: ISSUER, refreshToken: 'r', fetchImpl: cut }), { message: UNREACHED });
});

test('the browser comes back with the code, and it is exchanged for the tokens', async () => {
  const { fetchImpl, asked } = provider();
  let opened = null;
  const tokens = await signInWithBrowser({
    issuer: ISSUER,
    fetchImpl,
    open: async (url) => {
      opened = new URL(url);
      const back = new URL(opened.searchParams.get('redirect_uri'));
      back.searchParams.set('code', 'the-code');
      back.searchParams.set('state', opened.searchParams.get('state'));
      setImmediate(() => get(back.href).catch(() => {}));
    },
  });
  assert.equal(tokens.refresh_token, 'r1');
  assert.equal(opened.origin + opened.pathname, `${ISSUER}/authorize`);
  assert.equal(opened.searchParams.get('client_id'), CLIENT_ID);
  const exchange = asked.find((a) => a.url === `${ISSUER}/token`);
  assert.equal(exchange.init.body.get('code'), 'the-code');
  assert.equal(exchange.init.body.get('grant_type'), 'authorization_code');
});

test('the wait for the browser can be given up: the listener closes and nothing is exchanged', async () => {
  const { fetchImpl, asked } = provider();
  const stop = new AbortController();
  let redirect = null;
  const waiting = signInWithBrowser({
    issuer: ISSUER,
    fetchImpl,
    signal: stop.signal,
    open: (url) => { redirect = new URL(new URL(url).searchParams.get('redirect_uri')); setImmediate(() => stop.abort()); },
  });
  await assert.rejects(waiting, { code: 'cancelled', message: 'Sign-in was cancelled.' });
  assert.equal(redirect.hostname, '127.0.0.1');
  assert.equal(await free(Number(redirect.port)), true, 'nothing listens for the answer any more');
  assert.equal(asked.some((a) => a.url.endsWith('/token')), false);
  // Given up before the browser was opened.
  const before = new AbortController();
  before.abort();
  let opened = false;
  await assert.rejects(() => signInWithBrowser({ issuer: ISSUER, fetchImpl, signal: before.signal, open: () => { opened = true; } }), { code: 'cancelled' });
  assert.equal(opened, false);
});

test('Cancel while the browser is open puts the sign-in screen back as it was, with nothing said', async () => {
  const opened = [];
  const { svc, told } = talk({ fetchImpl: provider().fetchImpl, openBrowser: async (url) => { opened.push(url); } });
  assert.deepEqual(await svc.signInReopen(), { opened: false }, 'nothing waits yet');
  const waiting = svc.signInBrowser();
  waiting.catch(() => {});
  for (let i = 0; i < 500 && !opened.length; i++) await new Promise((r) => setImmediate(r));
  assert.equal(opened.length, 1);
  assert.equal(svc.state().phase, 'signing-in');
  // The browser's tab was closed: opened again on the same sign-in.
  assert.deepEqual(await svc.signInReopen(), { opened: true });
  assert.deepEqual(opened, [opened[0], opened[0]]);

  const before = told.length;
  assert.equal(svc.signInCancel().phase, 'signed-out', 'said at once');
  assert.ok(told.slice(before).some((t) => t.event === 'talk:state' && t.payload.phase === 'signed-out'), 'and every window is told');
  await assert.rejects(waiting, { code: 'cancelled' });
  assert.deepEqual([svc.state().phase, svc.state().error], ['signed-out', '']);
  assert.equal(await free(Number(new URL(new URL(opened[0]).searchParams.get('redirect_uri')).port)), true);
  assert.deepEqual(await svc.signInReopen(), { opened: false }, 'nothing waits any more');
  assert.equal(svc.signInCancel().phase, 'signed-out', 'cancelling nothing changes nothing');
});

test('asking to sign in again while one waits lets the earlier one go, and only the newer one is waited for', async () => {
  const opened = [];
  const { svc } = talk({ fetchImpl: provider().fetchImpl, openBrowser: async (url) => { opened.push(url); } });
  const first = svc.signInBrowser();
  first.catch(() => {});
  for (let i = 0; i < 500 && opened.length < 1; i++) await new Promise((r) => setImmediate(r));
  const second = svc.signInBrowser();
  second.catch(() => {});
  await assert.rejects(first, { code: 'cancelled' });
  for (let i = 0; i < 500 && opened.length < 2; i++) await new Promise((r) => setImmediate(r));
  assert.equal(svc.state().phase, 'signing-in', 'still waiting, for the second');
  assert.notEqual(opened[0], opened[1]);
  svc.signInCancel();
  await assert.rejects(second, { code: 'cancelled' });
  assert.equal(svc.state().phase, 'signed-out');
});

test('with no connection the sign-in screen is told what to check', async () => {
  const { svc } = talk({ fetchImpl: offline, openBrowser: async () => {} });
  await assert.rejects(() => svc.signInBrowser(), { message: UNREACHED });
  assert.deepEqual([svc.state().phase, svc.state().error], ['signed-out', UNREACHED]);
});

test('the windows can ask for each of these', () => {
  for (const name of ['signInBrowser', 'signInCancel', 'signInReopen', 'messageAgain', 'messageForget']) assert.ok(METHODS.talk.includes(name), name);
});
