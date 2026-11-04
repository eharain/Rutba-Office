// The sign-in clients a release carries: read out of Google's download at
// build time, and sent with each token request, the id with its secret.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createOAuthService, builtInClients, PROVIDERS } from '../apps/desktop/main/oauth.js';
import { googleClient } from '../apps/desktop/build/make-oauth-clients.js';
import { createMailService } from '../apps/desktop/main/mail.js';

const ID = '123456789012-abcdefghijklmnop.apps.googleusercontent.com';
const SECRET = 'GOCSPX-not-a-real-one';

const stores = (settings = {}, secrets = {}) => ({
  settings: { get: (k, d) => (k in settings ? settings[k] : d), set: (k, v) => { settings[k] = v; } },
  secrets: { get: (k) => secrets[k], set: (k, v) => { secrets[k] = v; }, delete: (k) => { delete secrets[k]; } },
});

/** Every token request the service makes, answered with a fresh access token. */
function recordFetch(t) {
  const seen = [];
  const was = globalThis.fetch;
  globalThis.fetch = async (url, init) => {
    seen.push({ url, body: Object.fromEntries(new URLSearchParams(init.body)) });
    return { ok: true, text: async () => JSON.stringify({ access_token: 'at', expires_in: 3600 }) };
  };
  t.after(() => { globalThis.fetch = was; });
  return seen;
}

test('a Desktop app download gives its id and secret; anything else is refused', () => {
  assert.deepEqual(googleClient(JSON.stringify({ installed: { client_id: ID, client_secret: SECRET, redirect_uris: ['http://localhost'] } })), { clientId: ID, clientSecret: SECRET });
  assert.throws(() => googleClient(JSON.stringify({ web: { client_id: ID, client_secret: SECRET } })), /Desktop app/);
  assert.throws(() => googleClient(JSON.stringify({ installed: { client_id: ID } })), /client_secret/);
  assert.throws(() => googleClient(JSON.stringify({ installed: { client_id: 'x', client_secret: SECRET } })), /client_id/);
});

test('the build\'s table is read from its file, and a missing or broken one is empty', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'oauth-clients-'));
  const file = path.join(dir, 'clients.json');
  fs.writeFileSync(file, JSON.stringify({ google: { clientId: ID, clientSecret: SECRET } }));
  assert.deepEqual(builtInClients(file), { google: { clientId: ID, clientSecret: SECRET } });
  assert.deepEqual(builtInClients(path.join(dir, 'none.json')), {});
  fs.writeFileSync(file, '{ not json');
  assert.deepEqual(builtInClients(file), {});
  assert.deepEqual(builtInClients(''), {});
  fs.rmSync(dir, { recursive: true, force: true });
});

test('a refresh sends the build\'s id with its secret', async (t) => {
  const seen = recordFetch(t);
  const svc = createOAuthService({ stores: stores({}, { 'oauth:google:a@gmail.com': 'rt' }), built: { google: { clientId: ID, clientSecret: SECRET } } });
  assert.equal(svc.provider({ email: 'a@gmail.com' }).configured, true);
  const { accessToken } = await svc.accessToken({ email: 'a@gmail.com' });
  assert.equal(accessToken, 'at');
  assert.equal(seen[0].url, PROVIDERS.google.token);
  assert.deepEqual(seen[0].body, { client_id: ID, client_secret: SECRET, refresh_token: 'rt', grant_type: 'refresh_token' });
});

test('an id set on this computer goes with its own secret, never the build\'s', async (t) => {
  const seen = recordFetch(t);
  const mine = 'mine-1.apps.googleusercontent.com';
  const secrets = { 'oauth:google:a@gmail.com': 'rt' };
  await createOAuthService({ stores: stores({ 'oauth.google.clientId': mine }, secrets), built: { google: { clientId: ID, clientSecret: SECRET } } }).accessToken({ email: 'a@gmail.com' });
  assert.deepEqual(seen[0].body, { client_id: mine, refresh_token: 'rt', grant_type: 'refresh_token' });
  await createOAuthService({ stores: stores({ 'oauth.google.clientId': mine, 'oauth.google.clientSecret': 's2' }, secrets), built: {} }).accessToken({ email: 'a@gmail.com' });
  assert.equal(seen[1].body.client_secret, 's2');
});

test('with no client anywhere, Gmail is not offered a sign-in', () => {
  assert.equal(createOAuthService({ stores: stores(), built: {} }).provider({ email: 'a@gmail.com' }).configured, false);
});

test('Gmail asks for the mailbox and who it belongs to', () => {
  assert.deepEqual(PROVIDERS.google.scope.split(' ').sort(), ['email', 'https://mail.google.com/', 'openid']);
});

test('removing a signed-in Gmail account deletes its token from this computer', () => {
  const settings = { 'mail.accounts': [
    { id: 'a1', email: 'A@gmail.com', provider: 'google', auth: 'oauth', imap: { auth: 'oauth' } },
    { id: 'a2', email: 'b@gmail.com', provider: 'google', auth: 'oauth', imap: { auth: 'oauth' } },
  ] };
  const secrets = { 'oauth:google:a@gmail.com': 'rt-a', 'oauth:google:b@gmail.com': 'rt-b' };
  const st = stores(settings, secrets);
  const oauth = createOAuthService({ stores: st, built: {} });
  const mail = createMailService({ stores: st, holdBlob: () => ({ url: '' }), broadcast: () => {}, userData: fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mail-oauth-')), oauth, autoFetch: false });
  mail.removeAccount({ id: 'a1' });
  assert.equal(secrets['oauth:google:a@gmail.com'], undefined);
  assert.equal(secrets['oauth:google:b@gmail.com'], 'rt-b');
});
