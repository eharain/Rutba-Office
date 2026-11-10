// A linked database's "encrypted connection" used to take any certificate at
// all. It is checked now, against a real handshake with a server whose
// certificate nobody vouches for.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import net from 'node:net';
import path from 'node:path';
import tls from 'node:tls';
import { tlsOptions, certificateProblem } from '../apps/desktop/main/engines/tls.js';
import { connect } from '../apps/desktop/main/engines/postgres.js';
import { makeIdentities } from './fixtures/smime-keys.js';

test('the options: checked by default, unchecked only when asked, nothing when not encrypted', () => {
  assert.deepEqual(tlsOptions({ ssl: true }), { rejectUnauthorized: true });
  assert.deepEqual(tlsOptions({ ssl: true, sslAnyCertificate: true }), { rejectUnauthorized: false });
  assert.equal(tlsOptions({ ssl: false, sslAnyCertificate: true }), false);
});

test('neither engine turns the check off by itself', () => {
  for (const f of ['postgres.js', 'mysql.js']) {
    const source = fs.readFileSync(path.join(import.meta.dirname, '..', 'apps', 'desktop', 'main', 'engines', f), 'utf8');
    assert.equal(/rejectUnauthorized/.test(source), false, f);
    assert.match(source, /tlsOptions\(\{ ssl, sslAnyCertificate \}\)/, f);
  }
});

test('the words say what to do, and only about certificates', () => {
  assert.match(certificateProblem({ code: 'SELF_SIGNED_CERT_IN_CHAIN', message: 'self-signed certificate in certificate chain' }, 'db.local'), /db\.local.*could not be checked.*Accept a certificate/s);
  assert.equal(certificateProblem({ code: 'ECONNREFUSED', message: 'connect ECONNREFUSED' }, 'db.local'), null);
});

/** A PostgreSQL server as far as the TLS handshake: says yes to SSL, offers a certificate, then hangs up. */
async function postgresThatEncrypts({ key, cert }) {
  let reached = false;
  const server = net.createServer((socket) => {
    socket.on('error', () => {});
    socket.once('data', () => {
      socket.write('S');
      const secure = new tls.TLSSocket(socket, { isServer: true, key, cert });
      secure.on('error', () => {});
      secure.on('secure', () => { reached = true; secure.destroy(); });
    });
  });
  await new Promise((r) => { server.listen(0, '127.0.0.1', r); });
  return { port: server.address().port, reached: () => reached, close: () => server.close() };
}

test('a server whose certificate cannot be checked is refused before anything is sent, and can be accepted', async () => {
  const { stranger } = makeIdentities();
  const server = await postgresThatEncrypts({ key: stranger.keyPem, cert: stranger.cert });
  const settings = { host: '127.0.0.1', port: server.port, database: 'x', user: 'x', ssl: true };
  try {
    await assert.rejects(() => connect(settings, 'a password'), /could not be checked/);
    assert.equal(server.reached(), false, 'the handshake never completed');
    await assert.rejects(() => connect({ ...settings, sslAnyCertificate: true }, 'a password'), (err) => !/could not be checked/.test(err.message));
    assert.equal(server.reached(), true, 'accepted, it went through');
  } finally {
    server.close();
  }
});
