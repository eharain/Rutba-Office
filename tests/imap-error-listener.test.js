// A server that drops the connection is reported to the command waiting on it, never thrown in the
// main process: every IMAP client Mail makes listens for its 'error' event before it connects.
// Found by the full gate on 11 Nov 2026 ("read ECONNRESET", uncaught, as a test server closed).
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import net from 'node:net';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const MAIN = path.join(here, '..', 'apps', 'desktop', 'main');

test('every ImapFlow client in Mail\'s main process listens for errors before it connects', () => {
  for (const name of fs.readdirSync(MAIN).filter((f) => f.endsWith('.js') && !f.startsWith('verify-'))) {
    const src = fs.readFileSync(path.join(MAIN, name), 'utf8');
    let at = src.indexOf('new ImapFlow(');
    while (at >= 0) {
      const connect = src.indexOf('.connect()', at);
      assert.ok(connect > at, `${name}: a client made at ${at} is connected`);
      assert.match(src.slice(at, connect), /\.on\('error'/, `${name}: the client made at ${at} listens for 'error' before connect()`);
      at = src.indexOf('new ImapFlow(', at + 1);
    }
  }
});

test('a connection reset while the client is signed in is not thrown in the process', async () => {
  const { ImapFlow } = await import('imapflow');
  // A server that greets, signs in and then resets the connection.
  const sockets = new Set();
  const server = net.createServer((s) => {
    sockets.add(s);
    s.on('error', () => {});
    s.write('* OK [CAPABILITY IMAP4rev1 AUTH=PLAIN] ready\r\n');
    s.on('data', (d) => {
      for (const line of String(d).split('\r\n').filter(Boolean)) {
        const [tag, cmd] = line.split(' ');
        if (/^(LOGIN|AUTHENTICATE)$/i.test(cmd)) s.write(`${tag} OK [CAPABILITY IMAP4rev1] signed in\r\n`);
        else if (/^CAPABILITY$/i.test(cmd)) s.write(`* CAPABILITY IMAP4rev1\r\n${tag} OK done\r\n`);
        else if (/^ID$/i.test(cmd)) s.write(`* ID NIL\r\n${tag} OK done\r\n`);
        else if (/^(NAMESPACE|ENABLE|COMPRESS)$/i.test(cmd)) s.write(`${tag} BAD no\r\n`);
        else s.write(`${tag} OK done\r\n`);
      }
    });
  });
  await new Promise((r) => server.listen(0, '127.0.0.1', r));
  const uncaught = [];
  const onUncaught = (err) => uncaught.push(err);
  process.on('uncaughtException', onUncaught);
  try {
    const client = new ImapFlow({ host: '127.0.0.1', port: server.address().port, secure: false, auth: { user: 'a', pass: 'b' }, logger: false });
    client.on('error', () => {});
    await client.connect();
    const closed = new Promise((r) => client.once('close', r));
    for (const s of sockets) s.resetAndDestroy();
    await closed;
    await new Promise((r) => setTimeout(r, 50));
    assert.deepEqual(uncaught.map((e) => e.message), []);
  } finally {
    process.off('uncaughtException', onUncaught);
    for (const s of sockets) s.destroy();
    server.close();
  }
});
