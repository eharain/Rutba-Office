// An account on a plain IMAP port, against a server that does not offer
// STARTTLS — which is what a server looks like once somebody between it and
// the client has taken the offer out. The password must not be sent.
//
// The server here is a few lines on the loopback address that answers the
// commands a client opens with and writes down every line it hears.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';
import { createMailService } from '../apps/desktop/main/mail.js';

function fakeStores() {
  const settings = new Map();
  const secrets = new Map();
  return {
    settings: {
      get: (k, fallback) => (settings.has(k) ? settings.get(k) : fallback),
      set: (k, v) => settings.set(k, v),
      delete: (k) => settings.delete(k),
      all: () => Object.fromEntries(settings),
    },
    secrets: {
      available: () => true,
      get: (k) => secrets.get(k) ?? null,
      set: (k, v) => secrets.set(k, v),
      delete: (k) => secrets.delete(k),
      keys: () => [...secrets.keys()],
    },
  };
}

/** An IMAP server with no STARTTLS that accepts any login, and remembers what it was told. */
async function strippedServer() {
  const heard = [];
  const sockets = new Set();
  const server = net.createServer((socket) => {
    sockets.add(socket);
    socket.on('close', () => sockets.delete(socket));
    socket.on('error', () => {});
    socket.write('* OK [CAPABILITY IMAP4rev1 AUTH=PLAIN] ready\r\n');
    socket.on('data', (chunk) => {
      for (const line of String(chunk).split('\r\n').filter(Boolean)) {
        heard.push(line);
        const [tag, command = ''] = line.split(' ');
        if (/^capability$/i.test(command)) socket.write(`* CAPABILITY IMAP4rev1 AUTH=PLAIN\r\n${tag} OK done\r\n`);
        else if (/^logout$/i.test(command)) socket.end(`* BYE\r\n${tag} OK bye\r\n`);
        else socket.write(`${tag} OK done\r\n`);
      }
    });
  });
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  return {
    heard,
    port: server.address().port,
    open: () => sockets.size,
    close: () => new Promise((resolve) => {
      for (const s of sockets) s.destroy();
      server.close(resolve);
    }),
  };
}

test('a plain IMAP port whose server offers no STARTTLS is refused before the password is sent', async () => {
  const server = await strippedServer();
  const service = createMailService({
    stores: fakeStores(),
    holdBlob: () => ({ url: '' }),
    broadcast: () => {},
    userData: fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mail-imap-')),
  });
  let result;
  let left;
  try {
    result = await service.testAccount({
      account: {
        email: 'one@checks.example',
        imap: { host: '127.0.0.1', port: server.port, secure: false, starttls: true },
        smtp: { host: '127.0.0.1', port: 1, secure: false, starttls: true },
      },
      password: 'opal-7-lantern',
    });
    // The close travels as a packet; give it a moment to land.
    for (let i = 0; i < 20 && server.open() > 0; i++) await new Promise((resolve) => setTimeout(resolve, 25));
    left = server.open();
  } finally {
    await server.close();
  }
  assert.equal(result.imap.ok, false, 'the account does not pass its test');
  assert.match(result.imap.message, /STARTTLS/, 'and the reason names what was missing');
  assert.ok(!server.heard.some((line) => /opal-7-lantern|\bLOGIN\b|\bAUTHENTICATE\b/i.test(line)), `the server never heard a login: ${JSON.stringify(server.heard)}`);
  assert.equal(left, 0, 'and the failed test does not leave its connection open');
});
