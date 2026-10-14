// Adding mail accounts, against what a hostile domain, network or paste can
// do: each test proves one fix from the security review. A password goes
// only to a server under the address's own domain, one the provider table
// knows, or one the person agreed to by name (RFC 6186, section 6); never
// unencrypted; the web lookups never follow a redirect off https; nothing is
// looked up for text that is not an address, or before there is a password;
// and a line of a pasted list that is not an address is never shown back.

import test from 'node:test';
import assert from 'node:assert/strict';
import { parseAccountRows, checkRows, isAddress, isDomain } from '@rutba/mailbox/account-rows';
import { discoverMailServers, parseAutoconfig, isForeign, fetchTextHttps } from '../apps/desktop/main/mail-discover.js';
import { setupAccount } from '../apps/desktop/main/mail-setup.js';
import { createMailService } from '../apps/desktop/main/mail.js';
import fs from 'node:fs';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';

const nothing = { resolveMx: async () => [], resolveSrv: async () => { throw new Error('ENODATA'); }, resolveCname: async () => [] };
const noFetch = async (url) => { throw new Error(`404 from ${new URL(url).host}`); };
const answering = (...pairs) => async (host, port, kind) => (pairs.includes(`${host}:${port}`) ? { banner: kind === 'imap' ? '* OK ready' : '220 ready' } : null);

// L1: an address's domain goes into DNS names and web addresses, so it must be a host name and nothing else.
test('L1: an address is a host name after the @, nothing that could change a lookup\'s address', () => {
  for (const ok of ['a.b+c@sub.example.co.uk', 'ali@northwind.example', 'ali@münchen.de', 'ali@xn--mnchen-3ya.de', 'o\'brien@example.ie', 'x@a-b.example']) assert.ok(isAddress(ok), ok);
  for (const bad of ['a@b', 'a@b.c', 'a@evil.example/x', 'a@evil.example?x=1', 'a@evil.example#.northwind.example', 'a@exa mple.com', 'a@-x.example', 'a@x-.example', 'a@x..example', 'a@1.2.3.4', 'a@[127.0.0.1]', 'a@evil.example:8080', 'a@evil.example\\x', '@example.com', `${'a'.repeat(65)}@example.com`, 'ali northwind'])
    assert.ok(!isAddress(bad), bad);
  assert.ok(isDomain('mail.northwind.example') && !isDomain('northwind') && !isDomain(`${'a'.repeat(64)}.example`));
});

test('L1: nothing is looked up for text that is not an address', async () => {
  const asked = [];
  const record = (what) => async (x) => {
    asked.push(`${what} ${x}`);
    throw new Error('ENODATA');
  };
  const dns = { resolveMx: record('mx'), resolveSrv: record('srv'), resolveCname: record('cname') };
  for (const address of ['a@evil.example/autodiscover', 'a@evil.example?', 'a@b', 'a@[10.0.0.1]']) {
    const r = await discoverMailServers(address, { dns, fetchText: record('fetch'), probe: record('probe') });
    assert.equal(r.imap, null);
    assert.equal(r.steps[0].status, 'failed');
  }
  assert.deepEqual(asked, []);
});

test('L1: a pasted line that is not an address and a password is never shown back, only its number', () => {
  const { rows } = parseAccountRows('ali@northwind.example one\nCorrectHorseBatteryStaple\n"my secret phrase", x\nbo@northwind\n');
  assert.deepEqual(rows.map((r) => [r.line, r.email, r.problem || null]), [
    [1, 'ali@northwind.example', null],
    [2, '', 'unreadable'],
    [3, '', 'unreadable'],
    // Where an address goes, and like one: shown, to be put right.
    [4, 'bo@northwind', 'not-an-address'],
  ]);
  assert.ok(rows.every((r) => !/CorrectHorse|secret phrase/.test(JSON.stringify({ ...r, password: '' }))), 'the text of an unread line is kept nowhere');
  // A table whose address column holds something else (the columns out of step).
  const table = parseAccountRows('email,password\nhunter2,ali@northwind.example\nbo@northwind.example,pw\n').rows;
  assert.deepEqual(table.map((r) => [r.email, r.problem || null]), [['', 'unreadable'], ['bo@northwind.example', null]]);
  // Read again (the box read twice), an unreadable row stays unreadable.
  assert.equal(checkRows(rows)[1].problem, 'unreadable');
});

test('L1: setting up looks nothing up for a bad address, and nothing on the network before there is a password', async () => {
  const calls = [];
  const discover = async (email, options = {}) => {
    calls.push({ email, offline: Boolean(options.offline) });
    return { provider: { id: 'google', label: 'Google' }, oauth: 'google', imap: null, smtp: null, candidates: [] };
  };
  const never = async () => assert.fail('nothing is tried');
  assert.equal((await setupAccount({ email: 'a@evil.example/x', password: 'pw', discover, tryImap: never, trySmtp: never })).outcome, 'not-an-address');
  assert.deepEqual(calls, []);
  assert.equal((await setupAccount({ email: 'a@gmail.com', password: '', discover, tryImap: never, trySmtp: never })).outcome, 'sign-in');
  assert.deepEqual(calls, [{ email: 'a@gmail.com', offline: true }], 'the table only');
});

// M3: a domain's records can name anybody's server; the password goes there only once the person agrees.
test('M3: a server outside the address\'s domain and the table is marked foreign; one under the domain or in the table is not', async () => {
  const domain = 'northwind.example';
  assert.equal(isForeign({ host: 'imap.northwind.example', source: 'srv' }, domain), false);
  assert.equal(isForeign({ host: 'northwind.example', source: 'guess' }, domain), false);
  assert.equal(isForeign({ host: 'imap.gmail.com', source: 'mx' }, domain), false, 'a host the table knows');
  assert.equal(isForeign({ host: 'mail.collector.example', source: 'srv' }, domain), true);
  assert.equal(isForeign({ host: 'evilnorthwind.example', source: 'autoconfig' }, domain), true, 'a suffix that is not a subdomain');
  assert.equal(isForeign({ host: 'mail.other.example', source: 'import' }, domain), false, 'from the person\'s own mail program');

  const dns = {
    ...nothing,
    resolveSrv: async (name) => {
      if (name === '_imaps._tcp.northwind.example') return [{ name: 'mail.collector.example', port: 993, priority: 0, weight: 0 }];
      if (name === '_submissions._tcp.northwind.example') return [{ name: 'smtp.northwind.example', port: 465, priority: 0, weight: 0 }];
      throw new Error('ENODATA');
    },
  };
  const r = await discoverMailServers('ali@northwind.example', { dns, fetchText: noFetch, probe: answering('mail.collector.example:993', 'smtp.northwind.example:465') });
  assert.equal(r.imap.host, 'mail.collector.example');
  assert.equal(r.imap.foreign, true);
  assert.equal(r.smtp.foreign, false);
  assert.deepEqual(r.foreign, ['mail.collector.example']);
});

test('M3: setting up asks before a password goes to a foreign server, and tries it only once agreed to', async () => {
  const found = {
    provider: null,
    imap: { host: 'mail.collector.example', port: 993, secure: true, foreign: true },
    smtp: { host: 'smtp.northwind.example', port: 465, secure: true, foreign: false },
    foreign: ['mail.collector.example'],
    candidates: [
      { kind: 'imap', host: 'mail.collector.example', port: 993, secure: true, confidence: 4, foreign: true },
      { kind: 'imap', host: 'imap.northwind.example', port: 993, secure: true, confidence: 1, foreign: false },
      { kind: 'smtp', host: 'smtp.northwind.example', port: 465, secure: true, confidence: 3, foreign: false },
    ],
  };
  const tried = [];
  const tryAt = (kind) => async (server, user) => {
    tried.push(`${kind} ${server.host} ${user}`);
    return server.host === 'mail.collector.example' || server.host === 'smtp.northwind.example' ? { ok: true } : { ok: false, message: 'AUTHENTICATIONFAILED' };
  };
  const first = await setupAccount({ email: 'ali@northwind.example', password: 'pw', discover: async () => found, tryImap: tryAt('imap'), trySmtp: tryAt('smtp') });
  assert.equal(first.outcome, 'confirm');
  assert.deepEqual(first.foreign, [{ kind: 'imap', host: 'mail.collector.example', port: 993 }]);
  assert.deepEqual(tried, [], 'no password went anywhere');
  const agreed = await setupAccount({ email: 'ali@northwind.example', password: 'pw', confirmed: ['MAIL.collector.example'], discover: async () => found, tryImap: tryAt('imap'), trySmtp: tryAt('smtp') });
  assert.equal(agreed.outcome, 'ready');
  assert.equal(agreed.account.imap.host, 'mail.collector.example');
  // A foreign server that was not the one chosen, and not agreed to, is never tried, even when the chosen one fails.
  tried.length = 0;
  const quiet = {
    ...found,
    imap: { host: 'imap.northwind.example', port: 993, secure: true, foreign: false },
    foreign: [],
  };
  const r = await setupAccount({ email: 'ali@northwind.example', password: 'pw', discover: async () => quiet, tryImap: tryAt('imap'), trySmtp: tryAt('smtp') });
  assert.ok(tried.every((t) => !t.includes('collector')), tried.join(' | '));
  assert.equal(r.outcome, 'password');
});

// M3: never plaintext.
test('M3: a server offering no encryption is never a candidate, and a server typed in is never tried in the clear', async () => {
  const xml = `<clientConfig><emailProvider id="plain.example">
    <incomingServer type="imap"><hostname>imap.plain.example</hostname><port>143</port><socketType>plain</socketType></incomingServer>
    <incomingServer type="imap"><hostname>imaps.plain.example</hostname><port>993</port><socketType>SSL</socketType></incomingServer>
    <outgoingServer type="smtp"><hostname>smtp.plain.example</hostname><port>25</port><socketType>plain</socketType></outgoingServer>
  </emailProvider></clientConfig>`;
  assert.deepEqual(parseAutoconfig(xml, 'a@plain.example').map((s) => s.host), ['imaps.plain.example']);
  const seen = [];
  const tryAt = async (server) => {
    seen.push(server);
    return { ok: true };
  };
  await setupAccount({
    email: 'a@plain.example',
    password: 'pw',
    imap: { host: 'imap.plain.example', port: 143, secure: false, starttls: false },
    smtp: { host: 'smtp.plain.example', port: 25, secure: false, starttls: false },
    discover: async () => ({ provider: null, imap: null, smtp: null, candidates: [{ kind: 'imap', host: 'old.plain.example', port: 143, secure: false, starttls: false, confidence: 9 }] }),
    tryImap: tryAt,
    trySmtp: tryAt,
  });
  assert.ok(seen.length >= 2);
  assert.ok(seen.every((s) => s.secure || s.starttls === true), JSON.stringify(seen));
});

test('M3: the web lookups follow a redirect only to https, and only a few times', async () => {
  const pages = {
    'https://autoconfig.northwind.example/a': { status: 302, location: 'https://config.northwind.example/b' },
    'https://config.northwind.example/b': { status: 200, body: '<clientConfig/>' },
    'https://autoconfig.northwind.example/down': { status: 301, location: 'http://autoconfig.northwind.example/a' },
    'https://loop.northwind.example/': { status: 302, location: '/' },
  };
  const asked = [];
  const fetchImpl = async (url, opts) => {
    asked.push(url);
    assert.equal(opts.redirect, 'manual', 'redirects are not followed by fetch itself');
    const p = pages[url] || { status: 404 };
    return { status: p.status, ok: p.status >= 200 && p.status < 300, headers: { get: (h) => (h === 'location' ? p.location || null : null) }, text: async () => p.body || '' };
  };
  assert.equal(await fetchTextHttps('https://autoconfig.northwind.example/a', { fetchImpl }), '<clientConfig/>');
  await assert.rejects(fetchTextHttps('https://autoconfig.northwind.example/down', { fetchImpl }), /not https/);
  assert.ok(!asked.some((u) => u.startsWith('http:')), 'the plain address was never asked');
  await assert.rejects(fetchTextHttps('http://autoconfig.northwind.example/a', { fetchImpl }), /not https/);
  asked.length = 0;
  await assert.rejects(fetchTextHttps('https://loop.northwind.example/', { fetchImpl }), /302/);
  assert.equal(asked.length, 4, 'the first ask and three redirects');
});

// M3: an account kept with no encryption (starttls: false, as a file from another client may say) is never used in the clear.
function fakeStores() {
  const settings = new Map();
  const secrets = new Map();
  return {
    settings: { get: (k, fallback) => (settings.has(k) ? settings.get(k) : fallback), set: (k, v) => settings.set(k, v), delete: (k) => settings.delete(k), all: () => Object.fromEntries(settings) },
    secrets: { available: () => true, get: (k) => secrets.get(k) ?? null, set: (k, v) => secrets.set(k, v), delete: (k) => secrets.delete(k), keys: () => [...secrets.keys()] },
  };
}

/** A server on the loopback address that offers no STARTTLS, takes any login and writes down every line it hears. */
async function strippedServer(kind) {
  const heard = [];
  const sockets = new Set();
  const server = net.createServer((socket) => {
    sockets.add(socket);
    socket.on('close', () => sockets.delete(socket));
    socket.on('error', () => {});
    socket.write(kind === 'imap' ? '* OK [CAPABILITY IMAP4rev1 AUTH=PLAIN] ready\r\n' : '220 relay.checks.example ESMTP ready\r\n');
    socket.on('data', (chunk) => {
      for (const line of String(chunk).split('\r\n').filter(Boolean)) {
        heard.push(line);
        if (kind === 'imap') {
          const [tag, command = ''] = line.split(' ');
          if (/^capability$/i.test(command)) socket.write(`* CAPABILITY IMAP4rev1 AUTH=PLAIN\r\n${tag} OK done\r\n`);
          else if (/^logout$/i.test(command)) socket.end(`* BYE\r\n${tag} OK bye\r\n`);
          else socket.write(`${tag} OK done\r\n`);
        } else if (/^(EHLO|HELO)/i.test(line)) socket.write('250-relay.checks.example\r\n250 AUTH PLAIN LOGIN\r\n');
        else if (/^QUIT/i.test(line)) socket.end('221 bye\r\n');
        else socket.write('235 ok\r\n');
      }
    });
  });
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  return { heard, port: server.address().port, close: () => new Promise((resolve) => { for (const s of sockets) s.destroy(); server.close(resolve); }) };
}

test('M3: an account saved with no encryption is upgraded or refused, never signed in to in the clear', async () => {
  const imap = await strippedServer('imap');
  const smtp = await strippedServer('smtp');
  const service = createMailService({ stores: fakeStores(), holdBlob: () => ({ url: '' }), broadcast: () => {}, userData: fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-mail-tls-')) });
  let result;
  try {
    result = await service.testAccount({
      account: {
        email: 'one@checks.example',
        imap: { host: '127.0.0.1', port: imap.port, secure: false, starttls: false },
        smtp: { host: '127.0.0.1', port: smtp.port, secure: false, starttls: false },
      },
      password: 'opal-7-lantern',
    });
  } finally {
    await imap.close();
    await smtp.close();
  }
  assert.equal(result.imap.ok, false);
  assert.equal(result.smtp.ok, false);
  const leaked = (lines) => lines.some((line) => /opal-7-lantern|\bLOGIN\b|\bAUTHENTICATE\b|^AUTH\b|b3BhbC03LWxhbnRlcm4/i.test(line));
  assert.ok(!leaked(imap.heard), `the IMAP server never heard a login: ${JSON.stringify(imap.heard)}`);
  assert.ok(!leaked(smtp.heard), `the SMTP server never heard a login: ${JSON.stringify(smtp.heard)}`);
});
