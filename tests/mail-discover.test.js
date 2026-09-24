// Finding a mail server from an address, with every lookup injected.
//
// The search asks DNS, autoconfig, Microsoft autodiscover and the servers
// themselves. None of that happens here: each source is a fake, so the test
// says exactly what the search does with what it is told.

import test from 'node:test';
import assert from 'node:assert/strict';
import { discoverMailServers, parseAutoconfig, knownProvider, providerTiles, KNOWN } from '../apps/desktop/main/mail-discover.js';

const nothing = { resolveMx: async () => [], resolveSrv: async () => { throw new Error('ENODATA'); }, resolveCname: async () => [] };
const noFetch = async (url) => { throw new Error(`404 from ${new URL(url).host}`); };
const silence = async () => null;
/** A prober that answers on the given host:port pairs only. */
const answering = (...pairs) => async (host, port, kind) => (pairs.includes(`${host}:${port}`) ? { banner: kind === 'imap' ? '* OK ready' : '220 ready' } : null);

test('a known provider needs no network, and names its browser sign-in', async () => {
  const r = await discoverMailServers('someone@gmail.com', { offline: true });
  assert.equal(r.imap.host, 'imap.gmail.com');
  assert.equal(r.smtp.host, 'smtp.gmail.com');
  assert.equal(r.oauth, 'google');
  assert.equal(r.provider.label, 'Google');
  assert.equal(r.steps.find((s) => s.name === 'network').status, 'none');
  assert.equal(knownProvider('hotmail.co.uk').id, 'microsoft');
});

test('SRV records published by the domain win, and a knock confirms them', async () => {
  const dns = {
    ...nothing,
    resolveSrv: async (name) => {
      if (name === '_imaps._tcp.example.org') return [{ name: 'mailbox.example.org', port: 993, priority: 0, weight: 0 }];
      if (name === '_submission._tcp.example.org') return [{ name: 'relay.example.org', port: 587, priority: 0, weight: 0 }];
      throw new Error('ENODATA');
    },
  };
  const r = await discoverMailServers('a@example.org', { dns, fetchText: noFetch, probe: answering('mailbox.example.org:993', 'relay.example.org:587') });
  assert.equal(r.imap.host, 'mailbox.example.org');
  assert.equal(r.imap.port, 993);
  assert.equal(r.imap.secure, true);
  assert.equal(r.imap.verified, true);
  assert.equal(r.imap.source, 'srv');
  assert.equal(r.smtp.host, 'relay.example.org');
  assert.equal(r.smtp.starttls, true, '587 is an upgrade, not TLS from the first byte');
  assert.match(r.note, /SRV records/);
  assert.equal(r.steps.find((s) => s.name === 'SRV records').status, 'ok');
});

test('an MX at Google means Google, whatever the domain', async () => {
  const dns = { ...nothing, resolveMx: async () => [{ exchange: 'aspmx.l.google.com', priority: 1 }] };
  const r = await discoverMailServers('a@company.example', { dns, fetchText: noFetch, probe: silence });
  assert.equal(r.imap.host, 'imap.gmail.com');
  assert.equal(r.oauth, 'google');
  assert.match(r.steps.find((s) => s.name === 'MX records').detail, /hosted at Google/);
});

test('Microsoft autodiscover answers for a mailbox on Microsoft 365', async () => {
  const fetchText = async (url) => {
    if (/autodiscover\.json/.test(url)) return JSON.stringify({ Protocol: 'IMAP', Url: 'outlook.office365.com:993' });
    throw new Error('404');
  };
  const r = await discoverMailServers('a@firm.example', { dns: nothing, fetchText, probe: silence });
  assert.equal(r.imap.host, 'outlook.office365.com');
  assert.equal(r.smtp.host, 'smtp.office365.com');
  assert.equal(r.smtp.port, 587);
  assert.equal(r.oauth, 'microsoft');
});

test('autoconfig XML is read, with the address filled into the username', async () => {
  const xml = `<?xml version="1.0"?><clientConfig version="1.1"><emailProvider id="example.net">
    <incomingServer type="imap"><hostname>imap.example.net</hostname><port>143</port><socketType>STARTTLS</socketType><username>%EMAILADDRESS%</username><authentication>password-cleartext</authentication></incomingServer>
    <incomingServer type="pop3"><hostname>pop.example.net</hostname><port>995</port><socketType>SSL</socketType></incomingServer>
    <outgoingServer type="smtp"><hostname>smtp.example.net</hostname><port>465</port><socketType>SSL</socketType><username>%EMAILLOCALPART%</username></outgoingServer>
  </emailProvider></clientConfig>`;
  const servers = parseAutoconfig(xml, 'kim@example.net');
  assert.deepEqual(servers.map((s) => [s.kind, s.host, s.port, s.secure, s.starttls, s.user]), [
    ['imap', 'imap.example.net', 143, false, true, 'kim@example.net'],
    ['smtp', 'smtp.example.net', 465, true, false, 'kim'],
  ]);
  let asked = [];
  const fetchText = async (url) => {
    asked.push(url);
    if (/autoconfig\.thunderbird\.net/.test(url)) return xml;
    throw new Error('404');
  };
  const r = await discoverMailServers('kim@example.net', { dns: nothing, fetchText, probe: silence });
  assert.equal(r.imap.host, 'imap.example.net');
  assert.equal(r.imap.user, 'kim@example.net');
  assert.equal(r.imap.source, 'autoconfig');
  assert.ok(asked[0].startsWith('https://autoconfig.example.net/'), 'the domain\'s own autoconfig is asked first');
  assert.match(r.steps.find((s) => s.name === 'autoconfig').detail, /Mozilla/);
});

test('with nothing published, the conventional names are knocked on and the one that answers is taken', async () => {
  const r = await discoverMailServers('a@plain.example', { dns: nothing, fetchText: noFetch, probe: answering('mail.plain.example:993', 'mail.plain.example:587') });
  assert.equal(r.imap.host, 'mail.plain.example');
  assert.equal(r.imap.verified, true);
  assert.equal(r.smtp.host, 'mail.plain.example');
  assert.equal(r.smtp.port, 587);
  assert.match(r.steps.find((s) => s.name === 'servers answering').detail, /mail\.plain\.example:993/);
  // Nothing at all: the guess stands, unverified, and the note says so.
  const none = await discoverMailServers('a@dark.example', { dns: nothing, fetchText: noFetch, probe: silence });
  assert.equal(none.imap.host, 'imap.dark.example');
  assert.equal(none.imap.verified, false);
  assert.equal(none.steps.find((s) => s.name === 'servers answering').status, 'none');
});

test('settings imported from another client are kept and checked like the rest', async () => {
  const seed = { imap: { host: 'old.example.com', port: 993, secure: true }, smtp: { host: 'old.example.com', port: 465, secure: true } };
  const r = await discoverMailServers('a@example.com', { dns: nothing, fetchText: noFetch, probe: answering('old.example.com:993'), seed });
  assert.equal(r.imap.host, 'old.example.com');
  assert.equal(r.imap.source, 'import');
  assert.equal(r.imap.verified, true);
  assert.equal(r.smtp.verified, false, 'the outgoing server did not answer, and the result says so');
  assert.match(r.note, /outgoing server has not answered/);
});

test('a lookup that hangs does not hang the search', async () => {
  const dns = { ...nothing, resolveMx: () => new Promise(() => {}) };
  const t = Date.now();
  const r = await discoverMailServers('a@slow.example', { dns, fetchText: noFetch, probe: silence });
  assert.ok(Date.now() - t < 12000, 'bounded');
  assert.equal(r.steps.find((s) => s.name === 'MX records').status, 'failed');
  assert.ok(r.imap, 'a guess is still offered');
});

test('not an address: says so, asks nothing', async () => {
  const r = await discoverMailServers('nobody', { dns: nothing, fetchText: noFetch, probe: silence });
  assert.equal(r.imap, null);
  assert.equal(r.steps[0].status, 'failed');
});

/* ── the table itself, and the tiles it offers ───────────────────────────── */

const PORTS = new Set([143, 993, 25, 465, 587, 1025, 1143]);

test('every known provider has an id, a label, and servers on a real port', () => {
  for (const k of KNOWN) {
    assert.ok(k.id, `an id: ${JSON.stringify(k)}`);
    assert.ok(k.label, `a label: ${k.id}`);
    for (const [kind, server] of [['imap', k.imap], ['smtp', k.smtp]]) {
      assert.ok(server?.host, `${k.id} ${kind} host`);
      assert.ok(PORTS.has(server.port), `${k.id} ${kind} port ${server.port} is not one mail actually uses`);
      assert.equal(typeof server.secure, 'boolean', `${k.id} ${kind} secure is a boolean`);
    }
    if (k.appPassword) assert.match(k.appPassword.url, /^https:\/\//, `${k.id}'s app password page is https`);
  }
});

test('providerTiles() is plain data — no RegExp, JSON round-trips it unchanged — and there are enough of them to be worth a choice', () => {
  const tiles = providerTiles();
  assert.ok(tiles.length >= 8, `only ${tiles.length} tiles`);
  const roundTripped = JSON.parse(JSON.stringify(tiles));
  assert.deepEqual(roundTripped, tiles);
  const search = (v) => (v instanceof RegExp ? true : v && typeof v === 'object' ? Object.values(v).some(search) : false);
  assert.equal(search(tiles), false, 'a RegExp cannot cross the IPC bridge');
  assert.ok(tiles.some((t) => t.id === 'yahoo'));
  assert.ok(tiles.every((t) => t.label && t.domain && t.imap && t.smtp));
});

test('a Yahoo, iCloud, Hotmail or Google address is recognised, whichever domain it uses', () => {
  assert.equal(knownProvider('yahoo.co.uk').id, 'yahoo');
  assert.equal(knownProvider('ymail.com').id, 'yahoo');
  assert.equal(knownProvider('hotmail.com').id, 'microsoft');
  assert.equal(knownProvider('icloud.com').id, 'icloud');
});

test('a Yahoo address carries its app password page, offline', async () => {
  const r = await discoverMailServers('a@yahoo.com', { offline: true });
  assert.equal(r.appPassword?.url, 'https://login.yahoo.com/myaccount/security/app-password');
});
