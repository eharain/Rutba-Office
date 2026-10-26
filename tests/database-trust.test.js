// A database from someone else reads its sources only when allowed here
// (A2.18): a .rbase sent by someone else, with tables linked to a file, a
// web address or an engine, reads nothing and connects to nothing, on open
// or on its schedule, until the person on this computer allows it; the
// question names each place. A database made here, or a link made here,
// asks nothing. What is allowed is this computer's (base-trust.js), never
// the file's. A password or token is kept with the host or address it was
// typed for, and a link whose host or address was changed in the file is
// sent none of it. The web addresses are stand-ins served here and the
// engine a stand-in that counts its connections.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import { Base, replaceRows } from '../packages/database/src/index.js';
import { createBaseService } from '../apps/desktop/main/base.js';
import { createTrust, memoryStore, addressOf, pathPlace, keptSecrets, destinationOf } from '../apps/desktop/main/base-trust.js';

const tmp = (name) => fs.mkdtempSync(path.join(os.tmpdir(), name));
const tableOf = (s, name) => s.tables.find((t) => t.name === name);
const secrets = (kept = new Map()) => ({ available: () => true, get: (k) => kept.get(k) ?? null, set: (k, v) => kept.set(k, v), delete: (k) => kept.delete(k) });

/** A stand-in web address answering JSON, counting what asks it and the keys it was sent. */
async function server(answer = () => [{ id: 1, name: 'Amina' }, { id: 2, name: 'Bilal' }]) {
  const asked = [];
  const keys = [];
  const srv = http.createServer((req, res) => {
    asked.push(req.url);
    keys.push(req.headers['x-api-key'] || null);
    res.writeHead(200, { 'content-type': 'application/json' });
    res.end(JSON.stringify(answer(req.url)));
  });
  await new Promise((r) => srv.listen(0, '127.0.0.1', r));
  return { url: `http://127.0.0.1:${srv.address().port}/people`, asked, keys, close: () => new Promise((r) => srv.close(r)) };
}

/** A stand-in engine: one table of two records, its connections counted with the host and password each was given. */
function engine() {
  const world = { connected: 0, given: [] };
  world.connect = async (kind, settings, password) => {
    world.connected++;
    world.given.push([settings.host, password]);
    return {
      tables: async () => [{ schema: 'public', name: 'stock', kind: 'table' }],
      describe: async () => ({ key: 'id', fields: [{ name: 'id', type: 'integer' }, { name: 'name', type: 'text' }] }),
      read: async () => [{ id: 1, name: 'Flour' }, { id: 2, name: 'Sugar' }],
      keys: async () => [1, 2],
      close: async () => {},
    };
  };
  return world;
}

/** A database as someone else made it: a CSV, a web address and an engine linked, the CSV's copy kept from before. */
function madeElsewhere(file, { csv, url }) {
  const b = Base.open(file);
  b.createTable({ name: 'Notes', fields: [{ name: 'Text', type: 'text' }] });
  const prices = b.saveLink({ name: 'Prices', engine: 'file', settings: { where: csv, header: 1, every: 15 } });
  b.linkTable({ table: 'Prices', link: prices, remote: { sheet: null }, key: null, fields: [{ name: 'Item', type: 'text' }, { name: 'Price', type: 'currency' }], mode: 'copy' });
  replaceRows(b, 'Prices', [{ Item: 'Kept flour', Price: 1 }], b.fields('Prices'));
  const people = b.saveLink({ name: 'People', engine: 'json', settings: { where: url, headers: [], records: '', paging: { kind: 'none' }, every: 15 } });
  b.linkTable({ table: 'People', link: people, remote: {}, key: null, fields: [{ name: 'id', type: 'integer' }, { name: 'name', type: 'text' }], mode: 'copy' });
  const shop = b.saveLink({ name: 'Head office', engine: 'postgres', settings: { host: 'db.example', port: 5432, database: 'shop', user: 'office' } });
  b.linkTable({ table: 'Stock', link: shop, remote: { schema: 'public', name: 'stock' }, key: 'id', fields: [{ name: 'id', type: 'integer' }, { name: 'name', type: 'text' }] });
  b.close();
}

test('a database from elsewhere reads nothing and connects to nothing, on open, on its schedule or on Refresh, until it is allowed; the question names each place', async () => {
  const dir = tmp('trust-');
  const file = path.join(dir, 'theirs.rbase');
  const csv = path.join(dir, 'prices.csv');
  fs.writeFileSync(csv, 'Item,Price\r\nFlour,1.80\r\nSugar,0.95\r\n');
  const web = await server();
  const world = engine();
  const trust = createTrust(memoryStore());
  const svc = createBaseService({ secrets: secrets(), connect: world.connect, trust });
  try {
    madeElsewhere(file, { csv, url: web.url });
    let s = await svc.open({ file });
    const id = s.id;
    assert.equal(web.asked.length, 0);
    assert.equal(world.connected, 0);
    // The copy kept from before, as it was: not read.
    assert.deepEqual(svc.rows({ id, table: 'Prices' }).rows.map((r) => r.Item), ['Kept flour']);
    assert.deepEqual(['Prices', 'People', 'Stock'].map((n) => tableOf(s, n).blocked), [true, true, true]);
    assert.equal(tableOf(s, 'Notes').blocked, undefined);
    // Each place named: the file by its path, the web address by its host and whole address, the engine by its host.
    const byKind = Object.fromEntries(s.ask.map((a) => [a.kind, a]));
    assert.equal(byKind.file.path, path.resolve(csv));
    assert.ok(['mine', 'outside'].includes(byKind.file.place));
    assert.deepEqual(byKind.file.tables, ['Prices']);
    assert.deepEqual([byKind.json.place, byKind.json.host, byKind.json.url], ['web', new URL(web.url).host, web.url]);
    assert.deepEqual([byKind.postgres.place, byKind.postgres.host, byKind.postgres.port, byKind.postgres.database, byKind.postgres.user, byKind.postgres.engine], ['engine', 'db.example', 5432, 'shop', 'office', 'PostgreSQL']);

    // Due on its schedule, and Refresh, and a change to the engine's table: still nothing.
    assert.equal(await svc.refreshDue({ id }), null);
    s = await svc.refreshLinked({ id, table: 'Prices', now: true });
    s = await svc.refreshLinked({ id, table: 'Stock', now: true });
    s = await svc.refreshLinked({ id });
    await assert.rejects(svc.insert({ id, table: 'Stock', values: { name: 'Salt' } }), /not read on this computer until you allow them/);
    await assert.rejects(svc.engineTables({ id, link: s.ask.find((a) => a.kind === 'postgres').id }), /until you allow them/);
    assert.equal(web.asked.length, 0);
    assert.equal(world.connected, 0);
    assert.deepEqual(svc.rows({ id, table: 'Prices' }).rows.map((r) => r.Item), ['Kept flour']);

    // Allowed, the file alone: it is read, the others still not.
    s = await svc.allowSources({ id, addresses: [byKind.file.address] });
    assert.deepEqual(svc.rows({ id, table: 'Prices' }).rows.map((r) => r.Item), ['Flour', 'Sugar']);
    assert.equal(tableOf(s, 'Prices').blocked, false);
    assert.deepEqual(s.ask.map((a) => a.kind).sort(), ['json', 'postgres']);
    assert.equal(web.asked.length, 0);
    assert.equal(world.connected, 0);
    // An address not shown is not allowed by being sent: only those the file links.
    s = await svc.allowSources({ id, addresses: ['json http://elsewhere.example/'] });
    assert.equal(s.ask.length, 2);
    // The rest allowed: read and connected to.
    s = await svc.allowSources({ id, addresses: s.ask.map((a) => a.address) });
    assert.equal(s.ask.length, 0);
    assert.ok(web.asked.length >= 1);
    assert.ok(world.connected >= 1);
    assert.equal(tableOf(s, 'People').count, 2);
    assert.equal(tableOf(s, 'Stock').count, 2);
    svc.close({ id });

    // Opened again on this computer: no question, read on open.
    fs.writeFileSync(csv, 'Item,Price\r\nFlour,1.80\r\nSugar,0.95\r\nSalt,0.40\r\n');
    s = await svc.open({ file });
    assert.equal(s.ask.length, 0);
    assert.equal(tableOf(s, 'Prices').count, 3);
    svc.close({ id: s.id });

    // On another computer (its own list): asked again, nothing read.
    const other = createBaseService({ secrets: secrets(), connect: world.connect });
    const was = { asked: web.asked.length, connected: world.connected };
    s = await other.open({ file });
    assert.equal(s.ask.length, 3);
    assert.deepEqual([web.asked.length, world.connected], [was.asked, was.connected]);
    other.closeAll();
  } finally {
    svc.closeAll();
    await web.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a database made here, and a source linked here, ask nothing; Save As keeps that, and a new database where one was allowed starts again', async () => {
  const dir = tmp('trust-mine-');
  const file = path.join(dir, 'mine.rbase');
  const csv = path.join(dir, 'prices.csv');
  fs.writeFileSync(csv, 'Item,Price\r\nFlour,1.80\r\n');
  const trust = createTrust(memoryStore());
  const svc = createBaseService({ secrets: secrets(), trust });
  try {
    let s = await svc.create({ file });
    assert.deepEqual(s.ask, []);
    s = (await svc.linkSource({ id: s.id, kind: 'file', name: 'Prices', settings: { where: csv } })).summary;
    assert.equal(tableOf(s, 'Prices').count, 1);
    assert.equal(tableOf(s, 'Prices').blocked, false);
    assert.deepEqual(s.ask, []);
    svc.close({ id: s.id });
    fs.writeFileSync(csv, 'Item,Price\r\nFlour,1.80\r\nSugar,0.95\r\n');
    s = await svc.open({ file });
    assert.deepEqual(s.ask, []);
    assert.equal(tableOf(s, 'Prices').count, 2);
    // Save As: the person's own copy reads as the database did.
    const copy = path.join(dir, 'copy.rbase');
    s = svc.saveAs({ id: s.id, file: copy });
    assert.equal(s.file, copy);
    assert.deepEqual(s.ask, []);
    svc.close({ id: s.id });
    s = await svc.open({ file: copy });
    assert.deepEqual(s.ask, []);
    svc.close({ id: s.id });
    // A template made here asks nothing either.
    s = svc.createFromTemplate({ file: path.join(dir, 'tpl.rbase'), template: 'inventory' });
    assert.deepEqual(s.ask, []);
    svc.close({ id: s.id });
    // A new database made where one was allowed: a file put there later with that source is asked about.
    s = await svc.create({ file });
    svc.close({ id: s.id });
    const b = Base.open(file);
    const l = b.saveLink({ name: 'Prices', engine: 'file', settings: { where: csv, every: 15 } });
    b.linkTable({ table: 'Prices', link: l, remote: { sheet: null }, key: null, fields: [{ name: 'Item', type: 'text' }], mode: 'copy' });
    b.close();
    s = await svc.open({ file });
    assert.equal(s.ask.length, 1);
    assert.equal(tableOf(s, 'Prices').count, 0);
  } finally {
    svc.closeAll();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('the allow list is not taken from the file: a link changed elsewhere to a share asks, nothing written in the file passes for allowed', async () => {
  const dir = tmp('trust-file-');
  const file = path.join(dir, 'shop.rbase');
  const csv = path.join(dir, 'prices.csv');
  fs.writeFileSync(csv, 'Item,Price\r\nFlour,1.80\r\n');
  const store = memoryStore();
  const svc = createBaseService({ secrets: secrets(), trust: createTrust(store) });
  try {
    let s = await svc.create({ file });
    s = (await svc.linkSource({ id: s.id, kind: 'file', name: 'Prices', settings: { where: csv } })).summary;
    svc.close({ id: s.id });
    // Someone else changes the same link (its id kept) to a share, and writes into the file that it is allowed.
    const share = '\\\\127.0.0.1\\no-such-share\\prices.csv';
    const b = Base.open(file);
    const link = b.links()[0];
    b.saveLink({ id: link.id, name: link.name, engine: 'file', settings: { ...link.settings, where: share } });
    b.setMeta('allowedSources', JSON.stringify([addressOf({ engine: 'file', settings: { where: share } })]));
    b.setMeta('trusted', '1');
    b.close();
    const before = JSON.stringify(store.get());
    s = await svc.open({ file });
    assert.equal(s.ask.length, 1);
    assert.deepEqual([s.ask[0].place, s.ask[0].server, s.ask[0].path], ['share', '127.0.0.1', share]);
    assert.equal(tableOf(s, 'Prices').blocked, true);
    // The copy as the person last read it, nothing tried at the share.
    assert.equal(tableOf(s, 'Prices').count, 1);
    assert.equal(tableOf(s, 'Prices').linkError, null);
    assert.equal(await svc.refreshDue({ id: s.id }), null);
    // Nothing of the file's went into this computer's list.
    assert.equal(JSON.stringify(store.get()), before);
    assert.ok(!before.includes('no-such-share'));
  } finally {
    svc.closeAll();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('an engine password is sent only to the host, port and database it was typed for: a link whose host was changed in the file asks for it again', async () => {
  const dir = tmp('trust-pw-');
  const file = path.join(dir, 'shop.rbase');
  const kept = new Map();
  const world = engine();
  const svc = createBaseService({ secrets: secrets(kept), connect: world.connect });
  try {
    let s = await svc.create({ file });
    const { link } = await svc.saveLink({ id: s.id, name: 'Head office', engine: 'postgres', settings: { host: 'db.example', port: 5432, database: 'shop', user: 'office' }, password: 'engine-pass' });
    s = (await svc.linkTables({ id: s.id, link, remotes: [{ schema: 'public', name: 'stock', as: 'Stock' }] })).summary;
    assert.deepEqual(world.given, [['db.example', 'engine-pass']]);
    svc.close({ id: s.id });
    // Someone else keeps the link's id and points it at their own host.
    const b = Base.open(file);
    const l = b.links()[0];
    b.saveLink({ id: l.id, name: l.name, engine: 'postgres', settings: { ...l.settings, host: 'evil.example' } });
    b.close();
    s = await svc.open({ file });
    assert.deepEqual(s.ask.map((a) => a.host), ['evil.example']);
    // Allowed (the person read the new host and allowed it): still, the password kept for db.example is not sent.
    s = await svc.allowSources({ id: s.id, addresses: s.ask.map((a) => a.address) });
    assert.ok(!world.given.some(([host]) => host === 'evil.example'));
    assert.match(tableOf(s, 'Stock').linkError, /saved for another engine\. Type it again for evil\.example/);
    assert.equal(s.links[0].password, false);
    // Typed again for the new host: kept for it, and sent there.
    s = (await svc.saveLink({ id: s.id, link: l.id, name: l.name, engine: 'postgres', settings: { ...l.settings, host: 'evil.example' }, password: 'typed-for-it' })).summary;
    assert.deepEqual(world.given.at(-1), ['evil.example', 'typed-for-it']);
    assert.ok(!world.given.some(([host, pw]) => host === 'evil.example' && pw === 'engine-pass'));
    assert.equal(s.links[0].password, true);
  } finally {
    svc.closeAll();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a key or token is sent only to the web address it was typed for: a source whose address was changed in the file reads nothing until it is typed again', async () => {
  const dir = tmp('trust-key-');
  const file = path.join(dir, 'shop.rbase');
  const theirs = await server();
  const other = await server();
  const svc = createBaseService({ secrets: secrets() });
  try {
    let s = await svc.create({ file });
    s = (await svc.linkSource({ id: s.id, kind: 'json', name: 'People', settings: { where: theirs.url, headers: [{ name: 'X-Api-Key', secret: true }] }, headers: { 'X-Api-Key': 'key-one' } })).summary;
    assert.ok(theirs.keys.includes('key-one'));
    svc.close({ id: s.id });
    const b = Base.open(file);
    const l = b.links()[0];
    b.saveLink({ id: l.id, name: l.name, engine: 'json', settings: { ...l.settings, where: other.url } });
    b.close();
    s = await svc.open({ file });
    s = await svc.allowSources({ id: s.id, addresses: s.ask.map((a) => a.address) });
    // Nothing asked of the new address, so no key went there.
    assert.equal(other.asked.length, 0);
    assert.equal(tableOf(s, 'People').token, true);
    assert.match(tableOf(s, 'People').linkError, /saved for another address\. Type it again for 127\.0\.0\.1:/);
    // Typed again for it: sent there, the old one never.
    s = await svc.sourceKeys({ id: s.id, table: 'People', headers: { 'X-Api-Key': 'key-two' } });
    assert.deepEqual(other.keys, ['key-two']);
    assert.equal(tableOf(s, 'People').token, false);
    assert.equal(tableOf(s, 'People').linkError, null);
  } finally {
    svc.closeAll();
    await theirs.close();
    await other.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a key is not sent to a next page on another address; a secret kept before destinations is bound to the first that reads it', async () => {
  const other = await server(() => ({ items: [{ id: 3, name: 'Cara' }] }));
  const first = await server(() => ({ items: [{ id: 1, name: 'Amina' }], next: other.url }));
  try {
    const { readJsonSource } = await import('../apps/desktop/main/base-sources.js');
    const got = await readJsonSource({ where: first.url, headers: [{ name: 'X-Api-Key', secret: true }], records: 'items', paging: { kind: 'next', next: 'next' } }, { secretOf: () => 'key-one' });
    assert.equal(got.rows.length, 2);
    assert.deepEqual(first.keys, ['key-one']);
    assert.deepEqual(other.keys, [null]);
  } finally {
    await first.close();
    await other.close();
  }
  // One kept as it was before 2.18: bound to the destination it is first read for, then not handed to another.
  const store = new Map([['base-link:L1', 'old-pass']]);
  const kept = keptSecrets(secrets(store));
  const link = { id: 'L1', engine: 'postgres', settings: { host: 'db.example', port: 5432, database: 'shop', user: 'office' } };
  assert.equal(kept.has('base-link:L1', link), true);
  assert.deepEqual(kept.read('base-link:L1', link), { value: 'old-pass', elsewhere: false });
  assert.notEqual(store.get('base-link:L1'), 'old-pass');
  assert.deepEqual(kept.read('base-link:L1', { ...link, settings: { ...link.settings, host: 'evil.example' } }), { value: null, elsewhere: true });
  assert.deepEqual(kept.read('base-link:L1', { ...link, settings: { ...link.settings, port: 6543 } }), { value: null, elsewhere: true });
  assert.deepEqual(kept.read('base-link:L1', link), { value: 'old-pass', elsewhere: false });
  // A web address by its origin: another path there is the same destination, another host or scheme is not.
  const at = (where) => destinationOf({ engine: 'json', settings: { where } });
  assert.equal(at('https://api.example/v1/people'), at('https://api.example/v2/orders?x=1'));
  assert.notEqual(at('https://api.example/people'), at('https://api.example.evil/people'));
  assert.notEqual(at('https://api.example/people'), at('http://api.example/people'));
});

test('where a path is, as the question says it: a network share, the person\'s own folders, or outside them', () => {
  const home = path.join(os.tmpdir(), 'home-of-someone');
  assert.deepEqual(pathPlace('\\\\server\\share\\a.csv', { home }), { place: 'share', server: 'server', path: '\\\\server\\share\\a.csv' });
  assert.equal(pathPlace('//server/share/a.csv', { home }).place, 'share');
  assert.deepEqual(pathPlace('\\\\?\\UNC\\files\\s\\a.csv', { home }).server, 'files');
  assert.equal(pathPlace(path.join(home, 'Documents', 'a.csv'), { home }).place, 'mine');
  assert.equal(pathPlace(path.join(os.tmpdir(), 'elsewhere', 'a.csv'), { home }).place, 'outside');
  assert.equal(pathPlace(home, { home }).place, 'outside');
  // One file, however it is spelt; on Windows in either case.
  const a = addressOf({ engine: 'file', settings: { where: path.join(home, 'x', '..', 'a.csv') } });
  assert.equal(a, addressOf({ engine: 'file', settings: { where: path.join(home, 'a.csv') } }));
  if (process.platform === 'win32') assert.equal(a, addressOf({ engine: 'file', settings: { where: path.join(home, 'A.CSV') } }));
  // An engine by host, port, database and account: another of any of them is another address.
  const pg = { engine: 'postgres', settings: { host: 'DB.example', port: 5432, database: 'shop', user: 'office' } };
  assert.equal(addressOf(pg), addressOf({ ...pg, settings: { ...pg.settings, host: 'db.example' } }));
  assert.notEqual(addressOf(pg), addressOf({ ...pg, settings: { ...pg.settings, host: 'evil.example' } }));
});
