// What the second review of Media and Data Files found.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Base } from '@rutba/database';
import { toStored } from '../packages/database/src/types.js';
import { readMp4 } from '../packages/media/src/mp4-read.js';
import { writeRvid, readRvid } from '../packages/studio/src/rvid.js';
import { readJsonSource } from '../apps/desktop/main/base-sources.js';
import { whileStopping } from '../apps/desktop/renderer/studio/stopping.js';

function roles() {
  const b = Base.open(':memory:');
  b.createTable({ name: 'Pay', fields: [{ name: 'Who', type: 'text' }, { name: 'Salary', type: 'number' }] });
  b.createTable({ name: 'Open', fields: [{ name: 'Note', type: 'text' }] });
  b.protect({ name: 'Boss', password: 'secret-for-the-test' });
  return b;
}
const asRole = (b, each) => { b.access = { locked: false, full: false, rights: { tables: { default: 'read', each }, hidden: {} } }; };
const design = (table, field) => ({ tables: [table], columns: [{ table, field }] });

test('a role cannot read the file\'s own lists of people and roles through a designed query', () => {
  const b = roles();
  asRole(b, {});
  for (const [table, field] of [['rutba_people', 'name'], ['rutba_roles', 'rights'], ['rutba_links', 'name']]) {
    assert.throws(() => b.runQuery({ design: design(table, field) }), /no table called/, table);
    assert.throws(() => b.queryRows({ design: design(table, field) }), /no table called/, table);
  }
  b.access = null;
  b.close();
});

test('a role cannot read a table hidden from it through a designed query', () => {
  const b = roles();
  asRole(b, { Pay: { access: 'hidden' } });
  assert.throws(() => b.runQuery({ design: design('Pay', 'Who') }), /no table called/);
  b.access = null;
  b.close();
});

test('a role still runs a query over what it may read, whatever the capitals', () => {
  const b = roles();
  b.insert('Open', { Note: 'hello' });
  asRole(b, { Pay: { access: 'hidden' } });
  assert.equal(b.runQuery({ design: design('Open', 'Note') }).rows[0].Note, 'hello');
  assert.doesNotThrow(() => b.runQuery({ design: { tables: ['open'], columns: [{ table: 'open', field: 'Note' }] } }));
  b.access = null;
  b.close();
});

test('the owner, with no role in force, is not held to any list', () => {
  const b = roles();
  assert.equal(b.runQuery({ design: design('Open', 'Note') }).rows.length, 0);
  b.close();
});

const f = { name: 'Key', type: 'integer' };

test('a whole number too long to keep exactly is refused, not rounded', () => {
  assert.throws(() => toStored(f, '9007199254740993'), /exactly/);
  assert.throws(() => toStored(f, 1e300), /exactly/);
  assert.equal(toStored(f, '9007199254740991'), 9007199254740991);
  assert.equal(toStored(f, '-12'), -12);
});

test('a read-only field given back as it stands is left alone, long or not', () => {
  const b = Base.open(':memory:');
  b.createTable({ name: 'T', fields: [{ name: 'Note', type: 'text' }, { name: 'Code', type: 'text' }] });
  const row = b.insert('T', { Note: 'a', Code: '9007199254740993' });
  const real = b.fields.bind(b);
  b.fields = (t) => real(t).map((x) => (x.name === 'Code' ? { ...x, readOnly: true, type: 'integer' } : x));
  b._known = b._known ? ((was) => (t) => was.call(b, t).map((x) => (x.name === 'Code' ? { ...x, readOnly: true, type: 'integer' } : x)))(b._known) : b._known;
  assert.doesNotThrow(() => b.update('T', row, { Note: 'b', Code: '9007199254740993' }));
  assert.throws(() => b.update('T', row, { Note: 'c', Code: '7' }), /read only/);
  b.close();
});

const box = (type, ...body) => {
  const inner = Buffer.concat(body.map((x) => Buffer.from(x)));
  const head = Buffer.alloc(8); head.writeUInt32BE(8 + inner.length); head.write(type, 4, 'latin1');
  return Buffer.concat([head, inner]);
};
const u = (...ns) => { const b = Buffer.alloc(ns.length * 4); ns.forEach((n, i) => b.writeUInt32BE(n >>> 0, i * 4)); return b; };

function fragments(runs, claimed, size = 1) {
  const hd = u(0, 0, 0, 1000, 1000);
  const moov = box('moov', box('mvhd', hd),
    box('trak', box('tkhd', u(0, 0, 0, 1)), box('mdia', box('mdhd', hd), box('hdlr', u(0, 0), Buffer.from('vide')), box('minf', box('stbl')))),
    box('mvex', box('trex', u(0, 1, 1, 33, size, 0))));
  const trun = box('trun', u(0, claimed));
  const moof = box('moof', box('traf', box('tfhd', u(0, 1)), ...Array(runs).fill(trun)));
  return Buffer.concat([box('ftyp', Buffer.from('isom'), u(0)), moov, moof]);
}

test('a small file of runs with nothing per sample cannot claim more samples than it has bytes', () => {
  const file = fragments(5000, 4000000000);
  const started = Date.now();
  const read = readMp4(file);
  assert.ok(read.tracks[0].samples.length <= file.length, read.tracks[0].samples.length + ' samples from ' + file.length + ' bytes');
  assert.ok(Date.now() - started < 3000);
});

test('an honest fragment with default sizes reads every sample', () => {
  const file = Buffer.concat([fragments(1, 10, 100), Buffer.alloc(2000)]);
  assert.equal(readMp4(file).tracks[0].samples.length, 10);
});

test('closing while a take is being finished waits for it', async () => {
  let phase = 'stopping';
  const seen = [];
  await whileStopping(() => phase, { wait: async () => { seen.push(1); if (seen.length === 3) phase = 'setup'; } });
  assert.equal(seen.length, 3);
});

test('closing when nothing is being finished does not wait, and a take that never ends does not hold the window for ever', async () => {
  let waits = 0;
  await whileStopping(() => 'setup', { wait: async () => { waits += 1; } });
  assert.equal(waits, 0);
  await whileStopping(() => 'stopping', { most: 500, every: 100, wait: async () => { waits += 1; } });
  assert.equal(waits, 5);
});

const project = (url) => ({ version: 1, sources: { [url]: { url, kind: 'video', duration: 1 } }, patches: [{ url, start: 0, end: 1 }], tracks: [] });

for (const [name, url] of [
  ['a plain share', '//evil/share/a.webm'],
  ['the long form of a share', '//?/UNC/evil/share/a.webm'],
  ['a share by its device name', '//?/GLOBALROOT/Device/Mup/evil/share/a.webm'],
  ['a share by its device name, backslashed', '\\\\.\\GLOBALROOT\\Device\\LanmanRedirector\\evil\\share\\a.webm'],
  ['the kernel form', '\\??\\UNC\\evil\\share\\a.webm'],
]) {
  test(`a project naming ${name} does not make this PC look at it`, async () => {
    const bytes = writeRvid(project(url), { projectPath: null });
    const asked = [];
    await readRvid(bytes, { projectPath: 'D:/film/cut.rvid', exists: async (p) => { asked.push(p); return false; } });
    assert.deepEqual(asked.filter((p) => /evil/i.test(p)), []);
  });
}

test('a project naming a file on this PC, in long form or not, is still looked for', async () => {
  for (const url of ['D:/film/a.webm', '//?/D:/film/a.webm']) {
    const asked = [];
    await readRvid(writeRvid(project(url), { projectPath: null }), { projectPath: 'E:/film/cut.rvid', exists: async (p) => { asked.push(p); return true; } });
    assert.ok(asked.includes(url), url);
  }
});

function redirecting(hops) {
  const calls = [];
  const fetchImpl = async (url, { headers }) => {
    calls.push({ url, headers });
    const to = hops[url];
    return to
      ? { status: 302, ok: false, headers: { get: () => to } }
      : { status: 200, ok: true, statusText: 'OK', headers: { get: () => 'application/json' }, text: async () => '[]', json: async () => [], arrayBuffer: async () => new TextEncoder().encode('[]').buffer };
  };
  return { calls, fetchImpl };
}
const headers = [{ name: 'X-Tenant', value: 'ours' }, { name: 'X-Key', secret: true }];
const fetched = (where, fetchImpl) => readJsonSource({ where, headers }, { secretOf: () => 'hush', fetchImpl }).catch(() => {});

test('a redirect to another server takes none of the headers a person set', async () => {
  const { calls, fetchImpl } = redirecting({ 'https://api.example.com/x': 'https://elsewhere.test/x' });
  await fetched('https://api.example.com/x', fetchImpl);
  const there = calls.find((c) => c.url.startsWith('https://elsewhere'));
  assert.ok(there, 'it was followed');
  assert.equal(there.headers['X-Key'], undefined);
  assert.equal(there.headers['X-Tenant'], undefined);
});

for (const [from, to] of [['http://example.com/x', 'https://example.com/x'], ['https://example.com/x', 'https://www.example.com/x'], ['https://www.example.com/x', 'https://example.com/x']]) {
  test(`a move from ${from} to ${to} keeps the key`, async () => {
    const { calls, fetchImpl } = redirecting({ [from]: to });
    await fetched(from, fetchImpl);
    assert.equal(calls.find((c) => c.url === to)?.headers['X-Key'], 'hush');
  });
}

test('a move down from https to http does not take the key', async () => {
  const { calls, fetchImpl } = redirecting({ 'https://example.com/x': 'http://example.com/x' });
  await fetched('https://example.com/x', fetchImpl);
  assert.equal(calls.find((c) => c.url === 'http://example.com/x')?.headers['X-Key'], undefined);
});
