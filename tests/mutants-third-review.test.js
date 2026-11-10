// Guards the third review found no test for: each was taken out of the code and every test still passed (10 Nov 2026).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { writeRvid, readRvid, unzip, zipStored, folderOf, resolvePath, relativePath } from '../packages/studio/src/rvid.js';
import { readMp4 } from '../packages/media/src/mp4-read.js';
import { readJsonSource } from '../apps/desktop/main/base-sources.js';
import { createRequire } from 'node:module';

const { lzw } = createRequire(import.meta.url)('../packages/pdf/src/reader.js');

const project = (url) => ({ version: 1, sources: { [url]: { url, kind: 'video', duration: 1 } }, patches: [{ url, start: 0, end: 1 }], tracks: [] });

test('a project on a network share keeps the share in its folder, and finds media moved with it there', async () => {
  assert.equal(folderOf('//home/share/film/cut.rvid'), '//home/share/film');
  assert.equal(resolvePath('//home/share/film', 'media/a.webm'), '//home/share/film/media/a.webm');
  assert.equal(relativePath('//home/share/film', '//home/share/film/media/a.webm'), 'media/a.webm');
  const bytes = writeRvid(project('//home/share/film/media/a.webm'), { projectPath: '//home/share/film/cut.rvid' });
  const there = '//home/share/copy/media/a.webm';
  const back = await readRvid(bytes, { projectPath: '//home/share/copy/cut.rvid', exists: async (p) => p === there });
  assert.deepEqual(back.moved.map((m) => m.to), [there]);
});

test('a relative path never climbs above the drive or the share it starts on', () => {
  assert.equal(resolvePath('D:/film', '../../../x/a.webm'), 'D:/x/a.webm');
  assert.equal(resolvePath('//home/share/film', '../../../../evil/share/a.webm'), '//home/share/evil/share/a.webm');
  assert.equal(relativePath('//home/share/film', '//other/share/a.webm'), null, 'no way from one share to another');
});

test('a project opened by a share\'s device name does not follow its relative paths to another server', async () => {
  const doc = project('D:/gone/a.webm');
  const bytes = writeRvid(doc, { projectPath: null });
  const files = await unzip(bytes);
  // What a project from elsewhere could say of where its media lies, beside it.
  const list = JSON.parse(new TextDecoder().decode(files['sources.json']));
  for (const s of Array.isArray(list) ? list : Object.values(list.sources ?? list)) s.relative = '../../../../../../evil/share/a.webm';
  const crafted = zipStored(Object.entries(files).map(([name, b]) => ({ name, bytes: name === 'sources.json' ? new TextEncoder().encode(JSON.stringify(list)) : b })));
  const asked = [];
  await readRvid(crafted, { projectPath: '//?/GLOBALROOT/Device/Mup/home/share/film/cut.rvid', exists: async (p) => { asked.push(p); return false; } });
  assert.deepEqual(asked.filter((p) => /evil/i.test(p)), []);
});

const box = (type, ...body) => {
  const inner = Buffer.concat(body.map((x) => Buffer.from(x)));
  const head = Buffer.alloc(8); head.writeUInt32BE(8 + inner.length); head.write(type, 4, 'latin1');
  return Buffer.concat([head, inner]);
};
const u = (...ns) => { const b = Buffer.alloc(ns.length * 4); ns.forEach((n, i) => b.writeUInt32BE(n >>> 0, i * 4)); return b; };

test('an edit list claiming more entries than its box holds reads only those it holds', () => {
  const hd = u(0, 0, 0, 1000, 1000);
  for (const [count, held] of [[0xffffffff, 0], [5, 1]]) {
    const elst = box('elst', u(0, count), ...Array.from({ length: held }, () => u(1000, 0, 0x10000)));
    const moov = box('moov', box('mvhd', hd), box('trak', box('tkhd', u(0, 0, 0, 1)), box('edts', elst),
      box('mdia', box('mdhd', hd), box('hdlr', u(0, 0), Buffer.from('vide')), box('minf', box('stbl')))));
    const file = Buffer.concat([box('ftyp', Buffer.from('isom'), u(0)), moov]);
    const started = Date.now();
    const read = readMp4(file);
    assert.equal(read.tracks[0].edits.length, held);
    assert.ok(Date.now() - started < 2000);
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

test('a move down to http on the same port does not take the key either', async () => {
  const { calls, fetchImpl } = redirecting({ 'https://example.com:8443/x': 'http://example.com:8443/x' });
  await readJsonSource({ where: 'https://example.com:8443/x', headers: [{ name: 'X-Key', secret: true }] }, { secretOf: () => 'hush', fetchImpl }).catch(() => {});
  assert.equal(calls.find((c) => c.url === 'http://example.com:8443/x')?.headers['X-Key'], undefined);
});

/** Codes as the reader takes them: widths grow as its table does. A run of one byte, each code one longer than the last. */
function lzwRun(codes) {
  let bits = 9;
  let size = 258;
  let acc = 0n;
  let count = 0;
  const out = [];
  codes.forEach((code, i) => {
    acc = (acc << BigInt(bits)) | BigInt(code);
    count += bits;
    while (count >= 8) { out.push(Number((acc >> BigInt(count - 8)) & 255n)); count -= 8; }
    if (i > 0) size += 1;
    if (size + 1 >= (1 << bits) && bits < 12) bits += 1;
  });
  if (count) out.push(Number((acc << BigInt(8 - count)) & 255n));
  return Buffer.from(out);
}

test('LZW that comes to more than the reader allows is refused, not held', () => {
  const codes = [97, ...Array.from({ length: 150 }, (_, i) => 258 + i)];
  const whole = lzw(lzwRun(codes));
  assert.equal(whole.length, (151 * 152) / 2, 'read whole when it fits');
  assert.ok(whole.every((b) => b === 97));
  assert.throws(() => lzw(lzwRun(codes), 1, 1000));
});
