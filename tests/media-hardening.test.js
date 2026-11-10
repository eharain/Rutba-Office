// Pictures, Video and Live Broadcast, given files that are odd or built to hurt.
import test from 'node:test';
import assert from 'node:assert/strict';
import zlib from 'node:zlib';
import { EventEmitter } from 'node:events';
import { planBatch, folderOf } from '@rutba/imaging/batch';
import { turnPng, decodePng } from '@rutba/imaging/png';
import { decodePicture } from '@rutba/imaging/decode';
import { readMp4 } from '@rutba/media/mp4-read';
import { readRvid, unzip, zipStored } from '@rutba/studio/rvid';
import { createChunkWriter } from '@rutba/studio/recorder';
import { RtmpPublisher, chunkMessage } from '@rutba/stream/rtmp';
import { createLiveService } from '../apps/desktop/main/live.js';

/* ---------- batch ---------- */

test('at a drive root a picture is written beside itself, and the plan says when that is itself', () => {
  assert.equal(folderOf('D:\\photo.jpg'), 'D:\\');
  assert.equal(folderOf('/photo.jpg'), '/');
  const plan = planBatch([{ path: 'D:\\photo.jpg', size: { width: 10, height: 10 } }], { format: 'same' });
  assert.equal(plan.items[0].target, 'D:\\photo.jpg');
  assert.equal(plan.items[0].original, true);
  assert.equal(plan.clashes.length, 1);
});

test('the same file written with the other slash is still that file', () => {
  const plan = planBatch([{ path: 'D:/photos\\a.jpg', size: { width: 10, height: 10 } }], { format: 'same', folder: 'd:\\photos' });
  assert.equal(plan.items[0].original, true);
});

test('a result never lands on another picture of the batch that is still to be read', () => {
  const size = { width: 10, height: 10 };
  const plan = planBatch([{ path: 'D:\\p\\a.jpg', size }, { path: 'D:\\p\\a.png', size }], { format: 'png' });
  assert.equal(plan.items[0].target, 'D:\\p\\a (2).png');
  assert.equal(plan.items[1].target, 'D:\\p\\a.png');
  assert.equal(plan.items[1].original, true);
});

/* ---------- png ---------- */

const crcTable = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
const crc = (b) => { let c = 0xffffffff; for (const v of b) c = crcTable[(c ^ v) & 255] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
const pngChunk = (type, body = Buffer.alloc(0)) => {
  const t = Buffer.concat([Buffer.from(type, 'latin1'), body]);
  const len = Buffer.alloc(4); len.writeUInt32BE(body.length);
  const sum = Buffer.alloc(4); sum.writeUInt32BE(crc(t));
  return Buffer.concat([len, t, sum]);
};
function makePng({ width, height, depth, colour, interlace = 0, raw, extra = [] }) {
  const head = Buffer.alloc(13);
  head.writeUInt32BE(width, 0); head.writeUInt32BE(height, 4); head[8] = depth; head[9] = colour; head[12] = interlace;
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 13, 10, 26, 10]), pngChunk('IHDR', head), ...extra, pngChunk('IDAT', zlib.deflateSync(raw)), pngChunk('IEND')]);
}
const chunksOf = (b) => {
  const out = [];
  for (let at = 8; at < b.length;) { const len = Buffer.from(b.buffer, b.byteOffset + at, 4).readUInt32BE(); out.push({ type: Buffer.from(b.subarray(at + 4, at + 8)).toString('latin1'), body: b.subarray(at + 8, at + 8 + len) }); at += 12 + len; }
  return out;
};
const rawOf = (png) => zlib.inflateSync(Buffer.from(chunksOf(png).filter((c) => c.type === 'IDAT').flatMap((c) => [...c.body])));

// 16-bit RGBA, 2 wide and 3 high, every sample its own value.
const deep = () => {
  const rows = [];
  let v = 1;
  for (let y = 0; y < 3; y++) { rows.push(0); for (let x = 0; x < 8; x++) { rows.push((v >> 8) & 255, v & 255); v += 257; } }
  return { width: 2, height: 3, depth: 16, colour: 6, raw: Buffer.from(rows) };
};

test('a 16-bit PNG turned stays 16-bit, its text and colour profile and density with it', () => {
  const pHYs = Buffer.alloc(9); pHYs.writeUInt32BE(100, 0); pHYs.writeUInt32BE(200, 4); pHYs[8] = 1;
  const before = makePng({ ...deep(), extra: [pngChunk('gAMA', Buffer.from([0, 0, 177, 143])), pngChunk('pHYs', pHYs), pngChunk('tEXt', Buffer.from('Title\0Kept', 'latin1'))] });
  const after = turnPng(before);
  const chunks = chunksOf(after);
  const head = Buffer.from(chunks[0].body);
  assert.deepEqual([head.readUInt32BE(0), head.readUInt32BE(4), head[8], head[9]], [3, 2, 16, 6]);
  assert.deepEqual(chunks.map((c) => c.type), ['IHDR', 'gAMA', 'pHYs', 'tEXt', 'IDAT', 'IEND']);
  const density = Buffer.from(chunks.find((c) => c.type === 'pHYs').body);
  assert.deepEqual([density.readUInt32BE(0), density.readUInt32BE(4)], [200, 100]);
});

test('four turns bring back every byte of the pixels', () => {
  const before = makePng(deep());
  let now = before;
  for (let i = 0; i < 4; i++) now = turnPng(now);
  const decoded = (b) => Buffer.from(decodePng(b).rgba).toString('hex');
  assert.equal(decoded(now), decoded(before));
});

test('a turn is clockwise', () => {
  // Red, green on top; blue, white below.
  const raw = Buffer.from([0, 255, 0, 0, 0, 255, 0, 0, 0, 0, 255, 255, 255, 255]);
  const { width, height, rgba } = decodePng(turnPng(makePng({ width: 2, height: 2, depth: 8, colour: 2, raw })));
  assert.deepEqual([width, height], [2, 2]);
  const at = (x, y) => [...rgba.subarray((y * 2 + x) * 4, (y * 2 + x) * 4 + 3)];
  assert.deepEqual(at(1, 0), [255, 0, 0], 'the top left goes to the top right');
  assert.deepEqual(at(0, 0), [0, 0, 255], 'the bottom left goes to the top left');
});

test('a one-bit palette PNG keeps its palette and its depth', () => {
  const raw = Buffer.from([0, 0b10000000, 0, 0b01000000]);
  const plte = pngChunk('PLTE', Buffer.from([1, 2, 3, 4, 5, 6]));
  const after = turnPng(makePng({ width: 2, height: 2, depth: 1, colour: 3, raw, extra: [plte] }));
  const chunks = chunksOf(after);
  assert.equal(chunks[0].body[8], 1);
  assert.deepEqual([...chunks.find((c) => c.type === 'PLTE').body], [1, 2, 3, 4, 5, 6]);
  assert.deepEqual([...rawOf(after)], [0, 0b01000000, 0, 0b10000000], 'the two set pixels, turned');
});

test('an interlaced PNG turns, and comes out plain', () => {
  // 1 by 1: only the first pass has a pixel.
  const after = turnPng(makePng({ width: 1, height: 1, depth: 8, colour: 2, interlace: 1, raw: Buffer.from([0, 9, 8, 7]) }));
  assert.equal(chunksOf(after)[0].body[12], 0);
  assert.deepEqual([...decodePng(after).rgba.subarray(0, 3)], [9, 8, 7]);
});

test('an animated PNG is not turned at all', () => {
  assert.equal(turnPng(makePng({ width: 1, height: 1, depth: 8, colour: 2, raw: Buffer.from([0, 1, 2, 3]), extra: [pngChunk('acTL', Buffer.alloc(8))] })), null);
});

/* ---------- limits on what a file claims ---------- */

const box = (type, ...body) => {
  const inner = Buffer.concat(body.map((b) => Buffer.from(b)));
  const head = Buffer.alloc(8); head.writeUInt32BE(8 + inner.length); head.write(type, 4, 'latin1');
  return Buffer.concat([head, inner]);
};
const u = (...ns) => { const b = Buffer.alloc(ns.length * 4); ns.forEach((n, i) => b.writeUInt32BE(n >>> 0, i * 4)); return b; };

function claimingMp4(stbl) {
  const hd = u(0, 0, 0, 1000, 1000);
  return Buffer.concat([box('ftyp', Buffer.from('isom'), u(0)), box('moov', box('mvhd', hd), box('trak', box('tkhd', u(0, 0, 0, 1)), box('mdia', box('mdhd', hd), box('hdlr', u(0, 0), Buffer.from('vide')), box('minf', box('stbl', ...stbl)))))]);
}

test('an MP4 that claims four billion samples in one kilobyte opens with what it holds', () => {
  const lots = 0xffffffff;
  const file = claimingMp4([box('stts', u(0, lots, 1, 1)), box('stsz', u(0, 0, lots)), box('stsc', u(0, lots, 1, 1, 1)), box('stco', u(0, lots, 40))]);
  assert.ok(file.length < 1024);
  const started = Date.now();
  const read = readMp4(file);
  assert.ok(Date.now() - started < 2000);
  assert.ok(read.tracks[0].samples.length <= 1);
});

test('fixed-size samples are bounded by the file too', () => {
  const read = readMp4(claimingMp4([box('stts', u(0, 1, 0xffffffff, 1)), box('stsz', u(0, 1, 0xffffffff)), box('stsc', u(0, 1, 1, 0xffffffff, 1)), box('stco', u(0, 1, 40))]));
  assert.ok(read.tracks[0].samples.length <= 2048);
});

test('a track with no header is passed over', () => {
  const file = Buffer.concat([box('ftyp', Buffer.from('isom'), u(0)), box('moov', box('mvhd', u(0, 0, 0, 1000, 1000)), box('trak', box('edts')))]);
  assert.deepEqual(readMp4(file).tracks, []);
});

const zipWith = (name, deflated, unpackedSize) => {
  const n = Buffer.from(name);
  const local = Buffer.alloc(30); local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(8, 8); local.writeUInt32LE(deflated.length, 18); local.writeUInt32LE(unpackedSize, 22); local.writeUInt16LE(n.length, 26);
  const central = Buffer.alloc(46); central.writeUInt32LE(0x02014b50, 0); central.writeUInt16LE(8, 10); central.writeUInt32LE(deflated.length, 20); central.writeUInt32LE(unpackedSize, 24); central.writeUInt16LE(n.length, 28);
  const at = local.length + n.length + deflated.length;
  const end = Buffer.alloc(22); end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(1, 8); end.writeUInt16LE(1, 10); end.writeUInt32LE(central.length + n.length, 12); end.writeUInt32LE(at, 16);
  return Buffer.concat([local, n, deflated, central, n, end]);
};

test('a project part that unpacks past what it declares is refused while it unpacks', async () => {
  const bomb = zlib.deflateRawSync(Buffer.alloc(64 * 1024 * 1024));
  await assert.rejects(unzip(zipWith('doc.json', bomb, 10)), /more than it says/);
});

test('a project part that declares too much is refused before it unpacks', async () => {
  await assert.rejects(unzip(zipWith('doc.json', zlib.deflateRawSync(Buffer.from('x')), 0xfffffff0)), /more than it says/);
});

test('an honest part still reads', async () => {
  const words = Buffer.from('a'.repeat(5000));
  assert.deepEqual(Buffer.from((await unzip(zipWith('doc.json', zlib.deflateRawSync(words), words.length)))['doc.json']), words);
});

const project = async (sources) => zipStored([
  ['mimetype', 'application/vnd.rutba.video-project+zip'],
  ['doc.json', JSON.stringify({ rvid: 1, patches: [], sources: {} })],
  ['sources.json', JSON.stringify({ sources })],
].map(([name, text]) => ({ name, bytes: new TextEncoder().encode(text) })));

test('a project never looks at a server it did not come from', async () => {
  const asked = [];
  const exists = async (p) => { asked.push(p); return false; };
  for (const url of ['\\\\attacker\\share\\a.mp4', '//attacker/share/a.mp4', 'file://attacker/share/a.mp4']) {
    asked.length = 0;
    let read;
    try {
      read = await readRvid(await project([{ url }]), { projectPath: 'C:\\work\\p.rvid', exists });
    } catch (err) {
      assert.fail(err.message);
    }
    assert.deepEqual(asked, []);
    assert.deepEqual(read.missing, [url]);
  }
});

test('a project on a share still finds its sources on that share, and local ones anywhere', async () => {
  const asked = [];
  const exists = async (p) => { asked.push(p); return true; };
  const read = await readRvid(await project([{ url: '\\\\nas\\media\\a.mp4' }, { url: 'C:\\clips\\b.mp4' }]), { projectPath: '\\\\NAS\\media\\p.rvid', exists });
  assert.deepEqual(asked, ['\\\\nas\\media\\a.mp4', 'C:\\clips\\b.mp4']);
  assert.deepEqual(read.missing, []);
});

test('a TIFF that claims sixteen floating-point samples a pixel is refused before anything is made', () => {
  const entry = (tag, type, count, value) => { const b = Buffer.alloc(12); b.writeUInt16LE(tag, 0); b.writeUInt16LE(type, 2); b.writeUInt32LE(count, 4); b.writeUInt32LE(value, 8); return b; };
  const entries = [entry(256, 4, 1, 10000), entry(257, 4, 1, 10000), entry(258, 3, 1, 32), entry(259, 3, 1, 1), entry(262, 3, 1, 1), entry(273, 4, 1, 8), entry(277, 3, 1, 16), entry(278, 4, 1, 10000), entry(279, 4, 1, 10), entry(339, 3, 1, 3)];
  const count = Buffer.alloc(2); count.writeUInt16LE(entries.length);
  const file = Buffer.concat([Buffer.from([0x49, 0x49, 42, 0]), u32le(8), count, ...entries, u32le(0)]);
  assert.throws(() => decodePicture(file, 'x.tif', {}), /MB of samples/);
});

test('a PSD whose channels would take gigabytes is refused before anything is unpacked', () => {
  const head = Buffer.alloc(26);
  head.write('8BPS', 0, 'latin1'); head.writeUInt16BE(1, 4); head.writeUInt16BE(56, 12); head.writeUInt32BE(10000, 14); head.writeUInt32BE(10000, 18); head.writeUInt16BE(16, 22); head.writeUInt16BE(3, 24);
  const file = Buffer.concat([head, Buffer.alloc(12), Buffer.from([0, 2]), Buffer.alloc(16)]);
  assert.throws(() => decodePicture(file, 'x.psd', {}), /MB of samples/);
});
function u32le(n) { const b = Buffer.alloc(4); b.writeUInt32LE(n); return b; }

/* ---------- recording and broadcast ---------- */

test('a write that fails is heard at once, and only once', async () => {
  const heard = [];
  const writer = createChunkWriter({ append: async () => { throw new Error('disk full'); }, onFail: (e) => heard.push(e.message) });
  const blob = new Blob([new Uint8Array(8)]);
  await writer.push(blob);
  await writer.push(blob);
  assert.deepEqual(heard, ['disk full']);
});

test('a reply too short to read fails the stream instead of throwing from the socket', () => {
  const p = new RtmpPublisher({ url: 'rtmp://example.invalid/live', key: 'k' });
  const failures = [];
  p.fail = (err) => failures.push(err);
  // WINDOW_ACK (5) says four bytes; a server sends two.
  p.state = 'connecting';
  assert.doesNotThrow(() => p.onData(Buffer.from(chunkMessage({ csid: 2, type: 5, body: Buffer.from([1, 2]) }))));
  assert.equal(failures.length, 1);
});

test('the broadcast ends with the window that began it', async () => {
  const closed = [];
  const makePublisher = () => Object.assign(new EventEmitter(), { start: async () => {}, connect: async () => {}, stop: () => closed.push('stopped'), close: () => closed.push('closed'), send: () => {} });
  const settings = new Map();
  const live = createLiveService({ stores: { settings: { get: (k, d) => settings.get(k) ?? d, set: (k, v) => settings.set(k, v) }, secrets: { get: () => null, set: () => {}, delete: () => {} } }, makePublisher, send: () => {} });
  const win = Object.assign(new EventEmitter(), { isDestroyed: () => false, webContents: { send: () => {} } });
  await live.start({ ids: [], video: { width: 1280, height: 720, fps: 30, kbps: 2000 }, audio: { kbps: 128, sampleRate: 48000, channels: 2 } }, win);
  assert.equal(live.status().live, true);
  win.emit('closed');
  assert.equal(live.status().live, false);
});
