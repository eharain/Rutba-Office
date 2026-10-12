// The Video app's project file (.rvid, docs/STUDIO.md D2 and V32): a zip of
// the project and its source references that comes back as it went, finds
// media moved along with it, says which is missing, refuses what is not one,
// and is registered as a kind the suite opens. And the container probe an
// export is read back with: a length, a picture and a sound in a WebM with
// no Duration (as a recorder writes it), and in an MP4 written in fragments.
import test from 'node:test';
import assert from 'node:assert/strict';
import { writeRvid, readRvid, unzip, zipStored, relativePath, resolvePath, folderOf, sourceList, RVID_MIME } from '../packages/studio/src/rvid.js';
import { emptyProject, addMedia, splitAt, setTransition, addTitle, addMarker, normaliseProject } from '../packages/studio/src/project.js';
import { probeMedia } from '../packages/media/src/probe.js';
import { sniff, kindFromExtension, KINDS } from '@rutba/office-formats/sniff';
import { APP_EXTENSIONS, WRITABLE, fileAssociations, saveFilters } from '@rutba/office-formats/registry';

const media = (url, extra = {}) => ({ url, name: url.split('/').pop(), kind: 'video', duration: 4, hasAudio: true, size: 1000, ...extra });

function project(dir = 'D:/work/film') {
  let doc = emptyProject();
  doc = addMedia(doc, media(`${dir}/media/a.webm`)).doc;
  doc = addMedia(doc, media(`${dir}/media/b.webm`)).doc;
  doc = addMedia(doc, { url: `${dir}/still.png`, name: 'still.png', kind: 'image' }).doc;
  const cut = splitAt(doc, 2);
  doc = setTransition(cut.doc, doc.patches.find((p) => p.url.endsWith('b.webm')).id, { kind: 'fade', seconds: 0.5 });
  doc = addTitle(doc, 'lowerThird', { at: 1, text: 'Ejaz Arain' }).doc;
  doc = addMarker(doc, 3, 'Chapter one');
  return doc;
}

test('V32: a project goes out as a zip whose first entry says what it is, and comes back exactly', async () => {
  const doc = project();
  const bytes = writeRvid(doc, { projectPath: 'D:/work/film/cut.rvid' });
  assert.equal(String.fromCharCode(...bytes.subarray(0, 4)), 'PK\u0003\u0004');
  assert.equal(new TextDecoder().decode(bytes.subarray(38, 38 + RVID_MIME.length)), RVID_MIME, 'mimetype first, stored, readable in the head');
  const parts = await unzip(bytes);
  assert.deepEqual(Object.keys(parts), ['mimetype', 'doc.json', 'sources.json']);
  const back = await readRvid(bytes);
  assert.deepEqual(normaliseProject(back.doc), normaliseProject(doc), 'the project round-trips');
  assert.equal(back.doc.patches.length, doc.patches.length);
  const sources = back.sources;
  assert.equal(sources.length, 3);
  const a = sources.find((s) => s.url.endsWith('a.webm'));
  assert.equal(a.relative, 'media/a.webm', 'relative to the project folder');
  assert.equal(a.duration, 4);
  assert.equal(a.kind, 'video');
});

test('V32: media moved with the project is found again by its relative path; what is gone is named', async () => {
  const doc = project('D:/work/film');
  const bytes = writeRvid(doc, { projectPath: 'D:/work/film/cut.rvid' });
  const there = new Set(['E:/copy/film/media/a.webm', 'E:/copy/film/still.png']);
  const back = await readRvid(bytes, { projectPath: 'E:/copy/film/cut.rvid', exists: async (p) => there.has(p) });
  assert.deepEqual(back.moved.map((m) => m.to).sort(), ['E:/copy/film/media/a.webm', 'E:/copy/film/still.png']);
  assert.deepEqual(back.missing, ['D:/work/film/media/b.webm']);
  assert.ok(back.doc.patches.filter((p) => p.url?.endsWith('a.webm')).every((p) => p.url === 'E:/copy/film/media/a.webm'), 'every clip of a moved source points at it');
  assert.ok(back.doc.sources['E:/copy/film/media/a.webm'], 'and so does the source list');
  assert.ok(!back.doc.sources['D:/work/film/media/a.webm']);
  const still = await readRvid(bytes, { projectPath: 'D:/work/film/cut.rvid', exists: () => true });
  assert.deepEqual([still.moved, still.missing], [[], []], 'nothing moved, nothing to do');
});

test('V32: what is not a project is refused in a sentence', async () => {
  await assert.rejects(readRvid(new Uint8Array(40)), /not a zip/);
  await assert.rejects(readRvid(zipStored([{ name: 'mimetype', bytes: new TextEncoder().encode('application/zip') }])), /not a Rutba Video project/);
  await assert.rejects(readRvid(zipStored([{ name: 'doc.json', bytes: new TextEncoder().encode('{') }])), /cannot be read/);
  await assert.rejects(readRvid(zipStored([{ name: 'doc.json', bytes: new TextEncoder().encode(JSON.stringify({ rvid: 99, patches: [] })) }])), /newer Rutba Office/);
});

test('paths: relative between folders, across drives none, resolved back, Windows and forward slashes alike', () => {
  assert.equal(relativePath('D:/work/film', 'D:/work/film/media/a.webm'), 'media/a.webm');
  assert.equal(relativePath('D:\\work\\film', 'D:\\work\\clips\\a.webm'), '../clips/a.webm');
  assert.equal(relativePath('D:/work', 'E:/a.webm'), null, 'another drive');
  assert.equal(relativePath('d:/Work', 'D:/work/a.webm'), 'a.webm', 'a Windows path ignores case');
  assert.equal(resolvePath('E:\\copy\\film', '../clips/a.webm'), 'E:\\copy\\clips\\a.webm');
  assert.equal(folderOf('/home/me/film/cut.rvid'), '/home/me/film');
  assert.equal(sourceList({ patches: [{ url: 'x/y.webm' }, { url: 'x/y.webm' }] })[0].name, 'y.webm');
});

test('V32: .rvid is a kind the suite opens in Video, is offered by the Open and Save As dialogs, and is associated with the app', () => {
  assert.equal(kindFromExtension('cut.rvid'), 'rvid');
  assert.equal(KINDS.rvid.app, 'video');
  assert.equal(APP_EXTENSIONS.video[0], 'rvid');
  assert.equal(WRITABLE.video[0].kind, 'rvid');
  assert.deepEqual(saveFilters('video')[0].extensions, ['rvid']);
  const assoc = fileAssociations().find((a) => a.ext === 'rvid');
  assert.equal(assoc.app, 'video');
  assert.equal(assoc.name, 'Rutba Video Project');
  const found = sniff(writeRvid(emptyProject()), 'cut.rvid');
  assert.equal(found.kind, 'rvid', 'a zip named .rvid is a project');
  assert.equal(sniff(writeRvid(emptyProject()), 'renamed.zip').kind, 'rvid', 'and one renamed is still known by its first entry');
});

/* ── the probe ───────────────────────────────────────────────────────── */

const vint = (n) => { const b = Buffer.alloc(8); b.writeBigUInt64BE(BigInt(n)); b[0] = 0x01; return b; };
const el = (id, ...body) => { const c = Buffer.concat(body); return Buffer.concat([Buffer.from(id), vint(c.length), c]); };
const uint = (id, n) => { const b = Buffer.alloc(8); b.writeBigUInt64BE(BigInt(n)); let i = 0; while (i < 7 && b[i] === 0) i++; return el(id, b.subarray(i)); };
const str = (id, s) => el(id, Buffer.from(s));
const UNKNOWN = Buffer.from([0x01, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff]);

/** A WebM as a recorder writes one: no Duration, Segment and Cluster of unknown size, a picture track and a sound track. */
function recorderWebm({ seconds = 3, fps = 25 } = {}) {
  const header = el([0x1a, 0x45, 0xdf, 0xa3], str([0x42, 0x82], 'webm'));
  const info = el([0x15, 0x49, 0xa9, 0x66], uint([0x2a, 0xd7, 0xb1], 1000000), str([0x4d, 0x80], 'test'));
  const tracks = el([0x16, 0x54, 0xae, 0x6b],
    el([0xae], uint([0xd7], 1), uint([0x83], 1), str([0x86], 'V_VP8'), el([0xe0], uint([0xb0], 320), uint([0xba], 240))),
    el([0xae], uint([0xd7], 2), uint([0x83], 2), str([0x86], 'A_OPUS'), el([0xe1], el([0xb5], (() => { const b = Buffer.alloc(8); b.writeDoubleBE(48000); return b; })()))));
  const clusters = [];
  for (let c = 0; c < seconds; c++) {
    const blocks = [];
    for (let f = 0; f < fps; f++) {
      const rel = Math.round((f * 1000) / fps);
      blocks.push(el([0xa3], Buffer.from([0x81, rel >> 8, rel & 0xff, 0x80, 1, 2, 3])));
      if (f % 5 === 0) blocks.push(el([0xa3], Buffer.from([0x82, rel >> 8, rel & 0xff, 0x80, 9])));
    }
    clusters.push(Buffer.concat([Buffer.from([0x1f, 0x43, 0xb6, 0x75]), UNKNOWN, uint([0xe7], c * 1000), ...blocks]));
  }
  return Buffer.concat([header, Buffer.from([0x18, 0x53, 0x80, 0x67]), UNKNOWN, info, tracks, ...clusters]);
}

const box = (type, ...body) => { const c = Buffer.concat(body); const h = Buffer.alloc(8); h.writeUInt32BE(8 + c.length); h.write(type, 4, 'latin1'); return Buffer.concat([h, c]); };
const full = (type, version, flags, ...body) => box(type, Buffer.from([version, flags >> 16, (flags >> 8) & 0xff, flags & 0xff]), ...body);
const u32 = (...ns) => { const b = Buffer.alloc(4 * ns.length); ns.forEach((n, i) => b.writeUInt32BE(n, i * 4)); return b; };

/** An MP4 as a recorder writes one: an empty movie header, a picture and a sound track, then fragments. */
function fragmentedMp4({ seconds = 2 } = {}) {
  const ftyp = box('ftyp', Buffer.from('isom'), u32(0x200), Buffer.from('isomiso6mp41'));
  const trak = (id, handler, timescale, codec, w, h) => box('trak',
    full('tkhd', 0, 3, u32(0, 0, id, 0, 0), Buffer.alloc(52), u32(w * 65536, h * 65536)),
    box('mdia', full('mdhd', 0, 0, u32(0, 0, timescale, 0), Buffer.alloc(4)), full('hdlr', 0, 0, u32(0), Buffer.from(handler), Buffer.alloc(13)),
      box('minf', box('stbl', full('stsd', 0, 0, u32(1), box(codec, Buffer.alloc(8)))))));
  const moov = box('moov', full('mvhd', 0, 0, u32(0, 0, 1000, 0), Buffer.alloc(80)), trak(1, 'vide', 90000, 'avc1', 640, 360), trak(2, 'soun', 48000, 'mp4a', 0, 0), box('mvex'));
  const frags = [];
  for (let s = 0; s < seconds; s++) {
    // One second of picture: 30 samples of 3000 ticks; of sound: 47 samples of 1024 at 48 kHz (about a second).
    const vtraf = box('traf', full('tfhd', 0, 0x08, u32(1, 3000)), full('tfdt', 0, 0, u32(s * 90000)), full('trun', 0, 0x001, u32(30, 0)));
    const atraf = box('traf', full('tfhd', 0, 0x00, u32(2)), full('tfdt', 0, 0, u32(s * 48128)), full('trun', 0, 0x101, u32(47, 0), u32(...Array(47).fill(1024))));
    frags.push(box('moof', vtraf, atraf), box('mdat', Buffer.alloc(16)));
  }
  return Buffer.concat([ftyp, moov, ...frags]);
}

test('probe: a recorder WebM has its length added up from its blocks, its picture and its sound found', () => {
  const p = probeMedia(recorderWebm({ seconds: 3 }), { complete: true });
  assert.equal(p.format, 'webm');
  assert.ok(Math.abs(p.duration - 3) < 0.05, `length ${p.duration}`);
  assert.equal(p.width, 320);
  assert.equal(p.hasVideo, true);
  assert.equal(p.hasAudio, true);
  assert.equal(p.sampleRate, 48000);
  assert.deepEqual(p.tracks.map((t) => t.codec), ['V_VP8', 'A_OPUS']);
  const head = probeMedia(recorderWebm({ seconds: 3 }).subarray(0, 600));
  assert.equal(head.duration, null, 'from a head alone, no length is guessed');
  assert.equal(head.hasAudio, true, 'but the tracks are known');
  const declared = Buffer.concat([el([0x1a, 0x45, 0xdf, 0xa3], str([0x42, 0x82], 'matroska')), el([0x18, 0x53, 0x80, 0x67], el([0x15, 0x49, 0xa9, 0x66], uint([0x2a, 0xd7, 0xb1], 1000000), el([0x44, 0x89], (() => { const b = Buffer.alloc(8); b.writeDoubleBE(4500); return b; })())))]);
  const d = probeMedia(declared);
  assert.equal(d.format, 'mkv');
  assert.equal(d.duration, 4.5, 'a stated Duration is read as it is');
});

test('probe: an MP4 in fragments has its length added up, per track timescale; its tracks and picture size read', () => {
  const p = probeMedia(fragmentedMp4({ seconds: 2 }), { complete: true });
  assert.equal(p.format, 'mp4');
  assert.ok(Math.abs(p.duration - 2.0) < 0.02, `length ${p.duration}`);
  assert.equal(p.hasVideo, true);
  assert.equal(p.hasAudio, true);
  assert.deepEqual(p.tracks.map((t) => t.codec), ['avc1', 'mp4a']);
  assert.equal(p.width, 640);
  assert.equal(p.height, 360);
  assert.equal(probeMedia(fragmentedMp4({ seconds: 2 })).duration, null, 'not from part of a file');
});
