// The viewer's arithmetic: what a folder holds, in what order, and which of
// it is drawn — the part that makes a folder of thousands cost the same to
// draw as a folder of fifty.
import test from 'node:test';
import assert from 'node:assert/strict';
import { arrange, countByFamily, familyOf, kindOf, gridWindow, stripWindow, stripCentre, neighbours, thumbUrl, fileUrl, fitColumns, foldersOf } from '../apps/desktop/renderer/apps/pictures/library.js';
import { cacheKey, createQueue, createThumbnailer } from '../packages/office-shell/src/electron/thumbs.js';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const entries = [
  { name: 'b.mp4', path: 'C:/f/b.mp4', size: 900, mtime: 30 },
  { name: 'a10.jpg', path: 'C:/f/a10.jpg', size: 100, mtime: 10 },
  { name: 'a2.jpg', path: 'C:/f/a2.jpg', size: 300, mtime: 20 },
  { name: 'notes.pdf', path: 'C:/f/notes.pdf', size: 50, mtime: 5 },
  { name: 'tone.wav', path: 'C:/f/tone.wav', size: 400, mtime: 40 },
  { name: 'sub', path: 'C:/f/sub', dir: true, size: 0, mtime: 0 },
];

test('a folder is arranged by family, name and order, directories left out', () => {
  assert.deepEqual(arrange(entries).map((e) => e.name), ['a2.jpg', 'a10.jpg', 'b.mp4', 'notes.pdf', 'tone.wav'], 'names sort numerically');
  assert.deepEqual(arrange(entries, { sort: 'newest' }).map((e) => e.name), ['tone.wav', 'b.mp4', 'a2.jpg', 'a10.jpg', 'notes.pdf']);
  assert.deepEqual(arrange(entries, { sort: 'oldest' }).map((e) => e.name), ['notes.pdf', 'a10.jpg', 'a2.jpg', 'b.mp4', 'tone.wav']);
  assert.deepEqual(arrange(entries, { sort: 'largest' }).map((e) => e.name), ['b.mp4', 'tone.wav', 'a2.jpg', 'a10.jpg', 'notes.pdf']);
  assert.deepEqual(arrange(entries, { sort: 'kind' }).map((e) => e.name), ['a2.jpg', 'a10.jpg', 'b.mp4', 'tone.wav', 'notes.pdf']);
  assert.deepEqual(arrange(entries, { family: 'video' }).map((e) => e.name), ['b.mp4']);
  assert.deepEqual(arrange(entries, { family: 'pictures', query: '10' }).map((e) => e.name), ['a10.jpg']);
  assert.deepEqual(countByFamily(entries), { all: 5, pictures: 2, video: 1, audio: 1, pdf: 1 });
  assert.equal(familyOf('x.webp'), 'pictures');
  assert.equal(kindOf('x.GIF'), 'maybe-animated');
  assert.equal(kindOf('C:\\a\\b.MOV'), 'video');
});

test('the grid draws the rows in view and holds the room for the rest', () => {
  // 1000 tiles, 4 columns of 116 px in a 500 px wide scroller 400 px tall.
  const w = gridWindow({ count: 1000, width: 500, tile: 116, gap: 6, scrollTop: 0, height: 400, overscan: 1, padding: 6 });
  assert.equal(w.columns, 4);
  assert.equal(w.rows, 250);
  assert.equal(w.first, 0);
  assert.equal(w.last, 5 * 4, 'four rows in view and one of overscan');
  assert.equal(w.top, 0);
  assert.equal(w.top + (w.last - w.first) / 4 * w.rowHeight + w.bottom, 250 * w.rowHeight, 'the spacers and the drawn rows add up to the whole folder');

  const deep = gridWindow({ count: 1000, width: 500, tile: 116, gap: 6, scrollTop: 12200, height: 400, overscan: 1, padding: 6 });
  assert.equal(deep.first, (100 - 1) * 4, 'row 100 is at the top, one row of overscan above');
  assert.ok(deep.last - deep.first <= 7 * 4, `${deep.last - deep.first} tiles drawn`);
  assert.equal(deep.top, 99 * 122);

  const end = gridWindow({ count: 1000, width: 500, tile: 116, gap: 6, scrollTop: 999999, height: 400, overscan: 1, padding: 6 });
  assert.equal(end.last, 1000);
  assert.equal(end.bottom, 0);
  assert.deepEqual(gridWindow({ count: 0, width: 500, tile: 116 }).last, 0);
  assert.equal(gridWindow({ count: 3, width: 50, tile: 116 }).columns, 1, 'never fewer than one column');
});

test('the strip draws the items in view, and knows where to scroll to centre one', () => {
  const s = stripWindow({ count: 500, itemWidth: 74, gap: 5, scrollLeft: 790, width: 400, overscan: 2 });
  assert.equal(s.step, 79);
  assert.equal(s.first, 8);
  assert.equal(s.last, 18);
  assert.equal(s.left, 8 * 79);
  assert.equal(s.right, (500 - 18) * 79);
  assert.equal(stripCentre({ index: 10, itemWidth: 74, gap: 5, width: 400 }), 10 * 79 + 37 - 200);
  assert.equal(stripCentre({ index: 0, itemWidth: 74, gap: 5, width: 400 }), 0);
});

test('the neighbours worth fetching are the next ones first, wrapping only when the folder loops', () => {
  assert.deepEqual(neighbours(5, 10), [6, 4, 7, 3]);
  assert.deepEqual(neighbours(0, 10), [1, 9, 2, 8]);
  assert.deepEqual(neighbours(0, 10, { loop: false }), [1, 2]);
  assert.deepEqual(neighbours(9, 10, { loop: false, radius: 3 }), [8, 7, 6]);
  assert.deepEqual(neighbours(0, 1), []);
  assert.deepEqual(neighbours(0, 2), [1]);
});

test('the URLs a tile uses are the platform\'s hosts with the path encoded once', () => {
  assert.equal(fileUrl('C:\\Users\\x\\a b.jpg'), 'rutba://file/QzpcVXNlcnNceFxhIGIuanBn');
  assert.equal(thumbUrl('C:\\Users\\x\\a b.jpg'), 'rutba://thumb/QzpcVXNlcnNceFxhIGIuanBn?s=256');
  assert.equal(thumbUrl('/home/x/ü.png', 128), `rutba://thumb/${Buffer.from('/home/x/ü.png').toString('base64url')}?s=128`);
});

test('a thumbnail is keyed by the file, its size and its time; the queue runs the newest first', async () => {
  assert.equal(cacheKey('C:/a.jpg', { size: 256, mtimeMs: 1000.4, fileSize: 5 }), cacheKey('C:/a.jpg', { size: 256, mtimeMs: 1000, fileSize: 5 }));
  assert.notEqual(cacheKey('C:/a.jpg', { size: 256, mtimeMs: 1000, fileSize: 5 }), cacheKey('C:/a.jpg', { size: 256, mtimeMs: 2000, fileSize: 5 }), 'a changed file is a new thumbnail');
  assert.notEqual(cacheKey('C:/a.jpg', { size: 256, mtimeMs: 1000, fileSize: 5 }), cacheKey('C:/a.jpg', { size: 128, mtimeMs: 1000, fileSize: 5 }));

  const ran = [];
  let release;
  const gate = new Promise((r) => (release = r));
  const queue = createQueue({ concurrency: 1, run: async (job) => { ran.push(job.id); if (job.id === 'first') await gate; return job.id; } });
  const first = queue.push({ id: 'first' });
  const second = queue.push({ id: 'second' });
  const third = queue.push({ id: 'third', signal: { aborted: true } });
  const fourth = queue.push({ id: 'fourth' });
  release();
  assert.deepEqual(await Promise.all([first, second, third, fourth]), ['first', 'second', null, 'fourth']);
  assert.deepEqual(ran, ['first', 'fourth', 'second'], 'the newest waiting job ran first; the abandoned one never ran');
});

test('the thumbnailer makes once, keeps on disk, remembers a refusal, and takes one handed in', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-thumbs-'));
  const file = path.join(dir, 'a.png');
  fs.writeFileSync(file, 'not really a png');
  const none = path.join(dir, 'b.webm');
  fs.writeFileSync(none, 'not a clip');
  let made = 0;
  const thumbs = createThumbnailer({
    dir: path.join(dir, 'cache'),
    make: async (target) => {
      made += 1;
      return target === file ? Buffer.from('JPEGBYTES') : null;
    },
  });
  assert.equal(String(await thumbs.bytesFor(file)), 'JPEGBYTES');
  assert.equal(String(await thumbs.bytesFor(file)), 'JPEGBYTES');
  assert.equal(made, 1, 'the second ask came from the cache');
  assert.equal(fs.readdirSync(path.join(dir, 'cache')).length, 1);

  assert.equal(await thumbs.bytesFor(none), null);
  assert.equal(await thumbs.bytesFor(none), null);
  assert.equal(made, 2, 'a refusal is remembered');
  const refused = await thumbs.respond(none, {});
  assert.equal(refused.status, 404);
  assert.equal(refused.headers.get('x-rutba-thumb'), 'none');

  assert.deepEqual(await thumbs.put(none, new Uint8Array([1, 2, 3])), { stored: true });
  assert.deepEqual([...(await thumbs.bytesFor(none))], [1, 2, 3], 'the frame the window drew is served from then on');
  const ok = await thumbs.respond(file, {});
  assert.equal(ok.status, 200);
  assert.equal(ok.headers.get('content-type'), 'image/jpeg');
  assert.equal(await thumbs.bytesFor(path.join(dir, 'missing.jpg')), null);
  fs.rmSync(dir, { recursive: true, force: true });
});

test('the columns fit the width to the nearest, and the tiles fill the row', () => {
  assert.deepEqual(fitColumns({ width: 260, tile: 84, gap: 6, padding: 8 }), { columns: 3, tile: 77 });
  assert.deepEqual(fitColumns({ width: 260, tile: 116, gap: 6, padding: 8 }), { columns: 2, tile: 119 });
  assert.deepEqual(fitColumns({ width: 260, tile: 168, gap: 6, padding: 8 }), { columns: 1, tile: 244 });
  assert.equal(fitColumns({ width: 1200, tile: 172, gap: 6, padding: 12 }).columns, 7);
  assert.equal(gridWindow({ count: 100, width: 260, tile: 77, columns: 3 }).columns, 3, 'a given column count is used as it is');
  assert.deepEqual(foldersOf([{ name: 'z', dir: true }, { name: 'a.jpg' }, { name: 'b', dir: true }]).map((e) => e.name), ['b', 'z']);
});
