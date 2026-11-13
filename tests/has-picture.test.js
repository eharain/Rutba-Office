// Export to Video never saves a video with no picture (09.0275, 13 Nov 2026:
// with every processor busy, a deck's video came out with its sound and none
// of its picture, the right length and black from end to end). The recording
// is looked at before it is saved: one with no picture is made again, and if
// that one has none either nothing is saved.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { isPicture, withPicture, pictureOrAbort, readFrame, NO_PICTURE } from '../apps/desktop/renderer/recording/has-picture.js';

const PICTURE = { width: 1280, height: 720, alpha: 255 };
const SOUND_ONLY = { width: 0, height: 0, alpha: 0 };
const NOTHING_DRAWN = { width: 1280, height: 720, alpha: 0 };

test('a picture is a size and something drawn on it', () => {
  assert.equal(isPicture(PICTURE), true);
  assert.equal(isPicture(SOUND_ONLY), false, 'the sound alone has no size');
  assert.equal(isPicture(NOTHING_DRAWN), false, 'a frame of the right size with nothing drawn is no picture');
  assert.equal(isPicture(null), false);
});

test('a recording with a picture is kept as it is, made once', async () => {
  let made = 0;
  const out = await withPicture(async () => ({ n: ++made }), async () => PICTURE);
  assert.deepEqual(out, { n: 1 });
  assert.equal(made, 1);
});

test('a recording with the sound and no picture is made again, and the second is kept', async () => {
  let made = 0;
  const seen = [SOUND_ONLY, PICTURE];
  const out = await withPicture(async () => ({ n: ++made }), async (r) => seen[r.n - 1]);
  assert.deepEqual(out, { n: 2 });
});

test('no picture twice over: nothing is handed back to be saved, and the error says why', async () => {
  let made = 0;
  await assert.rejects(withPicture(async () => ({ n: ++made }), async () => NOTHING_DRAWN), (err) => err.code === NO_PICTURE);
  assert.equal(made, 2, 'made twice, no more');
});

test('a recording that cannot be read is kept as before: a slow machine is not taken for a broken file', async () => {
  let made = 0;
  const out = await withPicture(async () => ({ n: ++made }), async () => null);
  assert.deepEqual(out, { n: 1 });
});

test('cancelled, the recording is handed back unlooked at and not made again', async () => {
  let looked = 0;
  const out = await withPicture(async () => 'part', async () => { looked += 1; return SOUND_ONLY; }, { cancelled: () => true });
  assert.equal(out, 'part');
  assert.equal(looked, 0);
});

test('cancelled while it was looked at, a recording with no picture is not made again', async () => {
  let made = 0;
  let stop = false;
  const out = await withPicture(async () => ({ n: ++made }), async () => { stop = true; return SOUND_ONLY; }, { cancelled: () => stop });
  assert.deepEqual(out, { n: 1 });
  assert.equal(made, 1);
});

test('for a caller that saves: a picture is handed back, and cancelled at any point nothing is', async () => {
  assert.deepEqual(await pictureOrAbort(async () => ({ n: 1 }), async () => PICTURE), { n: 1 });
  let stop = false;
  // Cancel pressed while the recording was looked at: an AbortError, not the recording.
  await assert.rejects(pictureOrAbort(async () => ({ n: 1 }), async () => { stop = true; return null; }, { cancelled: () => stop }), (err) => err.name === 'AbortError');
  // Cancelled before it was looked at, a recording with no picture is not handed back to be saved.
  let looked = 0;
  await assert.rejects(pictureOrAbort(async () => 'part', async () => { looked += 1; return SOUND_ONLY; }, { cancelled: () => true }), (err) => err.name === 'AbortError');
  assert.equal(looked, 0);
  await assert.rejects(pictureOrAbort(async () => 'x', async () => SOUND_ONLY), (err) => err.code === NO_PICTURE);
});

test('the reader counts only what it can open as unreadable: once opened, no size or nothing decoded is no picture', () => {
  const src = fs.readFileSync(new URL('../apps/desktop/renderer/recording/has-picture.js', import.meta.url), 'utf8');
  const body = src.slice(src.indexOf('export async function readFrame'));
  assert.equal((body.match(/return null;/g) || []).length, 4, 'null only for no document, not opened, cancelled, and a throw');
  assert.match(body, /if \(cancelled\(\)\) return null;/);
  assert.match(body, /if \(!\(width > 0 && height > 0\)\) return \{ width, height, alpha: 0 \};/);
  assert.match(body, /if \(v\.readyState < 2\) return 0;/);
  assert.match(body, /Math\.abs\(v\.currentTime - time\) < 0\.05/, 'a seek counts only where it was sent');
});

test('with no window to read it in, a frame cannot be read', async () => {
  assert.equal(await readFrame(new Uint8Array(4), 'video/mp4', { doc: null }), null);
});

test('the Export to Video box saves only what was looked at, and says when there was no picture', () => {
  const src = fs.readFileSync(new URL('../apps/desktop/renderer/apps/slides/export-video.js', import.meta.url), 'utf8');
  const go = src.slice(src.indexOf('const go = async'), src.indexOf('return (', src.indexOf('const go = async')));
  assert.ok(go.indexOf('withPicture(') >= 0 && go.indexOf('withPicture(') < go.indexOf('shell.fs.write('), 'looked at before the write');
  assert.match(go, /err\?\.code === NO_PICTURE \? t\('the video encoder gave the sound but no picture/);
  // The second try starts again at the first slide, and the box says why.
  assert.match(go, /setBusy\(\{ at: i \+ 1, of: slides\.length, again: again > 0 \}\)/);
  assert.match(src, /busy\.again\s*\n?\s*\? t\('The first try came out with no picture, so the video is being made again: slide \{at\} of \{of\}\.'/);
});

test('Video\'s own recorded export is looked at before it is written, and says when it starts again', () => {
  const src = fs.readFileSync(new URL('../apps/desktop/renderer/apps/video/export.js', import.meta.url), 'utf8').replace(/\r\n/g, '\n');
  const exp = src.slice(src.indexOf('export async function exportProject'));
  const recorded = exp.slice(exp.indexOf('// A recording plays through in real time'));
  assert.ok(recorded.indexOf('const record = (opts) => pictureOrAbort(') >= 0 && recorded.indexOf('pictureOrAbort(') < recorded.indexOf('shell.fs.write('), 'looked at before the write, and cancelled there nothing is written');
  assert.match(recorded, /if \(again\) options\.onNoPicture\?\.\(\);/);
  assert.equal((recorded.match(/recordProject\(/g) || []).length, 1, 'every recording, fitted or not, goes through the look');
  assert.equal((recorded.match(/await record\(/g) || []).length, 2, 'the fitted and the plain export');
  assert.match(recorded, /readFrame\(made\.blob, made\.blob\.type \|\| 'video\/webm'/, 'looked at as recorded, with no copy');
  const win = fs.readFileSync(new URL('../apps/desktop/renderer/apps/video/index.js', import.meta.url), 'utf8');
  assert.match(win, /onNoPicture: \(\) => setExporting\(\(e\) => \(e \? \{ \.\.\.e, progress: 0,/);
  assert.match(win, /err\?\.code === NO_PICTURE \? t\('the video encoder gave the sound but no picture/);
});
