// Sound or video at a web address fetched into a file, for Video's audio
// library (apps/desktop/main/remote-media.js): https only, redirects held to
// the same rule, only sound or video kept, a limit on its size, and no half
// file left behind.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createRemoteMedia, mediaEnding, mediaFileName } from '../apps/desktop/main/remote-media.js';
import { FROM_INTERNET } from '@rutba/office-shell/runs';

const folder = () => fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-remote-media-'));

/** A stand-in for fetch: answers by address from `pages`, and keeps what was asked. */
function fakeFetch(pages) {
  const asked = [];
  const fetchIt = async (url, opts) => {
    asked.push({ url: String(url), opts });
    const page = pages[String(url)];
    if (!page) return new Response('', { status: 404 });
    return new Response(page.body ?? null, { status: page.status || 200, headers: page.headers || {} });
  };
  return { fetchIt, asked };
}

test('a sound file at an https address is fetched into the folder, with no cookies and no referrer', async () => {
  const dir = folder();
  const bytes = new Uint8Array(5000).fill(7);
  const { fetchIt, asked } = fakeFetch({ 'https://example.com/sounds/Rain%20at%20night.mp3': { body: bytes, headers: { 'content-type': 'audio/mpeg' } } });
  const media = createRemoteMedia({ fetch: fetchIt });
  const made = await media.webMedia({ url: 'https://example.com/sounds/Rain%20at%20night.mp3', dir });
  assert.equal(made.name, 'Rain at night.mp3');
  assert.equal(made.path, path.join(dir, 'Rain at night.mp3'));
  assert.equal(made.size, 5000);
  assert.equal(made.type, 'audio/mpeg');
  assert.deepEqual(new Uint8Array(fs.readFileSync(made.path)), bytes);
  assert.equal(asked[0].opts.credentials, 'omit');
  assert.equal(asked[0].opts.referrerPolicy, 'no-referrer');
  assert.equal(asked[0].opts.redirect, 'manual');
  assert.deepEqual(fs.readdirSync(dir), ['Rain at night.mp3'], 'no part file is left');
  fs.rmSync(dir, { recursive: true, force: true });
});

test('the fetched file is marked as from the internet, as a browser marks a download', async () => {
  const wrote = [];
  // The disk as it is, but for the mark, which is kept here: only Windows' own disks hold one.
  const disk = { promises: { ...fs.promises, writeFile: async (file, data) => { wrote.push({ file, data }); } } };
  const page = { 'https://example.com/a.mp3': { body: new Uint8Array(10).fill(1), headers: { 'content-type': 'audio/mpeg' } } };
  const dir = 'C:\\Made by Video';
  const kept = { ...disk, promises: { ...disk.promises, stat: async () => null, mkdir: async () => {}, rename: async () => {}, rm: async () => {}, open: async () => ({ write: async () => {}, close: async () => {} }) } };
  const made = await createRemoteMedia({ fetch: fakeFetch(page).fetchIt, fs: kept, platform: 'win32' }).webMedia({ url: 'https://example.com/a.mp3', dir });
  assert.deepEqual(wrote, [{ file: `${made.path}:Zone.Identifier`, data: FROM_INTERNET }]);
  // A system with no such mark writes none, and a mark that cannot be written loses nothing.
  wrote.length = 0;
  await createRemoteMedia({ fetch: fakeFetch(page).fetchIt, fs: kept, platform: 'linux' }).webMedia({ url: 'https://example.com/a.mp3', dir });
  assert.deepEqual(wrote, []);
  const refuses = { ...kept, promises: { ...kept.promises, writeFile: async () => { throw new Error('no streams on this disk'); } } };
  assert.equal((await createRemoteMedia({ fetch: fakeFetch(page).fetchIt, fs: refuses, platform: 'win32' }).webMedia({ url: 'https://example.com/a.mp3', dir })).size, 10);
});

test('an address with no ending takes one from what the server says it is', async () => {
  const dir = folder();
  const { fetchIt } = fakeFetch({ 'https://example.com/stream/42': { body: new Uint8Array(10), headers: { 'content-type': 'video/mp4; codecs="avc1"' } } });
  const made = await createRemoteMedia({ fetch: fetchIt }).webMedia({ url: 'https://example.com/stream/42', dir });
  assert.equal(made.name, '42.mp4');
  fs.rmSync(dir, { recursive: true, force: true });
});

test('a web page, a picture or any other kind of file is refused, and nothing is written', async () => {
  const dir = folder();
  const { fetchIt } = fakeFetch({
    'https://example.com/watch?v=1': { body: '<html></html>', headers: { 'content-type': 'text/html; charset=utf-8' } },
    'https://example.com/a.png': { body: new Uint8Array(4), headers: { 'content-type': 'image/png' } },
    'https://example.com/page.mp3': { body: '<html></html>', headers: { 'content-type': 'text/html' } },
    'https://example.com/file': { body: new Uint8Array(4), headers: { 'content-type': 'application/octet-stream' } },
  });
  const media = createRemoteMedia({ fetch: fetchIt });
  for (const url of ['https://example.com/watch?v=1', 'https://example.com/a.png', 'https://example.com/page.mp3', 'https://example.com/file']) {
    await assert.rejects(media.webMedia({ url, dir }), /not a sound or a video file/, url);
  }
  await assert.rejects(media.webMedia({ url: 'https://example.com/gone.mp3', dir }), /answered 404/);
  assert.deepEqual(fs.readdirSync(dir), []);
  fs.rmSync(dir, { recursive: true, force: true });
});

test('only https is fetched, on the way in and at every redirect', async () => {
  const dir = folder();
  const { fetchIt, asked } = fakeFetch({
    'https://example.com/a.mp3': { status: 302, headers: { location: 'https://cdn.example.com/files/a.mp3' } },
    'https://cdn.example.com/files/a.mp3': { body: new Uint8Array(3), headers: { 'content-type': 'audio/mpeg' } },
    'https://example.com/down.mp3': { status: 301, headers: { location: 'http://plain.example.com/a.mp3' } },
    'https://example.com/loop.mp3': { status: 302, headers: { location: '/loop.mp3' } },
  });
  const media = createRemoteMedia({ fetch: fetchIt });
  const made = await media.webMedia({ url: 'https://example.com/a.mp3', dir });
  assert.equal(made.size, 3);
  assert.equal(made.name, 'a.mp3');
  await assert.rejects(media.webMedia({ url: 'https://example.com/down.mp3', dir }), /must start https/);
  assert.ok(!asked.some((a) => a.url.startsWith('http://')), 'the plain address is never asked');
  await assert.rejects(media.webMedia({ url: 'http://example.com/a.mp3', dir }), /must start https/);
  await assert.rejects(media.webMedia({ url: 'a.mp3', dir }), /not a web address/);
  await assert.rejects(media.webMedia({ url: 'https://example.com/loop.mp3', dir }), /too many times/);
  await assert.rejects(media.webMedia({ url: 'https://example.com/a.mp3', dir: 'made' }), /no folder/);
  // A check run's own local server is reached only when the run says so.
  const local = fakeFetch({ 'http://127.0.0.1:8123/a.wav': { body: new Uint8Array(2), headers: { 'content-type': 'audio/wav' } } });
  await assert.rejects(createRemoteMedia({ fetch: local.fetchIt }).webMedia({ url: 'http://127.0.0.1:8123/a.wav', dir }), /must start https/);
  assert.equal((await createRemoteMedia({ fetch: local.fetchIt, allowLoopback: true }).webMedia({ url: 'http://127.0.0.1:8123/a.wav', dir })).size, 2);
  fs.rmSync(dir, { recursive: true, force: true });
});

test('a file over the limit is refused, by what the server states or by what it sends, and no half file stays', async () => {
  const dir = folder();
  const { fetchIt } = fakeFetch({
    'https://example.com/stated.mp3': { body: new Uint8Array(10), headers: { 'content-type': 'audio/mpeg', 'content-length': '5000' } },
    'https://example.com/sent.mp3': { body: new Uint8Array(5000), headers: { 'content-type': 'audio/mpeg' } },
    'https://example.com/empty.mp3': { body: new Uint8Array(0), headers: { 'content-type': 'audio/mpeg' } },
  });
  const media = createRemoteMedia({ fetch: fetchIt, limit: 1000 });
  await assert.rejects(media.webMedia({ url: 'https://example.com/stated.mp3', dir }), /too large/);
  await assert.rejects(media.webMedia({ url: 'https://example.com/sent.mp3', dir }), /too large/);
  await assert.rejects(media.webMedia({ url: 'https://example.com/empty.mp3', dir }), /sent nothing/);
  assert.deepEqual(fs.readdirSync(dir), []);
  fs.rmSync(dir, { recursive: true, force: true });
});

test('a file already fetched, of the size the server states, is not fetched again', async () => {
  const dir = folder();
  fs.writeFileSync(path.join(dir, 'a.mp3'), new Uint8Array(6).fill(1));
  const { fetchIt } = fakeFetch({ 'https://example.com/a.mp3': { body: new Uint8Array(6).fill(2), headers: { 'content-type': 'audio/mpeg', 'content-length': '6' } } });
  const made = await createRemoteMedia({ fetch: fetchIt }).webMedia({ url: 'https://example.com/a.mp3', dir });
  assert.equal(made.size, 6);
  assert.equal(fs.readFileSync(made.path)[0], 1, 'the file there is kept');
  fs.rmSync(dir, { recursive: true, force: true });
});

test('the ending and the file name are worked out from the address and the content type', () => {
  assert.equal(mediaEnding('https://e.com/a.mp3', 'audio/mpeg'), 'mp3');
  assert.equal(mediaEnding('https://e.com/a.m4a', 'audio/mpeg'), 'm4a', 'the address\'s own ending is kept when it is a playable one');
  assert.equal(mediaEnding('https://e.com/a', 'audio/ogg'), 'ogg');
  assert.equal(mediaEnding('https://e.com/a.webm', ''), 'webm', 'a server that names no type is taken by the address');
  assert.equal(mediaEnding('https://e.com/a.webm', 'application/octet-stream'), 'webm');
  assert.equal(mediaEnding('https://e.com/a', 'audio/x-unknown'), null);
  assert.equal(mediaEnding('https://e.com/a.mp3', 'text/html'), null);
  assert.equal(mediaFileName('https://e.com/a/b/My%20song.final.mp3?token=1', 'mp3'), 'My song.final.mp3');
  assert.equal(mediaFileName('https://e.com/', 'mp4'), 'e.com.mp4');
  assert.equal(mediaFileName('https://e.com/a%3Ab%2Fc%3F.mp3', 'mp3'), 'a b c.mp3', 'nothing a folder cannot hold');
});

/** A stand-in for fetch whose file comes a piece at a time, each piece when the test lets it through; it keeps the signal it was given. */
function slowFetch({ pieces = 4, headers = { 'content-type': 'audio/mpeg' } } = {}) {
  const state = { signal: null, sent: 0, let: null, cancelled: false };
  const next = () => new Promise((resolve) => { state.let = resolve; });
  const fetchIt = async (url, opts) => {
    state.signal = opts.signal;
    const body = new ReadableStream({
      async pull(controller) {
        if (state.sent >= pieces) { controller.close(); return; }
        await next();
        state.sent += 1;
        controller.enqueue(new Uint8Array(100).fill(state.sent));
      },
      cancel() { state.cancelled = true; },
    });
    return new Response(body, { status: 200, headers });
  };
  return { fetchIt, state, pass: () => { const go = state.let; state.let = null; go?.(); } };
}
/** Waits until something is so, however busy the computer is, and fails plainly if it never is. */
async function until(so, what) {
  for (let i = 0; i < 1000; i += 1) {
    if (so()) return;
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
  assert.fail(`never came: ${what}`);
}

test('a fetch is stopped by the id it was started under: the request is let go of and no half file stays', async () => {
  const dir = folder();
  const { fetchIt, state, pass } = slowFetch();
  const media = createRemoteMedia({ fetch: fetchIt });
  const going = media.webMedia({ url: 'https://example.com/long.mp3', dir, id: 'm1' });
  going.catch(() => {});
  await until(() => state.let, 'the first piece is asked for');
  pass();
  await until(() => state.sent === 1 && state.let && fs.readdirSync(dir).length > 0, 'the first piece is written and the second asked for');
  assert.deepEqual(fs.readdirSync(dir), ['long.mp3.part'], 'the file is on its way');
  assert.deepEqual(media.webMediaStop({ id: 'm1' }), { stopped: true });
  assert.equal(state.signal.aborted, true, 'the request itself was told to stop');
  pass();
  await assert.rejects(going, /The fetch was stopped/);
  assert.deepEqual(fs.readdirSync(dir), [], 'no half file is left');
  assert.equal(state.cancelled, true, 'what was still coming was let go of');
  assert.deepEqual(media.webMediaStop({ id: 'm1' }), { stopped: false }, 'it had ended already');
  fs.rmSync(dir, { recursive: true, force: true });
});

test('stopping one fetch leaves another running, and a fetch with no id is fetched as before', async () => {
  const dir = folder();
  const one = slowFetch({ pieces: 1 });
  const two = slowFetch({ pieces: 1 });
  const pages = { 'https://example.com/one.mp3': one, 'https://example.com/two.mp3': two };
  const media = createRemoteMedia({ fetch: (url, opts) => pages[String(url)].fetchIt(url, opts) });
  const first = media.webMedia({ url: 'https://example.com/one.mp3', dir, id: 'm1' });
  first.catch(() => {});
  const second = media.webMedia({ url: 'https://example.com/two.mp3', dir, id: 'm2' });
  await until(() => one.state.let && two.state.let, 'both fetches wait for their piece');
  assert.deepEqual(media.webMediaStop({ id: 'm1' }), { stopped: true });
  assert.deepEqual(media.webMediaStop({ id: 'nobody' }), { stopped: false });
  assert.deepEqual(media.webMediaStop(), { stopped: false });
  one.pass();
  two.pass();
  await assert.rejects(first, /The fetch was stopped/);
  assert.equal((await second).size, 100);
  assert.equal(two.state.signal.aborted, false);
  // A fetch that ended is forgotten: its id stops nothing.
  assert.deepEqual(media.webMediaStop({ id: 'm2' }), { stopped: false });
  assert.deepEqual(fs.readdirSync(dir), ['two.mp3']);
  fs.rmSync(dir, { recursive: true, force: true });
});
