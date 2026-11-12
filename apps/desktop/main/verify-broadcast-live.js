// Live Broadcast → Go Live (owner, 2026-10-19: "stream videos live on youtube and
// facebook etc").
//
// A server on this computer answers as a platform's ingest does (the
// handshake, connect, createStream, publish) and keeps what it is sent. The
// check adds it as a destination in Go Live's box with a test key, adds a
// YouTube destination with no key (left out, and said so), and goes live on
// a Documents window this check opens (shown from the list on the right and
// recorded as the stand-in screen, withScreenStandIn), with the stand-in
// camera and microphone, with the icons on top of the Live panel choosing where
// it goes: the second stream started mid-broadcast, the first stopped while the
// second goes on, and every icon off with the broadcast still open. Take Picture
// still takes one. The ingest is given the metadata, the H.264 and AAC
// configurations and a key frame first, then frames on a steady clock, and
// the copy is recorded on disk at the same time. The connection is cut: the
// destination says it is reconnecting, comes back, and is given the
// configurations again before more frames. End broadcast asks, unpublishes and keeps
// the copy. The key is never handed back to the window, only its last four
// characters. Nothing leaves this computer's loopback.
// Run alone with RUTBA_VERIFY_ONLY=broadcast-live.

import fs from 'node:fs';
import path from 'node:path';
import net from 'node:net';
import crypto from 'node:crypto';
import { buildDocx } from '@rutba/ooxml/build';
import { probeMedia } from '@rutba/media/probe';
import { ChunkReader, chunkMessage, encodeAmf0, decodeAmf0 } from '@rutba/stream';
import { addItem, emptyStack } from '@rutba/studio/stack';
import { withScreenStandIn } from './verify-recorder.js';
import { seedSetup, restoreSetup, PANEL, IN_PANEL } from './verify-studio-seed.js';
import { letGoOfStudioWindows } from './verify-let-go.js';

/** A local ingest: answers a publisher as YouTube's does, and keeps what each connection sent. */
export function startIngest() {
  const seen = { connections: [], sockets: new Set() };
  const server = net.createServer((sock) => {
    const conn = { key: null, data: [], video: [], audio: [], commands: [], closedAt: 0, sock };
    seen.connections.push(conn);
    seen.sockets.add(sock);
    sock.on('close', () => { seen.sockets.delete(sock); conn.closedAt = Date.now(); });
    sock.on('error', () => {});
    let stage = 'c0c1';
    let buf = Buffer.alloc(0);
    const reader = new ChunkReader();
    const send = (csid, type, body, streamId = 0) => sock.write(chunkMessage({ csid, type, body, streamId, chunkSize: 128 }));
    sock.on('data', (d) => {
      buf = Buffer.concat([buf, d]);
      if (stage === 'c0c1') {
        if (buf.length < 1537) return;
        sock.write(Buffer.concat([Buffer.from([3]), crypto.randomBytes(1536), buf.subarray(1, 1537)]));
        buf = buf.subarray(1537);
        stage = 'c2';
      }
      if (stage === 'c2') {
        if (buf.length < 1536) return;
        buf = buf.subarray(1536);
        stage = 'chunks';
      }
      if (stage !== 'chunks' || !buf.length) return;
      const messages = reader.push(buf);
      buf = Buffer.alloc(0);
      for (const m of messages) {
        if (m.type === 20) {
          const [name, txn, , ...rest] = decodeAmf0(m.body);
          conn.commands.push(name);
          if (name === 'connect') send(3, 20, encodeAmf0('_result', txn, null, { code: 'NetConnection.Connect.Success' }));
          else if (name === 'createStream') send(3, 20, encodeAmf0('_result', txn, null, 1));
          else if (name === 'publish') { conn.key = rest[0]; conn.publishedAt = Date.now(); send(5, 20, encodeAmf0('onStatus', 0, null, { level: 'status', code: 'NetStream.Publish.Start' }), 1); }
        } else if (m.type === 18) conn.data.push(decodeAmf0(m.body));
        else if (m.type === 9) {
          const kind = m.body[1] === 0 ? 'config' : (m.body[0] >> 4) === 1 ? 'key' : 'inter';
          // Whether the bytes are what a decoder takes: an avcC record (version 1, a profile), or NAL units whose lengths add up exactly.
          let sound = false;
          if (kind === 'config') sound = m.body[5] === 1 && [66, 77, 100].includes(m.body[6]);
          else { let at = 5; let units = 0; while (at + 4 <= m.body.length) { const n = m.body.readUInt32BE(at); at += 4 + n; units += 1; } sound = at === m.body.length && units > 0; }
          conn.video.push({ ts: m.timestamp, kind, size: m.body.length, sound, at: Date.now(), avcc: kind === 'config' ? Buffer.from(m.body) : null });
        } else if (m.type === 8) conn.audio.push({ ts: m.timestamp, kind: m.body[1] === 0 ? 'config' : 'frame', sound: m.body[0] === 0xaf && (m.body[1] !== 0 || (m.body[2] >> 3) === 2) });
      }
    });
  });
  return new Promise((resolve) => server.listen(0, '127.0.0.1', () => resolve({ server, port: server.address().port, seen, cut: () => { for (const s of seen.sockets) s.destroy(); } })));
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder
 */
export function verifyBroadcastLive(h, args) {
  return withScreenStandIn(() => liveChecks(h, args));
}

async function liveChecks(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const folder = path.join(dir, 'live-recordings');
  const docFile = path.join(dir, 'live-fixture.docx');
  let ingest = null;
  let js = null;
  let kept = null;
  try {
    fs.mkdirSync(folder, { recursive: true });
    ingest = await startIngest();
    fs.writeFileSync(docFile, buildDocx({ styles: true, paragraphs: [{ text: 'Streamed by the Go Live check.' }] }));
    await open('word', docFile);
    const win = await open('broadcast');
    js = (code) => win.webContents.executeJavaScript(code);
    for (const d of await js(`window.rutbaOffice.live.destinations()`)) await js(`window.rutbaOffice.live.removeDestination({ id: ${JSON.stringify(d.id)} })`);
    await until(() => js(`window.rutbaOffice.capture.sources().then((l) => l.some((x) => x.name.includes('live-fixture')))`), 'the fixture window to be there', 10000).catch(() => {});
    const source = (await js(`window.rutbaOffice.capture.sources()`)).find((x) => x.name.includes('live-fixture'));
    let stack = { ...emptyStack(), sounds: [{ id: 'mic-1', kind: 'mic', ref: '', name: '', live: true, muted: false, gain: 1 }, { id: 'system-1', kind: 'system', ref: '', name: '', live: false, muted: false, gain: 1 }] };
    stack = addItem(stack, { kind: 'window', ref: source.id, name: source.name, on: false }).stack;
    stack = addItem(stack, { kind: 'camera', ref: '', name: '', on: false, shape: 'circle' }).stack;
    kept = await seedSetup(js, { app: 'broadcast', stack, studio: { size: 720, fps: 30, countdown: 0, folder, liveCopy: true, minimise: false, shot: 'program', shotTimer: 0 } });
    await win.webContents.reload();
    await until(() => js(`Boolean(document.querySelector('.st-canvas')) && Boolean(document.querySelector('.st-livebar'))`), 'the Live Broadcast window', 10000);
    // The Screen and camera scene up: the fixture window and the camera.
    await until(() => js(`Boolean(document.querySelector('.st-scene[data-scene="screen-camera"]'))`), 'the scenes', 8000).catch(() => {});
    await js(`document.querySelector('.st-scene[data-scene="screen-camera"]')?.click(), 1`);
    await until(() => js(`document.querySelector('.st-canvas')?.dataset.layers === 'window camera'`), 'the program', 25000).catch(() => {});

    // Nothing set up: Go live cannot be pressed and says why. My streams opens a list; Add a stream is its own short screen.
    const capture = async (name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      await wait(900);
      win.webContents.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await win.webContents.capturePage()).toPNG());
    };
    const setField = (sel, value) => js(`(() => { const n = document.querySelector(${JSON.stringify(sel)}); if (!n) return 'missing'; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(value)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    const click = (sel) => js(`(() => { const n = document.querySelector(${JSON.stringify(sel)}); if (!n) return 'no ' + ${JSON.stringify(sel)}; n.click(); return 1; })()`);
    const seen = (sel, what, ms = 8000) => ok(until(() => js(`Boolean(document.querySelector(${JSON.stringify(sel)}))`), what, ms));
    await seen('.lv-bar', 'the bar', 10000);
    const empty = await js(`(() => { const b = document.querySelector('.lv-go'); return { disabled: b?.disabled, tip: b?.title || b?.dataset.tip, note: document.querySelector('.lv-note')?.textContent, air: document.querySelector('.lv-air')?.textContent, oldButton: getComputedStyle(document.querySelector('.st-golive') || document.body).display === 'none' || !document.querySelector('.st-golive') }; })()`);
    check('broadcast live: with nothing on, Go live cannot be pressed and says "Turn on a stream to go live"',
      empty.disabled && empty.note === 'Turn on a stream to go live' && /Ready/.test(empty.air) && empty.oldButton, JSON.stringify(empty));

    const addStream = async ({ platform = null, url, key, name }) => {
      await click('.lv-add-stream');
      await seen('.lv-form', 'Add a stream');
      if (platform) { await click(`.lv-pt[data-platform="${platform}"]`); await wait(200); }
      if (url) await setField('.lv-form .lv-url', url);
      if (key) await setField('.lv-form .lv-keyfield', key);
      if (name) await setField('.lv-form .lv-name', name);
      await wait(200);
      const state = await js(`document.querySelector('.lv-form .lv-save')?.disabled`);
      await click('.lv-form .lv-save');
      await until(() => js(`Boolean(document.querySelector('.lv-mine')) && !document.querySelector('.lv-form')`), 'back at the list', 8000).catch(() => {});
      return state;
    };
    await click('.lv-cfg');
    const asked = await seen('.lv-mine', 'My streams');
    const url = `rtmp://127.0.0.1:${ingest.port}/live2`;
    await addStream({ platform: 'custom', url, key: 'check-key-7Q4z', name: 'Stream A' });
    await addStream({ platform: 'custom', url, key: 'second-key-9Z2x', name: 'Stream B' });
    await click('.lv-add-stream');
    await seen('.lv-form', 'Add a stream');
    await click('.lv-pt[data-platform="tiktok"]');
    await wait(200);
    const tallSuggested = await js(`document.querySelector('.lv-seg button[data-shape="tall"]')?.getAttribute('aria-checked') === 'true' && Boolean(document.querySelector('.lv-form .lv-url'))`);
    await capture('live-add-stream.png');
    await click('.lv-pt[data-platform="youtube"]');
    await wait(200);
    const noAddressAsked = await js(`!document.querySelector('.lv-form .lv-url') && Boolean(document.querySelector('.lv-more'))`);
    await click('.lv-form .lv-save');
    await until(() => js(`window.rutbaOffice.live.destinations().then((l) => l.some((d) => d.platform === 'youtube'))`), 'YouTube added', 5000).catch(() => {});
    await seen('.lv-mine .lv-dest[data-platform="youtube"] .lv-addkey', 'the Add key button');
    await capture('live-my-streams.png');
    const listed = await js(`window.rutbaOffice.live.destinations()`);
    const shown = await js(`document.querySelector('.lv-dialog')?.textContent || ''`);
    const a = listed.find((d) => d.name === 'Stream A');
    const b = listed.find((d) => d.name === 'Stream B');
    check('broadcast live: My streams is a list with no addresses or keys showing; each new stream is off and its key shown only by its last four characters; the one with no key says so and offers Add key; Tall is suggested for TikTok, which asks for its address; YouTube keeps its address out of the way',
      asked && a?.hasKey && a.keyHint === '••••7Q4z' && !a.on && !b.on && !JSON.stringify(listed).includes('check-key') && !shown.includes('rtmp:') && !shown.includes('7Q4z') && /Needs a key/.test(shown) && listed.some((d) => d.platform === 'youtube' && !d.hasKey && /youtube\.com/.test(d.url)) && tallSuggested && noAddressAsked,
      JSON.stringify({ asked, tallSuggested, noAddressAsked, listed }));
    await click('.lv-done');
    await until(() => js(`!document.querySelector('.lv-mine')`), 'the list closed', 5000).catch(() => {});

    const tile = (d) => `.lv-tile[data-id="${d.id}"]`;
    const tileState = (d) => js(`document.querySelector(${JSON.stringify(tile(d))})?.dataset.state || ''`);
    const youtube = listed.find((d) => d.platform === 'youtube');
    await seen(tile(a), 'the icon');
    await click(`${tile(a)} .lv-tile-go`);
    const turnedOn = await ok(until(async () => (await tileState(a)) === 'on', 'A on', 4000));
    await click(`${tile(youtube)} .lv-tile-go`);
    const keyAsked = await seen('.lv-form .lv-keyfield', 'the key screen from the icon');
    await click('.lv-form .lv-cancel');
    await wait(200);
    await click('.lv-mine .lv-done');
    await until(() => js(`!document.querySelector('.lv-dialog')`), 'closed', 5000).catch(() => {});
    const ready = await js(`({ note: document.querySelector('.lv-note')?.textContent, go: document.querySelector('.lv-go')?.disabled, yt: document.querySelector(${JSON.stringify(tile(youtube))})?.dataset.state, word: document.querySelector(${JSON.stringify(tile(youtube))})?.textContent })`);
    await capture('live-bar-ready.png');
    check('broadcast live: an icon turns a stream on, Go live can then be pressed, and the bar says where it goes; the stream with no key says Add key and its icon opens the key screen',
      turnedOn && keyAsked && ready.go === false && /Goes to 1/.test(ready.note) && ready.yt === 'nokey' && /Add key/.test(ready.word),
      JSON.stringify({ turnedOn, keyAsked, ready }));

    // Live on one of two: the ingest is given the key, the metadata, both configurations and a key frame first, then frames.
    await until(() => js(`(() => { const b = document.querySelector('.lv-go'); return Boolean(b) && !b.disabled; })()`), 'Go live ready', 8000).catch(() => {});
    await click('.lv-go');
    const went = await ok(until(async () => (await tileState(a)) === 'live', 'the stream live', 30000));
    // Until the ingest has been sent the frames the check reads, not a fixed
    // three and a half seconds: beside the gate's other lanes the encoder made
    // 23 pictures and 19 sounds in that time, where the check reads 40 of each.
    const sent = (c, video, audio) => Boolean(c) && c.video.length > video && c.audio.length > audio;
    await until(() => sent(ingest.seen.connections[0], 40, 40), 'the first frames at the ingest', 60000).catch(() => {});
    const c1 = ingest.seen.connections[0] || { video: [], audio: [], data: [] };
    const firstFrame = c1.video.findIndex((v) => v.kind !== 'config');
    const tsUp = c1.video.slice(firstFrame).every((v, i, x) => i === 0 || v.ts >= x[i - 1].ts);
    const keys = c1.video.filter((v) => v.kind === 'key').length;
    const wellFormed = c1.video.every((v) => v.sound) && c1.audio.every((x) => x.sound);
    const recordingCopy = await js(`document.querySelector('.st-status')?.dataset.phase === 'recording'`);
    const badge = await js(`document.querySelector('.st-live-badge')?.textContent || ''`);
    const barLive = await js(`({ air: document.querySelector('.lv-air')?.textContent, end: document.querySelector('.lv-end')?.textContent, go: Boolean(document.querySelector('.lv-go')), tip: document.querySelector(${JSON.stringify(tile(a))})?.textContent })`);
    await capture('live-bar-live.png');
    check('broadcast live: live on the one stream that is on, the other left alone; the metadata, the H.264 and AAC configurations (well formed: an avcC record, AAC-LC) and a key frame first, every frame whole NAL units, frames on a rising clock, sound alongside, a copy recording, LIVE on the picture and the bar, and the one button is End broadcast',
      went && /LIVE \d+:\d\d/.test(barLive.air) && /End broadcast/.test(barLive.end) && !barLive.go && c1.key === 'check-key-7Q4z' && ingest.seen.connections.length === 1 && c1.data.length === 1 && c1.data[0][2]?.width === 1280 && c1.video[0]?.kind === 'config' && c1.video[firstFrame]?.kind === 'key' && c1.video.length > 40 && keys >= 1 && tsUp && c1.audio[0]?.kind === 'config' && c1.audio.length > 40 && wellFormed && recordingCopy && /LIVE/.test(badge),
      JSON.stringify({ went, wellFormed, key: c1.key === 'check-key-7Q4z', conns: ingest.seen.connections.length, data: c1.data.length, video: c1.video.length, first: c1.video.slice(0, 3).map((v) => v.kind), keys, tsUp, audio: c1.audio.length, recordingCopy, badge, barLive }));

    // Take Picture while live: a PNG of the program, and the stream goes on.
    const pictures = () => fs.readdirSync(folder).filter((n) => /^Picture .*\.png$/.test(n));
    await js(`document.querySelector('.st-shoot')?.click(), 1`);
    const pictured = await ok(until(() => pictures().length === 1, 'a picture while live', 8000));
    const stillLive = (await tileState(a)) === 'live';
    check('broadcast live: Take Picture takes one while live, and the stream goes on', pictured && stillLive, JSON.stringify({ pictured, stillLive, pictures: pictures() }));

    // The line cut: retrying, then back, with the configurations again first.
    ingest.cut();
    const dropped = await ok(until(async () => (await tileState(a)) === 'retrying', 'retrying', 15000));
    const back = await ok(until(async () => (await tileState(a)) === 'live', 'back live', 30000));
    await until(() => sent(ingest.seen.connections[1], 20, 0), 'the frames after the reconnect', 60000).catch(() => {});
    const c2 = ingest.seen.connections[1] || { video: [], audio: [], data: [] };
    const again = c2.video.findIndex((v) => v.kind !== 'config');
    check('broadcast live: when the line is cut the stream says Retrying, comes back by itself, and is given the configurations and a key frame before more frames',
      dropped && back && c2.data.length === 1 && c2.video[0]?.kind === 'config' && c2.video[again]?.kind === 'key' && c2.audio[0]?.kind === 'config' && c2.video.length > 20,
      JSON.stringify({ dropped, back, data: c2.data.length, first: c2.video.slice(0, 3).map((v) => v.kind), video: c2.video.length, audio: c2.audio.length }));

    // A slow line: the server stops reading, the queue fills, and the picture's bitrate comes down
    // (owner, 2026-10-20: "lower the quality to keep it upto speed"); the sound is not touched.
    for (const sock of ingest.seen.sockets) sock.pause();
    const lowered = await ok(until(() => js(`/lowered to keep up/.test(document.querySelector('.st-live-rate')?.textContent || '')`), 'the bitrate lowered', 30000));
    const rateText = await js(`document.querySelector('.st-live-rate')?.textContent || ''`);
    for (const sock of ingest.seen.sockets) sock.resume();
    check('broadcast live: when the line cannot keep up, the picture\'s bitrate comes down and the bar says so',
      lowered && /\d+ kb\/s/.test(rateText), rateText.slice(0, 160));
    await wait(1500);

    // The second stream started mid-broadcast: a clean picture from its first frame, and soon.
    const named = (key) => ingest.seen.connections.filter((c) => c.key === key);
    const before = ingest.seen.connections.length;
    await click(`${tile(b)} .lv-tile-go`);
    const joined = await ok(until(async () => (await tileState(b)) === 'live', 'B live', 20000));
    await until(() => { const c = named('second-key-9Z2x')[0]; return sent(c, 20, 20) && c.video.some((v) => v.kind === 'key'); }, 'the second stream\'s frames', 60000).catch(() => {});
    const cB = named('second-key-9Z2x')[0] || { video: [], audio: [], data: [] };
    const bFirst = cB.video.findIndex((v) => v.kind !== 'config');
    const keyAfter = cB.publishedAt && cB.video[bFirst]?.at ? cB.video[bFirst].at - cB.publishedAt : -1;
    const aCarries = await ok(until(() => { const c = named('check-key-7Q4z').at(-1); const n = c.video.length; return new Promise((r) => setTimeout(() => r(c.video.length > n), 800)); }, 'the first stream still carried', 8000));
    // The time it took is in the detail, not the name: a name that changes each run is a new check each run, given a
    // new number (six of them by 12 Nov, 14.0045 the last), and the release refused the numbers file that record wrote.
    check('broadcast live: the second stream, started during the broadcast, is given the metadata and configurations first and starts on a key frame, within two seconds of joining, every frame well formed, and the first keeps going',
      joined && ingest.seen.connections.length === before + 1 && cB.data.length === 1 && cB.video[0]?.kind === 'config' && cB.video[bFirst]?.kind === 'key' && keyAfter >= 0 && keyAfter < 2000 && cB.audio[0]?.kind === 'config' && cB.video.every((v) => v.sound) && aCarries,
      JSON.stringify({ joined, keyAfter, conns: ingest.seen.connections.length, first: cB.video.slice(0, 3).map((v) => v.kind), aCarries }));

    // The first stopped while the second keeps going.
    const cA = named('check-key-7Q4z').at(-1);
    await click(`${tile(a)} .lv-tile-go`);
    const aOff = await ok(until(async () => (await tileState(a)) === 'off', 'the first off', 8000));
    const aGone = await ok(until(() => cA.closedAt > 0, 'the first closed', 8000));
    const aUnpublished = cA.commands.includes('FCUnpublish');
    const nB = cB.video.length;
    await wait(1500);
    const bCarries = cB.video.length > nB;
    check('broadcast live: a stream turned off is unpublished and closed, not tried again, and the other keeps going',
      aOff && aGone && aUnpublished && bCarries && (await tileState(b)) === 'live' && ingest.seen.connections.length === before + 1,
      JSON.stringify({ aOff, aGone, aUnpublished, bCarries, conns: ingest.seen.connections.length }));

    // Every icon off: offline, the broadcast still open, the copy still recording; one back on goes out again.
    await click(`${tile(b)} .lv-tile-go`);
    const offline = await ok(until(() => js(`document.querySelector('.lv-bar')?.dataset.air === 'offline'`), 'offline', 8000));
    await ok(until(() => cB.closedAt > 0, 'the second closed', 8000));
    const held = await js(`({ air: document.querySelector('.lv-air')?.textContent, end: Boolean(document.querySelector('.lv-end')), note: document.querySelector('.lv-note')?.textContent, rec: document.querySelector('.st-status')?.dataset.phase === 'recording' })`);
    await wait(1200);
    const clockRan = (await js(`document.querySelector('.lv-air')?.textContent`)) !== held.air;
    await capture('live-bar-offline.png');
    check('broadcast live: every icon off is OFFLINE with the broadcast still open, the clock running, the copy still recording, and End broadcast still there',
      offline && /OFFLINE \d+:\d\d/.test(held.air) && held.end && held.rec && /Nothing is going out/.test(held.note) && clockRan,
      JSON.stringify({ offline, held, clockRan }));
    const preCount = ingest.seen.connections.length;
    await click(`${tile(a)} .lv-tile-go`);
    const onAgain = await ok(until(async () => (await tileState(a)) === 'live', 'the first live again', 20000));
    await until(() => ingest.seen.connections.length > preCount && sent(ingest.seen.connections.at(-1), 10, 5), 'the frames again', 40000).catch(() => {});
    const cA2 = ingest.seen.connections.at(-1);
    const aFirst = cA2.video.findIndex((v) => v.kind !== 'config');
    check('broadcast live: an icon turned back on goes out again on a new connection, configurations and a key frame first',
      onAgain && cA2.key === 'check-key-7Q4z' && cA2.video[0]?.kind === 'config' && cA2.video[aFirst]?.kind === 'key',
      JSON.stringify({ onAgain, first: cA2.video.slice(0, 3).map((v) => v.kind) }));

    // Set up while live changes only Next: the picture on air stays as it is, and back on Live what changed waits for Take.
    const layersAir = await js(`document.querySelector('.st-air-canvas')?.dataset.layers`);
    await js(PANEL('setup'));
    await until(() => js(IN_PANEL('setup')), 'Set up while live', 4000).catch(() => {});
    const band = await js(`Boolean(document.querySelector('.st-band.live')) && Boolean(document.querySelector('.st-mirror'))`);
    await js(`document.querySelector('.st-scene[data-scene="camera"]')?.click(), 1`);
    const nextChanged = await ok(until(() => js(`document.querySelector('.st-preview-canvas')?.dataset.layers === 'camera'`), 'Next changed', 8000));
    if (process.env.RUTBA_VERIFY_CAPTURE) { win.webContents.invalidate(); await wait(700); fs.mkdirSync(process.env.RUTBA_VERIFY_CAPTURE, { recursive: true }); fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'live-setup-while-live.png'), (await win.webContents.capturePage()).toPNG()); }
    await js(PANEL('main'));
    await until(() => js(IN_PANEL('live')), 'back on Live', 4000).catch(() => {});
    const airHeld = await js(`document.querySelector('.st-air-canvas')?.dataset.layers`);
    const waits = await js(`Boolean(document.querySelector('.st-next-pane')) && Boolean(document.querySelector('.st-take-btn'))`);
    if (process.env.RUTBA_VERIFY_CAPTURE) { win.webContents.invalidate(); await wait(700); fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'live-on-air.png'), (await win.webContents.capturePage()).toPNG()); }
    await js(`document.querySelector('.st-take-btn')?.click(), 1`);
    const taken = await ok(until(() => js(`document.querySelector('.st-air-canvas')?.dataset.layers === 'camera'`), 'Take', 4000));
    check('broadcast live: Set up while live changes only Next and says so, the picture on air stays, and back on Live the change waits for Take',
      band && nextChanged && airHeld === layersAir && waits && taken, JSON.stringify({ band, nextChanged, layersAir, airHeld, waits, taken }));

    // End broadcast asks first, unpublishes, and keeps the copy.
    await click('.lv-end');
    const askedEnd = await seen('.lv-ask', 'the question');
    await click('.lv-ask .lv-end-now');
    const ended = await ok(until(() => js(`Boolean(document.querySelector('.lv-go')) && !document.querySelector('.lv-end') && /Saved/.test(document.querySelector('.st-status')?.textContent || '')`), 'ended and saved', 15000));
    await wait(500);
    const unpublished = (cA2.commands || []).includes('FCUnpublish');
    const copies = fs.readdirSync(folder).filter((n) => /\.webm$/.test(n));
    const probe = copies.length === 1 ? probeMedia(new Uint8Array(fs.readFileSync(path.join(folder, copies[0]))), { complete: true }) : null;
    check('broadcast live: End broadcast asks first, then unpublishes from the server and keeps the copy recorded meanwhile, picture and sound',
      askedEnd && ended && unpublished && copies.length === 1 && probe?.hasVideo && probe?.hasAudio && probe.duration > 5,
      JSON.stringify({ askedEnd, ended, unpublished, copies, probe: probe ? { v: probe.hasVideo, a: probe.hasAudio, d: probe.duration } : null }));

    const complaints = await errorsIn(win);
    check('broadcast live: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('broadcast live: the check ran to the end', false, err?.message || String(err));
  } finally {
    if (js) {
      await js(`window.rutbaOffice.live.stop()`).catch(() => {});
      for (const d of (await js(`window.rutbaOffice.live.destinations()`).catch(() => [])) || []) await js(`window.rutbaOffice.live.removeDestination({ id: ${JSON.stringify(d.id)} })`).catch(() => {});
      if (kept) await restoreSetup(js, kept).catch(() => {});
    }
    ingest?.cut();
    ingest?.server.close();
    letGoOfStudioWindows();
  }
}
