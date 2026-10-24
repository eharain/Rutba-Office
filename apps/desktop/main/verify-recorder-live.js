// Recorder → Go Live (owner, 2026-10-19: "stream videos live on youtube and
// facebook etc").
//
// A server on this computer answers as a platform's ingest does (the
// handshake, connect, createStream, publish) and keeps what it is sent. The
// check adds it as a destination in Go Live's box with a test key, adds a
// YouTube destination with no key (left out, and said so), and goes live on
// a Documents window this check opens (recorded as the stand-in screen,
// withScreenStandIn), with the stand-in camera and microphone. The ingest is given the metadata, the H.264 and AAC
// configurations and a key frame first, then frames on a steady clock, and
// the copy is recorded on disk at the same time. The connection is cut: the
// destination says it is reconnecting, comes back, and is given the
// configurations again before more frames. End Live unpublishes and keeps
// the copy. The key is never handed back to the window, only its last four
// characters. Nothing leaves this computer's loopback.
// Run alone with RUTBA_VERIFY_ONLY=recorder-live.

import fs from 'node:fs';
import path from 'node:path';
import net from 'node:net';
import crypto from 'node:crypto';
import { buildDocx } from '@rutba/ooxml/build';
import { probeMedia } from '@rutba/media/probe';
import { ChunkReader, chunkMessage, encodeAmf0, decodeAmf0 } from '@rutba/stream';
import { withScreenStandIn } from './verify-recorder.js';

/** A local ingest: answers a publisher as YouTube's does, and keeps what each connection sent. */
function startIngest() {
  const seen = { connections: [], sockets: new Set() };
  const server = net.createServer((sock) => {
    const conn = { key: null, data: [], video: [], audio: [], commands: [] };
    seen.connections.push(conn);
    seen.sockets.add(sock);
    sock.on('close', () => seen.sockets.delete(sock));
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
          else if (name === 'publish') { conn.key = rest[0]; send(5, 20, encodeAmf0('onStatus', 0, null, { level: 'status', code: 'NetStream.Publish.Start' }), 1); }
        } else if (m.type === 18) conn.data.push(decodeAmf0(m.body));
        else if (m.type === 9) {
          const kind = m.body[1] === 0 ? 'config' : (m.body[0] >> 4) === 1 ? 'key' : 'inter';
          // Whether the bytes are what a decoder takes: an avcC record (version 1, a profile), or NAL units whose lengths add up exactly.
          let sound = false;
          if (kind === 'config') sound = m.body[5] === 1 && [66, 77, 100].includes(m.body[6]);
          else { let at = 5; let units = 0; while (at + 4 <= m.body.length) { const n = m.body.readUInt32BE(at); at += 4 + n; units += 1; } sound = at === m.body.length && units > 0; }
          conn.video.push({ ts: m.timestamp, kind, size: m.body.length, sound });
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
export function verifyRecorderLive(h, args) {
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
    const win = await open('recorder');
    js = (code) => win.webContents.executeJavaScript(code);
    kept = await js(`window.rutbaOffice.store.get({ key: 'recorder.studio', fallback: null })`);
    for (const d of await js(`window.rutbaOffice.live.destinations()`)) await js(`window.rutbaOffice.live.removeDestination({ id: ${JSON.stringify(d.id)} })`);
    await js(`window.rutbaOffice.store.set({ key: 'recorder.studio', value: ${JSON.stringify({ scene: 'screen-camera', size: 720, fps: 30, countdown: 0, camera: true, mic: true, system: false, folder, liveCopy: true, minimise: false })} }).then(() => 1)`);
    await win.webContents.reload();
    await until(() => js(`Boolean(document.querySelector('.st-canvas'))`), 'the Recorder window', 10000);
    await js(`document.querySelector('.st-kinds button[data-value="window"]')?.click(), 1`);
    await until(() => js(`[...document.querySelectorAll('.st-sources .rec-source')].some((n) => (n.dataset.tip || '').includes('live-fixture'))`), 'the fixture window', 10000).catch(() => {});
    await js(`(() => { const b = [...document.querySelectorAll('.st-sources .rec-source')].find((n) => (n.dataset.tip || '').includes('live-fixture')); b?.click(); return 1; })()`);
    await until(() => js(`document.querySelector('.st-canvas')?.dataset.layers === 'screen camera'`), 'the program', 25000).catch(() => {});

    // Go Live with nowhere to go opens the box; a custom server is added with a test key, and YouTube with none.
    await until(() => js(`(() => { const b = document.querySelector('.st-golive'); return Boolean(b) && !b.disabled; })()`), 'Go Live ready', 8000).catch(() => {});
    await js(`document.querySelector('.st-golive')?.click(), 1`);
    const asked = await ok(until(() => js(`Boolean(document.querySelector('.lv-dialog'))`), 'the destinations box', 8000));
    await js(`document.querySelector('.lv-add-custom')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.lv-dest[data-platform="custom"] .lv-url'))`), 'the custom server fields', 8000).catch(() => {});
    const setField = (sel, value) => js(`(() => { const n = document.querySelector(${JSON.stringify(sel)}); if (!n) return 'missing'; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(value)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    const filled = [await setField('.lv-dest[data-platform="custom"] .lv-url', `rtmp://127.0.0.1:${ingest.port}/live2`), await setField('.lv-dest[data-platform="custom"] .lv-keyfield', 'check-key-7Q4z')];
    await wait(200);
    const saveState = await js(`(() => { const b = document.querySelector('.lv-dest[data-platform="custom"] .lv-save'); return b ? (b.disabled ? 'disabled' : 'enabled') : 'no button'; })()`);
    await js(`document.querySelector('.lv-dest[data-platform="custom"] .lv-save')?.click(), 1`);
    await until(() => js(`window.rutbaOffice.live.destinations().then((l) => l.some((d) => d.platform === 'custom' && d.hasKey))`), 'the destination kept', 5000).catch(() => {});
    await js(`document.querySelector('.lv-add-youtube')?.click(), 1`);
    await until(() => js(`window.rutbaOffice.live.destinations().then((l) => l.some((d) => d.platform === 'youtube'))`), 'YouTube added', 5000).catch(() => {});
    const listed = await js(`window.rutbaOffice.live.destinations()`);
    const shown = await js(`document.querySelector('.lv-dialog')?.textContent || ''`);
    const custom = listed.find((d) => d.platform === 'custom');
    check('recorder live: Go Live with nowhere to go opens the box; a server is added with its key, which is kept and shown only by its last four characters',
      asked && custom?.hasKey && custom.keyHint === '••••7Q4z' && !JSON.stringify(listed).includes('check-key') && !shown.includes('check-key') && listed.some((d) => d.platform === 'youtube' && !d.hasKey && /youtube\.com/.test(d.url)),
      JSON.stringify({ asked, filled, saveState, listed }));
    await js(`document.querySelector('.lv-done')?.click(), 1`);
    await wait(300);

    // Live: the ingest is given the key, the metadata, both configurations and a key frame first, then frames.
    await js(`document.querySelector('.st-golive')?.click(), 1`);
    const went = await ok(until(() => js(`[...document.querySelectorAll('.st-dest')].some((n) => n.dataset.state === 'live')`), 'the destination live', 30000));
    // Until the ingest has been sent the frames the check reads, not a fixed
    // three and a half seconds: beside the gate's other lanes the encoder made
    // 23 pictures and 19 sounds in that time, where the check reads 40 of each.
    const sent = (c, video, audio) => Boolean(c) && c.video.length > video && c.audio.length > audio;
    await until(() => sent(ingest.seen.connections[0], 40, 40), 'the first frames at the ingest', 60000).catch(() => {});
    const c1 = ingest.seen.connections[0] || { video: [], audio: [], data: [] };
    const firstFrame = c1.video.findIndex((v) => v.kind !== 'config');
    const tsUp = c1.video.slice(firstFrame).every((v, i, a) => i === 0 || v.ts >= a[i - 1].ts);
    const keys = c1.video.filter((v) => v.kind === 'key').length;
    const wellFormed = c1.video.every((v) => v.sound) && c1.audio.every((a) => a.sound);
    const recordingCopy = await js(`Boolean(document.querySelector('.st-stop'))`);
    const badge = await js(`document.querySelector('.st-live-badge')?.textContent || ''`);
    check('recorder live: live on the destination with its key, the metadata, the H.264 and AAC configurations (well formed: an avcC record, AAC-LC) and a key frame first, every frame whole NAL units, frames on a rising clock, sound alongside, a copy recording, and LIVE on the picture',
      went && c1.key === 'check-key-7Q4z' && c1.data.length === 1 && c1.data[0][2]?.width === 1280 && c1.video[0]?.kind === 'config' && c1.video[firstFrame]?.kind === 'key' && c1.video.length > 40 && keys >= 1 && tsUp && c1.audio[0]?.kind === 'config' && c1.audio.length > 40 && wellFormed && recordingCopy && /LIVE/.test(badge),
      JSON.stringify({ went, wellFormed, key: c1.key === 'check-key-7Q4z', data: c1.data.length, video: c1.video.length, first: c1.video.slice(0, 3).map((v) => v.kind), keys, tsUp, audio: c1.audio.length, recordingCopy, badge }));

    // The line cut: reconnecting, then back, with the configurations again first.
    ingest.cut();
    const dropped = await ok(until(() => js(`[...document.querySelectorAll('.st-dest')].some((n) => n.dataset.state === 'reconnecting')`), 'reconnecting', 15000));
    const back = await ok(until(() => js(`[...document.querySelectorAll('.st-dest')].some((n) => n.dataset.state === 'live')`), 'back live', 30000));
    await until(() => sent(ingest.seen.connections[1], 20, 0), 'the frames after the reconnect', 60000).catch(() => {});
    const c2 = ingest.seen.connections[1] || { video: [], audio: [], data: [] };
    const again = c2.video.findIndex((v) => v.kind !== 'config');
    check('recorder live: when the line is cut the destination says it is reconnecting, comes back by itself, and is given the configurations and a key frame before more frames',
      dropped && back && c2.data.length === 1 && c2.video[0]?.kind === 'config' && c2.video[again]?.kind === 'key' && c2.audio[0]?.kind === 'config' && c2.video.length > 20,
      JSON.stringify({ dropped, back, data: c2.data.length, first: c2.video.slice(0, 3).map((v) => v.kind), video: c2.video.length, audio: c2.audio.length }));

    // A slow line: the server stops reading, the queue fills, and the picture's bitrate comes down
    // (owner, 2026-10-20: "lower the quality to keep it upto speed"); the sound is not touched.
    for (const sock of ingest.seen.sockets) sock.pause();
    const lowered = await ok(until(() => js(`/lowered to keep up/.test(document.querySelector('.st-status')?.textContent || '')`), 'the bitrate lowered', 30000));
    const rateText = await js(`document.querySelector('.st-status')?.textContent || ''`);
    for (const sock of ingest.seen.sockets) sock.resume();
    check('recorder live: when the line cannot keep up, the picture\'s bitrate comes down and the status bar says so',
      lowered && /\d+ kb\/s/.test(rateText), rateText.slice(0, 160));
    await wait(1500);

    // End Live: unpublished, and the copy kept.
    await js(`document.querySelector('.st-endlive')?.click(), 1`);
    const ended = await ok(until(() => js(`Boolean(document.querySelector('.st-golive')) && Boolean(document.querySelector('.st-saved'))`), 'ended and saved', 15000));
    await wait(500);
    const unpublished = (ingest.seen.connections[1]?.commands || []).includes('FCUnpublish');
    const copies = fs.readdirSync(folder).filter((n) => /\.webm$/.test(n));
    const probe = copies.length === 1 ? probeMedia(new Uint8Array(fs.readFileSync(path.join(folder, copies[0]))), { complete: true }) : null;
    check('recorder live: End Live unpublishes from the server and keeps the copy recorded meanwhile, picture and sound',
      ended && unpublished && copies.length === 1 && probe?.hasVideo && probe?.hasAudio && probe.duration > 5,
      JSON.stringify({ ended, unpublished, copies, probe: probe ? { v: probe.hasVideo, a: probe.hasAudio, d: probe.duration } : null }));

    const complaints = await errorsIn(win);
    check('recorder live: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('recorder live: the check ran to the end', false, err?.message || String(err));
  } finally {
    if (js) {
      await js(`window.rutbaOffice.live.stop()`).catch(() => {});
      for (const d of (await js(`window.rutbaOffice.live.destinations()`).catch(() => [])) || []) await js(`window.rutbaOffice.live.removeDestination({ id: ${JSON.stringify(d.id)} })`).catch(() => {});
      await js(`window.rutbaOffice.store.set({ key: 'recorder.studio', value: ${JSON.stringify(kept ?? null)} }).then(() => 1)`).catch(() => {});
    }
    ingest?.cut();
    ingest?.server.close();
  }
}
