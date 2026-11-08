// Live Broadcast → wide and tall at once (plan A11; owner, 8 Nov: "enabling both
// sends both simultaneously").
//
// Two streams, one wide and one tall, go to a server on this computer that
// answers as a platform's ingest does. The sizes are read back at the ingest
// twice over: from the metadata each stream was sent, and from the H.264
// sequence header itself, which is the size a viewer's decoder will use. Each
// connection is then cut alone, and the other is seen to keep going; the tall
// stream is turned off and on during the broadcast, and the wide one off, so one
// shape runs alone. Nothing leaves this computer's loopback.
// Run alone with RUTBA_VERIFY_ONLY=broadcast-tall.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { withScreenStandIn } from './verify-recorder.js';
import { startIngest } from './verify-broadcast-live.js';
import { addItem, emptyStack } from '@rutba/studio/stack';
import { seedSetup, restoreSetup, readSetup, PANEL, IN_PANEL } from './verify-studio-seed.js';
import { probeMedia } from '@rutba/media/probe';
import { letGoOfStudioWindows } from './verify-let-go.js';

/** The picture size in an FLV AVC sequence header's first SPS. */
export function spsSize(body) {
  const len = body.readUInt16BE(11);
  const nal = body.subarray(13, 13 + len);
  const bytes = [];
  for (let i = 1; i < nal.length; i++) {
    if (i >= 3 && nal[i] === 3 && nal[i - 1] === 0 && nal[i - 2] === 0) continue;
    bytes.push(nal[i]);
  }
  let at = 0;
  const bit = () => { const v = (bytes[at >> 3] >> (7 - (at & 7))) & 1; at += 1; return v; };
  const bits = (n) => { let v = 0; for (let i = 0; i < n; i++) v = v * 2 + bit(); return v; };
  const ue = () => { let z = 0; while (!bit() && z < 32) z += 1; return 2 ** z - 1 + bits(z); };
  const se = () => { const k = ue(); return k % 2 ? (k + 1) / 2 : -k / 2; };
  const profile = bits(8);
  bits(16);
  ue();
  let chroma = 1;
  if ([100, 110, 122, 244, 44, 83, 86, 118, 128, 138, 139, 134, 135].includes(profile)) {
    chroma = ue();
    if (chroma === 3) bit();
    ue();
    ue();
    bit();
    if (bit()) {
      const lists = chroma === 3 ? 12 : 8;
      for (let i = 0; i < lists; i++) {
        if (!bit()) continue;
        let lastScale = 8;
        let nextScale = 8;
        for (let j = 0; j < (i < 6 ? 16 : 64); j++) {
          if (nextScale !== 0) nextScale = (lastScale + se() + 256) % 256;
          lastScale = nextScale === 0 ? lastScale : nextScale;
        }
      }
    }
  }
  ue();
  const poc = ue();
  if (poc === 0) ue();
  else if (poc === 1) { bit(); se(); se(); const n = ue(); for (let i = 0; i < n; i++) se(); }
  ue();
  bit();
  const wide = ue() + 1;
  const high = ue() + 1;
  const frameOnly = bit();
  if (!frameOnly) bit();
  bit();
  let crop = [0, 0, 0, 0];
  if (bit()) crop = [ue(), ue(), ue(), ue()];
  const unitY = 2 * (2 - frameOnly);
  return { width: wide * 16 - 2 * (crop[0] + crop[1]), height: (2 - frameOnly) * high * 16 - unitY * (crop[2] + crop[3]) };
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder
 */
export function verifyBroadcastTall(h, args) {
  return withScreenStandIn(() => tallChecks(h, args));
}

async function tallChecks(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const folder = path.join(dir, 'tall-recordings');
  const docFile = path.join(dir, 'tall-fixture.docx');
  let ingest = null;
  let js = null;
  let kept = null;
  try {
    fs.mkdirSync(folder, { recursive: true });
    ingest = await startIngest();
    fs.writeFileSync(docFile, buildDocx({ styles: true, paragraphs: [{ text: 'Streamed wide and tall by the check.' }] }));
    await open('word', docFile);
    const win = await open('broadcast');
    js = (code) => win.webContents.executeJavaScript(code);
    for (const d of await js(`window.rutbaOffice.live.destinations()`)) await js(`window.rutbaOffice.live.removeDestination({ id: ${JSON.stringify(d.id)} })`);
    const url = `rtmp://127.0.0.1:${ingest.port}/live2`;
    const save = (name, key, shape, platform = 'custom') => js(`window.rutbaOffice.live.saveDestination(${JSON.stringify({ platform, name, url, key, shape })})`);
    await save('Wide stream', 'wide-key-1A', 'wide');
    await save('Tall stream', 'tall-key-2B', 'tall', 'tiktok');
    await until(() => js(`window.rutbaOffice.capture.sources().then((l) => l.some((x) => x.name.includes('tall-fixture')))`), 'the fixture window', 10000).catch(() => {});
    const source = (await js(`window.rutbaOffice.capture.sources()`)).find((x) => x.name.includes('tall-fixture'));
    let stack = { ...emptyStack(), sounds: [{ id: 'mic-1', kind: 'mic', ref: '', name: '', live: true, muted: false, gain: 1 }] };
    stack = addItem(stack, { kind: 'window', ref: source.id, name: source.name, on: false }).stack;
    stack = addItem(stack, { kind: 'camera', ref: '', name: '', on: false, shape: 'circle' }).stack;
    kept = await seedSetup(js, { app: 'broadcast', stack, studio: { size: 720, fps: 30, countdown: 0, folder, liveCopy: false, minimise: false, shot: 'program', shotTimer: 0 } });
    await win.webContents.reload();
    await until(() => js(`Boolean(document.querySelector('.st-canvas')) && Boolean(document.querySelector('.st-livebar'))`), 'the Live Broadcast window', 10000);
    await until(() => js(`Boolean(document.querySelector('.st-scene[data-scene="screen-camera"]'))`), 'the scenes', 8000).catch(() => {});
    await js(`document.querySelector('.st-scene[data-scene="screen-camera"]')?.click(), 1`);
    await until(() => js(`document.querySelector('.st-canvas')?.dataset.layers === 'window camera'`), 'the program', 25000).catch(() => {});

    const click = (sel) => js(`(() => { const n = document.querySelector(${JSON.stringify(sel)}); if (!n) return 'no ' + ${JSON.stringify(sel)}; n.click(); return 1; })()`);

    // Set up on Tall: the scene's tall layout, made from the wide, and a picture dragged is kept for the scene.
    await js(PANEL('setup'));
    await until(() => js(IN_PANEL('setup')), 'Set up', 4000);
    await click('.st-frame-shape [data-value="tall"]');
    const tallShown = await ok(until(() => js(`(() => { const c = document.querySelector('.st-tall-frame canvas'); return Boolean(c) && Number(c.dataset.frames) > 3 && /window/.test(c.dataset.layers) && /camera/.test(c.dataset.layers); })()`), 'the tall frame', 12000));
    const frameRatio = await js(`(() => { const r = document.querySelector('.st-tall-frame').getBoundingClientRect(); return Math.round((r.width / r.height) * 1000) / 1000; })()`);
    const before = await js(`(() => { const b = document.querySelector('.st-tall-box'); return { id: b.dataset.id, top: b.style.top }; })()`);
    await js(`(() => { const b = document.querySelector('.st-tall-box'); const r = b.getBoundingClientRect(); const at = (x, y) => ({ clientX: x, clientY: y, bubbles: true, pointerId: 1, button: 0 }); b.dispatchEvent(new PointerEvent('pointerdown', at(r.left + 20, r.top + 20))); window.dispatchEvent(new PointerEvent('pointermove', at(r.left + 20, r.top + 80))); window.dispatchEvent(new PointerEvent('pointerup', at(r.left + 20, r.top + 80))); return 1; })()`);
    await wait(500);
    const keptTall = await ok(until(async () => { const d = await readSetup(js); return Boolean(d?.stack?.tall?.['screen-camera']?.length); }, 'the tall layout kept', 6000));
    const noted = await js(`document.querySelector('.st-tall-note')?.textContent || ''`);
    if (process.env.RUTBA_VERIFY_CAPTURE) { win.webContents.invalidate(); await wait(700); fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'live-setup-tall.png'), (await win.webContents.capturePage()).toPNG()); }
    check('broadcast tall: Set up has Wide and Tall over the frame; Tall shows the scene in a tall frame made from the wide layout; a picture dragged is kept for the scene',
      tallShown && Math.abs(frameRatio - 0.5625) < 0.01 && keptTall && /kept for this scene/.test(noted),
      JSON.stringify({ tallShown, frameRatio, before, keptTall, noted }));
    await click('.st-tall-reset');
    await wait(300);
    await click('.st-frame-shape [data-value="wide"]');
    await js(PANEL('main'));
    await until(() => js(IN_PANEL('live')), 'back on Live', 4000);

    const listed = await js(`window.rutbaOffice.live.destinations()`);
    const wide = listed.find((d) => d.name === 'Wide stream');
    const tall = listed.find((d) => d.name === 'Tall stream');
    const tile = (d) => `.lv-tile[data-id="${d.id}"]`;
    const tileState = (d) => js(`document.querySelector(${JSON.stringify(tile(d))})?.dataset.state || ''`);
    const go = (d) => click(`${tile(d)} .lv-tile-go`);
    await until(() => js(`Boolean(document.querySelector(${JSON.stringify(tile(tall))}))`), 'the icons', 8000);
    const grouped = await js(`({ wide: document.querySelector('.lv-grp[aria-label="Wide streams"]')?.textContent, tall: document.querySelector('.lv-grp[aria-label="Tall streams"]')?.textContent })`);
    await go(wide);
    await go(tall);
    await until(() => js(`(() => { const b = document.querySelector('.lv-go'); return Boolean(b) && !b.disabled; })()`), 'Go live ready', 8000);
    const planned = await js(`document.querySelector('.lv-note')?.textContent || ''`);
    await click('.lv-go');

    // Both at once: each stream has its own size, read back at the ingest.
    const wentW = await ok(until(async () => (await tileState(wide)) === 'live', 'the wide stream live', 30000));
    const wentT = await ok(until(async () => (await tileState(tall)) === 'live', 'the tall stream live', 30000));
    const named = (key) => ingest.seen.connections.filter((c) => c.key === key);
    const sent = (c, n) => Boolean(c) && c.video.length > n && c.video.some((v) => v.kind === 'key');
    await until(() => sent(named('wide-key-1A')[0], 30) && sent(named('tall-key-2B')[0], 30), 'both streams at the ingest', 60000).catch(() => {});
    const cW = named('wide-key-1A')[0] || { video: [], audio: [], data: [] };
    const cT = named('tall-key-2B')[0] || { video: [], audio: [], data: [] };
    const sizeOf = (c) => {
      const avcc = c.video.find((v) => v.kind === 'config')?.avcc;
      let sps = null;
      try { sps = avcc ? spsSize(avcc) : null; } catch { sps = null; }
      return { meta: c.data[0]?.[2] ? { width: c.data[0][2].width, height: c.data[0][2].height } : null, sps };
    };
    const sW = sizeOf(cW);
    const sT = sizeOf(cT);
    const firstIs = (c) => { const n = c.video.findIndex((v) => v.kind !== 'config'); return c.video[0]?.kind === 'config' && c.video[n]?.kind === 'key'; };
    const whole = (c) => c.video.every((v) => v.sound) && c.audio.every((x) => x.sound);
    check(`broadcast tall: a wide stream and a tall stream go out at once from two encoders, read back at the ingest as ${sW.sps?.width}x${sW.sps?.height} and ${sT.sps?.width}x${sT.sps?.height} in the H.264 headers themselves and in the metadata; each starts on its configurations and a key frame, with sound`,
      wentW && wentT && sW.sps?.width === 1280 && sW.sps?.height === 720 && sW.meta?.width === 1280 && sW.meta?.height === 720 && sT.sps?.width === 720 && sT.sps?.height === 1280 && sT.meta?.width === 720 && sT.meta?.height === 1280 && firstIs(cW) && firstIs(cT) && whole(cW) && whole(cT) && cW.audio.length > 5 && cT.audio.length > 5,
      JSON.stringify({ wentW, wentT, sW, sT, firstW: firstIs(cW), firstT: firstIs(cT), whole: [whole(cW), whole(cT)] }));

    // The tall picture is its own layout: drawn from the same sources, in the tall frame.
    const tallDrawn = await js(`(() => { const c = document.querySelector('.st-tall-canvas'); return c ? { w: c.width, h: c.height, layers: c.dataset.layers, frames: Number(c.dataset.frames) } : null; })()`);
    const note = await js(`document.querySelector('.lv-note')?.textContent || ''`);
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      const data = await js(`document.querySelector('.st-tall-canvas')?.toDataURL('image/png') || ''`);
      if (data) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'live-tall-picture.png'), Buffer.from(data.split(',')[1], 'base64'));
      await wait(900);
      win.webContents.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'live-bar-both.png'), (await win.webContents.capturePage()).toPNG());
    }
    check('broadcast tall: the streams sit under Wide and Tall; the tall picture is drawn from the same sources in its own frame; before going live the bar counted both uploads and during it says each shape\'s rate',
      /Wide stream|Wide/.test(grouped.wide || '') && Boolean(grouped.tall) && tallDrawn?.w === 720 && tallDrawn?.h === 1280 && /window/.test(tallDrawn?.layers || '') && /camera/.test(tallDrawn?.layers || '') && tallDrawn.frames > 10 && /Goes to 2/.test(planned) && /Wide .*Mb\/s · Tall .*Mb\/s/.test(note),
      JSON.stringify({ grouped, tallDrawn, planned, note }));

    // Each drops apart: the other's frames go on arriving.
    const carries = (c, ms = 1200) => { const n = c.video.length; return new Promise((r) => setTimeout(() => r(c.video.length > n), ms)); };
    const socketOf = (c) => [...ingest.seen.sockets].find((s) => s === c.sock);
    const cut = (c) => { const s = socketOf(c); if (s) s.destroy(); return Boolean(s); };
    const cutTall = cut(cT);
    const tRetry = await ok(until(async () => (await tileState(tall)) === 'retrying', 'tall retrying', 15000));
    const wKept = (await tileState(wide)) === 'live' && await carries(cW);
    const tBack = await ok(until(async () => (await tileState(tall)) === 'live', 'tall back', 30000));
    await until(() => sent(named('tall-key-2B').at(-1), 10) && named('tall-key-2B').length === 2, 'the tall stream again', 40000).catch(() => {});
    const cT2 = named('tall-key-2B').at(-1);
    const cutWide = cut(cW);
    const wRetry = await ok(until(async () => (await tileState(wide)) === 'retrying', 'wide retrying', 15000));
    const tKept = (await tileState(tall)) === 'live' && await carries(cT2);
    const wBack = await ok(until(async () => (await tileState(wide)) === 'live', 'wide back', 30000));
    await until(() => sent(named('wide-key-1A').at(-1), 10) && named('wide-key-1A').length === 2, 'the wide stream again', 40000).catch(() => {});
    const cW2 = named('wide-key-1A').at(-1);
    check('broadcast tall: a tall connection cut leaves the wide one going and comes back by itself, and the same the other way, each with its configurations and a key frame first',
      cutTall && cutWide && tRetry && wKept && tBack && wRetry && tKept && wBack && firstIs(cT2) && firstIs(cW2),
      JSON.stringify({ cutTall, cutWide, tRetry, wKept, tBack, wRetry, tKept, wBack, firstT: firstIs(cT2), firstW: firstIs(cW2) }));

    // Tall off during the broadcast, and on: its encoder stops and starts, the wide stream unmoved.
    await go(tall);
    const tOff = await ok(until(async () => (await tileState(tall)) === 'off', 'tall off', 8000));
    const tClosed = await ok(until(() => cT2.closedAt > 0, 'tall closed', 8000));
    await wait(300);
    const onlyWide = await js(`({ note: document.querySelector('.lv-note')?.textContent, canvas: Boolean(document.querySelector('.st-tall-canvas')) })`);
    const wStill = await carries(cW2);
    await go(tall);
    const tOn = await ok(until(async () => (await tileState(tall)) === 'live', 'tall again', 20000));
    await until(() => named('tall-key-2B').length === 3 && sent(named('tall-key-2B').at(-1), 10), 'the third tall connection', 40000).catch(() => {});
    const cT3 = named('tall-key-2B').at(-1);
    const sT3 = sizeOf(cT3);
    check('broadcast tall: the tall stream turned off stops its picture and the wide stream carries on; turned on again its encoder starts then, the same size, on a key frame',
      tOff && tClosed && !onlyWide.canvas && /Wide/.test(onlyWide.note) && !/Tall/.test(onlyWide.note) && wStill && tOn && sT3.sps?.width === 720 && sT3.sps?.height === 1280 && firstIs(cT3),
      JSON.stringify({ tOff, tClosed, onlyWide, wStill, tOn, sT3, first: firstIs(cT3) }));

    // Wide off: the tall stream alone.
    await go(wide);
    const wOff = await ok(until(async () => (await tileState(wide)) === 'off', 'wide off', 8000));
    await wait(500);
    const tAlone = (await tileState(tall)) === 'live' && await carries(cT3);
    const tallNote = await js(`document.querySelector('.lv-note')?.textContent || ''`);
    check('broadcast tall: with the wide stream off the tall one goes on alone', wOff && tAlone && /Tall/.test(tallNote) && !/Wide/.test(tallNote), JSON.stringify({ wOff, tAlone, tallNote }));

    await click('.lv-end');
    await until(() => js(`Boolean(document.querySelector('.lv-ask'))`), 'the question', 8000);
    await click('.lv-ask .lv-end-now');
    const ended = await ok(until(() => js(`Boolean(document.querySelector('.lv-go')) && !document.querySelector('.lv-end')`), 'ended', 15000));

    // Keep a copy, chosen Tall: the file is the tall picture, whatever goes out.
    await until(() => js(`Boolean(document.querySelector('.st-keep-copy'))`), 'Keep a copy', 4000);
    await click('.st-keep-copy');
    await until(() => js(`Boolean(document.querySelector('.st-copy-shape'))`), 'the shape of the copy', 4000);
    await click('.st-copy-shape [data-value="tall"]');
    await wait(200);
    await click('.lv-go');
    await until(async () => (await tileState(tall)) === 'live', 'live with a tall copy', 30000);
    await until(() => js(`document.querySelector('.st-status')?.dataset.phase === 'recording'`), 'the copy recording', 8000);
    await wait(3000);
    // The window's sound outlives a broadcast: the second one in the same window has it too.
    const secondHeard = await ok(until(() => (named('tall-key-2B').at(-1)?.audio.length || 0) > 5 && named('tall-key-2B').length > 1, 'sound on the second broadcast', 15000));
    await click('.lv-end');
    await until(() => js(`Boolean(document.querySelector('.lv-ask'))`), 'the question', 8000);
    await click('.lv-ask .lv-end-now');
    const copySaved =await ok(until(() => js(`/Saved/.test(document.querySelector('.st-status')?.textContent || '')`), 'the copy saved', 15000));
    const copies = fs.readdirSync(folder).filter((n) => /\.webm$/.test(n));
    const probe = copies.length === 1 ? probeMedia(new Uint8Array(fs.readFileSync(path.join(folder, copies[0]))), { complete: true }) : null;
    check('broadcast tall: a second broadcast from the same window has its sound, and Keep a copy set to Tall records the tall picture with that sound, sized as the tall stream is',
      secondHeard && copySaved && copies.length === 1 && probe?.hasVideo && probe.hasAudio && probe.width === 720 && probe.height === 1280,
      JSON.stringify({ secondHeard, copySaved, copies, probe: probe ? { v: probe.hasVideo, a: probe.hasAudio, w: probe.width, h: probe.height, d: probe.duration } : null }));
    const complaints = await errorsIn(win);
    check('broadcast tall: End broadcast ends both, and none of that reported an error', ended && complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('broadcast tall: the check ran to the end', false, err?.message || String(err));
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
