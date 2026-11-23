// A camera on the network as a source, seen in the Recorder's window (Office
// 2.27.0): added by its address as a person adds one, read over RTSP with the
// sign-in the camera asks for, and drawn in the frame from its own H.264.
//
// The camera is the stand-in of the unit tests (tests/netcam-server.js), here
// sending a real picture: the window's own encoder makes a short H.264 film of
// a red band over a bar that travels, and the stand-in plays it round. So the
// picture in the frame was packed by a camera's rules, carried over RTSP,
// put back into frames by Office's reader and decoded by the window.
//
// Then the cameras on the network: a second stand-in answers the question
// Office sends to find cameras (tests/onvif-camera.js, on the loopback, asked
// through `globalThis.__rutbaCameraFind` so nothing is sent to the network the
// computer is on). It is listed by its name, paired with a sign-in, and added
// as the one by address was, its stream address never typed by anyone.

import path from 'node:path';
import { emptyStack } from '@rutba/studio/stack';
import { seedSetup, restoreSetup, readSetup, CHECK_SOUNDS, PANEL, IN_PANEL } from './verify-studio-seed.js';
import { letGoOfStudioWindows } from './verify-let-go.js';
import { discover } from '@rutba/netcam';
import { startCamera } from '../../../tests/netcam-server.js';
import { startOnvifCamera } from '../../../tests/onvif-camera.js';

const USER = 'viewer';
// A sign-in for the stand-in camera of this check only.
const PASS = 'check-camera-27';
const FPS = 15;

/** A short film, encoded by the window: [{ key, b64 }] of Annex B frames, a key frame every second. */
const FILM = `(async () => {
  const W = 640, H = 360, N = ${FPS * 3};
  const c = new OffscreenCanvas(W, H);
  const g = c.getContext('2d');
  const out = [];
  let fault = '';
  const enc = new VideoEncoder({
    output: (chunk) => {
      const b = new Uint8Array(chunk.byteLength);
      chunk.copyTo(b);
      let s = '';
      for (let i = 0; i < b.length; i += 8192) s += String.fromCharCode.apply(null, b.subarray(i, i + 8192));
      out.push({ key: chunk.type === 'key', b64: btoa(s) });
    },
    error: (e) => { fault = String(e?.message || e); },
  });
  enc.configure({ codec: 'avc1.42E01F', width: W, height: H, bitrate: 1200000, framerate: ${FPS}, latencyMode: 'realtime', avc: { format: 'annexb' } });
  for (let i = 0; i < N; i += 1) {
    g.fillStyle = 'rgb(210,25,25)';
    g.fillRect(0, 0, W, H / 2);
    g.fillStyle = 'rgb(20,20,160)';
    g.fillRect(0, H / 2, W, H / 2);
    g.fillStyle = 'rgb(25,210,25)';
    g.fillRect((i * 14) % (W - 80), H / 2, 80, H / 2);
    const f = new VideoFrame(c, { timestamp: Math.round(i * 1e6 / ${FPS}) });
    enc.encode(f, { keyFrame: i % ${FPS} === 0 });
    f.close();
  }
  await enc.flush();
  enc.close();
  return { frames: out, fault };
})()`;

/** The NAL units of an Annex B frame, start codes taken off. */
function nalsOf(bytes) {
  const starts = [];
  for (let i = 0; i + 3 <= bytes.length; i += 1) {
    if (bytes[i] === 0 && bytes[i + 1] === 0 && bytes[i + 2] === 1) { starts.push([i, i + 3]); i += 2; }
  }
  return starts.map(([, from], n) => {
    let to = n + 1 < starts.length ? starts[n + 1][0] : bytes.length;
    // A four-byte start code leaves its first zero at the end of the unit before it.
    while (to > from && n + 1 < starts.length && bytes[to - 1] === 0) to -= 1;
    return bytes.subarray(from, to);
  }).filter((n) => n.length);
}

export async function verifyStudioNetcam(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  let camera = null;
  let yard = null;
  let onvif = null;
  let feed = null;
  let feedYard = null;
  let found = '';
  const standIn = globalThis.__rutbaCameraFind;
  let js = null;
  let seeded = null;
  let added = '';
  try {
    const win = await open('recorder');
    js = (code) => win.webContents.executeJavaScript(code);
    const q = (sel) => `document.querySelector(${JSON.stringify(sel)})`;
    const click = (sel) => js(`(${q(sel)}?.click(), 1)`);
    const seen = (sel, what, ms = 8000) => ok(until(() => js(`Boolean(${q(sel)})`), what, ms));
    const gone = (sel, what, ms = 8000) => ok(until(() => js(`!${q(sel)}`), what, ms));
    const type = (sel, value) => js(`(() => { const n = ${q(sel)}; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(value)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);

    // The film the camera will play, made by this window's encoder.
    const film = await js(FILM);
    const frames = (film.frames || []).map((f) => ({ key: f.key, nals: nalsOf(Buffer.from(f.b64, 'base64')) }));
    const first = frames[0]?.nals || [];
    const sps = first.find((n) => (n[0] & 0x1f) === 7);
    const pps = first.find((n) => (n[0] & 0x1f) === 8);
    const filmed = frames.length === FPS * 3 && !film.fault && Boolean(sps) && Boolean(pps) && frames.filter((f) => f.key).length === 3;

    camera = await startCamera({ auth: { scheme: 'digest', user: USER, password: PASS, algorithm: 'MD5', qop: true }, sprop: sps && pps ? [Buffer.from(sps), Buffer.from(pps)] : null });
    let at = 0;
    let stamp = 0;
    feed = setInterval(() => {
      if (!frames.length || !camera.players.size) { at = 0; return; }
      camera.send(frames[at % frames.length].nals.map((n) => Buffer.from(n)), { timestamp: stamp });
      at += 1;
      stamp = (stamp + Math.round(90000 / FPS)) >>> 0;
    }, Math.round(1000 / FPS));

    // A setup of the check's own, with nothing in the frame.
    seeded = await seedSetup(js, { app: 'recorder', stack: { ...emptyStack(), sounds: CHECK_SOUNDS(false) }, studio: { size: 720, fps: 30, countdown: 0, folder: path.join(dir, 'netcam-recordings'), liveCopy: false, minimise: false, shot: 'program', shotTimer: 0 } });
    await win.webContents.reload();
    await until(() => js(`Boolean(${q('.st-canvas')})`), 'the Recorder window', 15000).catch(() => {});
    await js(PANEL('setup'));
    await until(() => js(IN_PANEL('setup')), 'Set up', 6000).catch(() => {});

    // Add, Camera by address: the address, the sign-in and a name, as a person types them.
    await click('.st-add');
    await seen('.rw-menu', 'the Add menu', 12000);
    const inMenu = await js(`(() => { const b = document.querySelector('.rw-menu [data-key="camera-by-address"]'); if (!b) return [...document.querySelectorAll('.rw-menu button')].map((n) => (n.dataset.scope || '') + '.' + (n.dataset.key || '')).join(' '); b.click(); return 'pressed'; })()`);
    const asked = await seen('.st-cam-dialog .st-cam-url', 'the camera\'s details');
    await type('.st-cam-url', camera.url);
    await type('.st-cam-user', USER);
    await type('.st-cam-pass', PASS);
    await type('.st-cam-name', 'Door camera');
    await ok(until(() => js(`${q('.st-cam-ok')}?.disabled === false`), 'Add ready', 5000));
    await click('.st-cam-ok');
    const closed = await gone('.st-cam-dialog', 'the dialog closed', 10000);
    const CARD = '.st-card[data-kind="netcam"]';
    const carded = await seen(CARD, 'the camera among the sources');
    const live = await ok(until(() => js(`${q(CARD)}?.dataset.state === 'live'`), 'the camera live', 25000));
    const card = await js(`(() => { const n = ${q(CARD)}; return n ? { state: n.dataset.state, on: n.dataset.on, says: n.querySelector('.st-err')?.textContent || '' } : null; })()`);
    if (card?.on === '0') await click(`${CARD} .st-eye`);
    const framed = await ok(until(() => js(`/netcam/.test(${q('.st-canvas')}?.dataset.layers || '')`), 'the camera in the frame', 10000));
    // The frame as the window draws it: the red band above, and where the bar is below.
    const look = () => js(`(() => {
      const c = ${q('.st-canvas')};
      if (!c?.width) return null;
      const t = document.createElement('canvas');
      t.width = 64; t.height = 36;
      const g = t.getContext('2d', { willReadFrequently: true });
      g.drawImage(c, 0, 0, 64, 36);
      const top = [...g.getImageData(16, 8, 1, 1).data].slice(0, 3);
      const row = g.getImageData(0, 27, 64, 1).data;
      let bar = -1, most = 0;
      for (let x = 0; x < 64; x += 1) { const green = row[x * 4 + 1] - row[x * 4] - row[x * 4 + 2]; if (green > most) { most = green; bar = x; } }
      return { top, bar };
    })()`);
    const red = (p) => Boolean(p) && p.top[0] > 150 && p.top[1] < 90 && p.top[2] < 90;
    const drawn = await ok(until(() => look().then((p) => red(p) && p.bar >= 0), 'the camera\'s picture in the frame', 15000));
    const one = await look();
    const moved = await ok(until(() => look().then((p) => red(p) && p.bar >= 0 && p.bar !== one?.bar), 'the bar somewhere else', 6000));
    const two = await look();
    check('network camera: Add, Camera by address takes an address, a sign-in and a name, the camera is read and shows live among the sources, and its picture is drawn in the frame and moves',
      filmed && inMenu === 'pressed' && asked && closed && carded && live && framed && drawn && moved,
      JSON.stringify({ filmed, fault: film.fault || '', inMenu, asked, closed, carded, live, card, framed, drawn, one, moved, two }));

    // The sign-in: the camera asked for it and was answered by Digest, never with the password in the clear; the
    // window is told the camera has a password and is never given it.
    const list = await js(`window.rutbaOffice.capture.cameraList()`);
    const mine = list.find((c) => c.url === camera.url || c.name === 'Door camera');
    added = mine?.id || '';
    const sent = camera.requests.filter((r) => r.headers?.authorization).map((r) => r.headers.authorization);
    const everything = JSON.stringify(list) + JSON.stringify(await js(`(() => { const out = []; for (let i = 0; i < localStorage.length; i += 1) out.push(localStorage.getItem(localStorage.key(i))); return out; })()`));
    const inSetup = JSON.stringify(await readSetup(js)) + JSON.stringify(await js(`[...document.querySelectorAll(${JSON.stringify(CARD)})].map((n) => n.outerHTML)`));
    check('network camera: the camera asked for a sign-in and was answered by Digest with the name and never the password itself; the window\'s list says a password is kept and does not carry it, nor do the source\'s card and the saved setup',
      Boolean(mine) && mine.hasPassword === true && mine.user === USER && !('password' in mine) && sent.length > 0 && sent.every((a) => /^Digest /i.test(a) && a.includes(`username="${USER}"`) && !a.includes(PASS)) && !everything.includes(PASS) && !inSetup.includes(PASS) && !String(mine.url).includes(PASS),
      JSON.stringify({ mine, answers: sent.length, digest: sent.every((a) => /^Digest /i.test(a)) }));

    // The camera loses its connection, as one that loses power does: the source says it waits, and is live again
    // by itself when the camera is back.
    const before = camera.requests.length;
    camera.cut();
    const waits = await ok(until(() => js(`${q(CARD)}?.dataset.state !== 'live'`), 'the camera not live', 15000));
    const lost = await js(`${q(CARD)}?.dataset.state || ''`);
    const back = await ok(until(() => js(`${q(CARD)}?.dataset.state === 'live'`), 'the camera live again', 40000));
    const again = camera.requests.length > before;
    const drawnAgain = await ok(until(() => look().then((p) => red(p) && p.bar >= 0), 'its picture again', 15000));
    check('network camera: a camera that drops its connection is no longer shown as live, is asked again without the person doing anything, and its picture comes back',
      waits && back && again && drawnAgain, JSON.stringify({ waits, lost, back, again, drawnAgain }));

    // Stopped on its card, the camera is let go: nothing is read from it while no source needs it.
    await click(`${CARD} .st-run`);
    const stopped = await ok(until(() => js(`${q(CARD)}?.dataset.state !== 'live'`), 'the camera stopped', 8000));
    const letGo = await ok(until(() => camera.players.size === 0 || null, 'no connection to the camera', 10000));
    await click(`${CARD} .st-run`);
    const started = await ok(until(() => js(`${q(CARD)}?.dataset.state === 'live'`), 'the camera live once more', 25000));
    check('network camera: Stop on the camera\'s card lets go of the camera, with no connection kept to it, and Start reads it again',
      stopped && letGo && started, JSON.stringify({ stopped, letGo, started, players: camera.players.size }));

    /* ---- cameras found on the network ---- */
    // The camera added by its address goes first: Remove on its card, and Office forgets it. A found camera is
    // known as one added before by its host alone, and both stand-ins are this computer.
    await click(`${CARD} .st-more`);
    await seen('.rw-menu [data-key="remove"]', 'the menu of the card', 8000);
    await click('.rw-menu [data-key="remove"]');
    const cardGone = await gone(CARD, 'the camera off the list', 8000);
    const doorLetGo = await ok(until(() => camera.players.size === 0 || null, 'no connection to the removed camera', 10000));
    if (added) await js(`window.rutbaOffice.capture.cameraRemove({ id: ${JSON.stringify(added)} })`).catch(() => {});
    added = '';
    // The camera that answers: its stream is a second stand-in playing the same film, at the address it gives.
    const YARD_PATH = '/Streaming/Channels/yard?transportmode=unicast&profile=yard';
    yard = await startCamera({ path: YARD_PATH, auth: { scheme: 'digest', user: USER, password: PASS, algorithm: 'MD5', qop: true }, sprop: sps && pps ? [Buffer.from(sps), Buffer.from(pps)] : null });
    let atYard = 0;
    let stampYard = 0;
    feedYard = setInterval(() => {
      if (!frames.length || !yard.players.size) { atYard = 0; return; }
      yard.send(frames[atYard % frames.length].nals.map((n) => Buffer.from(n)), { timestamp: stampYard });
      atYard += 1;
      stampYard = (stampYard + Math.round(90000 / FPS)) >>> 0;
    }, Math.round(1000 / FPS));
    const yardPort = Number(new URL(yard.url).port);
    onvif = await startOnvifCamera({ name: 'Yard camera', hardware: 'IPC-T1', user: USER, password: PASS, profiles: [{ token: 'yard', name: 'yard', encoding: 'H264', width: 640, height: 360 }], rtspPort: yardPort });
    let answers = [onvif.target];
    globalThis.__rutbaCameraFind = () => discover({ ms: 700, targets: answers, interfaces: ['127.0.0.1'], loopback: true });
    const camerasNow = () => js(`window.rutbaOffice.capture.cameraList()`);
    const openFind = async () => {
      await click('.st-add');
      await seen('.rw-menu', 'the Add menu', 12000);
      await click('.rw-menu [data-key="cameras-network"]');
      return seen('.st-cam-find', 'Cameras on the network');
    };
    const PICK = '.st-cam-find .st-cam-pick[data-host="127.0.0.1"]';
    const small = () => js(`${q(PICK)}?.querySelector('small')?.textContent || ''`);

    // Add, Cameras on the network: it says it is looking, then lists the camera by the name it gave.
    const findOpened = await openFind();
    const looking = await js(`Boolean(${q('.st-cam-looking')}) || Boolean(${q('.st-cam-pick')})`);
    const listed = await seen(PICK, 'the camera that answered', 8000);
    const row = await js(`(() => { const all = [...document.querySelectorAll('.st-cam-find .st-cam-pick')]; return { count: all.length, title: all[0]?.querySelector('b')?.textContent || '', small: all[0]?.querySelector('small')?.textContent || '', address: document.querySelector('.st-cam-find').textContent.includes('onvif') }; })()`);
    check('network camera: Remove on the card of a camera takes it off the list and lets go of it; Add, Cameras on the network asks which cameras are there and lists the one that answers by the name and the model it gave, with no service address shown',
      cardGone && doorLetGo && findOpened && looking && listed && row.count === 1 && row.title === 'Yard camera' && row.small.includes('IPC-T1') && row.small.endsWith('127.0.0.1') && !row.address && onvif.probes.length > 0,
      JSON.stringify({ cardGone, doorLetGo, findOpened, looking, listed, row, probes: onvif.probes.length }));

    // Picked, it asks for the camera's sign-in and offers its name. A wrong password is said under the boxes and
    // nothing is added; the right one adds the camera and closes the dialog.
    await click(PICK);
    const pairing = await seen('.st-cam-find .st-cam-pair', 'the sign-in');
    const offered = await js(`${q('.st-cam-find .st-cam-name')}?.value || ''`);
    const had = (await camerasNow()).length;
    await type('.st-cam-find .st-cam-user', USER);
    await type('.st-cam-find .st-cam-pass', 'not-the-password');
    await click('.st-cam-find .st-cam-ok');
    const refused = await seen('.st-cam-find .st-cam-err', 'the refusal', 12000);
    const refusal = await js(`${q('.st-cam-find .st-cam-err')}?.textContent || ''`);
    const stillAsking = await js(`Boolean(${q('.st-cam-find .st-cam-pair')})`);
    const noneAdded = (await camerasNow()).length === had;
    await type('.st-cam-find .st-cam-pass', PASS);
    await ok(until(() => js(`${q('.st-cam-find .st-cam-ok')}?.disabled === false`), 'Add ready', 5000));
    await click('.st-cam-find .st-cam-ok');
    const findClosed = await gone('.st-cam-find', 'the dialog closed', 15000);
    const after = await camerasNow();
    const paired = after.find((c) => String(c.url).startsWith(`rtsp://127.0.0.1:${yardPort}/`)) || null;
    found = paired?.id || '';
    const foundLive = await ok(until(() => js(`[...document.querySelectorAll(${JSON.stringify(CARD)})].filter((n) => n.dataset.state === 'live').length === 1`), 'the found camera live', 25000));
    const yardAsked = yard.requests.filter((r) => r.headers?.authorization).map((r) => r.headers.authorization);
    check('network camera: a camera picked from the list asks for its sign-in and offers its name; a wrong password is refused in words under the boxes and nothing is added; the right one adds the camera with the stream address it gave, and it is read and live among the sources with the password kept out of the window',
      pairing && offered === 'Yard camera' && refused && refusal.trim().length > 15 && stillAsking && noneAdded && findClosed && Boolean(paired) && after.length === had + 1 && paired.name === 'Yard camera' && paired.user === USER && paired.hasPassword === true && !('password' in paired) && !JSON.stringify(after).includes(PASS) && foundLive && yardAsked.length > 0 && yardAsked.every((a) => /^Digest /i.test(a) && !a.includes(PASS)),
      JSON.stringify({ pairing, offered, refused, refusal, stillAsking, noneAdded, findClosed, paired, cameras: after.length, foundLive, answers: yardAsked.length }));

    // Asked again, the camera is marked as one added before, and picking it takes it with no sign-in asked.
    const cardsBefore = await js(`document.querySelectorAll(${JSON.stringify(CARD)}).length`);
    await openFind();
    await seen(PICK, 'the camera again', 8000);
    const smallNow = await small();
    const marked = smallNow.startsWith(row.small) && smallNow.length > row.small.length + 5;
    await click(PICK);
    const noAsking = !(await js(`Boolean(${q('.st-cam-find .st-cam-pair')})`));
    const closedAtOnce = await gone('.st-cam-find', 'the dialog closed at the pick', 8000);
    const sameCameras = (await camerasNow()).length === after.length;
    const cardsAfter = await js(`document.querySelectorAll(${JSON.stringify(CARD)}).length`);
    check('network camera: found again, a camera added before is marked so, and picking it takes it with no sign-in asked and no second camera kept',
      marked && noAsking && closedAtOnce && sameCameras && cardsAfter >= cardsBefore, JSON.stringify({ smallNow, marked, noAsking, closedAtOnce, sameCameras, cardsBefore, cardsAfter }));

    // No camera answers: the dialog says so and offers to look again or to add one by its address.
    answers = [];
    await openFind();
    const none = await seen('.st-cam-find .st-cam-none', 'no camera answered', 8000);
    const offers = await js(`Boolean(${q('.st-cam-find .st-cam-again')}) && Boolean(${q('.st-cam-find .st-cam-by-address')}) && !${q('.st-cam-find .st-cam-pick')}`);
    await click('.st-cam-find .st-cam-by-address');
    const byAddress = await seen('.st-cam-dialog .st-cam-url', 'the camera by address', 8000);
    await js(`([...document.querySelectorAll('.st-cam-dialog .rw-dialog-foot button')].find((b) => !b.classList.contains('st-cam-ok'))?.click(), 1)`);
    const cancelled = await gone('.st-cam-dialog', 'the dialog closed', 8000);
    check('network camera: when no camera answers the dialog says so, and offers to look again or to add a camera by its address, which opens the address dialog',
      none && offers && byAddress && cancelled, JSON.stringify({ none, offers, byAddress, cancelled }));

    await wait(300);
    const complaints = await errorsIn(win);
    check('network camera: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('network camera: the check ran to the end', false, err?.message || String(err));
  } finally {
    clearInterval(feed);
    clearInterval(feedYard);
    if (standIn === undefined) delete globalThis.__rutbaCameraFind;
    else globalThis.__rutbaCameraFind = standIn;
    if (js) {
      if (added) await js(`window.rutbaOffice.capture.cameraRemove({ id: ${JSON.stringify(added)} })`).catch(() => {});
      if (found) await js(`window.rutbaOffice.capture.cameraRemove({ id: ${JSON.stringify(found)} })`).catch(() => {});
      if (seeded) await restoreSetup(js, seeded).catch(() => {});
    }
    await camera?.close().catch(() => {});
    await yard?.close().catch(() => {});
    await onvif?.close().catch(() => {});
    letGoOfStudioWindows();
  }
}
