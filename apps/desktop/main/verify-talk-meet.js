// Small meetings in Talk, seen in its window: Amira is Rutba Office, Ben is
// plain web code in a hidden window, and the stand-in introducer knows
// meetings the way the real one does (workers/introducer src/meetings.js:
// the same messages, the same refusals, the newcomer offering to whoever is
// there). What is checked is what a person sees and what the other side
// receives: the meeting set up from Meet now with the camera's picture
// looked at first and no microphone opened before joining; the two
// computers connected with each other's picture; mute, a raised hand and
// what is written passing between them; what only the person who set the
// meeting up can do; the question asked when they hang up; being asked into
// someone else's meeting and coming in muted as chosen; and a meeting that is
// locked, full, closed to this person or over saying so with nothing to
// press but Close, before any microphone or camera opens.
// Run alone with RUTBA_VERIFY_ONLY=talk-meet.

import crypto from 'node:crypto';
import { BrowserWindow } from 'electron';
import { sendEvent } from '@rutba/office-shell/electron/main';
import { standInIntroducer } from './verify-talk.js';
import { MEETING_CODE } from './talk-meetings.js';

const MOST = 4;

/**
 * Meetings for the stand-in introducer (its `more`): what the real one
 * answers to each `meet.*` message, for one connection a person. `seat`
 * puts people in a meeting who have no computer here, to fill it; `bar`
 * marks someone as taken out of it.
 */
export function standInMeetings() {
  const meetings = new Map(); // id -> { id, name, starter, at, created, live, recorded, locked, barred: Set }
  const rooms = new Map(); // id -> Set of user ids, in the order they came in
  const asked = []; // every meet.create as it arrived
  let tools = null;
  const person = (id) => ({ id, name: Object.values(tools.people).find((p) => p.id === id)?.name || '' });
  const roomOf = (id) => rooms.get(id) || new Set();
  const view = (m, viewer) => ({
    id: m.id, name: m.name, starter: person(m.starter), at: m.at, created: m.created, live: m.live, recorded: m.recorded, locked: m.locked, groupId: null,
    link: `${tools.base()}/m/${m.id}`, most: MOST,
    people: [...roomOf(m.id)].map((id) => ({ ...person(id), contact: id !== viewer && tools.isContact(viewer, id) })),
  });
  const tellRoom = (id, make, except = null) => { for (const who of roomOf(id)) if (who !== except) tools.to(who, make(who)); };
  const out = (id, who, event) => {
    const room = rooms.get(id);
    if (!room?.has(who)) return;
    room.delete(who);
    if (!room.size) rooms.delete(id);
    tellRoom(id, () => ({ type: 'meet.person', id, event, person: person(who) }));
  };
  const refusal = (m, who) => (m.barred.has(who) ? 'removed' : who === m.starter || tools.isContact(m.starter, who) ? null : 'not-allowed');

  const more = (msg, me, t) => {
    tools = t;
    if (msg.type === ':closed') { for (const id of [...rooms.keys()]) out(id, me.id, 'left'); return true; }
    if (!String(msg.type || '').startsWith('meet.')) return false;
    const answer = (m) => t.conn.send(JSON.stringify(m));
    const fail = (code, extra = {}) => answer({ type: 'error', for: msg.type, code, ref: msg.ref, ...extra });
    if (msg.type === 'meet.create') {
      const name = String(msg.name || '').trim();
      if (!name) { fail('name'); return true; }
      const invite = [].concat(msg.invite || []).map(String);
      if (invite.some((id) => !t.isContact(me.id, id))) { fail('not-a-contact'); return true; }
      const m = { id: `${msg.live === true ? 'live' : 'meet'}-${crypto.randomBytes(15).toString('base64url')}`, name, starter: me.id, at: msg.at == null ? null : Number(msg.at), created: Date.now(), live: msg.live === true, recorded: msg.recorded === true, locked: false, barred: new Set() };
      meetings.set(m.id, m);
      asked.push({ by: me.id, id: m.id, name, invite, tell: msg.tell === true, live: m.live, recorded: m.recorded });
      answer({ type: 'meet.created', ref: msg.ref, meeting: view(m, me.id) });
      if (msg.tell === true) for (const id of invite) t.to(id, { type: 'meet.invited', meeting: view(m, id), by: person(me.id) });
      return true;
    }
    const id = String(msg.id || '');
    if (msg.type === 'meet.leave') { out(id, me.id, 'left'); answer({ type: 'meet.left', ref: msg.ref, id }); return true; }
    if (msg.type === 'meet.signal') {
      const to = String(msg.to || '');
      if (roomOf(id).has(me.id) && roomOf(id).has(to) && to !== me.id) t.to(to, { type: 'meet.signal', id, from: me.id, data: msg.data ?? null });
      return true;
    }
    const m = meetings.get(id);
    if (!m) { fail('no-meeting', { id }); return true; }
    if (msg.type === 'meet.peek' || msg.type === 'meet.join') {
      const why = refusal(m, me.id);
      if (why) { fail(why, { id, starter: { name: person(m.starter).name } }); return true; }
      if (msg.type === 'meet.peek') { answer({ type: 'meet.info', ref: msg.ref, meeting: view(m, me.id) }); return true; }
      const room = roomOf(id);
      if (room.has(me.id)) { answer({ type: 'meet.joined', ref: msg.ref, meeting: view(m, me.id) }); return true; }
      if (m.locked && me.id !== m.starter) { fail('locked', { id }); return true; }
      if (room.size >= MOST) { fail('full', { id, most: MOST }); return true; }
      const before = view(m, me.id);
      if (!rooms.has(id)) rooms.set(id, new Set());
      rooms.get(id).add(me.id);
      tellRoom(id, () => ({ type: 'meet.person', id, event: 'joined', person: person(me.id) }), me.id);
      answer({ type: 'meet.joined', ref: msg.ref, meeting: { ...view(m, me.id), people: before.people } });
      return true;
    }
    // The rest is only for the person who set the meeting up.
    if (m.starter !== me.id) { fail('not-yours', { id }); return true; }
    if (msg.type === 'meet.invite') {
      const invite = [].concat(msg.invite || []).map(String).filter((x) => x && x !== me.id);
      let told = 0;
      for (const who of invite) if (!roomOf(id).has(who) && !m.barred.has(who)) { t.to(who, { type: 'meet.invited', meeting: view(m, who), by: person(me.id) }); told += 1; }
      answer({ type: 'meet.asked', ref: msg.ref, id, told });
    } else if (msg.type === 'meet.lock') {
      m.locked = msg.locked !== false;
      tellRoom(id, (who) => ({ type: 'meet.changed', meeting: view(m, who) }), me.id);
      answer({ type: 'meet.changed', ref: msg.ref, meeting: view(m, me.id) });
    } else if (msg.type === 'meet.remove') {
      const who = String(msg.who || '');
      if (who === me.id || !roomOf(id).has(who)) { fail('nobody', { id }); return true; }
      m.barred.add(who);
      out(id, who, 'removed');
      t.to(who, { type: 'meet.removed', id, why: 'removed', by: person(me.id) });
      answer({ type: 'meet.changed', ref: msg.ref, meeting: view(m, me.id) });
    } else if (msg.type === 'meet.end') {
      tellRoom(id, () => ({ type: 'meet.ended', id, by: person(me.id) }), me.id);
      rooms.delete(id);
      meetings.delete(id);
      answer({ type: 'meet.ended', ref: msg.ref, id });
    }
    return true;
  };

  return {
    more,
    asked,
    get: (id) => meetings.get(id) || null,
    inside: (id) => [...roomOf(id)],
    /** People with no computer here put in a meeting, or taken out of it again. */
    seat: (id, ids) => { if (!rooms.has(id)) rooms.set(id, new Set()); for (const who of ids) rooms.get(id).add(who); },
    unseat: (id, ids) => { for (const who of ids) rooms.get(id)?.delete(who); },
    /** Someone marked as taken out of a meeting, or let back. */
    bar: (id, who, on = true) => { const m = meetings.get(id); if (on) m?.barred.add(who); else m?.barred.delete(who); },
  };
}

/* ── Ben in a meeting: plain web code in a hidden window ──────────────── */

// The rule of the call (renderer/apps/talk/mesh.js): whoever comes in offers to everyone already there, with an
// audio and a video line and a line named 'call' beside them that carries each one's state and what is written.
const BEN = (url) => `<!doctype html><meta charset="utf-8"><body><script>
const log = []; const wire = []; const peers = {};
const mine = { muted: false, camera: true, sharing: false, hand: false };
window.ben = { log, wire, peers, remoteWidth: 0, connected: false, over: '' };
const ws = new WebSocket(${JSON.stringify(url)});
const send = (m) => ws.send(JSON.stringify(m));
let media = null; let meeting = null;
const open = () => Object.values(peers).filter((p) => p.channel && p.channel.readyState === 'open');
const tellAll = () => { for (const p of open()) p.channel.send(JSON.stringify({ t: 'state', muted: mine.muted, camera: mine.camera, sharing: mine.sharing, hand: mine.hand })); };
function wireUp(p, ch) {
  p.channel = ch;
  ch.onopen = () => tellAll();
  ch.onmessage = (e) => {
    let m = null; try { m = JSON.parse(e.data); } catch (err) { return; }
    wire.push(m);
    // Muted by whoever may ask: done, and said.
    if (m.t === 'mute') { mine.muted = true; tellAll(); }
  };
  if (ch.readyState === 'open') tellAll();
}
async function link(id) {
  if (!media) media = await navigator.mediaDevices.getUserMedia({ audio: true, video: true });
  const pc = new RTCPeerConnection({ iceServers: [] });
  const p = peers[id] = { pc, channel: null, pending: [] };
  pc.onicecandidate = (ev) => send({ type: 'meet.signal', id: meeting, to: id, data: { kind: 'candidate', candidate: ev.candidate ? ev.candidate.toJSON() : null } });
  pc.ontrack = (ev) => {
    if (ev.track.kind !== 'video') return;
    const v = document.createElement('video'); v.muted = true; v.autoplay = true; v.srcObject = new MediaStream([ev.track]); document.body.appendChild(v);
    v.onresize = () => { window.ben.remoteWidth = Math.max(window.ben.remoteWidth, v.videoWidth); };
  };
  pc.onconnectionstatechange = () => { window.ben.connected = pc.connectionState === 'connected'; };
  pc.ondatachannel = (ev) => wireUp(p, ev.channel);
  return p;
}
async function give(pc) {
  for (const t of pc.getTransceivers()) { const tr = media.getTracks().find((x) => x.kind === t.receiver.track.kind); if (tr) { await t.sender.replaceTrack(tr); t.direction = 'sendrecv'; } }
}
function drop(id) { const p = peers[id]; if (!p) return; delete peers[id]; try { p.pc.close(); } catch (err) { /* closed */ } window.ben.connected = false; }
async function heard(m) {
  log.push(m);
  if (m.type === 'meet.joined') {
    meeting = m.meeting.id;
    for (const person of m.meeting.people) {
      const p = await link(person.id);
      p.pc.addTransceiver('audio', { direction: 'sendrecv' });
      p.pc.addTransceiver('video', { direction: 'sendrecv' });
      wireUp(p, p.pc.createDataChannel('call'));
      await give(p.pc);
      await p.pc.setLocalDescription(await p.pc.createOffer());
      send({ type: 'meet.signal', id: meeting, to: person.id, data: { kind: 'offer', sdp: p.pc.localDescription.sdp } });
    }
  } else if (m.type === 'meet.signal') {
    const d = m.data || {};
    if (d.kind === 'offer') {
      const p = await link(m.from);
      await p.pc.setRemoteDescription({ type: 'offer', sdp: d.sdp });
      await give(p.pc);
      for (const c of p.pending.splice(0)) await p.pc.addIceCandidate(c).catch(() => {});
      await p.pc.setLocalDescription(await p.pc.createAnswer());
      send({ type: 'meet.signal', id: meeting, to: m.from, data: { kind: 'answer', sdp: p.pc.localDescription.sdp } });
    } else if (d.kind === 'answer' && peers[m.from]) {
      const p = peers[m.from];
      await p.pc.setRemoteDescription({ type: 'answer', sdp: d.sdp });
      for (const c of p.pending.splice(0)) await p.pc.addIceCandidate(c).catch(() => {});
    } else if (d.kind === 'candidate' && peers[m.from]) {
      const p = peers[m.from];
      if (p.pc.remoteDescription) await p.pc.addIceCandidate(d.candidate).catch(() => {}); else p.pending.push(d.candidate);
    }
  } else if (m.type === 'meet.person' && m.event !== 'joined') {
    drop(m.person.id);
  } else if (m.type === 'meet.ended' || m.type === 'meet.removed') {
    window.ben.over = m.type === 'meet.ended' ? 'ended' : 'removed';
    for (const id of Object.keys(peers)) drop(id);
  }
}
// One thing at a time, in the order it came.
let line = Promise.resolve();
ws.onopen = () => send({ type: 'hello', token: 'ben' });
ws.onmessage = (e) => { const m = JSON.parse(e.data); line = line.then(() => heard(m)).catch((err) => { log.push({ type: 'page-error', message: String(err && err.message || err) }); }); };
window.ben.send = send;
window.ben.tell = (change) => { Object.assign(mine, change); tellAll(); return open().length; };
window.ben.say = (id, text) => { for (const p of open()) p.channel.send(JSON.stringify({ t: 'say', id, text })); return open().length; };
</script></body>`;

export async function verifyTalkMeet(h) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const meetings = standInMeetings();
  const intro = standInIntroducer({
    page: BEN,
    more: meetings.more,
    contacts: [['u-amira', 'u-ben'], ['u-ben', 'u-cara'], ['u-ben', 'u-dev'], ['u-ben', 'u-eli']],
    people: { cara: { id: 'u-cara', name: 'Cara Diaz', email: 'cara@example.com' }, dev: { id: 'u-dev', name: 'Dev Patel', email: 'dev@example.com' }, eli: { id: 'u-eli', name: 'Eli Stone', email: 'eli@example.com' } },
  });
  let ben = null;
  let win = null;
  try {
    const port = await intro.listen();
    const url = `ws://127.0.0.1:${port}/ws`;
    win = await open('talk');
    const js = (code) => win.webContents.executeJavaScript(code);
    const q = (sel) => `document.querySelector(${JSON.stringify(sel)})`;
    const click = (sel) => js(`(${q(sel)}?.click(), 1)`);
    const seen = (sel, what, ms = 8000) => ok(until(() => js(`Boolean(${q(sel)})`), what, ms));
    const gone = (sel, what, ms = 8000) => ok(until(() => js(`!${q(sel)}`), what, ms));
    const set = (cls, value) => js(`(() => { const n = document.querySelector('.${cls}'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(value)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    // The last button of the dialog that is open: the one a dialog with nothing else to press has.
    const lastButton = () => js(`([...document.querySelectorAll('.rw-dialog button')].at(-1)?.click(), 1)`);

    // Every microphone and camera this window opens, in order, as the page asks for them.
    await js(`(() => { window.__opened = []; const md = navigator.mediaDevices; const real = md.getUserMedia.bind(md); md.getUserMedia = (c) => { window.__opened.push(c && c.audio ? 'microphone' : 'camera'); return real(c); }; return 1; })()`);
    const opened = () => js(`window.__opened.slice()`);
    const fresh = () => js(`(window.__opened = [], 1)`);

    // Amira signs in with a test server; Ben is her contact already.
    await js(`window.rutbaOffice.talk.signOut().then(() => 1, () => 1)`);
    await until(() => js(`Boolean(document.querySelector('.tk-signin'))`), 'the sign-in', 8000).catch(() => {});
    await js(`document.querySelector('.tk-test')?.setAttribute('open', ''), 1`);
    await set('tk-server', url);
    await set('tk-token', 'amira');
    await click('.tk-connect');
    await until(() => js(`/Signed in as Amira Khan/.test(document.querySelector('.tk-statusbar')?.textContent || '')`), 'signed in', 8000).catch(() => {});
    ben = new BrowserWindow({ show: false, webPreferences: { contextIsolation: true, nodeIntegration: false, backgroundThrottling: false, partition: 'talk-check-ben-meet' } });
    await ben.loadURL(`http://127.0.0.1:${port}/ben`);
    const bjs = (code) => ben.webContents.executeJavaScript(code);
    const benHeard = (test, what, ms = 8000) => ok(until(() => bjs(`window.ben.log.some((m) => ${test})`), what, ms));
    const benWire = (test, what, ms = 8000) => ok(until(() => bjs(`window.ben.wire.some((m) => ${test})`), what, ms));
    await until(() => bjs(`window.ben.log.some((m) => m.type === 'welcome')`), 'Ben signed in', 6000).catch(() => {});
    await seen('.tk-contact[data-contact="u-ben"]', 'Ben among the contacts');

    // Meet now: nothing can start until someone is asked; the camera's picture is looked at first, and no microphone opens yet.
    await click('.tk-meet-now');
    const setUp = await seen('.tk-meet-start', 'the set-up');
    const before = await js(`({ mic: ${q('.tk-device-mic')}?.checked, camera: ${q('.tk-device-camera')}?.checked, start: ${q('.tk-meet-start')}?.disabled, name: ${q('.tk-meet-name')}?.value || '' })`);
    await click('.tk-pick[data-pick="u-ben"] input');
    await click('.tk-device-camera');
    const ownPicture = await ok(until(() => js(`(${q('.tk-devices-picture')}?.videoWidth || 0) > 0`), 'the camera\'s picture before joining', 10000));
    const beforeJoining = await opened();
    const ready = await js(`${q('.tk-meet-start')}?.disabled === false`);
    check('talk meetings: Meet now cannot start with nobody asked; the camera is looked at first and no microphone opens before joining',
      setUp && before.mic === true && before.camera === false && before.start === true && Boolean(before.name) && ownPicture && ready && beforeJoining.join() === 'camera',
      JSON.stringify({ setUp, before, ownPicture, ready, beforeJoining }));

    await click('.tk-meet-start');
    const inIt = await seen('.tk-meeting[data-state="in"][data-meeting]', 'in the meeting', 12000);
    const id = await js(`${q('.tk-meeting')}?.dataset.meeting || ''`);
    const mine = await ok(until(() => js(`(${q('.tk-tile[data-person="u-amira"] video.tk-tile-video')}?.videoWidth || 0) > 0`), 'her own picture in the meeting', 8000));
    const afterJoining = await opened();
    const made = meetings.asked[0] || {};
    const benAsked = await benHeard(`m.type === 'meet.invited' && m.meeting.id === ${JSON.stringify(id)} && m.by.name === 'Amira Khan' && m.meeting.live === false`, 'Ben asked in');
    const plain = await js(`({ live: ${q('.tk-meeting')}?.dataset.live, recorded: ${q('.tk-meeting')}?.dataset.recorded, notice: Boolean(${q('.tk-meeting .tk-live')}), asks: ${q('.tk-meet-leave')}?.dataset.asks })`);
    check('talk meetings: Start sets the meeting up with the person asked told at once, comes in with the camera already looked at (opened once) and the microphone opened only now, and a meeting not streamed or recorded shows no such notice',
      inIt && MEETING_CODE.test(id) && id.startsWith('meet-') && mine && afterJoining.filter((x) => x === 'camera').length === 1 && afterJoining.filter((x) => x === 'microphone').length === 1 && afterJoining[0] === 'camera'
        && made.tell === true && made.live === false && made.invite?.join() === 'u-ben' && benAsked && plain.live === 'no' && plain.recorded === 'no' && !plain.notice && plain.asks === 'yes',
      JSON.stringify({ inIt, id, mine, afterJoining, made, benAsked, plain }));

    // Ben comes in and offers; each sees the other's picture, and each is told how the other is.
    await bjs(`window.ben.send({ type: 'meet.join', id: ${JSON.stringify(id)}, ref: 'b-join-1' }), 1`);
    const connected = await seen('.tk-tile[data-person="u-ben"][data-connected="yes"]', 'Ben connected', 20000);
    const benPicture = await ok(until(() => js(`(${q('.tk-tile[data-person="u-ben"] video.tk-tile-video')}?.videoWidth || 0) > 0`), 'Ben\'s picture', 10000));
    const benSees = await ok(until(() => bjs(`window.ben.remoteWidth > 0`), 'Amira\'s picture at Ben', 10000));
    const benTold = await benWire(`m.t === 'state' && m.camera === true && m.muted === false`, 'Amira\'s state at Ben');
    const tiles = await js(`document.querySelectorAll('.tk-tiles .tk-tile').length`);
    check('talk meetings: someone who comes in is connected straight to the person there, each sees the other\'s picture, and each is told the other\'s microphone and camera',
      connected && benPicture && benSees && benTold && tiles === 2, JSON.stringify({ connected, benPicture, benSees, benTold, tiles }));

    // Mute and a raised hand go to the other side; theirs show here; what is written passes both ways.
    await click('.tk-meet-mic');
    const mutedShown = await ok(until(() => js(`${q('.tk-meet-mic')}?.getAttribute('aria-pressed') === 'true'`), 'mute pressed', 4000));
    const benKnowsMuted = await benWire(`m.t === 'state' && m.muted === true`, 'muted at Ben');
    await click('.tk-meet-hand');
    const benKnowsHand = await benWire(`m.t === 'state' && m.hand === true`, 'the hand at Ben');
    await bjs(`window.ben.tell({ hand: true })`);
    const benHand = await seen('.tk-tile[data-person="u-ben"].hand', 'Ben\'s hand');
    await bjs(`window.ben.say('s-ben-1', 'Can everyone see the agenda?')`);
    const counted = await ok(until(() => js(`${q('.tk-meet-messages .tk-round-count')}?.textContent === '1'`), 'one message waiting', 6000));
    await click('.tk-meet-messages');
    const read = await ok(until(() => js(`[...document.querySelectorAll('.tk-meet-says .tk-bubble:not(.mine) .tk-text')].some((n) => n.textContent === 'Can everyone see the agenda?') && !${q('.tk-meet-messages .tk-round-count')}`), 'the message read', 6000));
    const from = await js(`${q('.tk-meet-says .tk-bubble:not(.mine) .tk-from')}?.textContent || ''`);
    await js(`(() => { const n = ${q('.tk-meet-input')}; Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(n, 'Yes, page two.'); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await click('.tk-meet-send');
    const benRead = await benWire(`m.t === 'say' && m.text === 'Yes, page two.'`, 'the reply at Ben');
    const wrote = await js(`[...document.querySelectorAll('.tk-meet-says .tk-bubble.mine .tk-text')].map((n) => n.textContent)`);
    check('talk meetings: mute and a raised hand reach the other side, theirs show on their tile, and what is written passes both ways with a count while unread',
      mutedShown && benKnowsMuted && benKnowsHand && benHand && counted && read && from === 'Ben Ortiz' && benRead && wrote.join() === 'Yes, page two.',
      JSON.stringify({ mutedShown, benKnowsMuted, benKnowsHand, benHand, counted, read, from, benRead, wrote }));

    // What only the person who set the meeting up can do: mute someone, lock the meeting; and the link copied is the meeting's.
    await bjs(`window.ben.tell({ hand: false })`);
    await click('.tk-meet-people');
    const row = await seen('.tk-meet-row[data-person="u-ben"] .tk-meet-mute', 'Ben\'s row');
    const canTakeOut = await js(`Boolean(${q('.tk-meet-row[data-person="u-ben"] .tk-meet-remove')}) && Boolean(${q('.tk-meet-end')}) && Boolean(${q('.tk-meet-lock input')})`);
    await click('.tk-meet-row[data-person="u-ben"] .tk-meet-mute');
    const benMuted = await benWire(`m.t === 'mute'`, 'Ben asked to mute');
    const shownMuted = await ok(until(() => js(`${q('.tk-meet-row[data-person="u-ben"] .tk-meet-mute')}?.disabled === true`), 'Ben shown muted', 6000));
    await click('.tk-meet-lock input');
    const locked = await seen('.tk-meet-locked', 'the meeting shown locked');
    const benKnowsLocked = await benHeard(`m.type === 'meet.changed' && m.meeting.locked === true`, 'Ben told it is locked');
    await click('.tk-meet-copy');
    const link = meetings.get(id) ? await js(`window.rutbaOffice.clipboard.readText()`).catch(() => '') : '';
    check('talk meetings: the person who set it up can mute someone and lock the meeting, the others are told, and Copy meeting link copies its link',
      row && canTakeOut && benMuted && shownMuted && locked && meetings.get(id)?.locked === true && benKnowsLocked && String(link).endsWith(`/m/${id}`),
      JSON.stringify({ row, canTakeOut, benMuted, shownMuted, locked, benKnowsLocked, link }));

    // Hanging up asks: Cancel stays in, End for everyone ends it for the other side too.
    await click('.tk-meet-leave');
    const asks = await seen('.tk-meet-end-all', 'the question');
    const choices = await js(`({ end: Boolean(${q('.tk-meet-end-all')}), leave: Boolean(${q('.tk-meet-leave-only')}), cancel: Boolean(${q('.tk-meet-end-all')}?.previousElementSibling) })`);
    await js(`(${q('.tk-meet-end-all')}?.previousElementSibling?.click(), 1)`);
    const stayed = await gone('.tk-meet-end-all', 'the question closed') && await js(`${q('.tk-meeting')}?.dataset.state === 'in'`) && meetings.inside(id).length === 2;
    await click('.tk-meet-leave');
    await seen('.tk-meet-end-all', 'the question again');
    await click('.tk-meet-end-all');
    const benEnded = await ok(until(() => bjs(`window.ben.over === 'ended'`), 'ended at Ben', 8000));
    const left = await gone('.tk-meeting', 'out of the meeting');
    check('talk meetings: the person who set it up is asked when they hang up; Cancel stays in, and End for everyone ends it for the others and takes its link away',
      asks && choices.end && choices.leave && choices.cancel && stayed && benEnded && left && !meetings.get(id), JSON.stringify({ asks, choices, stayed, benEnded, left, kept: Boolean(meetings.get(id)) }));

    // The other way: Ben sets one up and asks Amira in. A card says who asks; before joining she sees whose it is and that
    // it is not streamed or recorded; she turns her microphone off, and comes in muted, offering to Ben.
    await bjs(`window.ben.send({ type: 'meet.create', ref: 'b-make-1', name: 'Budget review', live: false, invite: ['u-amira'], tell: true }), 1`);
    await benHeard(`m.type === 'meet.created' && m.ref === 'b-make-1'`, 'Ben\'s meeting');
    const id2 = await bjs(`window.ben.log.find((m) => m.type === 'meet.created' && m.ref === 'b-make-1')?.meeting.id || ''`);
    await bjs(`window.ben.send({ type: 'meet.join', id: ${JSON.stringify(id2)}, ref: 'b-join-2' }), 1`);
    const card = await seen(`.tk-meet-card[data-meeting="${id2}"]`, 'the card');
    const cardSays = await js(`${q('.tk-meet-card')}?.textContent || ''`);
    await fresh();
    await click('.tk-meet-card-join');
    const looked = await seen(`.tk-meet-seen[data-meeting="${id2}"][data-can="yes"]`, 'the meeting before joining');
    const lookedAt = await js(`({ says: ${q('.tk-meet-seen')}?.textContent || '', plain: Boolean(${q('.tk-meet-seen .tk-meet-plain')}), notice: Boolean(${q('.tk-meet-seen .tk-live')}), devices: Boolean(${q('.tk-meet-seen .tk-devices')}) })`);
    const untouched = await opened();
    await click('.tk-device-mic');
    const micOff = await seen('.tk-devices[data-mic="off"]', 'the microphone off');
    await bjs(`(window.ben.wire.length = 0, 1)`);
    await click('.tk-meet-join-go');
    const inBens = await seen(`.tk-meeting[data-state="in"][data-meeting="${id2}"]`, 'in Ben\'s meeting', 12000);
    const benThere = await seen('.tk-tile[data-person="u-ben"][data-connected="yes"]', 'connected to Ben', 20000);
    const offered = await benHeard(`m.type === 'meet.signal' && m.from === 'u-amira' && m.data && m.data.kind === 'offer'`, 'Amira\'s offer at Ben');
    const cameMuted = await js(`${q('.tk-meet-mic')}?.getAttribute('aria-pressed') === 'true'`);
    const benKnows = await benWire(`m.t === 'state' && m.muted === true && m.camera === false`, 'Amira muted at Ben');
    await click('.tk-meet-people');
    await seen('.tk-meet-row[data-person="u-ben"]', 'the people');
    const notHers = await js(`({ asks: ${q('.tk-meet-leave')}?.dataset.asks, mute: Boolean(${q('.tk-meet-mute')}), remove: Boolean(${q('.tk-meet-remove')}), end: Boolean(${q('.tk-meet-end')}), lock: Boolean(${q('.tk-meet-lock')}), copy: Boolean(${q('.tk-meet-copy')}) })`);
    check('talk meetings: asked into someone\'s meeting, a card says who asks; before joining it says whose it is and that it is not streamed or recorded, with nothing opened; she comes in muted as chosen and is connected',
      card && /Budget review/.test(cardSays) && /Ben Ortiz/.test(cardSays) && looked && /Ben Ortiz/.test(lookedAt.says) && lookedAt.plain && !lookedAt.notice && lookedAt.devices && untouched.length === 0 && micOff
        && inBens && benThere && offered && cameMuted && benKnows,
      JSON.stringify({ card, cardSays, looked, lookedAt, untouched, micOff, inBens, benThere, offered, cameMuted, benKnows }));
    check('talk meetings: in someone else\'s meeting there is no muting others, taking out, locking or ending, and hanging up leaves without a question',
      notHers.asks === 'no' && !notHers.mute && !notHers.remove && !notHers.end && !notHers.lock && notHers.copy, JSON.stringify(notHers));
    await click('.tk-meet-leave');
    const leftBens = await gone('.tk-meeting', 'out of Ben\'s meeting');
    const noQuestion = await js(`!${q('.tk-meet-end-all')}`);
    const benToldLeft = await benHeard(`m.type === 'meet.person' && m.event === 'left' && m.person.id === 'u-amira'`, 'Ben told she left');
    check('talk meetings: leaving someone else\'s meeting goes at once, it goes on for them, and they are told', leftBens && noQuestion && benToldLeft && meetings.inside(id2).join() === 'u-ben', JSON.stringify({ leftBens, noQuestion, benToldLeft, inside: meetings.inside(id2) }));

    // A meeting she cannot come into says why with nothing to press but Close, before any microphone or camera opens.
    const refusedAs = async (why) => {
      await fresh();
      sendEvent(win, 'app:link', { meet: id2 });
      const said = await seen(`.tk-meet-refused[data-refused="${why}"]`, `refused as ${why}`);
      const shown = await js(`({ words: ${q('.tk-meet-refused')}?.textContent || '', devices: Boolean(${q('.tk-devices')}), join: Boolean(${q('.tk-meet-join-go')}), opened: window.__opened.slice() })`);
      await lastButton();
      const closed = await gone('.tk-meet-refused', 'the dialog closed');
      return { why, ok: said && shown.words.length > 10 && !shown.devices && !shown.join && shown.opened.length === 0 && closed, said, ...shown, closed };
    };
    await bjs(`window.ben.send({ type: 'meet.lock', id: ${JSON.stringify(id2)}, locked: true, ref: 'b-lock' }), 1`);
    await benHeard(`m.type === 'meet.changed' && m.ref === 'b-lock'`, 'locked');
    const whenLocked = await refusedAs('locked');
    await bjs(`window.ben.send({ type: 'meet.lock', id: ${JSON.stringify(id2)}, locked: false, ref: 'b-unlock' }), 1`);
    await benHeard(`m.type === 'meet.changed' && m.ref === 'b-unlock'`, 'unlocked');
    meetings.seat(id2, ['u-cara', 'u-dev', 'u-eli']);
    const whenFull = await refusedAs('full');
    meetings.unseat(id2, ['u-cara', 'u-dev', 'u-eli']);
    meetings.bar(id2, 'u-amira');
    const whenBarred = await refusedAs('removed');
    meetings.bar(id2, 'u-amira', false);
    check('talk meetings: a meeting that is locked, full, or closed to this person says so before joining, offers nothing but Close, and opens no microphone or camera',
      whenLocked.ok && whenFull.ok && /4/.test(whenFull.words) && whenBarred.ok, JSON.stringify({ whenLocked, whenFull, whenBarred }));

    // Ended while she was out of it: its line in the conversation still offers Join until she asks, then says it has ended.
    await bjs(`window.ben.send({ type: 'meet.end', id: ${JSON.stringify(id2)}, ref: 'b-end' }), 1`);
    await benHeard(`m.type === 'meet.ended' && m.ref === 'b-end'`, 'ended');
    await click('.tk-contact[data-contact="u-ben"]');
    const lineThere = await seen('.tk-line[data-line="meeting"]', 'the meeting\'s line');
    const offeredStill = await js(`Boolean(${q('.tk-line[data-line="meeting"][data-ended="no"] .tk-line-join')})`);
    await fresh();
    await click('.tk-line[data-line="meeting"] .tk-line-join');
    const over = await seen('.tk-meet-refused[data-refused="no-meeting"]', 'told it is over');
    const nothingOpened = (await opened()).length === 0;
    await lastButton();
    await gone('.tk-meet-refused', 'the dialog closed');
    const lineEnded = await ok(until(() => js(`Boolean(${q('.tk-line[data-line="meeting"][data-ended="yes"] .tk-line-ended')}) && !${q('.tk-line[data-line="meeting"] .tk-line-join')}`), 'the line says ended', 8000));
    check('talk meetings: a meeting ended while she was out of it says so when she asks to join, and its line stops offering Join', lineThere && over && nothingOpened && lineEnded, JSON.stringify({ lineThere, offeredStill, over, nothingOpened, lineEnded }));

    await wait(300);
    const complaints = await errorsIn(win);
    const benFaults = await bjs(`window.ben.log.filter((m) => m.type === 'page-error').map((m) => m.message)`);
    check('talk meetings: none of that reported an error, in Office or on the other side', complaints.length === 0 && benFaults.length === 0, [...complaints, ...benFaults].join(' | ') || 'nothing reported');
    await js(`window.rutbaOffice.talk.signOut().then(() => 1)`).catch(() => {});
  } catch (err) {
    check('talk meetings: the check ran to the end', false, err?.message || String(err));
  } finally {
    try { ben?.destroy(); } catch { /* gone */ }
    await intro.close().catch(() => {});
  }
}
