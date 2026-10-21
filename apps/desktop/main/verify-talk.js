// Talk (owner, 2026-10-20 and -21: calls and messages between contacts on
// both sides, kept plain), checked end to end on this computer.
//
// A stand-in introducer runs here, speaking the introducer's protocol
// (workers/introducer README) over a WebSocket. Office signs in to it as
// Amira. The other person, Ben, is a hidden window of plain web code on the
// same introducer, with Chromium's stand-in camera and microphone. Ben asks
// to be in contact, Amira accepts; they write to each other; Amira calls Ben
// with video, Ben answers, and the picture comes through, the same safety
// code on both sides; Amira hangs up and Ben is told. Run alone with
// RUTBA_VERIFY_ONLY=talk.

import http from 'node:http';
import crypto from 'node:crypto';
import { BrowserWindow } from 'electron';

/* ── a WebSocket end, as small as the stand-in needs ─────────────────── */

function accept(req, socket) {
  const key = req.headers['sec-websocket-key'];
  socket.write(['HTTP/1.1 101 Switching Protocols', 'Upgrade: websocket', 'Connection: Upgrade',
    `Sec-WebSocket-Accept: ${crypto.createHash('sha1').update(`${key}258EAFA5-E914-47DA-95CA-C5AB0DC85B11`).digest('base64')}`, '', ''].join('\r\n'));
  let buf = Buffer.alloc(0);
  const conn = { onmessage: null, onclose: null };
  conn.send = (text) => {
    const body = Buffer.from(text);
    const head = body.length < 126 ? Buffer.from([0x81, body.length]) : body.length < 65536 ? Buffer.from([0x81, 126, body.length >> 8, body.length & 255]) : null;
    if (!head) { const h = Buffer.alloc(10); h[0] = 0x81; h[1] = 127; h.writeBigUInt64BE(BigInt(body.length), 2); socket.write(Buffer.concat([h, body])); return; }
    socket.write(Buffer.concat([head, body]));
  };
  socket.on('data', (chunk) => {
    buf = Buffer.concat([buf, chunk]);
    for (;;) {
      if (buf.length < 2) return;
      const op = buf[0] & 15;
      let len = buf[1] & 127;
      let at = 2;
      if (len === 126) { if (buf.length < 4) return; len = buf.readUInt16BE(2); at = 4; } else if (len === 127) { if (buf.length < 10) return; len = Number(buf.readBigUInt64BE(2)); at = 10; }
      if (buf.length < at + 4 + len) return;
      const mask = buf.subarray(at, at + 4);
      const data = Buffer.from(buf.subarray(at + 4, at + 4 + len)).map((b, i) => b ^ mask[i & 3]);
      buf = buf.subarray(at + 4 + len);
      if (op === 8) { socket.end(); conn.onclose?.(); return; }
      if (op === 1) conn.onmessage?.(data.toString('utf8'));
    }
  });
  socket.on('close', () => conn.onclose?.());
  socket.on('error', () => {});
  return conn;
}

/* ── the stand-in introducer: the protocol, for two people ─────────────── */

/**
 * The stand-in introducer. For the checks, two people and nothing else. For
 * the screenshots (screen-features.js), more: `people` added to Amira and
 * Ben by their token, `contacts` already made ([id, id] pairs), the
 * `presence` of those not connected ({ id: 'online' | 'busy' | 'away' }),
 * and a `history` of messages ({ from, to, text, at }) handed to whoever
 * signs in, theirs as copies from another device.
 */
export function standInIntroducer({ page = null, people = {}, contacts: made = [], presence: shown = {}, history = [], mailPerDay = 10, mailPerWeek = 30 } = {}) {
  const PEOPLE = {
    amira: { id: 'u-amira', name: 'Amira Khan', email: 'amira@example.com' },
    ben: { id: 'u-ben', name: 'Ben Ortiz', email: 'ben@example.com' },
    ...people,
  };
  const online = new Map(); // id -> conn
  const presence = new Map();
  const contacts = new Set(made.map(([a, b]) => [a, b].sort().join('|'))); // 'a|b' sorted
  const requests = new Map();
  const calls = new Map();
  // Invitations (D-2026-10-21-27): every request has a code, and links have their own.
  const invites = new Map(); // code -> { code, kind, from, note, requestId?, uses, maxUses, created, expires }
  const mailed = []; // the emails Rutba sent: { by, to, link, at }
  const code = () => crypto.randomBytes(18).toString('base64url');
  let base = '';
  const linkOf = (c) => `${base}/i/${c}`;
  const allowanceOf = (id) => {
    const now = Date.now();
    const mine = mailed.filter((m) => m.by === id);
    return { day: { used: mine.filter((m) => now - m.at < 24 * 3600 * 1000).length, limit: mailPerDay }, week: { used: mine.filter((m) => now - m.at < 7 * 24 * 3600 * 1000).length, limit: mailPerWeek } };
  };
  const inviteView = (i) => ({ code: i.code, url: linkOf(i.code), note: i.note, created: i.created, expires: i.expires, uses: i.uses, maxUses: i.maxUses });
  const linksOf = (id) => [...invites.values()].filter((i) => i.kind === 'link' && i.from.id === id).map(inviteView);
  const pair = (a, b) => [a, b].sort().join('|');
  const isContact = (a, b) => contacts.has(pair(a, b));
  const to = (id, msg) => online.get(id)?.send(JSON.stringify(msg));
  const view = (p) => ({ ...p, presence: online.has(p.id) ? (presence.get(p.id) || 'online') === 'hidden' ? 'offline' : presence.get(p.id) || 'online' : shown[p.id] || 'offline' });
  const contactsOf = (id) => Object.values(PEOPLE).filter((p) => p.id !== id && isContact(p.id, id)).map(view);
  const tellContacts = (id) => { for (const c of contactsOf(id)) to(c.id, { type: 'presence', id, presence: view(Object.values(PEOPLE).find((p) => p.id === id)).presence }); };

  // The other person's page is served from here too: 127.0.0.1 is a secure origin, as a camera needs.
  const server = http.createServer((req, res) => {
    const invitation = /^\/i\/([\w-]+)$/.exec(req.url || '');
    if (invitation) { const i = invites.get(invitation[1]); res.writeHead(i ? 200 : 404, { 'content-type': 'text/html; charset=utf-8' }); res.end(i ? `<p>${i.from.name} invites you</p>` : '<p>This invitation has ended.</p>'); return; }
    if (req.url === '/ben' && page) { res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' }); res.end(page(`ws://127.0.0.1:${server.address().port}/ws`)); return; }
    res.writeHead(404);
    res.end();
  });
  // What the sign-in check asks of it: the next sign-in turned away as run out.
  const options = { expireNext: false };
  /** A JWT-shaped token's claims, as the stand-in sign-in issues them; null for a plain test token. */
  const claimsOf = (token) => { try { return JSON.parse(Buffer.from(String(token).split('.')[1], 'base64url').toString('utf8')); } catch { return null; } };
  const sockets = new Set();
  server.on('upgrade', (req, socket) => {
    sockets.add(socket);
    socket.on('close', () => sockets.delete(socket));
    const conn = accept(req, socket);
    conn.socket = socket;
    let me = null;
    conn.onmessage = (text) => {
      const msg = JSON.parse(text);
      if (!me) {
        const claims = claimsOf(msg.token);
        if (claims && (options.expireNext || claims.exp * 1000 < Date.now())) {
          options.expireNext = false;
          conn.send(JSON.stringify({ type: 'error', for: 'hello', code: 'expired', message: 'The sign-in has run out.' }));
          socket.end();
          return;
        }
        me = PEOPLE[msg.token] || (claims && PEOPLE[claims.sub]) || null;
        if (msg.type !== 'hello' || !me) { conn.send(JSON.stringify({ type: 'error', for: 'hello', code: 'token', message: 'That is not a sign-in token.' })); socket.end(); return; }
        online.set(me.id, conn);
        presence.set(me.id, msg.presence || 'online');
        conn.send(JSON.stringify({ type: 'welcome', me, contacts: contactsOf(me.id), requests: { incoming: [...requests.values()].filter((r) => r.toEmail === me.email).map((r) => ({ id: r.id, from: r.from, note: r.note, at: r.at })), outgoing: [...requests.values()].filter((r) => r.from.id === me.id).map((r) => ({ id: r.id, email: r.toEmail, note: r.note, at: r.at, link: linkOf(r.code) })) }, blocked: [], relays: [], invites: linksOf(me.id), mailAllowance: allowanceOf(me.id) }));
        // The history: messages to this person, and their own as copies from another device.
        history.forEach((h, i) => {
          const message = { id: `h-${i + 1}`, from: h.from, to: h.to, text: h.text, at: h.at };
          if (h.to === me.id) conn.send(JSON.stringify({ type: 'messages', messages: [message] }));
          else if (h.from === me.id) conn.send(JSON.stringify({ type: 'msg.copy', message }));
        });
        tellContacts(me.id);
        return;
      }
      const m = msg;
      if (m.type === 'presence') { presence.set(me.id, m.presence); tellContacts(me.id); return; }
      if (m.type === 'contact.request') {
        if (String(m.email).toLowerCase() === me.email) { conn.send(JSON.stringify({ type: 'error', for: 'contact.request', code: 'yourself' })); return; }
        const them = Object.values(PEOPLE).find((p) => p.email === String(m.email).toLowerCase());
        const r = { id: `req-${requests.size + 1}`, from: me, toEmail: String(m.email).toLowerCase(), note: m.note || '', at: Date.now(), code: code() };
        requests.set(r.id, r);
        invites.set(r.code, { code: r.code, kind: 'request', from: me, note: r.note, requestId: r.id, uses: 0, maxUses: 1, created: r.at, expires: r.at + 30 * 24 * 3600 * 1000 });
        let mail;
        if (m.notify === 'rutba') {
          const a = allowanceOf(me.id);
          const over = a.day.used >= a.day.limit || a.week.used >= a.week.limit;
          if (!over) mailed.push({ by: me.id, to: r.toEmail, link: linkOf(r.code), at: Date.now() });
          mail = { by: 'rutba', sent: !over, ...(over ? { reason: 'limit' } : {}), allowance: allowanceOf(me.id) };
        }
        conn.send(JSON.stringify({ type: 'contact.request.sent', ref: m.ref, request: { id: r.id, email: r.toEmail, note: r.note, at: r.at, link: linkOf(r.code) }, ...(mail ? { mail } : {}) }));
        if (them) to(them.id, { type: 'contact.requested', request: { id: r.id, from: me, note: r.note, at: r.at } });
        return;
      }
      if (m.type === 'contact.accept') {
        const r = requests.get(m.requestId);
        if (!r || r.toEmail !== me.email) return;
        requests.delete(r.id);
        contacts.add(pair(me.id, r.from.id));
        to(me.id, { type: 'contact.added', contact: view(r.from), requestId: r.id });
        to(r.from.id, { type: 'contact.added', contact: view(me), requestId: r.id });
        return;
      }
      if (m.type === 'mail.allowance') { conn.send(JSON.stringify({ type: 'mail.allowance', ref: m.ref, allowance: allowanceOf(me.id) })); return; }
      if (m.type === 'invite.create') {
        const at = Date.now();
        const i = { code: code(), kind: 'link', from: me, note: String(m.note || ''), uses: 0, maxUses: 20, created: at, expires: at + 30 * 24 * 3600 * 1000 };
        invites.set(i.code, i);
        conn.send(JSON.stringify({ type: 'invite.created', ref: m.ref, invite: inviteView(i) }));
        return;
      }
      if (m.type === 'invite.list') { conn.send(JSON.stringify({ type: 'invites', ref: m.ref, invites: linksOf(me.id) })); return; }
      if (m.type === 'invite.revoke') {
        const i = invites.get(m.code);
        if (i?.kind === 'link' && i.from.id === me.id) invites.delete(m.code);
        conn.send(JSON.stringify({ type: 'invite.revoked', ref: m.ref, code: m.code }));
        return;
      }
      if (m.type === 'invite.peek' || m.type === 'invite.accept') {
        const i = invites.get(m.code);
        if (!i || i.uses >= i.maxUses || i.expires < Date.now()) { conn.send(JSON.stringify({ type: 'error', for: m.type, code: 'no-invite' })); return; }
        if (i.from.id === me.id) { conn.send(JSON.stringify({ type: 'error', for: m.type, code: 'yourself' })); return; }
        if (m.type === 'invite.peek') { conn.send(JSON.stringify({ type: 'invite.peeked', ref: m.ref, code: i.code, from: i.from, note: i.note, kind: i.kind })); return; }
        i.uses += 1;
        if (i.kind === 'request') { requests.delete(i.requestId); invites.delete(i.code); }
        contacts.add(pair(me.id, i.from.id));
        to(me.id, { type: 'contact.added', contact: view(i.from), requestId: i.requestId || null });
        to(i.from.id, { type: 'contact.added', contact: view(me), requestId: i.requestId || null });
        return;
      }
      if (m.type === 'msg.send') {
        if (!isContact(me.id, m.to)) return;
        const message = { id: m.id, from: me.id, to: m.to, text: m.text, at: Date.now() };
        to(m.to, { type: 'messages', messages: [message] });
        conn.send(JSON.stringify({ type: 'msg.sent', id: m.id, at: message.at, delivered: online.has(m.to) }));
        return;
      }
      if (m.type === 'msg.ack') return;
      if (m.type === 'call.offer') {
        if (!isContact(me.id, m.to)) { conn.send(JSON.stringify({ type: 'error', for: 'call.offer', code: 'not-a-contact', callId: m.callId })); return; }
        calls.set(m.callId, { from: me.id, to: m.to });
        to(m.to, { type: 'call.incoming', callId: m.callId, from: me, sdp: m.sdp, video: Boolean(m.video), relay: m.relay || null });
        conn.send(JSON.stringify({ type: 'call.ringing', callId: m.callId }));
        return;
      }
      const call = calls.get(m.callId);
      if (!call) return;
      const other = call.from === me.id ? call.to : call.from;
      if (m.type === 'call.answer') to(other, { type: 'call.answered', callId: m.callId, sdp: m.sdp });
      else if (m.type === 'call.candidate') to(other, { type: 'call.candidate', callId: m.callId, candidate: m.candidate });
      else if (m.type === 'call.decline') { calls.delete(m.callId); to(other, { type: 'call.declined', callId: m.callId, reason: m.reason || 'declined' }); }
      else if (m.type === 'call.end') { calls.delete(m.callId); to(other, { type: 'call.ended', callId: m.callId, reason: m.reason || 'hung-up' }); }
      else if (m.type === 'call.update') to(other, { type: 'call.update', callId: m.callId, kind: m.kind, sdp: m.sdp });
    };
    conn.onclose = () => { if (me && online.get(me.id) === conn) { online.delete(me.id); tellContacts(me.id); } };
  });
  return {
    options,
    listen: () => new Promise((resolve) => server.listen(0, '127.0.0.1', () => { base = `http://127.0.0.1:${server.address().port}`; resolve(server.address().port); })),
    /** The invitation emails Rutba sent. */
    mailed: () => mailed.slice(),
    /** One person's connection ended from this side, as a server restart would. */
    drop: (id) => online.get(id)?.socket?.end(),
    online: (id) => online.has(id),
    // Its WebSockets are cut too: an upgraded socket is not the server's to wait for, and Talk keeps one open.
    close: () => new Promise((resolve) => { for (const c of online.values()) { try { c.onclose = null; } catch { /* gone */ } } for (const s of sockets) s.destroy(); server.closeAllConnections?.(); server.close(() => resolve()); }),
  };
}

/* ── Ben: plain web code in a hidden window ───────────────────────────── */

const BEN = (url) => `<!doctype html><meta charset="utf-8"><body><script>
const log = []; window.ben = { log, code: '', remoteWidth: 0, state: 'new' };
const ws = new WebSocket(${JSON.stringify(url)});
let pc = null; let call = null; const pending = [];
const send = (m) => ws.send(JSON.stringify(m));
async function code() {
  const p = [pc.localDescription.sdp, pc.remoteDescription.sdp].map((s) => (/a=fingerprint:\\S+ ([0-9A-F:]+)/i.exec(s)?.[1] || '').toUpperCase()).sort().join('|');
  const h = new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(p)));
  let n = 0n; for (let i = 0; i < 8; i++) n = (n << 8n) | BigInt(h[i]);
  const d = (n % 1000000000000n).toString().padStart(12, '0');
  return d.slice(0, 4) + ' ' + d.slice(4, 8) + ' ' + d.slice(8);
}
ws.onopen = () => send({ type: 'hello', token: 'ben' });
ws.onmessage = async (e) => {
  const m = JSON.parse(e.data); log.push(m);
  if (m.type === 'call.incoming') {
    call = m.callId;
    const media = await navigator.mediaDevices.getUserMedia({ audio: true, video: true });
    pc = new RTCPeerConnection({ iceServers: [] });
    pc.onicecandidate = (ev) => send({ type: 'call.candidate', callId: call, candidate: ev.candidate ? ev.candidate.toJSON() : null });
    pc.ontrack = (ev) => { const v = document.createElement('video'); v.muted = true; v.autoplay = true; v.srcObject = ev.streams[0] || new MediaStream([ev.track]); document.body.appendChild(v); v.onresize = () => { window.ben.remoteWidth = Math.max(window.ben.remoteWidth, v.videoWidth); }; };
    pc.onconnectionstatechange = async () => { window.ben.state = pc.connectionState; if (pc.connectionState === 'connected') window.ben.code = await code(); };
    await pc.setRemoteDescription({ type: 'offer', sdp: m.sdp });
    for (const t of pc.getTransceivers()) { const tr = media.getTracks().find((x) => x.kind === t.receiver.track.kind); if (tr) { await t.sender.replaceTrack(tr); t.direction = 'sendrecv'; } }
    for (const c of pending.splice(0)) await pc.addIceCandidate(c).catch(() => {});
    await pc.setLocalDescription(await pc.createAnswer());
    send({ type: 'call.answer', callId: call, sdp: pc.localDescription.sdp });
  } else if (m.type === 'call.candidate') {
    if (pc && pc.remoteDescription) await pc.addIceCandidate(m.candidate).catch(() => {}); else pending.push(m.candidate);
  } else if (m.type === 'call.ended') {
    window.ben.state = 'ended'; try { pc.close(); } catch {}
  }
};
window.ben.send = send;
</script></body>`;

export async function verifyTalk(h) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const intro = standInIntroducer({ page: BEN });
  let ben = null;
  let win = null;
  try {
    const port = await intro.listen();
    const url = `ws://127.0.0.1:${port}/ws`;
    win = await open('talk');
    const js = (code) => win.webContents.executeJavaScript(code);
    const set = (cls, value) => js(`(() => { const n = document.querySelector('.${cls}'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(value)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);

    // Signed out first; a test server and a token sign Amira in.
    const signedOut = await ok(until(() => js(`Boolean(document.querySelector('.tk-signin'))`), 'the sign-in', 8000));
    await js(`document.querySelector('.tk-test')?.setAttribute('open', ''), 1`);
    await set('tk-server', url);
    await set('tk-token', 'amira');
    await js(`document.querySelector('.tk-connect')?.click(), 1`);
    const signedIn = await ok(until(() => js(`/Signed in as Amira Khan/.test(document.querySelector('.tk-statusbar')?.textContent || '')`), 'signed in', 8000));
    check('talk: signed out it offers the sign-in, and a test server signs in with a token', signedOut && signedIn, JSON.stringify({ signedOut, signedIn }));

    // Ben, in a hidden window, asks to be in contact; Amira accepts.
    ben = new BrowserWindow({ show: false, webPreferences: { contextIsolation: true, nodeIntegration: false, backgroundThrottling: false, partition: 'talk-check-ben' } });
    await ben.loadURL(`http://127.0.0.1:${port}/ben`);
    const bjs = (code) => ben.webContents.executeJavaScript(code);
    await until(() => bjs(`window.ben.log.some((m) => m.type === 'welcome')`), 'Ben signed in', 6000);
    await bjs(`window.ben.send({ type: 'contact.request', email: 'amira@example.com', note: 'From the board meeting' }), 1`);
    const asked = await ok(until(() => js(`Boolean(document.querySelector('.tk-request .tk-accept'))`), 'the request', 5000));
    const note = await js(`document.querySelector('.tk-request')?.textContent || ''`);
    await js(`document.querySelector('.tk-request .tk-accept')?.click(), 1`);
    const added = await ok(until(() => js(`Boolean(document.querySelector('.tk-contact[data-contact="u-ben"]'))`), 'Ben among the contacts', 5000));
    const online = await js(`Boolean(document.querySelector('.tk-contact[data-contact="u-ben"] .tk-dot.online'))`);
    check('talk: a request shows who asked and their note; Accept makes them a contact, shown online',
      asked && /Ben Ortiz/.test(note) && /board meeting/.test(note) && added && online, JSON.stringify({ asked, note, added, online }));

    // A message each way.
    await js(`document.querySelector('.tk-contact[data-contact="u-ben"]')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.tk-input'))`), 'the conversation', 4000).catch(() => {});
    await js(`(() => { const n = document.querySelector('.tk-input'); Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(n, 'The board pack is ready.'); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await js(`document.querySelector('.tk-send')?.click(), 1`);
    const benGot = await ok(until(() => bjs(`window.ben.log.some((m) => m.type === 'messages' && m.messages[0]?.text === 'The board pack is ready.')`), 'the message at Ben', 5000));
    await bjs(`window.ben.send({ type: 'msg.send', to: 'u-amira', id: 'm-from-ben-1', text: 'Thanks, reading it now.' }), 1`);
    const amiraGot = await ok(until(() => js(`[...document.querySelectorAll('.tk-bubble:not(.mine) .tk-text')].some((n) => n.textContent === 'Thanks, reading it now.')`), 'the reply in the conversation', 5000));
    const mine = await js(`[...document.querySelectorAll('.tk-bubble.mine .tk-text')].map((n) => n.textContent)`);
    check('talk: a message sent reaches the contact, and theirs comes back into the conversation',
      benGot && amiraGot && mine.includes('The board pack is ready.'), JSON.stringify({ benGot, amiraGot, mine }));

    // A video call: Ben answers, the picture comes through, the same safety code on both sides.
    await js(`document.querySelector('.tk-call-video')?.click(), 1`);
    const rang = await ok(until(() => bjs(`window.ben.log.some((m) => m.type === 'call.incoming' && m.video)`), 'the call at Ben', 8000));
    const connected = await ok(until(() => js(`document.querySelector('.tk-callview')?.dataset.state === 'connected'`), 'the call connected', 15000));
    const picture = await ok(until(() => js(`(document.querySelector('.tk-remote')?.videoWidth || 0) > 0`), 'Ben\'s picture', 8000));
    await until(() => bjs(`Boolean(window.ben.code)`), 'Ben\'s code', 5000).catch(() => {});
    const amiraCode = await js(`document.querySelector('.tk-code')?.dataset.code || ''`);
    const benCode = await bjs(`window.ben.code`);
    const benSees = await ok(until(() => bjs(`window.ben.remoteWidth > 0`), 'Amira\'s picture at Ben', 6000));
    check('talk: a video call reaches the contact, connects straight between the two, shows each one the other\'s picture, and gives both the same safety code',
      rang && connected && picture && benSees && /^\d{4} \d{4} \d{4}$/.test(amiraCode) && amiraCode === benCode, JSON.stringify({ rang, connected, picture, benSees, amiraCode, benCode }));

    // Mute, then hang up: Ben is told.
    await js(`document.querySelector('.tk-controls .tk-round')?.click(), 1`);
    const muted = await js(`document.querySelector('.tk-controls .tk-round')?.getAttribute('aria-pressed') === 'true'`);
    await js(`document.querySelector('.tk-round.end')?.click(), 1`);
    const benEnded = await ok(until(() => bjs(`window.ben.state === 'ended'`), 'Ben told the call ended', 5000));
    const back = await ok(until(() => js(`!document.querySelector('.tk-callview') && Boolean(document.querySelector('.tk-conversation'))`), 'back to the conversation', 5000));
    check('talk: mute shows as pressed, and hanging up ends the call for the other side and returns to the conversation', muted && benEnded && back, JSON.stringify({ muted, benEnded, back }));

    // Relays: one's own relay is kept and used for the next call.
    await js(`document.querySelector('.tk-relays-btn')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.tk-own-urls'))`), 'the relays', 4000).catch(() => {});
    await set('tk-own-urls', 'turn:relay.example.com:3478');
    await set('tk-own-user', 'amira');
    await set('tk-own-pass', 'relay-password');
    await js(`document.querySelector('.tk-own-save')?.click(), 1`);
    await wait(500);
    const setup = await js(`window.rutbaOffice.talk.callSetup()`);
    const relays = await js(`window.rutbaOffice.talk.relays()`);
    check('talk: one\'s own relay is kept, its password hidden, and offered for the next call',
      setup.offerRelay?.urls?.[0] === 'turn:relay.example.com:3478' && setup.offerRelay.credential === 'relay-password' && setup.iceServers.some((s) => s.urls.includes('turn:relay.example.com:3478')) && relays.own?.credential === '••••' && relays.choice === 'own',
      JSON.stringify({ offer: setup.offerRelay?.urls, own: relays.own?.urls, choice: relays.choice }));
    await js(`[...document.querySelectorAll('.rw-dialog button')].find((b) => b.textContent.trim() === 'Done')?.click(), 1`);

    const complaints = await errorsIn(win);
    check('talk: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    await js(`window.rutbaOffice.talk.signOut().then(() => 1)`).catch(() => {});
  } catch (err) {
    check('talk: the check ran to the end', false, err?.message || String(err));
  } finally {
    try { ben?.destroy(); } catch { /* gone */ }
    await intro.close().catch(() => {});
  }
}
