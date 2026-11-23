// A stand-in for a Talk Server's door, for the tests of Office's side
// (talk-door.test.js, talk-servers.test.js): a certificate it makes itself,
// as a server with no name does, `/admit`, `/control`, and a WebSocket under
// `/rtc` that says back what it is sent. Nothing here is the server's own
// code; it answers only what the contract says a door answers.

import crypto from 'node:crypto';
import http from 'node:http';
import https from 'node:https';
import { OP, acceptKey, closePayload, encodeFrame, frameReader } from '../apps/desktop/main/talk-frames.js';

// ── a certificate signed by its own key, written out by hand ─────────────────
const len = (n) => (n < 128 ? Buffer.from([n]) : n < 256 ? Buffer.from([0x81, n]) : Buffer.from([0x82, n >> 8, n & 255]));
const tlv = (tag, ...parts) => { const body = Buffer.concat(parts); return Buffer.concat([Buffer.from([tag]), len(body.length), body]); };
const seq = (...parts) => tlv(0x30, ...parts);
const oid = (dotted) => {
  const a = dotted.split('.').map(Number);
  const out = [a[0] * 40 + a[1]];
  for (const v of a.slice(2)) {
    const b = [v & 0x7f];
    for (let x = v >> 7; x; x >>= 7) b.unshift((x & 0x7f) | 0x80);
    out.push(...b);
  }
  return tlv(0x06, Buffer.from(out));
};
const int = (buf) => tlv(0x02, buf[0] & 0x80 ? Buffer.concat([Buffer.from([0]), buf]) : buf);
const time = (d) => tlv(0x17, Buffer.from(`${d.toISOString().replace(/[-:T]/g, '').slice(2, 14)}Z`));
const ipBytes = (ip) => {
  if (ip.includes('.')) return Buffer.from(ip.split('.').map(Number));
  const [head, tail = ''] = ip.split('::');
  const h = head ? head.split(':') : [];
  const t = tail ? tail.split(':') : [];
  const groups = ip.includes('::') ? [...h, ...Array(8 - h.length - t.length).fill('0'), ...t] : h;
  return Buffer.concat(groups.map((g) => { const b = Buffer.alloc(2); b.writeUInt16BE(parseInt(g, 16)); return b; }));
};
const pem = (label, der) => `-----BEGIN ${label}-----\n${der.toString('base64').match(/.{1,64}/g).join('\n')}\n-----END ${label}-----\n`;

/**
 * A certificate for these names and addresses, signed by its own key.
 * { key, cert } in PEM, and `pin`: the SHA-256 of the certificate, in hex.
 */
export function selfSigned({ names = [], addresses = ['127.0.0.1'], days = 2 } = {}) {
  const { publicKey, privateKey } = crypto.generateKeyPairSync('ec', { namedCurve: 'prime256v1' });
  const name = seq(tlv(0x31, seq(oid('2.5.4.3'), tlv(0x0c, Buffer.from(names[0] || 'talk-server stand-in')))));
  const algorithm = seq(oid('1.2.840.10045.4.3.2'));
  const alt = seq(...names.map((n) => tlv(0x82, Buffer.from(n))), ...addresses.map((a) => tlv(0x87, ipBytes(a))));
  const from = new Date(Date.now() - 3600 * 1000);
  const to = new Date(Date.now() + days * 86400 * 1000);
  const tbs = seq(
    tlv(0xa0, int(Buffer.from([2]))),
    int(crypto.randomBytes(8)),
    algorithm,
    name,
    seq(time(from), time(to)),
    name,
    publicKey.export({ type: 'spki', format: 'der' }),
    tlv(0xa3, seq(seq(oid('2.5.29.17'), tlv(0x04, alt)))),
  );
  const der = seq(tbs, algorithm, tlv(0x03, Buffer.from([0]), crypto.sign('sha256', tbs, privateKey)));
  return { key: privateKey.export({ type: 'pkcs8', format: 'pem' }), cert: pem('CERTIFICATE', der), pin: crypto.createHash('sha256').update(der).digest('hex') };
}

/**
 * A door on this computer.
 * @param {object} [o]
 * @param {{ key: string, cert: string }|null} [o.tls] with none it speaks in the clear, as a server tried on one's own computer may
 * @param {(body: object) => { status: number, body: object }} [o.admit]
 * @param {(body: object) => { status: number, body: object }} [o.control]
 * @param {'take'|'refuse'} [o.rtc]
 * @returns {Promise<{ port: number, url: string, hits: object[], sockets: object[], close: () => Promise<void> }>}
 */
export async function standInDoor({ tls = null, admit = () => ({ status: 200, body: { token: 'join-token', room: 'room-1', identity: 'u-a', role: 'member', e2ee: true, live: false } }), control = () => ({ status: 200, body: { people: [] } }), rtc = 'take', host = '127.0.0.1' } = {}) {
  /** What was asked: { method, path, body }. */
  const hits = [];
  /** The WebSockets taken up: { target, protocols, send(data, binary), close(code), got: [] }. */
  const taken = [];
  const live = new Set();
  const handler = (req, res) => {
    const chunks = [];
    req.on('data', (c) => chunks.push(c));
    req.on('end', () => {
      let body = null;
      try { body = JSON.parse(Buffer.concat(chunks).toString('utf8') || 'null'); } catch { body = null; }
      hits.push({ method: req.method, path: req.url, body, host: req.headers.host });
      const answer = (a) => { res.writeHead(a.status, { 'content-type': 'application/json', ...(a.headers || {}) }); res.end(JSON.stringify(a.body ?? {})); };
      if (req.method === 'POST' && req.url === '/admit') return answer(admit(body || {}));
      if (req.method === 'POST' && req.url === '/control') return answer(control(body || {}));
      return answer({ status: 404, body: { error: 'not-found' } });
    });
  };
  const server = tls ? https.createServer({ key: tls.key, cert: tls.cert }, handler) : http.createServer(handler);
  server.on('connection', (s) => { live.add(s); s.on('close', () => live.delete(s)); });
  server.on('secureConnection', (s) => { live.add(s); s.on('close', () => live.delete(s)); });
  server.on('upgrade', (req, socket) => {
    hits.push({ method: 'GET', path: req.url, upgrade: true });
    if (rtc === 'refuse' || !req.url.startsWith('/rtc')) {
      socket.end('HTTP/1.1 401 Unauthorized\r\ncontent-type: text/plain\r\ncontent-length: 13\r\nconnection: close\r\n\r\ninvalid token');
      return;
    }
    const protocols = String(req.headers['sec-websocket-protocol'] || '').split(',').map((p) => p.trim()).filter(Boolean);
    socket.write(`HTTP/1.1 101 Switching Protocols\r\nupgrade: websocket\r\nconnection: Upgrade\r\nsec-websocket-accept: ${acceptKey(req.headers['sec-websocket-key'])}\r\n${protocols.length ? `sec-websocket-protocol: ${protocols[0]}\r\n` : ''}\r\n`);
    const one = {
      target: req.url,
      protocols,
      got: [],
      closed: null,
      send: (data, binary = true) => socket.write(encodeFrame(binary ? OP.binary : OP.text, data, { mask: false })),
      ping: (data = Buffer.from('p')) => socket.write(encodeFrame(OP.ping, data, { mask: false })),
      close: (code = 1000, reason = '') => { socket.write(encodeFrame(OP.close, closePayload(code, reason), { mask: false })); },
      pongs: [],
    };
    const reader = frameReader({
      // Said back as it came, so both directions are seen.
      onMessage: (data, binary) => { one.got.push({ data, binary }); one.send(data, binary); },
      onPong: (data) => one.pongs.push(data),
      onClose: (code, reason) => { one.closed = { code, reason }; socket.end(); },
    });
    socket.on('data', (c) => reader.push(c));
    socket.on('error', () => {});
    taken.push(one);
  });
  await new Promise((resolve) => server.listen(0, host, resolve));
  const { port } = server.address();
  return {
    port,
    url: `${tls ? 'wss' : 'ws'}://${host.includes(':') ? `[${host}]` : host}:${port}`,
    hits,
    sockets: taken,
    close: () => new Promise((resolve) => { for (const s of live) s.destroy(); server.close(() => resolve()); }),
  };
}
