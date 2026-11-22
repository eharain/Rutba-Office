// Finding cameras on the local network and asking one for its stream address,
// as ONVIF lays it down: WS-Discovery to find them, then SOAP over HTTP to the
// camera for its profiles and the RTSP address of one.
//
// Finding: a probe is sent to the discovery group (239.255.255.250, port 3702)
// from a socket of Office's own, and the answers to that probe are read for a
// few seconds on the same socket, which is then closed. Office joins no group
// and waits for no announcement: it hears only what answers its own question.
//
// Office then speaks only to the device that answered, at the address the
// answer came from: an answer naming a service anywhere else is not followed,
// since any device on the network can answer a probe and a sign-in must not be
// sent where it points. For the same reason only an answer from a private
// address of the local network is heard, and the stream address a camera gives
// is read as being on the camera itself.
//
// Sign-in is the WS-Security user name token every ONVIF camera takes (a
// digest of the password with a nonce and the time, never the password), and
// HTTP Digest where a camera asks for that instead. Cameras with a wrong clock
// are common and refuse a token stamped with another time, so the camera is
// asked its time first and the token is stamped with that.

import crypto from 'node:crypto';
import dgramModule from 'node:dgram';
import http from 'node:http';
import https from 'node:https';
import os from 'node:os';
import { authorization, chooseChallenge, parseChallenges } from './auth.js';
import { parseCameraAddress } from './address.js';
import { CameraError } from './rtsp.js';
import { child, escapeXml, find, findAll, parseXml, textOf } from './xml.js';

/** Where a probe goes. */
export const DISCOVERY = Object.freeze({ address: '239.255.255.250', port: 3702 });

const MAX_ANSWER = 1024 * 1024;
const MAX_FOUND = 64;
/** The largest picture chosen without being asked: full HD. */
const FULL_HD = 1920 * 1080;

const SOAP = 'http://www.w3.org/2003/05/soap-envelope';

/** The probe that asks every network video device to answer. */
export function probeMessage(messageId) {
  return `<?xml version="1.0" encoding="UTF-8"?>`
    + `<e:Envelope xmlns:e="${SOAP}" xmlns:w="http://schemas.xmlsoap.org/ws/2004/08/addressing" xmlns:d="http://schemas.xmlsoap.org/ws/2005/04/discovery" xmlns:dn="http://www.onvif.org/ver10/network/wsdl">`
    + `<e:Header><w:MessageID>${messageId}</w:MessageID><w:To e:mustUnderstand="true">urn:schemas-xmlsoap-org:ws:2005:04:discovery</w:To>`
    + `<w:Action e:mustUnderstand="true">http://schemas.xmlsoap.org/ws/2005/04/discovery/Probe</w:Action></e:Header>`
    + `<e:Body><d:Probe><d:Types>dn:NetworkVideoTransmitter</d:Types></d:Probe></e:Body></e:Envelope>`;
}

const tidy = (text, max = 80) => String(text || '').replace(/[\u0000-\u001f\u007f]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, max);

function scopeValue(scopes, kind) {
  const head = `onvif://www.onvif.org/${kind}/`;
  const got = scopes.find((s) => s.toLowerCase().startsWith(head));
  if (!got) return '';
  const raw = got.slice(head.length);
  try {
    return tidy(decodeURIComponent(raw));
  } catch {
    return tidy(raw);
  }
}

/** Whether an address is one of the local network's own: private or link-local IPv4. Loopback only when asked. */
export function isLocalAddress(address, { loopback = false } = {}) {
  const m = /^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/.exec(String(address || ''));
  if (!m) return false;
  const [a, b] = [Number(m[1]), Number(m[2])];
  if (m.slice(1).some((n) => Number(n) > 255)) return false;
  if (a === 127) return loopback;
  return a === 10 || (a === 172 && b >= 16 && b <= 31) || (a === 192 && b === 168) || (a === 169 && b === 254);
}

/**
 * The cameras one answer to a probe names: `{ key, name, hardware, host,
 * service }`. Only an answer to this probe counts (`messageId`), and of the
 * services it names only one at the address the answer came from (`from`).
 */
export function parseProbeMatches(xml, { from, messageId = null }) {
  const doc = parseXml(xml);
  if (!doc || doc.name !== 'Envelope') return [];
  if (messageId && textOf(child(doc, 'Header'), 'RelatesTo') !== messageId) return [];
  const out = [];
  for (const match of findAll(child(doc, 'Body'), 'ProbeMatch')) {
    const types = textOf(match, 'Types');
    if (types && !/NetworkVideoTransmitter/i.test(types)) continue;
    const scopes = textOf(match, 'Scopes').split(/\s+/).filter(Boolean);
    let service = '';
    for (const address of textOf(match, 'XAddrs').split(/\s+/)) {
      let url;
      try {
        url = new URL(address);
      } catch {
        continue;
      }
      if ((url.protocol !== 'http:' && url.protocol !== 'https:') || url.hostname !== from || url.username || url.password) continue;
      // The plain service before the secure one: a camera's own certificate cannot be checked.
      if (!service || (service.startsWith('https:') && url.protocol === 'http:')) service = url.href;
    }
    if (!service) continue;
    const urn = tidy(textOf(child(match, 'EndpointReference'), 'Address'), 200);
    out.push({ key: urn || service, name: scopeValue(scopes, 'name'), hardware: scopeValue(scopes, 'hardware'), host: from, service });
  }
  return out;
}

/** This computer's own addresses on its networks. */
export function localAddresses() {
  const out = [];
  for (const list of Object.values(os.networkInterfaces())) {
    for (const a of list || []) if (a.family === 'IPv4' && !a.internal) out.push(a.address);
  }
  return out;
}

/**
 * Asks the local network for cameras and gives those that answered within
 * `ms`. The probe goes out of every network this computer is on, twice, since
 * a datagram can be lost. Nothing is thrown: a computer on no network finds
 * none.
 * @param {object} [o]
 * @param {number} [o.ms]  how long answers are waited for
 * @param {Array<{ address: string, port: number }>} [o.targets]  where the probe goes (the tests send it to their own camera)
 * @param {string[]} [o.interfaces]  the local addresses to send from
 * @param {boolean} [o.loopback]  hear an answer from this computer itself (the tests)
 * @param {object} [o.dgram]  node:dgram, or a stand-in
 * @returns {Promise<Array<{ key: string, name: string, hardware: string, host: string, service: string }>>}
 */
export function discover({ ms = 3000, targets = [DISCOVERY], interfaces = localAddresses(), loopback = false, dgram = dgramModule } = {}) {
  return new Promise((resolve) => {
    const messageId = `uuid:${crypto.randomUUID()}`;
    const probe = Buffer.from(probeMessage(messageId), 'utf8');
    const found = new Map();
    const sockets = [];
    const timers = [];
    let done = false;
    const finish = () => {
      if (done) return;
      done = true;
      for (const t of timers) clearTimeout(t);
      for (const s of sockets) {
        try {
          s.close();
        } catch {
          // Closed already.
        }
      }
      resolve([...found.values()]);
    };
    for (const address of interfaces) {
      let socket;
      try {
        socket = dgram.createSocket({ type: 'udp4' });
      } catch {
        continue;
      }
      sockets.push(socket);
      socket.on('error', () => {});
      socket.on('message', (bytes, rinfo) => {
        if (done || bytes.length > 64 * 1024 || !isLocalAddress(rinfo.address, { loopback })) return;
        for (const cam of parseProbeMatches(bytes.toString('utf8'), { from: rinfo.address, messageId })) {
          if (found.size < MAX_FOUND && !found.has(cam.key)) found.set(cam.key, cam);
        }
      });
      const send = () => {
        if (done) return;
        for (const t of targets) socket.send(probe, t.port, t.address, () => {});
      };
      try {
        socket.bind(0, address, () => {
          send();
          timers.push(setTimeout(send, Math.min(400, ms / 3)));
        });
      } catch {
        // This network is left out.
      }
    }
    timers.push(setTimeout(finish, sockets.length ? ms : 0));
  });
}

// ---- the camera's own service ----

/** The WS-Security header that signs a request in: a digest of the password with a nonce and the time. */
export function usernameToken({ user, password, created = new Date(), nonce = crypto.randomBytes(16) }) {
  const stamp = created.toISOString().replace(/\.\d{3}Z$/, 'Z');
  const digest = crypto.createHash('sha1').update(Buffer.concat([nonce, Buffer.from(stamp, 'utf8'), Buffer.from(String(password), 'utf8')])).digest('base64');
  return '<wsse:Security e:mustUnderstand="true" xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd" xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">'
    + `<wsse:UsernameToken><wsse:Username>${escapeXml(user)}</wsse:Username>`
    + `<wsse:Password Type="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordDigest">${digest}</wsse:Password>`
    + `<wsse:Nonce EncodingType="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-soap-message-security-1.0#Base64Binary">${nonce.toString('base64')}</wsse:Nonce>`
    + `<wsu:Created>${stamp}</wsu:Created></wsse:UsernameToken></wsse:Security>`;
}

const envelope = (body, header = '') => `<?xml version="1.0" encoding="UTF-8"?><e:Envelope xmlns:e="${SOAP}" xmlns:tds="http://www.onvif.org/ver10/device/wsdl" xmlns:trt="http://www.onvif.org/ver10/media/wsdl" xmlns:tt="http://www.onvif.org/ver10/schema">`
  + `<e:Header>${header}</e:Header><e:Body>${body}</e:Body></e:Envelope>`;

/**
 * One HTTP POST to a camera. Nothing is followed anywhere else, the answer is
 * bounded, and a camera that says nothing is given up on.
 * @returns {Promise<{ status: number, headers: object, text: string }>}
 */
export function httpPost({ url, body, headers = {}, timeoutMs = 5000 }) {
  return new Promise((resolve, reject) => {
    let target;
    try {
      target = new URL(url);
    } catch {
      reject(new CameraError('address', 'That is not a camera service address.'));
      return;
    }
    const data = Buffer.from(body, 'utf8');
    const lib = target.protocol === 'https:' ? https : http;
    let settled = false;
    const fail = (code, message) => {
      if (settled) return;
      settled = true;
      reject(new CameraError(code, message));
    };
    const req = lib.request({
      protocol: target.protocol,
      hostname: target.hostname,
      port: target.port || (target.protocol === 'https:' ? 443 : 80),
      path: `${target.pathname}${target.search}`,
      method: 'POST',
      agent: false,
      headers: { 'Content-Type': 'application/soap+xml; charset=utf-8', 'Content-Length': data.length, 'User-Agent': 'Rutba Office', Connection: 'close', ...headers },
    }, (res) => {
      const parts = [];
      let size = 0;
      res.on('data', (chunk) => {
        size += chunk.length;
        if (size > MAX_ANSWER) {
          req.destroy();
          fail('protocol', 'The camera answered with more than can be read.');
          return;
        }
        parts.push(chunk);
      });
      res.on('end', () => {
        if (settled) return;
        settled = true;
        resolve({ status: res.statusCode || 0, headers: res.headers, text: Buffer.concat(parts).toString('utf8') });
      });
      res.on('error', () => fail('closed', 'The camera closed the connection.'));
    });
    req.setTimeout(timeoutMs, () => {
      req.destroy();
      fail('timeout', 'The camera did not answer in time.');
    });
    req.on('error', (e) => {
      const code = /CERT|SELF_SIGNED|UNABLE_TO_VERIFY/i.test(String(e?.code)) ? 'certificate' : 'unreachable';
      fail(code, code === 'certificate' ? 'The camera certificate could not be checked.' : 'The camera could not be reached.');
    });
    req.end(data);
  });
}

/** The same address on the camera itself: whatever host an answer names, the one that answered the probe is spoken to. */
function onDevice(address, device, { rtsp = false } = {}) {
  if (rtsp) {
    const cam = parseCameraAddress(address);
    if (!cam) return null;
    const usual = cam.secure ? 322 : 554;
    return `${cam.secure ? 'rtsps' : 'rtsp'}://${device.hostname}${cam.port === usual ? '' : `:${cam.port}`}${cam.path}`;
  }
  let url;
  try {
    url = new URL(address);
  } catch {
    return null;
  }
  if (url.protocol !== 'http:' && url.protocol !== 'https:') return null;
  url.hostname = device.hostname;
  url.username = '';
  url.password = '';
  return url.href;
}

const refusedSignIn = (answer) => answer.status === 401 || /NotAuthorized|FailedAuthentication|not authorized|Unauthorized/i.test(answer.text.slice(0, 20000));

/**
 * What a found camera offers and the address of the stream Office reads.
 * @param {object} o
 * @param {string} o.service  the camera's device service, as its answer to the probe gave it
 * @param {string} [o.user]
 * @param {string} [o.password]
 * @param {Function} [o.post]  httpPost, or a stand-in
 * @param {() => number} [o.now]
 * @returns {Promise<{ url: string, profile: object, profiles: object[] }>} `url` is rtsp:// with no sign-in in it
 * @throws {CameraError} sign-in-needed, sign-in-refused, unreachable, timeout, certificate, protocol, no-video, h265
 */
export async function readCameraService({ service, user = '', password = '', post = httpPost, now = Date.now, timeoutMs = 5000 }) {
  let device;
  try {
    device = new URL(service);
  } catch {
    throw new CameraError('address', 'That is not a camera service address.');
  }
  let offset = 0;
  let challenge = null;
  let count = 0;

  /** One question to the camera, signed in when there is a sign-in; its body's first answer element. */
  const ask = async (url, body, { signed = true } = {}) => {
    const header = signed && user ? usernameToken({ user, password, created: new Date(now() + offset) }) : '';
    const text = envelope(body, header);
    const path = (() => { const u = new URL(url); return `${u.pathname}${u.search}`; })();
    const send = () => {
      const headers = {};
      if (challenge && user) {
        count += 1;
        headers.Authorization = authorization({ challenge, user, password, method: 'POST', uri: path, count });
      }
      return post({ url, body: text, headers, timeoutMs });
    };
    let answer = await send();
    if (answer.status === 401 && signed && user && !challenge) {
      // A camera that wants HTTP Digest: answered once, and kept for the questions after.
      const lines = [].concat(answer.headers?.['www-authenticate'] || []);
      challenge = chooseChallenge(parseChallenges(lines));
      // The password itself is never sent to a camera over a plain connection it could be read from.
      if (challenge?.scheme === 'basic' && new URL(url).protocol !== 'https:') challenge = null;
      if (challenge) answer = await send();
    }
    if (refusedSignIn(answer)) {
      throw new CameraError(user ? 'sign-in-refused' : 'sign-in-needed', user ? 'The camera did not accept the user name and password.' : 'The camera asks for a user name and password.');
    }
    const doc = parseXml(answer.text);
    const inside = doc && doc.name === 'Envelope' ? child(doc, 'Body') : null;
    if (!inside) throw new CameraError('protocol', 'What answered is not a camera service.');
    if (find(inside, 'Fault')) return { fault: inside };
    return { body: inside };
  };

  // The camera's time, which needs no sign-in, so a token is stamped as the camera expects.
  try {
    const { body } = await ask(device.href, '<tds:GetSystemDateAndTime/>', { signed: false });
    const utc = find(body, 'UTCDateTime');
    const n = (name) => Number(textOf(utc, name));
    const at = utc ? Date.UTC(n('Year'), n('Month') - 1, n('Day'), n('Hour'), n('Minute'), n('Second')) : NaN;
    if (Number.isFinite(at) && n('Year') > 2000) offset = at - now();
  } catch (e) {
    if (e instanceof CameraError && ['unreachable', 'timeout', 'certificate'].includes(e.code)) throw e;
    // A camera that will not tell its time is asked the rest with this computer's.
  }

  // Where its media service is.
  let media = '';
  const caps = await ask(device.href, '<tds:GetCapabilities><tds:Category>Media</tds:Category></tds:GetCapabilities>');
  if (caps.body) media = textOf(find(caps.body, 'Media'), 'XAddr');
  if (!media) {
    const services = await ask(device.href, '<tds:GetServices><tds:IncludeCapability>false</tds:IncludeCapability></tds:GetServices>');
    for (const s of findAll(services.body, 'Service')) {
      if (textOf(s, 'Namespace') === 'http://www.onvif.org/ver10/media/wsdl') media = textOf(s, 'XAddr');
    }
  }
  const mediaUrl = media ? onDevice(media, device) : null;
  if (!mediaUrl) throw new CameraError('protocol', 'The camera did not say where its pictures are asked for.');

  const got = await ask(mediaUrl, '<trt:GetProfiles/>');
  if (!got.body) throw new CameraError('protocol', 'The camera did not list what it sends.');
  const profiles = findAll(got.body, 'Profiles').map((p) => {
    const video = child(p, 'VideoEncoderConfiguration');
    return {
      token: p.attrs.token || '',
      name: tidy(textOf(p, 'Name')),
      encoding: video ? textOf(video, 'Encoding').toUpperCase() : '',
      width: Number(textOf(child(video, 'Resolution'), 'Width')) || 0,
      height: Number(textOf(child(video, 'Resolution'), 'Height')) || 0,
    };
  }).filter((p) => p.token && p.token.length <= 200);
  const profile = chooseProfile(profiles);
  if (!profile) {
    const other = profiles.some((p) => p.encoding === 'H265' || p.encoding === 'HEVC');
    throw new CameraError(other ? 'h265' : 'no-video', other ? 'The camera sends H.265 only.' : 'The camera sends no H.264 picture.');
  }

  const uri = await ask(mediaUrl, '<trt:GetStreamUri><trt:StreamSetup><tt:Stream>RTP-Unicast</tt:Stream><tt:Transport><tt:Protocol>RTSP</tt:Protocol></tt:Transport></trt:StreamSetup>'
    + `<trt:ProfileToken>${escapeXml(profile.token)}</trt:ProfileToken></trt:GetStreamUri>`);
  const url = uri.body ? onDevice(textOf(uri.body, 'Uri'), device, { rtsp: true }) : null;
  if (!url) throw new CameraError('protocol', 'The camera did not give a stream address.');
  return { url, profile, profiles };
}

/**
 * The stream Office reads of those a camera offers: H.264, the largest picture
 * up to full HD; when every one is larger, the smallest. Null when none is
 * H.264.
 */
export function chooseProfile(profiles) {
  const usable = (profiles || []).filter((p) => p.encoding === 'H264');
  if (!usable.length) return null;
  const area = (p) => p.width * p.height;
  const within = usable.filter((p) => area(p) <= FULL_HD);
  if (within.length) return within.reduce((best, p) => (area(p) > area(best) ? p : best));
  return usable.reduce((best, p) => (area(p) < area(best) ? p : best));
}
