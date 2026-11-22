// The sign-ins a camera asks for: Basic and Digest (RFC 7617, RFC 7616, and
// the older RFC 2069 form many cameras still speak, with no qop).
//
// A camera answers 401 with one challenge or several; the strongest it offers
// is the one answered. Digest never sends the password. Basic does, and over
// rtsp:// that is in the clear on the local network, so it is answered only
// when the camera offers nothing else.

import crypto from 'node:crypto';

const hash = (name, text) => crypto.createHash(name).update(text).digest('hex');

/** `key=value, key="value"` as an object, keys in lower case. */
function params(text) {
  const out = {};
  const re = /([a-z0-9_-]+)\s*=\s*(?:"((?:[^"\\]|\\.)*)"|([^\s,]*))/gi;
  for (let m = re.exec(text); m; m = re.exec(text)) out[m[1].toLowerCase()] = m[2] !== undefined ? m[2].replace(/\\(.)/g, '$1') : m[3];
  return out;
}

/** Every WWW-Authenticate line a 401 carried, as challenges: { scheme: 'digest' | 'basic', realm, nonce, qop, algorithm, opaque, stale }. */
export function parseChallenges(lines) {
  const out = [];
  for (const line of [].concat(lines || [])) {
    const m = /^\s*(digest|basic)\b\s*(.*)$/i.exec(String(line));
    if (!m) continue;
    const p = params(m[2]);
    const scheme = m[1].toLowerCase();
    if (scheme === 'basic') out.push({ scheme, realm: p.realm || '' });
    else if (p.nonce) {
      const qops = String(p.qop || '').split(',').map((q) => q.trim().toLowerCase()).filter(Boolean);
      out.push({
        scheme,
        realm: p.realm || '',
        nonce: p.nonce,
        qop: qops.includes('auth') ? 'auth' : '',
        // A camera that will take only a kind of answer this reader does not give (auth-int).
        unsupportedQop: qops.length > 0 && !qops.includes('auth'),
        algorithm: String(p.algorithm || 'MD5').toUpperCase(),
        opaque: p.opaque,
        stale: /^true$/i.test(p.stale || ''),
      });
    }
  }
  return out;
}

const DIGESTS = { MD5: 'md5', 'MD5-SESS': 'md5', 'SHA-256': 'sha256', 'SHA-256-SESS': 'sha256' };

/** The challenge to answer: Digest with SHA-256 before Digest with MD5 before Basic; null when none can be answered. */
export function chooseChallenge(challenges) {
  const usable = (challenges || []).filter((c) => c.scheme === 'basic' || (DIGESTS[c.algorithm] && !c.unsupportedQop));
  const rank = (c) => (c.scheme === 'basic' ? 0 : c.algorithm.startsWith('SHA-256') ? 2 : 1);
  return [...usable].sort((a, b) => rank(b) - rank(a))[0] || null;
}

/**
 * The Authorization line for one request. `count` is how many requests this
 * nonce has answered, this one included; `cnonce` is the client's own, made
 * here unless handed in (the tests hand one in).
 */
export function authorization({ challenge, user = '', password = '', method, uri, count = 1, cnonce = crypto.randomBytes(8).toString('hex') }) {
  if (challenge.scheme === 'basic') return `Basic ${Buffer.from(`${user}:${password}`, 'utf8').toString('base64')}`;
  const name = DIGESTS[challenge.algorithm];
  const h = (text) => hash(name, text);
  let ha1 = h(`${user}:${challenge.realm}:${password}`);
  if (challenge.algorithm.endsWith('-SESS')) ha1 = h(`${ha1}:${challenge.nonce}:${cnonce}`);
  const ha2 = h(`${method}:${uri}`);
  const quoted = (s) => String(s).replace(/(["\\])/g, '\\$1');
  const parts = [`username="${quoted(user)}"`, `realm="${quoted(challenge.realm)}"`, `nonce="${quoted(challenge.nonce)}"`, `uri="${quoted(uri)}"`];
  if (challenge.qop) {
    const nc = count.toString(16).padStart(8, '0');
    parts.push(`response="${h(`${ha1}:${challenge.nonce}:${nc}:${cnonce}:auth:${ha2}`)}"`, 'qop=auth', `nc=${nc}`, `cnonce="${cnonce}"`);
  } else {
    parts.push(`response="${h(`${ha1}:${challenge.nonce}:${ha2}`)}"`);
  }
  if (challenge.algorithm !== 'MD5') parts.push(`algorithm=${challenge.algorithm}`);
  if (challenge.opaque !== undefined) parts.push(`opaque="${quoted(challenge.opaque)}"`);
  return `Digest ${parts.join(', ')}`;
}
