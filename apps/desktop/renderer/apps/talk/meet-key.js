// The key of a meeting held on a Talk Server. Sound, picture and what is
// written in such a meeting are encrypted by the apps of the people in it;
// the server only passes them on. The key is made on the computer of the
// person who starts the meeting and reaches each other person sealed for
// them alone, inside the meeting's set-up: it is never given to the
// introducer's `/servers/`, to the server, or put in a pass.
//
// Sealed means: the one who asks makes a pair of keys for this one asking and
// says the public half; the one who has the meeting's key makes a pair of
// their own, and the two agree on a secret (ECDH, P-256) from which the key
// that locks the box is drawn (HKDF, SHA-256; AES-256-GCM). What the box is
// for (which meeting, from whom, to whom, with which two public halves) is
// bound into it, so a box made for one person or one meeting opens for no
// other.
//
// What this does not do: prove whose public half it is. The set-up passes
// through the introducer, as the set-up of a call does, and the two people
// are known to each other only by what it says of them.

const CURVE = { name: 'ECDH', namedCurve: 'P-256' };
const subtle = () => globalThis.crypto.subtle;

const toText = (bytes) => {
  let s = '';
  for (const b of bytes) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
};
const toBytes = (text, most = 512) => {
  if (typeof text !== 'string' || !text || text.length > most || !/^[\w-]+$/.test(text)) return null;
  try {
    const s = atob(text.replace(/-/g, '+').replace(/_/g, '/'));
    return Uint8Array.from(s, (c) => c.charCodeAt(0));
  } catch { return null; }
};

/** A public half as it travels: 65 bytes of an uncompressed P-256 point, in base64url. */
const isPub = (text) => { const b = toBytes(text, 120); return Boolean(b) && b.length === 65 && b[0] === 4; };

/** A new key for a meeting: 32 random bytes. */
export function newMeetingKey() { return globalThis.crypto.getRandomValues(new Uint8Array(32)); }

const about = ({ meeting, from, to }, askPub, sealPub) => new TextEncoder().encode(['rutba-talk meeting key 1', meeting, from, to, askPub, sealPub].map((x) => String(x ?? '')).join('\n'));

async function lockOf(privateKey, theirPub, bound) {
  const their = await subtle().importKey('raw', toBytes(theirPub, 120), CURVE, false, []);
  const shared = await subtle().deriveBits({ name: 'ECDH', public: their }, privateKey, 256);
  const base = await subtle().importKey('raw', shared, 'HKDF', false, ['deriveKey']);
  return subtle().deriveKey({ name: 'HKDF', hash: 'SHA-256', salt: new Uint8Array(32), info: bound }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
}

/**
 * The asking side: a pair made for this one asking. `pub` is what is said to
 * whoever has the key; `open` reads what comes back, and answers the
 * meeting's key or null for anything that was not sealed for this asking.
 * @param {{ meeting: string, from: string, to: string }} who the meeting, who has the key (`from`), and who asks (`to`)
 */
export async function askForKey(who) {
  const pair = await subtle().generateKey(CURVE, false, ['deriveBits']);
  const pub = toText(new Uint8Array(await subtle().exportKey('raw', pair.publicKey)));
  return {
    pub,
    async open(sealed) {
      if (!sealed || !isPub(sealed.pub)) return null;
      const iv = toBytes(sealed.iv, 32);
      const box = toBytes(sealed.box, 256);
      if (!iv || iv.length !== 12 || !box) return null;
      try {
        const bound = about(who, pub, sealed.pub);
        const lock = await lockOf(pair.privateKey, sealed.pub, bound);
        const key = new Uint8Array(await subtle().decrypt({ name: 'AES-GCM', iv, additionalData: bound }, lock, box));
        return key.length === 32 ? key : null;
      } catch { return null; }
    },
  };
}

/**
 * The meeting's key sealed for one person who asked with `askPub`:
 * { pub, iv, box } as text, or null when what they asked with is not a
 * public half.
 * @param {string} askPub
 * @param {Uint8Array} key
 * @param {{ meeting: string, from: string, to: string }} who
 */
export async function sealKeyFor(askPub, key, who) {
  if (!isPub(askPub) || !(key instanceof Uint8Array) || key.length !== 32) return null;
  try {
    const pair = await subtle().generateKey(CURVE, false, ['deriveBits']);
    const pub = toText(new Uint8Array(await subtle().exportKey('raw', pair.publicKey)));
    const bound = about(who, askPub, pub);
    const lock = await lockOf(pair.privateKey, askPub, bound);
    const iv = globalThis.crypto.getRandomValues(new Uint8Array(12));
    const box = new Uint8Array(await subtle().encrypt({ name: 'AES-GCM', iv, additionalData: bound }, lock, key));
    return { pub, iv: toText(iv), box: toText(box) };
  } catch { return null; }
}

/**
 * The line beside sound and picture in a meeting on a server (who is muted,
 * whose hand is up, what is written), sealed with the meeting's key by the
 * app that says it and opened by the apps that hear it. The server passes
 * the words on and names who sent them; it holds no key, so it can neither
 * read them nor write any, and words sealed by one person do not open as
 * another's, nor words meant for one person as meant for everyone.
 *
 * A message as it travels: one byte for this form (1), one for whom it is
 * meant (0 everyone, 1 the one who receives it), eight that are this
 * sender's own for as long as they are in the meeting, six that count their
 * messages, twelve for the seal, then the sealed words (AES-256-GCM, the key
 * drawn from the meeting's with HKDF). Each receiver takes from one sender
 * only a count higher than the last, so words passed on a second time are
 * not heeded.
 *
 * What it does not stop: a server that makes a person seem to leave and come
 * back can pass on again, once, what that person said earlier in the same
 * meeting.
 *
 * @param {Uint8Array} key the meeting's key
 * @param {{ meeting: string, me: string }} who
 */
export async function wireSeal(key, { meeting, me }) {
  const base = await subtle().importKey('raw', key, 'HKDF', false, ['deriveKey']);
  const lock = await subtle().deriveKey({ name: 'HKDF', hash: 'SHA-256', salt: new Uint8Array(32), info: new TextEncoder().encode('rutba-talk wire 1') }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
  const run = globalThis.crypto.getRandomValues(new Uint8Array(8));
  let count = 0;
  const bound = (from, to, head) => {
    const words = new TextEncoder().encode(['rutba-talk wire 1', meeting, from, to, ''].join('\n'));
    const all = new Uint8Array(words.length + head.length);
    all.set(words);
    all.set(head, words.length);
    return all;
  };
  return {
    /** A message sealed, for everyone or for one person (`to`, their id). */
    async seal(message, to = null) {
      count += 1;
      const head = new Uint8Array(28);
      head[0] = 1;
      head[1] = to ? 1 : 0;
      head.set(run, 2);
      for (let i = 0, n = count; i < 6; i += 1) { head[15 - i] = n % 256; n = Math.floor(n / 256); }
      const iv = globalThis.crypto.getRandomValues(new Uint8Array(12));
      head.set(iv, 16);
      const box = new Uint8Array(await subtle().encrypt({ name: 'AES-GCM', iv, additionalData: bound(me, to || '*', head.subarray(0, 16)) }, lock, new TextEncoder().encode(JSON.stringify(message))));
      const out = new Uint8Array(28 + box.length);
      out.set(head);
      out.set(box, 28);
      return out;
    },
    /**
     * What arrived, opened: the words as text, or null for anything not
     * sealed with the meeting's key by `from` for this person, or passed on
     * a second time. `last` is what was last taken from that sender
     * ({ run, seq }), kept by whoever calls.
     */
    async open(bytes, from, last) {
      if (!(bytes instanceof Uint8Array) || bytes.length < 29 + 16 || bytes[0] !== 1 || bytes[1] > 1) return null;
      const runText = toText(bytes.subarray(2, 10));
      let seq = 0;
      for (let i = 10; i < 16; i += 1) seq = seq * 256 + bytes[i];
      if (last.run && (last.run !== runText || seq <= last.seq)) return null;
      try {
        const text = new TextDecoder().decode(await subtle().decrypt({ name: 'AES-GCM', iv: bytes.subarray(16, 28), additionalData: bound(from, bytes[1] ? me : '*', bytes.subarray(0, 16)) }, lock, bytes.subarray(28)));
        // Taken in order: another that was opened meanwhile may have moved the count on.
        if (last.run && (last.run !== runText || seq <= last.seq)) return null;
        last.run = runText;
        last.seq = seq;
        return text;
      } catch { return null; }
    },
  };
}

const MID = /^[\w-]{8,64}$/;
const WHY = new Set(['not-starter', 'refused', 'no-server', 'server-gone', 'offline', 'failed', 'full']);

/**
 * What the apps of a meeting on a server say to each other in its set-up,
 * checked as it arrives; anything else is null.
 *   { t: 'srv-ask', pub }                 someone coming in asks the starter to be let in and for the key
 *   { t: 'srv-let', mid, pub, iv, box }   let in at the server under `mid`, and the key sealed for them
 *   { t: 'srv-no', why }                  not let in
 *   { t: 'srv-moved' }                    the meeting is on another server now: ask again
 */
export function readSetUp(data) {
  if (!data || typeof data !== 'object') return null;
  if (data.t === 'srv-ask') return isPub(data.pub) ? { t: 'srv-ask', pub: data.pub } : null;
  if (data.t === 'srv-let') {
    if (typeof data.mid !== 'string' || !MID.test(data.mid) || !isPub(data.pub)) return null;
    const iv = toBytes(data.iv, 32);
    const box = toBytes(data.box, 256);
    return iv && iv.length === 12 && box ? { t: 'srv-let', mid: data.mid, pub: data.pub, iv: data.iv, box: data.box } : null;
  }
  if (data.t === 'srv-no') return { t: 'srv-no', why: WHY.has(data.why) ? data.why : 'refused' };
  if (data.t === 'srv-moved') return { t: 'srv-moved' };
  return null;
}
