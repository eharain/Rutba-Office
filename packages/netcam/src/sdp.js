// What a camera says it sends (SDP, RFC 8866), read for the reader: each
// stream's kind, its codec, its payload number and the address it is set up
// by. Nothing here is trusted further than being text: a field that is not
// what it should be is left out, never guessed.

/** `a=fmtp` parameters as an object, keys in lower case. */
function fmtp(text) {
  const out = {};
  for (const part of String(text || '').split(';')) {
    const at = part.indexOf('=');
    if (at > 0) out[part.slice(0, at).trim().toLowerCase()] = part.slice(at + 1).trim();
  }
  return out;
}

/**
 * The address a stream is set up or played by: its own `control` joined to
 * the base the camera gave (Content-Base, or the address asked for). A
 * control that is a whole address is used as it is; `*` or none is the base.
 */
export function controlUrl(base, control) {
  const c = String(control || '').trim();
  if (!c || c === '*') return base;
  if (/^rtsps?:\/\//i.test(c)) return c;
  if (c.startsWith('/')) {
    const m = /^(rtsps?:\/\/[^/]+)/i.exec(base);
    return m ? `${m[1]}${c}` : base;
  }
  return `${base}${base.endsWith('/') ? '' : '/'}${c}`;
}

const FIXED = { 0: ['PCMU', 8000], 8: ['PCMA', 8000], 26: ['JPEG', 90000] };

/**
 * { control, tracks: [{ kind: 'video' | 'audio', payload, codec, clock,
 * channels, fmtp, control, transport }] }. A stream with no rtpmap for its
 * payload number takes the number's fixed meaning where it has one (0 and 8:
 * G.711).
 */
export function parseSdp(text) {
  const out = { control: '', tracks: [] };
  let track = null;
  // Once the first stream starts, nothing after it speaks for the whole session, a stream left out here included.
  let streams = false;
  for (const raw of String(text || '').split(/\r?\n/)) {
    const line = raw.trim();
    if (line.length < 2 || line[1] !== '=') continue;
    const value = line.slice(2);
    if (line[0] === 'm') {
      const m = /^(\w+)\s+\d+(?:\/\d+)?\s+(\S+)\s+(.*)$/.exec(value);
      track = null;
      streams = true;
      if (!m || !['video', 'audio'].includes(m[1])) continue;
      const payload = Number(m[3].split(/\s+/)[0]);
      if (!Number.isInteger(payload) || payload < 0 || payload > 127) continue;
      const fixed = FIXED[payload];
      track = { kind: m[1], payload, codec: fixed?.[0] || '', clock: fixed?.[1] || 0, channels: m[1] === 'audio' ? 1 : 0, fmtp: {}, control: '', transport: m[2] };
      out.tracks.push(track);
    } else if (line[0] === 'a') {
      const at = value.indexOf(':');
      const name = at < 0 ? value : value.slice(0, at);
      const rest = at < 0 ? '' : value.slice(at + 1);
      if (name === 'control') {
        if (track) track.control = rest.trim();
        else if (!streams) out.control = rest.trim();
      } else if (track && name === 'rtpmap') {
        const m = /^(\d+)\s+([^/\s]+)\/(\d+)(?:\/(\d+))?/.exec(rest);
        if (m && Number(m[1]) === track.payload) {
          track.codec = m[2].toUpperCase();
          track.clock = Number(m[3]);
          if (track.kind === 'audio') track.channels = Number(m[4]) || 1;
        }
      } else if (track && name === 'fmtp') {
        const m = /^(\d+)\s+(.*)$/.exec(rest);
        if (m && Number(m[1]) === track.payload) track.fmtp = fmtp(m[2]);
      }
    }
  }
  return out;
}

/** The parameter sets a camera names in its SDP (sprop-parameter-sets), as bytes: the SPS and PPS a decoder needs before the first frame. */
export function spropSets(track) {
  const out = [];
  for (const part of String(track?.fmtp?.['sprop-parameter-sets'] || '').split(',')) {
    const clean = part.trim();
    if (!clean || !/^[A-Za-z0-9+/]+=*$/.test(clean)) continue;
    const bytes = Buffer.from(clean, 'base64');
    if (bytes.length) out.push(bytes);
  }
  return out;
}
