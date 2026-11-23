// The frames of a WebSocket (RFC 6455), for the one connection Office's main
// side makes by hand: the signalling of a meeting on a Talk Server, which goes
// through the server's door on a connection whose certificate was checked
// here first (talk-door.js). No extensions are offered, so none are read.
//
// Kept apart from the connection so both directions can be tested with no
// network: what one side writes, the other reads.

import crypto from 'node:crypto';

const GUID = '258EAFA5-E914-47DA-95CA-C5AB0DC85B11';

export const OP = { more: 0, text: 1, binary: 2, close: 8, ping: 9, pong: 10 };

/** What a server answers in `sec-websocket-accept` for the key it was sent. */
export const acceptKey = (key) => crypto.createHash('sha1').update(String(key) + GUID).digest('base64');

/**
 * One frame. A client masks what it sends and a server does not.
 * @param {number} opcode
 * @param {Buffer|Uint8Array|string} [payload]
 * @param {{ mask?: boolean, fin?: boolean }} [o]
 */
export function encodeFrame(opcode, payload = Buffer.alloc(0), { mask = true, fin = true } = {}) {
  const data = Buffer.isBuffer(payload) ? payload : typeof payload === 'string' ? Buffer.from(payload, 'utf8') : Buffer.from(payload.buffer ?? payload, payload.byteOffset ?? 0, payload.byteLength ?? payload.length);
  const n = data.length;
  const head = n < 126 ? 2 : n < 65536 ? 4 : 10;
  const out = Buffer.alloc(head + (mask ? 4 : 0) + n);
  out[0] = (fin ? 0x80 : 0) | (opcode & 0x0f);
  if (n < 126) out[1] = n;
  else if (n < 65536) { out[1] = 126; out.writeUInt16BE(n, 2); } else { out[1] = 127; out.writeBigUInt64BE(BigInt(n), 2); }
  if (!mask) { data.copy(out, head); return out; }
  out[1] |= 0x80;
  const key = crypto.randomBytes(4);
  key.copy(out, head);
  for (let i = 0; i < n; i += 1) out[head + 4 + i] = data[i] ^ key[i & 3];
  return out;
}

/** The payload of a close frame: its code and a short reason. */
export function closePayload(code = 1000, reason = '') {
  const text = Buffer.from(String(reason || ''), 'utf8').subarray(0, 120);
  const out = Buffer.alloc(2 + text.length);
  out.writeUInt16BE(code, 0);
  text.copy(out, 2);
  return out;
}

/**
 * Reads frames out of what arrives, in whatever pieces it arrives: whole
 * messages (their parts joined), pings, pongs and the close.
 * @param {object} o
 * @param {(data: Buffer, binary: boolean) => void} o.onMessage
 * @param {(data: Buffer) => void} [o.onPing]
 * @param {(data: Buffer) => void} [o.onPong]
 * @param {(code: number, reason: string) => void} [o.onClose]
 * @param {(why: string) => void} [o.onError] what was read is not a WebSocket's, or is larger than `most`; nothing more is read
 * @param {number} [o.most] the largest message taken, in bytes
 */
export function frameReader({ onMessage, onPing = () => {}, onPong = () => {}, onClose = () => {}, onError = () => {}, most = 8 * 1024 * 1024 }) {
  let held = Buffer.alloc(0);
  /** The parts of a message still arriving: its opcode and what came so far. */
  let parts = null;
  let partsSize = 0;
  let dead = false;
  const fail = (why) => { dead = true; held = Buffer.alloc(0); parts = null; onError(why); };

  return {
    push(chunk) {
      if (dead) return;
      held = held.length ? Buffer.concat([held, chunk]) : Buffer.from(chunk);
      for (;;) {
        if (dead || held.length < 2) return;
        const fin = Boolean(held[0] & 0x80);
        if (held[0] & 0x70) return fail('extension');
        const opcode = held[0] & 0x0f;
        const masked = Boolean(held[1] & 0x80);
        let n = held[1] & 0x7f;
        let at = 2;
        if (n === 126) {
          if (held.length < 4) return;
          n = held.readUInt16BE(2);
          at = 4;
        } else if (n === 127) {
          if (held.length < 10) return;
          const big = held.readBigUInt64BE(2);
          if (big > BigInt(most)) return fail('too-large');
          n = Number(big);
          at = 10;
        }
        if (n > most) return fail('too-large');
        const control = opcode >= 8;
        if (control && (!fin || n > 125)) return fail('control');
        const need = at + (masked ? 4 : 0) + n;
        if (held.length < need) return;
        let data = held.subarray(at + (masked ? 4 : 0), need);
        if (masked) {
          const key = held.subarray(at, at + 4);
          const clear = Buffer.alloc(n);
          for (let i = 0; i < n; i += 1) clear[i] = data[i] ^ key[i & 3];
          data = clear;
        } else data = Buffer.from(data);
        held = held.subarray(need);

        if (opcode === OP.ping) { onPing(data); continue; }
        if (opcode === OP.pong) { onPong(data); continue; }
        if (opcode === OP.close) {
          dead = true;
          onClose(data.length >= 2 ? data.readUInt16BE(0) : 1005, data.length > 2 ? data.subarray(2).toString('utf8') : '');
          return;
        }
        if (opcode === OP.more) {
          if (!parts) return fail('continuation');
          parts.list.push(data);
        } else if (opcode === OP.text || opcode === OP.binary) {
          if (parts) return fail('interleaved');
          parts = { binary: opcode === OP.binary, list: [data] };
          partsSize = 0;
        } else return fail('opcode');
        partsSize += n;
        if (partsSize > most) return fail('too-large');
        if (fin) {
          const whole = parts.list.length === 1 ? parts.list[0] : Buffer.concat(parts.list);
          const binary = parts.binary;
          parts = null;
          onMessage(whole, binary);
        }
      }
    },
  };
}
