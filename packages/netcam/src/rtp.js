// One RTP packet (RFC 3550), as its fields and its payload.

/** { payloadType, marker, sequence, timestamp, ssrc, payload }, or null for what is not an RTP packet. */
export function parseRtp(buf) {
  if (!buf || buf.length < 12 || buf[0] >> 6 !== 2) return null;
  const padding = Boolean(buf[0] & 0x20);
  const extension = Boolean(buf[0] & 0x10);
  let at = 12 + (buf[0] & 0x0f) * 4;
  if (extension) {
    if (buf.length < at + 4) return null;
    at += 4 + buf.readUInt16BE(at + 2) * 4;
  }
  let end = buf.length;
  if (padding) {
    const pad = buf[buf.length - 1];
    if (pad === 0 || pad > end - at) return null;
    end -= pad;
  }
  if (at > end) return null;
  return {
    payloadType: buf[1] & 0x7f,
    marker: Boolean(buf[1] & 0x80),
    sequence: buf.readUInt16BE(2),
    timestamp: buf.readUInt32BE(4),
    ssrc: buf.readUInt32BE(8),
    payload: buf.subarray(at, end),
  };
}
