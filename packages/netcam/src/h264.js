// A camera's H.264 picture, from the RTP packets it arrives in (RFC 6184) to
// whole frames a decoder takes.
//
// A frame is one NAL unit or several, and a NAL unit is one packet, several
// in one packet (STAP-A) or one cut across many (FU-A). They are put back
// together here and handed on as Annex B: each unit after a start code, which
// is the form a decoder is given when it has no separate description.
//
// A decoder can start only at a key frame and only when it has the stream's
// parameter sets (SPS and PPS). Cameras send those in their SDP, in the
// stream, or both, and not always beside the key frame; the last ones seen
// are kept here and put in front of every key frame that lacks them, so each
// key frame can be decoded by itself. Nothing is handed on before the first
// key frame, and after a gap in the packets nothing until the next: a frame
// built on a missing one is a smeared picture, and a short stop is better.

const START = Buffer.from([0, 0, 0, 1]);

// More than any camera sends in one frame; a stream that claims more is not a picture.
const MAX_NAL = 8 * 1024 * 1024;
const MAX_UNIT = 16 * 1024 * 1024;

const SLICE = 1;
const IDR = 5;
const SPS = 7;
const PPS = 8;
const AUD = 9;
const FILLER = 12;
const STAP_A = 24;
const FU_A = 28;

const typeOf = (nal) => nal[0] & 0x1f;

/** Reads the bits of a parameter set: the bytes after the NAL header, with the escape bytes (00 00 03) taken out. */
class Bits {
  constructor(nal) {
    const out = [];
    for (let i = 1; i < nal.length; i += 1) {
      if (i >= 3 && nal[i] === 3 && nal[i - 1] === 0 && nal[i - 2] === 0) continue;
      out.push(nal[i]);
    }
    this.bytes = out;
    this.at = 0;
  }

  bit() {
    const byte = this.bytes[this.at >> 3];
    if (byte === undefined) throw new RangeError('The parameter set ends early.');
    const b = (byte >> (7 - (this.at & 7))) & 1;
    this.at += 1;
    return b;
  }

  bits(n) {
    let v = 0;
    for (let i = 0; i < n; i += 1) v = v * 2 + this.bit();
    return v;
  }

  /** Unsigned Exp-Golomb. */
  ue() {
    let zeros = 0;
    while (this.bit() === 0) {
      zeros += 1;
      if (zeros > 32) throw new RangeError('The parameter set is not readable.');
    }
    return 2 ** zeros - 1 + this.bits(zeros);
  }

  /** Signed Exp-Golomb. */
  se() {
    const v = this.ue();
    return v % 2 ? (v + 1) / 2 : -(v / 2);
  }
}

const HIGH_PROFILES = new Set([100, 110, 122, 244, 44, 83, 86, 118, 128, 138, 139, 134, 135]);

function skipScalingList(r, size) {
  let last = 8;
  let next = 8;
  for (let j = 0; j < size; j += 1) {
    if (next !== 0) next = (last + r.se() + 256) % 256;
    last = next === 0 ? last : next;
  }
}

/**
 * What a sequence parameter set says about the picture: { codec, width,
 * height }. `codec` is the name a decoder is configured by ("avc1.64001f":
 * profile, constraints and level). The size is 0 by 0 when the set cannot be
 * read that far; null when it is not a parameter set at all.
 */
export function spsInfo(sps) {
  if (!sps || sps.length < 4 || typeOf(sps) !== SPS) return null;
  const hex = (n) => n.toString(16).padStart(2, '0');
  const info = { codec: `avc1.${hex(sps[1])}${hex(sps[2])}${hex(sps[3])}`, width: 0, height: 0 };
  try {
    const r = new Bits(sps);
    const profile = r.bits(8);
    r.bits(16);
    r.ue();
    let chroma = 1;
    let separate = 0;
    if (HIGH_PROFILES.has(profile)) {
      chroma = r.ue();
      if (chroma === 3) separate = r.bit();
      r.ue();
      r.ue();
      r.bit();
      if (r.bit()) {
        const lists = chroma !== 3 ? 8 : 12;
        for (let i = 0; i < lists; i += 1) if (r.bit()) skipScalingList(r, i < 6 ? 16 : 64);
      }
    }
    r.ue();
    const orderType = r.ue();
    if (orderType === 0) r.ue();
    else if (orderType === 1) {
      r.bit();
      r.se();
      r.se();
      const cycle = r.ue();
      if (cycle > 255) throw new RangeError('The parameter set is not readable.');
      for (let i = 0; i < cycle; i += 1) r.se();
    }
    r.ue();
    r.bit();
    const widthMbs = r.ue() + 1;
    const heightUnits = r.ue() + 1;
    const frameOnly = r.bit();
    if (!frameOnly) r.bit();
    r.bit();
    let crop = [0, 0, 0, 0];
    if (r.bit()) crop = [r.ue(), r.ue(), r.ue(), r.ue()];
    const arrayType = separate ? 0 : chroma;
    const unitX = arrayType === 0 ? 1 : arrayType === 3 ? 1 : 2;
    const unitY = (arrayType === 1 ? 2 : 1) * (2 - frameOnly);
    const width = widthMbs * 16 - unitX * (crop[0] + crop[1]);
    const height = (2 - frameOnly) * heightUnits * 16 - unitY * (crop[2] + crop[3]);
    if (width > 0 && height > 0 && width <= 16384 && height <= 16384) {
      info.width = width;
      info.height = height;
    }
  } catch {
    // The name alone is enough to start a decoder; it finds the size itself.
  }
  return info;
}

/** NAL units as Annex B: each after a start code. */
export function annexB(nals) {
  const parts = [];
  for (const nal of nals) parts.push(START, nal);
  return Buffer.concat(parts);
}

/** The NAL units inside Annex B bytes (what a test or an encoder hands over), without their start codes. */
export function splitAnnexB(bytes) {
  const out = [];
  let from = -1;
  let i = 0;
  while (i + 2 < bytes.length) {
    if (bytes[i] === 0 && bytes[i + 1] === 0 && bytes[i + 2] === 1) {
      if (from >= 0) {
        let end = i;
        while (end > from && bytes[end - 1] === 0) end -= 1;
        if (end > from) out.push(bytes.subarray(from, end));
      }
      from = i + 3;
      i += 3;
    } else i += 1;
  }
  if (from >= 0 && from < bytes.length) out.push(bytes.subarray(from));
  return out;
}

export class H264Assembler {
  #clock;
  #sps = null;
  #pps = null;
  #announced = '';
  #nals = [];
  #size = 0;
  #stamp = null;
  #pieces = null;
  #piecesSize = 0;
  #lastSequence = null;
  #waiting = true;
  #lastStamp = null;
  #extended = 0;

  /** `sets` are the parameter sets the camera named in its SDP; `clock` is the stream's (90000 for video). */
  constructor({ sets = [], clock = 90000 } = {}) {
    this.#clock = clock > 0 ? clock : 90000;
    for (const set of sets) this.#keep(set);
  }

  #keep(nal) {
    if (!nal?.length) return;
    if (typeOf(nal) === SPS) this.#sps = Buffer.from(nal);
    else if (typeOf(nal) === PPS) this.#pps = Buffer.from(nal);
  }

  /** The stream's time on a clock that does not wrap, in microseconds from its first packet. */
  #micros(stamp) {
    if (this.#lastStamp === null) this.#extended = 0;
    else this.#extended += (stamp - this.#lastStamp) | 0;
    this.#lastStamp = stamp;
    return Math.round((this.#extended * 1e6) / this.#clock);
  }

  #lose() {
    this.#nals = [];
    this.#size = 0;
    this.#pieces = null;
    this.#piecesSize = 0;
    this.#waiting = true;
  }

  #add(nal) {
    if (!nal.length || nal.length > MAX_NAL) return;
    const type = typeOf(nal);
    if (type === AUD || type === FILLER || type === 0 || type > 23) return;
    this.#keep(nal);
    this.#size += nal.length;
    if (this.#size > MAX_UNIT) {
      this.#lose();
      return;
    }
    this.#nals.push(nal);
  }

  #flush(out) {
    const nals = this.#nals;
    const stamp = this.#stamp;
    this.#nals = [];
    this.#size = 0;
    if (!nals.length) return;
    const key = nals.some((n) => typeOf(n) === IDR);
    if (!key && !nals.some((n) => typeOf(n) === SLICE)) return;
    if (this.#waiting && !key) return;
    const timestamp = this.#micros(stamp);
    const unit = { timestamp, key, data: null };
    if (key) {
      const head = [];
      if (!nals.some((n) => typeOf(n) === SPS)) {
        if (!this.#sps) return;
        head.push(this.#sps);
      }
      if (!nals.some((n) => typeOf(n) === PPS)) {
        if (!this.#pps) return;
        head.push(this.#pps);
      }
      // Parameter sets lead, whatever order the camera sent them in.
      const sets = nals.filter((n) => typeOf(n) === SPS || typeOf(n) === PPS);
      const rest = nals.filter((n) => typeOf(n) !== SPS && typeOf(n) !== PPS);
      unit.data = annexB([...head, ...sets, ...rest]);
      this.#waiting = false;
      const info = spsInfo(this.#sps);
      const told = this.#sps.toString('hex');
      if (info && told !== this.#announced) {
        this.#announced = told;
        unit.config = info;
      }
    } else unit.data = annexB(nals);
    out.push(unit);
  }

  /**
   * One RTP packet in (as parseRtp gives it), the frames it finished out:
   * [{ timestamp, key, data, config? }]. `timestamp` is in microseconds,
   * `data` is Annex B, and `config` ({ codec, width, height }) rides on the
   * first key frame and on each one after the stream's parameters change.
   */
  push(packet) {
    const out = [];
    if (!packet?.payload?.length) return out;
    if (this.#lastSequence !== null) {
      if (packet.sequence === this.#lastSequence) return out;
      if (packet.sequence !== ((this.#lastSequence + 1) & 0xffff)) this.#lose();
    }
    this.#lastSequence = packet.sequence;
    if (this.#stamp !== null && packet.timestamp !== this.#stamp) {
      this.#flush(out);
      this.#pieces = null;
      this.#piecesSize = 0;
    }
    this.#stamp = packet.timestamp;

    const payload = packet.payload;
    const type = payload[0] & 0x1f;
    if (type >= 1 && type <= 23) this.#add(payload);
    else if (type === STAP_A) {
      let at = 1;
      while (at + 2 <= payload.length) {
        const size = payload.readUInt16BE(at);
        at += 2;
        if (size === 0 || at + size > payload.length) break;
        this.#add(payload.subarray(at, at + size));
        at += size;
      }
    } else if (type === FU_A && payload.length >= 2) {
      const start = Boolean(payload[1] & 0x80);
      const end = Boolean(payload[1] & 0x40);
      const piece = payload.subarray(2);
      if (start) {
        this.#pieces = [Buffer.from([(payload[0] & 0xe0) | (payload[1] & 0x1f)]), piece];
        this.#piecesSize = 1 + piece.length;
      } else if (this.#pieces) {
        this.#pieces.push(piece);
        this.#piecesSize += piece.length;
        if (this.#piecesSize > MAX_NAL) {
          this.#lose();
        }
      }
      if (end && this.#pieces) {
        const whole = Buffer.concat(this.#pieces);
        this.#pieces = null;
        this.#piecesSize = 0;
        this.#add(whole);
      }
    }
    if (packet.marker) this.#flush(out);
    return out;
  }
}
