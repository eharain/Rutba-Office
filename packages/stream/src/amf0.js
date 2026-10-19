// AMF0, the encoding RTMP's commands and a stream's metadata are written in.
//
// Only what publishing needs: numbers, booleans, strings, objects, an ECMA
// array (the metadata), null and undefined, and strict arrays when reading.

const NUMBER = 0x00;
const BOOLEAN = 0x01;
const STRING = 0x02;
const OBJECT = 0x03;
const NULL = 0x05;
const UNDEFINED = 0x06;
const ECMA_ARRAY = 0x08;
const OBJECT_END = 0x09;
const STRICT_ARRAY = 0x0a;
const LONG_STRING = 0x0c;

/** A value marked as an ECMA array (an object with a count), as onMetaData is written. */
export class EcmaArray {
  constructor(entries = {}) { this.entries = entries; }
}

const utf8 = (s) => Buffer.from(String(s), 'utf8');

function key(name) {
  const b = utf8(name);
  const len = Buffer.alloc(2);
  len.writeUInt16BE(b.length);
  return Buffer.concat([len, b]);
}

function props(obj) {
  const parts = [];
  for (const [k, v] of Object.entries(obj)) {
    if (v === undefined) continue;
    parts.push(key(k), encodeValue(v));
  }
  parts.push(Buffer.from([0x00, 0x00, OBJECT_END]));
  return Buffer.concat(parts);
}

/** One value as AMF0. */
export function encodeValue(v) {
  if (v === null) return Buffer.from([NULL]);
  if (v === undefined) return Buffer.from([UNDEFINED]);
  if (typeof v === 'number') {
    const b = Buffer.alloc(9);
    b[0] = NUMBER;
    b.writeDoubleBE(v, 1);
    return b;
  }
  if (typeof v === 'boolean') return Buffer.from([BOOLEAN, v ? 1 : 0]);
  if (typeof v === 'string') {
    const s = utf8(v);
    if (s.length > 0xffff) {
      const h = Buffer.alloc(5);
      h[0] = LONG_STRING;
      h.writeUInt32BE(s.length, 1);
      return Buffer.concat([h, s]);
    }
    const h = Buffer.alloc(3);
    h[0] = STRING;
    h.writeUInt16BE(s.length, 1);
    return Buffer.concat([h, s]);
  }
  if (v instanceof EcmaArray) {
    const h = Buffer.alloc(5);
    h[0] = ECMA_ARRAY;
    h.writeUInt32BE(Object.keys(v.entries).length, 1);
    return Buffer.concat([h, props(v.entries)]);
  }
  if (typeof v === 'object') return Buffer.concat([Buffer.from([OBJECT]), props(v)]);
  throw new TypeError(`AMF0 cannot write ${typeof v}`);
}

/** Several values one after another, as a command is written. */
export function encode(...values) {
  return Buffer.concat(values.map(encodeValue));
}

/** Every value in a buffer, in order. */
export function decode(buf) {
  const out = [];
  let at = 0;
  const readString = (long = false) => {
    const n = long ? buf.readUInt32BE(at) : buf.readUInt16BE(at);
    at += long ? 4 : 2;
    const s = buf.toString('utf8', at, at + n);
    at += n;
    return s;
  };
  const readProps = () => {
    const obj = {};
    for (;;) {
      if (at + 3 <= buf.length && buf.readUInt16BE(at) === 0 && buf[at + 2] === OBJECT_END) { at += 3; return obj; }
      const k = readString();
      obj[k] = readValue();
    }
  };
  function readValue() {
    const type = buf[at++];
    switch (type) {
      case NUMBER: { const n = buf.readDoubleBE(at); at += 8; return n; }
      case BOOLEAN: return buf[at++] !== 0;
      case STRING: return readString();
      case LONG_STRING: return readString(true);
      case OBJECT: return readProps();
      case NULL: return null;
      case UNDEFINED: return undefined;
      case ECMA_ARRAY: at += 4; return readProps();
      case STRICT_ARRAY: {
        const n = buf.readUInt32BE(at);
        at += 4;
        const list = [];
        for (let i = 0; i < n; i++) list.push(readValue());
        return list;
      }
      default: throw new Error(`AMF0 type ${type} is not read here`);
    }
  }
  while (at < buf.length) out.push(readValue());
  return out;
}
