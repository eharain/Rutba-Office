// Pictures with EXIF and XMP in each container, made here byte by byte, for
// the EXIF tests: a TIFF block with a camera, an Exif directory, a GPS
// directory and a thumbnail; and that block inside a JPEG, a PNG, a WebP, a
// HEIC (item located in the file) and an AVIF (item in the meta box's idat).
import zlib from 'node:zlib';

const enc = (s) => new TextEncoder().encode(s);
const ascii = (s) => [...s].map((c) => c.charCodeAt(0));

export function concat(parts) {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
}

/**
 * A TIFF block. `ifd0`, `exif`, `gps` are lists of [tag, type, values]:
 * type 2 takes a string, 1 and 7 a list of bytes, 3/4 numbers, 5 pairs.
 * `thumb` adds IFD1 with a JPEG thumbnail of those bytes.
 */
export function tiffBlock({ little = true, ifd0 = [], exif = null, gps = null, thumb = null } = {}) {
  const buf = new Uint8Array(65536);
  const dv = new DataView(buf.buffer);
  let len = 8;
  buf.set(little ? [0x49, 0x49, 42, 0] : [0x4d, 0x4d, 0, 42], 0);
  const u16 = (o, v) => dv.setUint16(o, v, little);
  const u32 = (o, v) => dv.setUint32(o, v, little);
  const valueBytes = (type, values) => {
    if (type === 2) return [...enc(values), 0];
    if (type === 1 || type === 7) return [...values];
    const size = { 3: 2, 4: 4, 5: 8 }[type];
    const list = Array.isArray(values) ? values : [values];
    const out = new Uint8Array(list.length * size);
    const v = new DataView(out.buffer);
    list.forEach((x, i) => {
      if (type === 3) v.setUint16(i * 2, x, little);
      else if (type === 4) v.setUint32(i * 4, x, little);
      else { v.setUint32(i * 8, x[0], little); v.setUint32(i * 8 + 4, x[1], little); }
    });
    return [...out];
  };
  const count = (type, values) => (type === 2 ? enc(values).length + 1 : type === 1 || type === 7 ? values.length : Array.isArray(values) ? values.length : 1);
  const writeDir = (entries) => {
    if (len & 1) len++;
    const at = len;
    len += 2 + entries.length * 12 + 4;
    u16(at, entries.length);
    entries.sort((a, b) => a[0] - b[0]).forEach(([tag, type, values], i) => {
      const p = at + 2 + i * 12;
      u16(p, tag);
      if (type === 'ifd') { u16(p + 2, 4); u32(p + 4, 1); u32(p + 8, values()); return; }
      u16(p + 2, type);
      u32(p + 4, count(type, values));
      const bytes = valueBytes(type, values);
      if (bytes.length <= 4) buf.set(bytes, p + 8);
      else { if (len & 1) len++; buf.set(bytes, len); u32(p + 8, len); len += bytes.length; }
    });
    return at;
  };
  const entries = [...ifd0];
  if (exif) entries.push([0x8769, 'ifd', () => writeDir([...exif])]);
  if (gps) entries.push([0x8825, 'ifd', () => writeDir([...gps])]);
  const first = writeDir(entries);
  u32(4, first);
  if (thumb) {
    const thumbAt = len;
    buf.set(thumb, len);
    len += thumb.length;
    const ifd1 = writeDir([[0x0103, 3, 6], [0x0201, 4, thumbAt], [0x0202, 4, thumb.length]]);
    u32(first + 2 + entries.length * 12, ifd1);
  }
  return buf.slice(0, len);
}

/** XPKeywords as Windows writes them: UTF-16LE, nul-ended. */
export const xp = (s) => [...new Uint8Array(new Uint16Array([...s].map((c) => c.charCodeAt(0)).concat([0])).buffer)];

/** The camera most tests use: a phone, on its side, rated, tagged and placed. */
export const THUMB = new Uint8Array([0xff, 0xd8, 0xff, 0xdb, 0x13, 0x37, 0xbe, 0xef, 0xca, 0xfe, 0xff, 0xd9]);
export const LAT = [[51, 1], [30, 1], [2650, 100]];
export const LON = [[0, 1], [7, 1], [3999, 100]];
export function cameraBlock({ little = true, thumb = true, gps = true } = {}) {
  return tiffBlock({
    little,
    ifd0: [[0x010f, 2, 'Rutba'], [0x0110, 2, 'Test Phone 7'], [0x0112, 3, 6], [0x4746, 3, 4], [0x4749, 3, 75], [0x9c9e, 1, xp('beach;family')]],
    exif: [[0x9003, 2, '2026:10:27 09:15:00'], [0x829d, 5, [[18, 10]]], [0x8827, 3, 100], [0xa002, 4, 4032], [0xa003, 4, 3024], [0xa434, 2, 'Wide 26mm']],
    gps: gps ? [[0x0001, 2, 'N'], [0x0002, 5, LAT], [0x0003, 2, 'W'], [0x0004, 5, LON]] : null,
    thumb: thumb ? THUMB : null,
  });
}

export const XMP = `<?xpacket begin="" id="W5M0MpCehiHzreSzNTczkc9d"?>
<x:xmpmeta xmlns:x="adobe:ns:meta/"><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
<rdf:Description rdf:about="" xmlns:xmp="http://ns.adobe.com/xap/1.0/" xmlns:MicrosoftPhoto="http://ns.microsoft.com/photo/1.0/" xmp:Rating="3" MicrosoftPhoto:Rating="50"/>
<rdf:Description rdf:about="" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:subject><rdf:Bag><rdf:li>Holiday</rdf:li><rdf:li>Beach</rdf:li><rdf:li>Tom &amp; Ann</rdf:li></rdf:Bag></dc:subject></rdf:Description>
</rdf:RDF></x:xmpmeta><?xpacket end="w"?>`;

const JFIF = [0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00];
const SCAN = [0xff, 0xda, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3f, 0x00, 0x12, 0x34, 0x56, 0xff, 0xd9];
const segment = (marker, payload) => [0xff, marker, (payload.length + 2) >> 8, (payload.length + 2) & 0xff, ...payload];

export function jpegWith({ block = null, xmp = null } = {}) {
  return new Uint8Array([0xff, 0xd8, ...JFIF,
    ...(block ? segment(0xe1, [...ascii('Exif'), 0, 0, ...block]) : []),
    ...(xmp ? segment(0xe1, [...ascii('http://ns.adobe.com/xap/1.0/'), 0, ...enc(xmp)]) : []),
    ...SCAN]);
}

const pngChunk = (type, data) => {
  const body = concat([new Uint8Array(ascii(type)), data]);
  const out = new Uint8Array(12 + data.length);
  const dv = new DataView(out.buffer);
  dv.setUint32(0, data.length);
  out.set(body, 4);
  dv.setUint32(8 + data.length, zlib.crc32(body));
  return out;
};

/** A real PNG of w × h grey pixels; the EXIF and XMP before or after the pixels. */
export function pngWith({ block = null, xmp = null, after = false, width = 2, height = 2 } = {}) {
  const ihdr = new Uint8Array(13);
  const hv = new DataView(ihdr.buffer);
  hv.setUint32(0, width);
  hv.setUint32(4, height);
  ihdr.set([8, 0, 0, 0, 0], 8);
  const raw = new Uint8Array((width + 1) * height).fill(128);
  for (let y = 0; y < height; y++) raw[y * (width + 1)] = 0;
  const meta = [];
  if (block) meta.push(pngChunk('eXIf', block));
  if (xmp) meta.push(pngChunk('iTXt', concat([new Uint8Array([...ascii('XML:com.adobe.xmp'), 0, 0, 0, 0, 0]), enc(xmp)])));
  const idat = pngChunk('IDAT', new Uint8Array(zlib.deflateSync(raw)));
  return concat([new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), pngChunk('IHDR', ihdr), ...(after ? [idat, ...meta] : [...meta, idat]), pngChunk('IEND', new Uint8Array(0))]);
}

const riffChunk = (type, data) => {
  const out = new Uint8Array(8 + data.length + (data.length & 1));
  out.set(ascii(type), 0);
  new DataView(out.buffer).setUint32(4, data.length, true);
  out.set(data, 8);
  return out;
};
const riff = (chunks) => {
  const body = concat([new Uint8Array(ascii('WEBP')), ...chunks]);
  const out = new Uint8Array(8 + body.length);
  out.set(ascii('RIFF'), 0);
  new DataView(out.buffer).setUint32(4, body.length, true);
  out.set(body, 8);
  return out;
};
/** A VP8L chunk's start for a w × h picture (the rest of the bitstream is filler: only the header is read). */
export function vp8l(width, height, alpha = false, filler = 16) {
  const data = new Uint8Array(5 + filler);
  data[0] = 0x2f;
  new DataView(data.buffer).setUint32(1, ((width - 1) & 0x3fff) | (((height - 1) & 0x3fff) << 14) | ((alpha ? 1 : 0) << 28), true);
  return riffChunk('VP8L', data);
}
/** A simple WebP, as a canvas writes. */
export const simpleWebp = (width = 3, height = 2, alpha = false, filler = 16) => riff([vp8l(width, height, alpha, filler)]);
/** An extended WebP with its EXIF and XMP after the picture, as the format puts them. */
export function webpWith({ block = null, xmp = null, filler = 16, lead = false } = {}) {
  const vp8x = new Uint8Array(10);
  vp8x[0] = (block ? 0x08 : 0) | (xmp ? 0x04 : 0);
  vp8x.set([2, 0, 0, 1, 0, 0], 4);
  return riff([riffChunk('VP8X', vp8x), vp8l(3, 2, false, filler),
    ...(block ? [riffChunk('EXIF', lead ? concat([new Uint8Array([...ascii('Exif'), 0, 0]), block]) : block)] : []),
    ...(xmp ? [riffChunk('XMP ', enc(xmp))] : [])]);
}

const box = (type, ...parts) => {
  const body = concat(parts.map((p) => (p instanceof Uint8Array ? p : new Uint8Array(p))));
  const out = new Uint8Array(8 + body.length);
  new DataView(out.buffer).setUint32(0, out.length);
  out.set(ascii(type), 4);
  out.set(body, 8);
  return out;
};
const full = (type, version, ...parts) => box(type, [version, 0, 0, 0], ...parts);
const u16 = (v) => [v >> 8, v & 0xff];
const u32 = (v) => [(v >>> 24) & 0xff, (v >> 16) & 0xff, (v >> 8) & 0xff, v & 0xff];
const infe = (id, type, extra = []) => full('infe', 2, u16(id), u16(0), ascii(type), [0], extra);

/**
 * A HEIC (or, with `avif`, an AVIF) whose meta box lists a picture, an Exif
 * item and an XMP item. HEIC: the items located in mdat (iloc version 0,
 * construction method 0). AVIF: iloc version 1, the Exif and XMP in the
 * meta box's own idat (construction method 1).
 */
export function heifWith({ block = null, xmp = null, avif = false } = {}) {
  const exifItem = block ? concat([new Uint8Array([0, 0, 0, 6, ...ascii('Exif'), 0, 0]), block]) : null;
  const xmpItem = xmp ? enc(xmp) : null;
  const picture = new Uint8Array(64).fill(0x5a);
  const ftyp = box('ftyp', ascii(avif ? 'avif' : 'heic'), u32(0), ascii('mif1'), ascii(avif ? 'avif' : 'heic'));
  const items = [infe(1, avif ? 'av01' : 'hvc1')];
  if (exifItem) items.push(infe(2, 'Exif'));
  if (xmpItem) items.push(infe(3, 'mime', [...ascii('application/rdf+xml'), 0]));
  const iinf = full('iinf', 0, u16(items.length), ...items);
  const hdlr = full('hdlr', 0, u32(0), ascii('pict'), new Array(12).fill(0), [0]);
  const make = (offsets) => {
    const list = [[1, picture.length, offsets.picture, 0]];
    if (exifItem) list.push([2, exifItem.length, offsets.exif, avif ? 1 : 0]);
    if (xmpItem) list.push([3, xmpItem.length, offsets.xmp, avif ? 1 : 0]);
    const v = avif ? 1 : 0;
    const rows = list.map(([id, length, offset, method]) => [...u16(id), ...(v ? u16(method) : []), ...u16(0), ...u16(1), ...u32(offset), ...u32(length)]);
    return full('iloc', v, [0x44, 0x00], u16(list.length), rows.flat());
  };
  if (avif) {
    const idatData = concat([exifItem || new Uint8Array(0), xmpItem || new Uint8Array(0)]);
    const offsets = { picture: 0, exif: 0, xmp: exifItem ? exifItem.length : 0 };
    // The picture's offset is in the file: lay out once to learn where mdat lands.
    let meta = full('meta', 0, hdlr, iinf, make(offsets), box('idat', idatData));
    offsets.picture = ftyp.length + meta.length + 8;
    meta = full('meta', 0, hdlr, iinf, make(offsets), box('idat', idatData));
    return concat([ftyp, meta, box('mdat', picture)]);
  }
  const offsets = { picture: 0, exif: 0, xmp: 0 };
  let meta = full('meta', 0, hdlr, iinf, make(offsets));
  const start = ftyp.length + meta.length + 8;
  offsets.picture = start;
  offsets.exif = start + picture.length;
  offsets.xmp = offsets.exif + (exifItem ? exifItem.length : 0);
  meta = full('meta', 0, hdlr, iinf, make(offsets));
  return concat([ftyp, meta, box('mdat', picture, exifItem || new Uint8Array(0), xmpItem || new Uint8Array(0))]);
}

/** A one-page uncompressed grey TIFF of w × h, as a simple TIFF writer makes it: pixels first, the directory after. */
export function plainTiff(width = 4, height = 3, { little = true, value = 200 } = {}) {
  const pixels = new Uint8Array(width * height).fill(value);
  const head = 8;
  const dirAt = head + pixels.length + (pixels.length & 1);
  const entries = [[0x0100, 3, width], [0x0101, 3, height], [0x0102, 3, 8], [0x0103, 3, 1], [0x0106, 3, 1], [0x0111, 4, head], [0x0115, 3, 1], [0x0116, 3, height], [0x0117, 4, pixels.length], [0x0131, 2, 'Rutba Image']];
  const extra = [];
  const buf = new Uint8Array(dirAt + 2 + entries.length * 12 + 4 + 64);
  const dv = new DataView(buf.buffer);
  buf.set(little ? [0x49, 0x49, 42, 0] : [0x4d, 0x4d, 0, 42], 0);
  dv.setUint32(4, dirAt, little);
  buf.set(pixels, head);
  dv.setUint16(dirAt, entries.length, little);
  let tail = dirAt + 2 + entries.length * 12 + 4;
  entries.forEach(([tag, type, v], i) => {
    const p = dirAt + 2 + i * 12;
    dv.setUint16(p, tag, little);
    dv.setUint16(p + 2, type, little);
    if (type === 2) {
      const bytes = [...enc(v), 0];
      dv.setUint32(p + 4, bytes.length, little);
      buf.set(bytes, tail);
      dv.setUint32(p + 8, tail, little);
      tail += bytes.length;
    } else {
      dv.setUint32(p + 4, 1, little);
      if (type === 3) dv.setUint16(p + 8, v, little);
      else dv.setUint32(p + 8, v, little);
    }
  });
  void extra;
  return buf.slice(0, tail);
}
