// Photographs with a camera's EXIF, made here for the EXIF checks: the
// camera, lens, date and place a phone writes, and the rating and keywords
// Windows Photos and Lightroom write (EXIF Rating, XPKeywords and an XMP
// packet with xmp:Rating and dc:subject). Shared by verify-exif-location.js
// and verify-pictures-ratings.js.
import { jpegWithExif } from '@rutba/imaging/exif';

const enc = (s) => new TextEncoder().encode(s);

/**
 * A little-endian TIFF block: IFD0 entries, an Exif directory and a GPS
 * directory, each a list of [tag, type, value] (2 text, 1 bytes, 3 and 4
 * numbers, 5 a list of [numerator, denominator]).
 */
export function tiffBlock({ ifd0 = [], exif = null, gps = null }) {
  const buf = new Uint8Array(16384);
  const dv = new DataView(buf.buffer);
  let len = 8;
  buf.set([0x49, 0x49, 42, 0], 0);
  const bytesOf = (type, v) => {
    if (type === 2) return [...enc(v), 0];
    if (type === 1 || type === 7) return [...v];
    const list = Array.isArray(v) ? v : [v];
    const size = { 3: 2, 4: 4, 5: 8 }[type];
    const out = new Uint8Array(list.length * size);
    const ov = new DataView(out.buffer);
    list.forEach((x, i) => {
      if (type === 3) ov.setUint16(i * 2, x, true);
      else if (type === 4) ov.setUint32(i * 4, x, true);
      else { ov.setUint32(i * 8, x[0], true); ov.setUint32(i * 8 + 4, x[1], true); }
    });
    return [...out];
  };
  const countOf = (type, v) => (type === 2 ? enc(v).length + 1 : type === 1 || type === 7 ? v.length : Array.isArray(v) ? v.length : 1);
  const dir = (entries) => {
    if (len & 1) len++;
    const at = len;
    len += 2 + entries.length * 12 + 4;
    dv.setUint16(at, entries.length, true);
    entries.sort((a, b) => a[0] - b[0]).forEach(([tag, type, v], i) => {
      const p = at + 2 + i * 12;
      dv.setUint16(p, tag, true);
      if (type === 'ifd') { dv.setUint16(p + 2, 4, true); dv.setUint32(p + 4, 1, true); dv.setUint32(p + 8, v(), true); return; }
      dv.setUint16(p + 2, type, true);
      dv.setUint32(p + 4, countOf(type, v), true);
      const bytes = bytesOf(type, v);
      if (bytes.length <= 4) buf.set(bytes, p + 8);
      else { if (len & 1) len++; buf.set(bytes, len); dv.setUint32(p + 8, len, true); len += bytes.length; }
    });
    return at;
  };
  const entries = [...ifd0];
  if (exif) entries.push([0x8769, 'ifd', () => dir([...exif])]);
  if (gps) entries.push([0x8825, 'ifd', () => dir([...gps])]);
  dv.setUint32(4, dir(entries), true);
  return buf.slice(0, len);
}

/** XPKeywords as Windows writes them: UTF-16LE, nul-ended. */
const xp = (s) => [...new Uint8Array(new Uint16Array([...s].map((c) => c.charCodeAt(0)).concat([0])).buffer)];

/** A phone's block: camera, lens, date, and a place in London; optionally Windows' own rating and tags. */
export function cameraBlock({ rating = null, keywords = null } = {}) {
  const ifd0 = [[0x010f, 2, 'Rutba'], [0x0110, 2, 'Test Phone 7'], [0x0112, 3, 1]];
  if (rating != null) ifd0.push([0x4746, 3, rating], [0x4749, 3, [0, 1, 25, 50, 75, 99][rating]]);
  if (keywords) ifd0.push([0x9c9e, 1, xp(keywords.join(';'))]);
  return tiffBlock({
    ifd0,
    exif: [[0x9003, 2, '2026:10:27 09:15:00'], [0x829d, 5, [[18, 10]]], [0x8827, 3, 100], [0xa434, 2, 'Wide 26mm']],
    gps: [[0x0001, 2, 'N'], [0x0002, 5, [[51, 1], [30, 1], [2650, 100]]], [0x0003, 2, 'W'], [0x0004, 5, [[0, 1], [7, 1], [3999, 100]]]],
  });
}

/** An XMP packet with a rating and keywords, as Windows Photos writes one. */
export function xmpPacket({ rating, keywords }) {
  const items = keywords.map((k) => `<rdf:li>${k.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</rdf:li>`).join('');
  return `<?xpacket begin="" id="W5M0MpCehiHzreSzNTczkc9d"?><x:xmpmeta xmlns:x="adobe:ns:meta/"><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">`
    + `<rdf:Description rdf:about="" xmlns:xmp="http://ns.adobe.com/xap/1.0/"><xmp:Rating>${rating}</xmp:Rating></rdf:Description>`
    + `<rdf:Description rdf:about="" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:subject><rdf:Bag>${items}</rdf:Bag></dc:subject></rdf:Description>`
    + '</rdf:RDF></x:xmpmeta><?xpacket end="w"?>';
}

/** A JPEG given the camera's block, and an XMP packet after it when one is given. */
export function withCamera(jpeg, { block = cameraBlock(), xmp = null } = {}) {
  let out = jpegWithExif(new Uint8Array(jpeg), block);
  if (xmp) {
    const payload = new Uint8Array([...enc('http://ns.adobe.com/xap/1.0/'), 0, ...enc(xmp)]);
    const segment = new Uint8Array([0xff, 0xe1, (payload.length + 2) >> 8, (payload.length + 2) & 0xff, ...payload]);
    // After the EXIF segment: SOI, then APPn segments in turn.
    let at = 2;
    while (out[at] === 0xff && out[at + 1] >= 0xe0 && out[at + 1] <= 0xef) at += 2 + ((out[at + 2] << 8) | out[at + 3]);
    const next = new Uint8Array(out.length + segment.length);
    next.set(out.subarray(0, at), 0);
    next.set(segment, at);
    next.set(out.subarray(at), at + segment.length);
    out = next;
  }
  return Buffer.from(out);
}
