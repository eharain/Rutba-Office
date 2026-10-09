// EXIF, and the one field that changes what you see.
//
// Orientation is the reason this exists. A phone writes every photograph in the
// sensor's own orientation and records how to turn it; a viewer that ignores
// the tag shows a third of a holiday sideways. The rest — camera, lens,
// exposure, when and where — is what the Details panel shows.

const TAGS = {
  0x010f: 'make',
  0x0110: 'model',
  0x0112: 'orientation',
  0x011a: 'xResolution',
  0x011b: 'yResolution',
  0x0131: 'software',
  0x0132: 'dateTime',
  0x013b: 'artist',
  0x8298: 'copyright',
  0x829a: 'exposureTime',
  0x829d: 'fNumber',
  0x8827: 'iso',
  0x9003: 'dateTimeOriginal',
  0x9004: 'dateTimeDigitized',
  0x920a: 'focalLength',
  0x9209: 'flash',
  0xa002: 'pixelWidth',
  0xa003: 'pixelHeight',
  0xa405: 'focalLength35mm',
  0xa434: 'lens',
  0x8825: '@gps',
  0x8769: '@exif',
};

const GPS_TAGS = {
  0x0001: 'latRef',
  0x0002: 'lat',
  0x0003: 'lonRef',
  0x0004: 'lon',
  0x0006: 'altitude',
};

/** How the image should be turned, as the operations a canvas would apply. */
export const ORIENTATION = {
  1: { rotate: 0, flip: false, label: 'as shot' },
  2: { rotate: 0, flip: true, label: 'mirrored' },
  3: { rotate: 180, flip: false, label: 'upside down' },
  4: { rotate: 180, flip: true, label: 'mirrored, upside down' },
  5: { rotate: 90, flip: true, label: 'mirrored, quarter turn' },
  6: { rotate: 90, flip: false, label: 'quarter turn clockwise' },
  7: { rotate: 270, flip: true, label: 'mirrored, three-quarter turn' },
  8: { rotate: 270, flip: false, label: 'quarter turn anticlockwise' },
};

function readIfd(view, start, base, little, into, tagMap = TAGS, depth = 0) {
  if (depth > 3 || start + 2 > view.byteLength) return into;
  const count = view.getUint16(start, little);
  for (let i = 0; i < count; i++) {
    const at = start + 2 + i * 12;
    if (at + 12 > view.byteLength) break;
    const tag = view.getUint16(at, little);
    const type = view.getUint16(at + 2, little);
    const length = view.getUint32(at + 4, little);
    const name = tagMap[tag];
    if (!name) continue;

    const sizes = { 1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 7: 1, 9: 4, 10: 8 };
    const size = (sizes[type] || 1) * length;
    const offset = size > 4 ? base + view.getUint32(at + 8, little) : at + 8;
    if (offset + Math.min(size, 4) > view.byteLength) continue;

    if (name === '@exif' || name === '@gps') {
      const sub = base + (size > 4 ? view.getUint32(at + 8, little) : view.getUint32(at + 8, little));
      readIfd(view, sub, base, little, into, name === '@gps' ? GPS_TAGS : TAGS, depth + 1);
      continue;
    }

    let value = null;
    if (type === 2) {
      const bytes = new Uint8Array(view.buffer, view.byteOffset + offset, Math.min(size, view.byteLength - offset));
      value = new TextDecoder('utf-8', { fatal: false }).decode(bytes).replace(/\0+$/, '').trim();
    } else if (type === 3) {
      value = view.getUint16(offset, little);
    } else if (type === 4) {
      value = view.getUint32(offset, little);
    } else if (type === 5 || type === 10) {
      const parts = [];
      for (let n = 0; n < Math.min(length, 3); n++) {
        const numerator = type === 5 ? view.getUint32(offset + n * 8, little) : view.getInt32(offset + n * 8, little);
        const denominator = type === 5 ? view.getUint32(offset + n * 8 + 4, little) : view.getInt32(offset + n * 8 + 4, little);
        parts.push(denominator ? numerator / denominator : 0);
      }
      value = length > 1 ? parts : parts[0];
    }
    if (value !== null && value !== '') into[name] = value;
  }
  return into;
}

/**
 * Read the EXIF block of a JPEG or TIFF.
 * @returns {object} the fields that were found; `{}` when there are none
 */
export function readExif(bytes) {
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  let tiffAt = -1;

  if (b[0] === 0xff && b[1] === 0xd8) {
    // Find APP1, which holds "Exif\0\0" then a whole TIFF header.
    let at = 2;
    while (at + 4 < b.length) {
      if (b[at] !== 0xff) break;
      const marker = b[at + 1];
      const length = (b[at + 2] << 8) | b[at + 3];
      if (marker === 0xe1 && String.fromCharCode(b[at + 4], b[at + 5], b[at + 6], b[at + 7]) === 'Exif') {
        tiffAt = at + 10;
        break;
      }
      if (marker === 0xda) break; // start of scan: no EXIF above this
      at += 2 + length;
    }
  } else if ((b[0] === 0x49 && b[1] === 0x49) || (b[0] === 0x4d && b[1] === 0x4d)) {
    tiffAt = 0;
  }

  if (tiffAt < 0 || tiffAt + 8 > b.length) return {};
  const view = new DataView(b.buffer, b.byteOffset, b.byteLength);
  const little = b[tiffAt] === 0x49;
  const ifd = view.getUint32(tiffAt + 4, little);
  const out = readIfd(view, tiffAt + ifd, tiffAt, little, {});

  if (out.lat && out.latRef) {
    const toDegrees = (parts) => (Array.isArray(parts) ? parts[0] + (parts[1] || 0) / 60 + (parts[2] || 0) / 3600 : parts);
    out.latitude = toDegrees(out.lat) * (String(out.latRef).toUpperCase() === 'S' ? -1 : 1);
    out.longitude = toDegrees(out.lon) * (String(out.lonRef).toUpperCase() === 'W' ? -1 : 1);
  }
  return out;
}

/** The details panel's rows: only what is there, phrased the way a camera would. */
export function describeExif(exif) {
  const rows = [];
  const add = (label, value) => value != null && value !== '' && rows.push({ label, value: String(value) });

  add('Camera', [exif.make, exif.model].filter(Boolean).join(' ').trim());
  add('Lens', exif.lens);
  if (exif.exposureTime) {
    const t = exif.exposureTime;
    add('Exposure', t >= 1 ? `${t}s` : `1/${Math.round(1 / t)}s`);
  }
  if (exif.fNumber) add('Aperture', `f/${Math.round(exif.fNumber * 10) / 10}`);
  add('ISO', exif.iso);
  if (exif.focalLength) add('Focal length', `${Math.round(exif.focalLength)}mm${exif.focalLength35mm ? ` (${exif.focalLength35mm}mm equivalent)` : ''}`);
  const taken = exif.dateTimeOriginal || exif.dateTime;
  if (taken) add('Taken', String(taken).replace(/^(\d{4}):(\d{2}):(\d{2})/, '$1-$2-$3'));
  if (exif.latitude != null) add('Location', `${exif.latitude.toFixed(5)}, ${exif.longitude.toFixed(5)}`);
  add('Software', exif.software);
  add('Artist', exif.artist);
  add('Copyright', exif.copyright);
  if (exif.orientation && exif.orientation !== 1) add('Orientation', ORIENTATION[exif.orientation]?.label);
  return rows;
}

/** The orientation a quarter turn clockwise more makes, whatever turn and mirror one says already. */
export const TURN_CLOCKWISE = { 1: 6, 6: 3, 3: 8, 8: 1, 2: 7, 7: 4, 4: 5, 5: 2 };

/**
 * A JPEG with its EXIF orientation set to `orientation`, its pixels left
 * exactly as they were — the lossless turn Windows Photos makes. The tag is
 * rewritten where the file has one; a file with no EXIF gets a small block
 * holding only it, after its JFIF header. Null where the file is not a JPEG
 * or has EXIF without an orientation tag, which cannot be added without
 * moving everything after it.
 */
export function withOrientation(bytes, orientation) {
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (!(orientation >= 1 && orientation <= 8)) throw new Error('an orientation is 1 to 8');
  if (b.length < 4 || b[0] !== 0xff || b[1] !== 0xd8) return null;
  let at = 2;
  let afterHeader = 2;
  while (at + 4 <= b.length && b[at] === 0xff) {
    const marker = b[at + 1];
    const length = (b[at + 2] << 8) | b[at + 3];
    if (length < 2) return null;
    if (marker === 0xe0 && at === afterHeader) afterHeader = at + 2 + length;
    if (marker === 0xe1 && at + 10 <= b.length && String.fromCharCode(b[at + 4], b[at + 5], b[at + 6], b[at + 7]) === 'Exif') {
      const tiff = at + 10;
      const end = Math.min(b.length, at + 2 + length);
      if (tiff + 8 > end) return null;
      const view = new DataView(b.buffer, b.byteOffset, b.byteLength);
      const little = b[tiff] === 0x49;
      const ifd = tiff + view.getUint32(tiff + 4, little);
      if (ifd + 2 > end) return null;
      const count = view.getUint16(ifd, little);
      for (let i = 0; i < count; i++) {
        const entry = ifd + 2 + i * 12;
        if (entry + 12 > end) return null;
        if (view.getUint16(entry, little) !== 0x0112) continue;
        if (view.getUint16(entry + 2, little) !== 3) return null;
        const out = new Uint8Array(b);
        new DataView(out.buffer).setUint16(entry + 8, orientation, little);
        return out;
      }
      return null;
    }
    if (marker === 0xda) break;
    at += 2 + length;
  }
  // No EXIF at all: one holding the orientation and nothing else.
  const tiff = new Uint8Array(26);
  const tv = new DataView(tiff.buffer);
  tiff.set([0x49, 0x49], 0);
  tv.setUint16(2, 42, true);
  tv.setUint32(4, 8, true);
  tv.setUint16(8, 1, true);
  tv.setUint16(10, 0x0112, true);
  tv.setUint16(12, 3, true);
  tv.setUint32(14, 1, true);
  tv.setUint16(18, orientation, true);
  tv.setUint32(22, 0, true);
  const payload = new Uint8Array(6 + tiff.length);
  payload.set([0x45, 0x78, 0x69, 0x66, 0, 0], 0);
  payload.set(tiff, 6);
  const app1 = new Uint8Array(4 + payload.length);
  app1.set([0xff, 0xe1, (payload.length + 2) >> 8, (payload.length + 2) & 0xff], 0);
  app1.set(payload, 4);
  const out = new Uint8Array(b.length + app1.length);
  out.set(b.subarray(0, afterHeader), 0);
  out.set(app1, afterHeader);
  out.set(b.subarray(afterHeader), afterHeader + app1.length);
  return out;
}

/** What the viewer must do to show the image the right way up. */
export function orientationOf(exif) {
  return ORIENTATION[exif?.orientation] || ORIENTATION[1];
}
