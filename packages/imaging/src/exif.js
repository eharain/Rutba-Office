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

/** What the viewer must do to show the image the right way up. */
export function orientationOf(exif) {
  return ORIENTATION[exif?.orientation] || ORIENTATION[1];
}
