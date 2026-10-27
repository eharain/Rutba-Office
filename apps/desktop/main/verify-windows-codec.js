// Pictures through Windows' own codecs (2.21 C5), in the main process where
// the protocol will call it: a camera RAW (a DNG made here, its raw data
// nothing a codec can develop) shows the full-size JPEG its camera put
// inside it, at the screen's size and turned upright by the RAW's
// orientation; a HEIC that is not a picture is null, never a throw. Run
// alone with RUTBA_VERIFY_ONLY=windows-codec.
import fs from 'node:fs';
import path from 'node:path';
import { nativeImage } from 'electron';
import { decodeWithSystem } from '@rutba/office-shell/electron/windows-codec';
import { noisyPicture } from './verify-image-edits.js';

/** A DNG-shaped file whose IFD0 is the camera's preview: one JPEG strip, the orientation given. */
function dngWithPreview(jpeg, { width, height, orientation }) {
  const entries = [[0x00fe, 4, 1, 1], [0x0100, 3, 1, width], [0x0101, 3, 1, height], [0x0103, 3, 1, 7], [0x0106, 3, 1, 6],
    [0x0111, 4, 1, -1], [0x0112, 3, 1, orientation], [0x0115, 3, 1, 3], [0x0117, 4, 1, jpeg.length], [0xc612, 1, 4, 0x00000401]];
  const at = 8;
  const head = Buffer.alloc(at + 2 + entries.length * 12 + 4);
  head.write('II', 0, 'latin1');
  head.writeUInt16LE(42, 2);
  head.writeUInt32LE(at, 4);
  head.writeUInt16LE(entries.length, at);
  entries.forEach(([tag, type, count, value], i) => {
    const p = at + 2 + i * 12;
    head.writeUInt16LE(tag, p);
    head.writeUInt16LE(type, p + 2);
    head.writeUInt32LE(count, p + 4);
    const v = value === -1 ? head.length : value;
    if (type === 3) head.writeUInt16LE(v, p + 8);
    else head.writeUInt32LE(v, p + 8);
  });
  return Buffer.concat([head, jpeg]);
}

/** The average colour of a 16 × 16 patch of a PNG at a fraction of its width and height. */
function patch(png, fx, fy) {
  const image = nativeImage.createFromBuffer(png);
  const { width, height } = image.getSize();
  const bgra = image.toBitmap();
  const x0 = Math.min(width - 16, Math.max(0, Math.round(width * fx) - 8));
  const y0 = Math.min(height - 16, Math.max(0, Math.round(height * fy) - 8));
  const s = [0, 0, 0];
  for (let y = y0; y < y0 + 16; y++) for (let x = x0; x < x0 + 16; x++) { const i = (y * width + x) * 4; s[0] += bgra[i + 2]; s[1] += bgra[i + 1]; s[2] += bgra[i]; }
  return s.map((v) => Math.round(v / 256));
}

/**
 * @param {object} h the harness: check
 * @param {{ dir: string }} args where the files are made
 */
export async function verifyWindowsCodec({ check }, { dir }) {
  const root = path.join(dir, 'windows-codec');
  try {
    fs.mkdirSync(root, { recursive: true });
    // Blue at the top left, orange at the bottom right, green at the bottom left; 3000 × 2000.
    const shot = noisyPicture(3000, 2000, 5);
    const dng = path.join(root, 'camera.dng');
    fs.writeFileSync(dng, dngWithPreview(shot.toJPEG(92), { width: 3000, height: 2000, orientation: 6 }));
    const started = Date.now();
    const png = await decodeWithSystem(dng, 2048);
    const ms = Date.now() - started;
    const size = png ? nativeImage.createFromBuffer(png).getSize() : null;
    // A quarter turn clockwise: the green bottom left is now the top left, the blue top left the top right.
    const topLeft = png ? patch(png, 0.02, 0.02) : null;
    const topRight = png ? patch(png, 0.98, 0.02) : null;
    const green = topLeft && topLeft[1] > topLeft[2] && topLeft[1] > 120;
    const blue = topRight && topRight[2] > topRight[0] + 60;
    check('pictures: a camera RAW with no codec shows the JPEG its camera put inside it, at 2048 on its long side, turned upright by its orientation',
      size?.width === 1365 && size.height === 2048 && green && blue && ms < 15000,
      JSON.stringify({ size, topLeft, topRight, ms }));
    const fake = path.join(root, 'not-a-picture.heic');
    fs.writeFileSync(fake, 'words, not a picture');
    const none = await decodeWithSystem(fake, 2048).then((v) => v, (err) => `threw ${err?.message}`);
    check('pictures: a HEIC Windows cannot decode is null from the system decoder, never a throw', none === null, String(none));
  } catch (err) {
    check('pictures: the Windows codec checks ran', false, err?.message || String(err));
  }
}
