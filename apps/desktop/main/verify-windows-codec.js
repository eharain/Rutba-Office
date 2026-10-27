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
import { NAME_OF } from './verify-pictures-tiff.js';

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

/**
 * Pictures shows what only the main process can draw (2.21): an EMF and a
 * WMF drawn as SVG by the picture helper the application lends the file
 * host, and a camera RAW through Windows' own codecs, the JPEG inside it at
 * the screen's size; a HEIC that is not a picture is said in a sentence.
 * Run alone with RUTBA_VERIFY_ONLY=pictures-drawn.
 * @param {object} h the harness: open, check, until, wait, press, errorsIn
 * @param {{ dir: string }} args
 */
export async function verifyPicturesDrawn({ open, check, until, wait, press, errorsIn }, { dir }) {
  const root = path.join(dir, 'pictures-drawn');
  try {
    fs.mkdirSync(root, { recursive: true });
    const { showcaseEmf, showcaseWmf } = await import('../../../tests/fixtures/metafiles.js');
    const shot = noisyPicture(3000, 2000, 5);
    const files = [
      ['1 chart.emf', showcaseEmf()],
      ['2 clip art.wmf', showcaseWmf()],
      ['3 camera.dng', dngWithPreview(shot.toJPEG(92), { width: 3000, height: 2000, orientation: 1 })],
    ];
    for (const [name, bytes] of files) fs.writeFileSync(path.join(root, name), bytes);
    const win = await open('pictures', path.join(root, files[0][0]));
    const js = (code) => win.webContents.executeJavaScript(code);
    const seen = [];
    for (let k = 0; k < files.length; k++) {
      const name = files[k][0];
      const shown = `(() => { ${NAME_OF} const i = document.querySelector('.pv-image'); return i && i.complete && nameOf(i.src) === ${JSON.stringify(name)} ? { w: i.naturalWidth, h: i.naturalHeight } : null; })()`;
      let size = null;
      await until(async () => { size = await js(shown); return Boolean(size?.w); }, `${name} drawn`, 20000).catch(() => {});
      seen.push({ name, size });
      if (k + 1 < files.length) {
        win.focus();
        await js(`document.querySelector('.pv-stage')?.focus?.(), 1`);
        await press(win.webContents, 'Right');
      }
    }
    const [emf, wmf, raw] = seen;
    check('pictures: an EMF and a WMF open drawn, at the size each drawing says',
      emf.size?.w > 100 && emf.size?.h > 50 && wmf.size?.w > 100 && wmf.size?.h > 50, JSON.stringify(seen.slice(0, 2)));
    check('pictures: a camera RAW opens as the picture its camera put inside it, brought to 2048 on its long side',
      raw.size?.w === 2048 && raw.size?.h === 1365, JSON.stringify(raw));
    // Their tiles in the filmstrip: Windows has no thumbnail for a metafile,
    // so the window draws one from the SVG it is served (frames.js); a tile
    // left as the picture icon, or drawn blank, fails.
    const tileOf = (name) => `(() => {
      const item = [...document.querySelectorAll('.pv-strip-item')].find((b) => b.title === ${JSON.stringify(name)});
      const img = item?.querySelector('img.pv-thumb.loaded');
      if (!img || !img.complete || !img.naturalWidth) return { drawn: false, icon: Boolean(item?.querySelector('.pv-tile-icon')) };
      let colours = null;
      try {
        const c = document.createElement('canvas');
        c.width = 64; c.height = 64;
        const g = c.getContext('2d');
        g.drawImage(img, 0, 0, 64, 64);
        const d = g.getImageData(0, 0, 64, 64).data;
        const seen = new Set();
        for (let i = 0; i < d.length; i += 4) seen.add((d[i] >> 4) + ',' + (d[i + 1] >> 4) + ',' + (d[i + 2] >> 4));
        colours = seen.size;
      } catch {}
      return { drawn: true, w: img.naturalWidth, colours };
    })()`;
    const tiles = {};
    for (const name of [files[0][0], files[1][0]]) {
      await until(async () => (tiles[name] = await js(tileOf(name)))?.drawn, `${name}'s tile`, 15000).catch(() => {});
    }
    check('pictures: an EMF\'s and a WMF\'s tiles show their drawings, not a blank tile or the picture icon',
      Object.values(tiles).length === 2 && Object.values(tiles).every((t) => t?.drawn && (t.colours == null || t.colours > 2)), JSON.stringify(tiles));
    // What a window gets for a HEIC Windows cannot read: a sentence naming what to add.
    const fake = path.join(root, 'not a picture.heic');
    fs.writeFileSync(fake, 'words, not a picture');
    const said = await js(`fetch(${JSON.stringify('rutba://file/')} + ${JSON.stringify(Buffer.from(fake, 'utf8').toString('base64url'))}).then(async (r) => ({ status: r.status, text: await r.text() }), (e) => ({ error: String(e) }))`);
    check('pictures: a HEIC Windows cannot read is answered with what to add from the Microsoft Store, not a broken picture',
      said?.status === 415 && /HEIF Image Extensions/.test(said.text || ''), JSON.stringify(said));
    const complaints = await errorsIn(win);
    check('pictures: those pictures report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    win.close();
  } catch (err) {
    check('pictures: the drawn-picture checks ran', false, err?.message || String(err));
  }
}
