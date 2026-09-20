// The Pictures viewer on a large, mixed folder.
//
// The viewer crawled on a folder of hundreds of files with clips among them,
// and stuttered as tiles came and went: every tile decoded the whole file,
// every clip was a live <video>, nothing was drawn only where it could be
// seen, and moving on emptied the stage before the next picture arrived.
// This opens a folder of three hundred pictures with two clips, a sound
// file, a PDF and a folder inside it, and reads back what the window did.
//
// Everything here is generated: the pictures are gradients, the clips are a
// second of canvas encoded in one of its windows, and nothing comes from the
// machine the run happens to be on. Run alone with RUTBA_VERIFY_ONLY=viewer.

import fs from 'node:fs';
import path from 'node:path';
import { gradientPng } from './sample-picture.js';

/**
 * A second of colour: forty VP8 frames encoded in a window, muxed here.
 *
 * MediaRecorder did this at first and captured nothing whenever another
 * window was open — a canvas nobody composites hands its stream no frames,
 * and the WebM came out as a header with no pictures in it, playable by
 * nothing. WebCodecs takes the frames straight from the canvas, so the
 * clip is the same forty frames on every machine, painted or not; the
 * container is written here, the minimum Matroska Chromium and the shell
 * both read.
 */
async function encodeFrames(webContents) {
  const chunks = await webContents.executeJavaScript(`(async () => {
    if (typeof VideoEncoder === 'undefined') return null;
    const c = document.createElement('canvas'); c.width = 320; c.height = 240;
    const g = c.getContext('2d');
    const out = [];
    const enc = new VideoEncoder({
      output: (chunk) => { const b = new Uint8Array(chunk.byteLength); chunk.copyTo(b); out.push({ key: chunk.type === 'key', ts: chunk.timestamp, b64: btoa(String.fromCharCode(...b)) }); },
      error: () => {},
    });
    enc.configure({ codec: 'vp8', width: 320, height: 240, bitrate: 600000, framerate: 25 });
    for (let t = 0; t < 40; t++) {
      g.fillStyle = 'hsl(' + ((t * 9) % 360) + ', 70%, 50%)'; g.fillRect(0, 0, 320, 240);
      g.fillStyle = '#fff'; g.fillRect((t * 7) % 280, 100, 40, 40);
      const frame = new VideoFrame(c, { timestamp: t * 40000 });
      enc.encode(frame, { keyFrame: t % 10 === 0 });
      frame.close();
    }
    await enc.flush(); enc.close();
    return out;
  })()`);
  return chunks ? chunks.map((c) => ({ key: c.key, ms: Math.round(c.ts / 1000), bytes: Buffer.from(c.b64, 'base64') })) : null;
}

/** Matroska's variable-size integers, written at a fixed eight bytes so nothing here has to be measured twice. */
const ebmlSize = (n) => { const out = Buffer.alloc(8); out.writeBigUInt64BE(BigInt(n)); out[0] = 0x01; return out; };
const ebml = (id, ...payload) => { const body = Buffer.concat(payload); return Buffer.concat([Buffer.from(id), ebmlSize(body.length), body]); };
const ebmlUint = (id, n) => { const b = Buffer.alloc(8); b.writeBigUInt64BE(BigInt(n)); let i = 0; while (i < 7 && b[i] === 0) i++; return ebml(id, b.subarray(i)); };
const ebmlFloat = (id, d) => { const b = Buffer.alloc(8); b.writeDoubleBE(d); return ebml(id, b); };
const ebmlStr = (id, str) => ebml(id, Buffer.from(str, 'utf8'));

/** A WebM of VP8 frames: one track, one cluster, every frame a SimpleBlock. */
export function muxWebm({ width, height, frames }) {
  const header = ebml([0x1a, 0x45, 0xdf, 0xa3],
    ebmlUint([0x42, 0x86], 1), ebmlUint([0x42, 0xf7], 1), ebmlUint([0x42, 0xf2], 4), ebmlUint([0x42, 0xf3], 8),
    ebmlStr([0x42, 0x82], 'webm'), ebmlUint([0x42, 0x87], 2), ebmlUint([0x42, 0x85], 2));
  const last = frames[frames.length - 1];
  const info = ebml([0x15, 0x49, 0xa9, 0x66], ebmlUint([0x2a, 0xd7, 0xb1], 1000000), ebmlFloat([0x44, 0x89], last.ms + 40), ebmlStr([0x4d, 0x80], 'rutba-checks'), ebmlStr([0x57, 0x41], 'rutba-checks'));
  const tracks = ebml([0x16, 0x54, 0xae, 0x6b], ebml([0xae],
    ebmlUint([0xd7], 1), ebmlUint([0x73, 0xc5], 1), ebmlUint([0x83], 1), ebmlStr([0x86], 'V_VP8'),
    ebml([0xe0], ebmlUint([0xb0], width), ebmlUint([0xba], height))));
  const blocks = frames.map((fr) => {
    const head = Buffer.from([0x81, (fr.ms >> 8) & 0xff, fr.ms & 0xff, fr.key ? 0x80 : 0x00]);
    return ebml([0xa3], head, fr.bytes);
  });
  const cluster = ebml([0x1f, 0x43, 0xb6, 0x75], ebmlUint([0xe7], 0), ...blocks);
  const segment = ebml([0x18, 0x53, 0x80, 0x67], info, tracks, cluster);
  return Buffer.concat([header, segment]);
}

/** The clips, encoded in `webContents` and written into the folder. */
async function recordClips(root, webContents) {
  const clips = {};
  try {
    const frames = await encodeFrames(webContents);
    if (process.env.RUTBA_VERIFY_DEBUG) console.log(`     [debug] encoded ${frames ? frames.length : 'no'} frames, ${frames ? frames.reduce((n, f) => n + f.bytes.length, 0) : 0} bytes`);
    if (frames && frames.length >= 10) {
      const bytes = muxWebm({ width: 320, height: 240, frames });
      fs.writeFileSync(path.join(root, 'frame.webm'), bytes);
      clips['frame.webm'] = true;
      // The same Matroska bytes under an extension no platform has a
      // thumbnail handler for: Chromium plays it by its content, so this
      // is the clip whose frame the window must draw itself.
      fs.writeFileSync(path.join(root, 'frame.ogv'), bytes);
      clips['frame.ogv'] = true;
    }
  } catch {
    /* a machine with no VP8 encoder: the check says so */
  }
  return clips;
}

/** A JPEG that says it was shot a quarter turn clockwise (EXIF orientation 6). */
async function sidewaysJpeg(png) {
  const { nativeImage } = await import('electron');
  const base = nativeImage.createFromBuffer(png).toJPEG(85);
  const tiff = Buffer.alloc(26);
  tiff.write('II', 0);
  tiff.writeUInt16LE(42, 2);
  tiff.writeUInt32LE(8, 4);
  tiff.writeUInt16LE(1, 8); // one entry
  tiff.writeUInt16LE(0x0112, 10); // Orientation
  tiff.writeUInt16LE(3, 12); // SHORT
  tiff.writeUInt32LE(1, 14);
  tiff.writeUInt16LE(6, 18); // a quarter turn clockwise
  tiff.writeUInt32LE(0, 22); // no next IFD
  const payload = Buffer.concat([Buffer.from('Exif\0\0'), tiff]);
  const app1 = Buffer.concat([Buffer.from([0xff, 0xe1, (payload.length + 2) >> 8, (payload.length + 2) & 0xff]), payload]);
  return Buffer.concat([base.subarray(0, 2), app1, base.subarray(2)]);
}

/** The folder: 300 pictures, a sound, a PDF, a sideways photo and a folder inside; the clips come after, from a window. */
export async function makeViewerFolder(dir, wav) {
  const root = path.join(dir, 'viewer');
  fs.mkdirSync(path.join(root, 'sub'), { recursive: true });
  const shades = [
    gradientPng(640, 400, [16, 74, 92], [126, 214, 214]),
    gradientPng(400, 640, [14, 26, 56], [86, 110, 178]),
    gradientPng(640, 400, [122, 92, 52], [232, 214, 178]),
  ];
  for (let i = 1; i <= 300; i++) fs.writeFileSync(path.join(root, `img-${String(i).padStart(3, '0')}.png`), shades[i % 3]);
  fs.writeFileSync(path.join(root, 'sub', 'inside-1.png'), shades[0]);
  fs.writeFileSync(path.join(root, 'sub', 'inside-2.png'), shades[2]);
  if (wav) fs.copyFileSync(wav, path.join(root, 'tone.wav'));
  fs.writeFileSync(path.join(root, 'notes.pdf'), '%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj\n3 0 obj<</Type/Page/Parent 2 0 R/MediaBox[0 0 200 200]>>endobj\ntrailer<</Root 1 0 R>>\n%%EOF');
  fs.writeFileSync(path.join(root, 'sideways.jpg'), await sidewaysJpeg(gradientPng(480, 300, [104, 32, 62], [226, 154, 176])));
  // One picture newer than the rest, for the order.
  const soon = new Date(Date.now() + 3600 * 1000);
  fs.utimesSync(path.join(root, 'img-150.png'), soon, soon);
  return { root, clips: {}, first: path.join(root, 'img-001.png'), sideways: path.join(root, 'sideways.jpg') };
}

/**
 * @param {object} h the harness: open, check, until, wait, js, clickIn, errorsIn, capture
 */
export async function verifyViewer(h, { dir, wav }) {
  const { open, check, until, wait, errorsIn, capture } = h;
  // A button by what it says or what its tip says, pressed the way a person presses it.
  const clickIn = async (win, title) => {
    const find = `[...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(title)}) && !n.disabled)`;
    await until(() => win.webContents.executeJavaScript(`Boolean(${find})`), `the ${title} button`, 3000).catch(() => {});
    return win.webContents.executeJavaScript(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(title)}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
  };
  const thumbOf = (p) => `rutba://thumb/${Buffer.from(p).toString('base64url')}?s=256`;
  let fixture;
  try {
    fixture = await makeViewerFolder(dir, wav);
  } catch (err) {
    return check('viewer: the folder could be made', false, err.message);
  }

  try {
    // The clips are encoded in a window of the suite's own, then the folder
    // is opened afresh so its listing has them.
    const studio = await open('pictures', fixture.first);
    fixture.clips = await recordClips(fixture.root, studio.webContents);
    const win = await open('pictures', fixture.first);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.pv-image')) && document.querySelector('.pv-image').naturalWidth > 0`), 'the first picture', 8000);

    // Escape: the folder has the whole window.
    await clickIn(win, 'Back to the folder');
    await until(() => js(`Boolean(document.querySelector('.pv-grid-scroll.browse'))`), 'the folder grid', 5000);
    await until(() => js(`document.querySelectorAll('.pv-tile').length > 10`), 'the tiles', 5000);
    const grid = await js(`(() => {
      const s = document.querySelector('.pv-grid-scroll');
      const imgs = [...document.querySelectorAll('.pv-tile img')];
      return { drawn: Number(s.dataset.drawn), count: Number(s.dataset.count), tiles: document.querySelectorAll('.pv-tile').length,
        scrollHeight: s.scrollHeight, clientHeight: s.clientHeight, videos: document.querySelectorAll('.pv-grid video').length,
        fromThumbHost: imgs.length > 0 && imgs.every((i) => i.src.startsWith('rutba://thumb/')), folders: document.querySelectorAll('.pv-tile.folder').length,
        status: document.querySelector('.rw-status')?.textContent || '' };
    })()`);
    const expected = 300 + 1 + 1 + 1 + Object.keys(fixture.clips).length + 1; // pictures, sideways, tone, pdf, clips, the folder
    check('viewer: a folder of three hundred files draws only the tiles in view, and the folder inside it first',
      grid.count === expected && grid.drawn < 120 && grid.tiles === grid.drawn && grid.scrollHeight > grid.clientHeight * 2 && grid.folders === 1,
      `${grid.drawn} of ${grid.count} drawn (expected ${expected}), ${grid.folders} folder, scroll ${grid.scrollHeight} in ${grid.clientHeight}`);
    check('viewer: every tile is the platform\'s thumbnail — never the file itself, never a live video',
      grid.fromThumbHost && grid.videos === 0, `${grid.videos} video elements; from the thumb host: ${grid.fromThumbHost}`);

    const answered = await until(() => js(`document.querySelectorAll('.pv-thumb.loaded').length >= 24`), 'the thumbnails to arrive', 20000).catch(() => false);
    const loaded = await js(`({ loaded: document.querySelectorAll('.pv-thumb.loaded').length, natural: [...document.querySelectorAll('.pv-thumb.loaded')].slice(0, 5).map((i) => i.naturalWidth) })`);
    check('viewer: the thumbnails arrive, small', answered === true && loaded.natural.every((w) => w > 0 && w <= 256), `${loaded.loaded} loaded, widths ${loaded.natural.join(',')}`);

    // Scrolling far down draws a different window of tiles, the same size.
    await js(`document.querySelector('.pv-grid-scroll').scrollTo(0, 5000)`);
    await wait(400);
    const deep = await js(`(() => { const s = document.querySelector('.pv-grid-scroll'); const names = [...document.querySelectorAll('.pv-tile')].map((t) => t.dataset.name || t.title); return { drawn: Number(s.dataset.drawn), first: names[0], top: s.scrollTop }; })()`);
    check('viewer: scrolled deep into the folder, a different window of tiles is drawn, no larger', deep.top > 4000 && deep.drawn < 120 && deep.first !== 'sub' && /img-\d+/.test(deep.first || ''), `${deep.drawn} drawn from ${deep.first} at ${deep.top}`);
    await js(`document.querySelector('.pv-grid-scroll').scrollTo(0, 0)`);

    // Videos: the MP4's frame comes from the platform; the WebM's, which
    // Windows has no thumbnail for, is drawn by the window and handed back.
    await js(`document.querySelector('.pv-kind[data-family="video"]')?.click()`);
    const clipNames = Object.keys(fixture.clips);
    const t0 = Date.now();
    const timeline = [];
    const framed = await until(async () => {
      const now = await js(`[...document.querySelectorAll('.pv-tile[data-kind="video"]')].map((t) => t.dataset.name + ':' + ((t.querySelector('img')?.src || 'icon').split(':')[0]) + (t.querySelector('img.loaded') ? '+' : '-')).join(' ')`);
      if (timeline[timeline.length - 1]?.state !== now) timeline.push({ at: Date.now() - t0, state: now });
      return (await js(`document.querySelectorAll('.pv-tile[data-kind="video"] img.pv-thumb.loaded').length`)) === clipNames.length;
    }, 'the clips to get frames', 45000).catch(() => false);
    if (process.env.RUTBA_VERIFY_DEBUG) console.log(`     [debug] clips ${clipNames.join(',')}; timeline ${timeline.map((e) => `${e.at}ms ${e.state}`).join(' | ')}`);
    const clips = await js(`[...document.querySelectorAll('.pv-tile[data-kind="video"]')].map((t) => ({ name: t.dataset.name, src: (t.querySelector('img')?.src || '').split(':')[0], w: t.querySelector('img')?.naturalWidth || 0 }))`);
    if (process.env.RUTBA_VERIFY_DEBUG) {
      const ogv = path.join(fixture.root, 'frame.ogv');
      const probe = await js(`(async () => {
        const img = document.querySelector('.pv-tile[data-name="frame.ogv"] img');
        const play = await new Promise((resolve) => {
          const v = document.createElement('video'); v.muted = true; v.preload = 'metadata';
          const timer = setTimeout(() => resolve('timeout'), 6000);
          v.onerror = () => { clearTimeout(timer); resolve('error ' + (v.error && v.error.code) + ' ' + (v.error && v.error.message)); };
          v.onloadedmetadata = () => { clearTimeout(timer); resolve('metadata ' + v.videoWidth + 'x' + v.videoHeight + ' ' + v.duration); };
          v.src = 'rutba://file/' + ${JSON.stringify(Buffer.from(ogv).toString('base64url'))};
        });
        return { img: img ? { src: img.src.slice(0, 40), complete: img.complete, w: img.naturalWidth, cls: img.className } : null, play };
      })()`);
      console.log(`     [debug] frame.ogv: ${JSON.stringify(probe)}`);
    }
    const mp4 = clips.find((c) => c.name === 'clip.mp4');
    const drawn = clips.filter((c) => c.src === 'blob');
    // A frame the window drew is handed to the platform, which serves it from then on.
    const kept = drawn.length ? await js(`fetch(${JSON.stringify(thumbOf(path.join(fixture.root, drawn[0].name)))}).then((r) => r.status).catch((e) => String(e))`) : 'none needed';
    // The WebM and its twin are always made; the MP4 needs an H.264 encoder
    // the machine may not lend a hidden window, and is checked when it is there.
    check('viewer: every clip\'s tile is a frame of it — the platform\'s when it has one, the window\'s own when it has not, and that one is kept for next time',
      framed === true && clipNames.length >= 2 && clips.length === clipNames.length && clips.every((c) => c.w > 0 && (c.src === 'rutba' || c.src === 'blob')) && (!mp4 || mp4.src === 'rutba') && (!drawn.length || kept === 200),
      `${clips.map((c) => `${c.name}: ${c.src === 'blob' ? 'drawn by the window' : 'from the platform'} (${c.w}px)`).join('; ')}; kept by the platform: ${kept}`);
    const chips = await js(`[...document.querySelectorAll('.pv-kind')].map((b) => b.textContent.replace(/\\s+/g, ' ').trim() + (b.getAttribute('aria-pressed') === 'true' ? '*' : ''))`);
    check('viewer: the kind chips count the folder and the pressed one filters it', chips.some((c) => c === `Videos${clipNames.length}*`) && chips.some((c) => c === `All${303 + clipNames.length}`) && chips.some((c) => c === 'Pictures301') && chips.some((c) => c === 'Audio1') && chips.some((c) => c === 'PDF1'), chips.join(' | '));
    await js(`document.querySelector('.pv-kind[data-family="all"]')?.click()`);

    // Newest first: the picture touched an hour into the future leads.
    await js(`(() => { const s = document.querySelector('select[data-role="sort"]'); s.value = 'newest'; s.dispatchEvent(new Event('change', { bubbles: true })); })()`);
    const newest = await until(() => js(`document.querySelector('.pv-tile:not(.folder)')?.dataset.name === 'img-150.png'`), 'the newest first', 4000).catch(() => false);
    await js(`(() => { const s = document.querySelector('select[data-role="sort"]'); s.value = 'name'; s.dispatchEvent(new Event('change', { bubbles: true })); })()`);
    const firstByName = [...clipNames, 'img-001.png'].sort()[0];
    const byName = await until(() => js(`document.querySelector('.pv-tile:not(.folder)')?.dataset.name === ${JSON.stringify(firstByName)}`), 'names again', 4000).catch(() => false);
    check('viewer: the order can be newest first, and name again', newest === true && byName === true, `newest ${newest}, by name ${byName}`);

    // Larger tiles: fewer per row, each wider.
    const before = await js(`document.querySelector('.pv-tile').getBoundingClientRect().width`);
    await js(`document.querySelector('.pv-size[data-size="l"]').click()`);
    const larger = await until(() => js(`document.querySelector('.pv-tile').getBoundingClientRect().width > ${before} + 30`), 'larger tiles', 4000).catch(() => false);
    const large = await js(`document.querySelector('.pv-tile').getBoundingClientRect().width`);
    await js(`document.querySelector('.pv-size[data-size="m"]').click()`);
    check('viewer: the tiles can be made larger', larger === true, `${before} → ${large}`);

    await until(() => js(`document.querySelectorAll('.pv-thumb.loaded').length >= 24`), 'the thumbnails again', 10000).catch(() => {});
    await wait(500);
    await capture(win, 'viewer-grid.png');

    // Into the folder inside, and up again.
    await js(`document.querySelector('.pv-tile.folder').click()`);
    const inside = await until(() => js(`document.querySelector('.pv-grid-scroll')?.dataset.count === '2' && /sub$/.test(document.querySelector('.pv-browse-name')?.textContent || '')`), 'the folder inside', 5000).catch(() => false);
    await clickIn(win, 'Up a folder');
    const upAgain = await until(() => js(`Number(document.querySelector('.pv-grid-scroll')?.dataset.count) === ${expected}`), 'the folder above', 5000).catch(() => false);
    check('viewer: a folder inside is entered from its tile, and Up comes back', inside === true && upAgain === true, `inside ${inside}, up ${upAgain}`);

    const complaints = await errorsIn(win);
    check('viewer: the folder reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('viewer: the folder checks ran', false, err.message);
  }

  try {
    const win = await open('pictures', fixture.sideways);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.pv-image')) && document.querySelector('.pv-image').naturalWidth > 0`), 'the sideways picture', 8000);
    const turned = await js(`(() => { const i = document.querySelector('.pv-image'); return { w: i.naturalWidth, h: i.naturalHeight, transform: i.style.transform, orientation: document.querySelector('.pv-details')?.textContent.includes('quarter turn') }; })()`);
    check('viewer: a phone\'s sideways photo is turned once — by the decoder, not again by the view', turned.w < turned.h && !/rotate\((?!0deg)/.test(turned.transform) && turned.orientation === true, JSON.stringify(turned));

    // Next: the stage always holds a decoded picture — the last one until
    // the next has decoded — and the details panel stays where it is. The
    // scrub bar first, to a run of pictures: the sideways photo's neighbours
    // by name are a PDF and a sound file.
    await js(`(() => { const s = document.querySelector('.pv-scrub'); const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; set.call(s, 100); s.dispatchEvent(new Event('input', { bubbles: true })); })()`);
    await until(() => js(`/^101 \\//.test(document.querySelector('.pv-counter')?.textContent || '') && document.querySelector('.pv-image')?.naturalWidth > 0`), 'the scrub to move', 6000);
    const from = await js(`document.querySelector('.pv-image')?.alt || ''`);
    const samples = [];
    await clickIn(win, 'Next');
    for (let i = 0; i < 16; i++) {
      samples.push(await js(`(() => { const i = document.querySelector('.pv-image'); return { has: Boolean(i), complete: i ? i.naturalWidth > 0 : false, panels: document.querySelectorAll('.rw-panel.right').length, src: i ? i.src.slice(-30) : '' }; })()`));
      await wait(25);
    }
    const moved = await until(() => js(`/^102 \\//.test(document.querySelector('.pv-counter')?.textContent || '') && /img-\\d+\\.png/.test(document.querySelector('.pv-image')?.alt || '') && document.querySelector('.pv-image')?.alt !== ${JSON.stringify(from)}`), 'the counter and the next picture', 4000).catch(() => false);
    const to = await js(`document.querySelector('.pv-image')?.alt || ''`);
    const steady = samples.every((s) => s.has && s.complete && s.panels === 1);
    check('viewer: moving on never shows an empty frame, and the details panel stays put', steady && moved === true, `${from} → ${to}; ${samples.filter((s) => !(s.has && s.complete)).length} empty of ${samples.length}, panels ${[...new Set(samples.map((s) => s.panels))].join('/')}`);

    const strip = await js(`(() => { const s = document.querySelector('.pv-strip'); return s ? { drawn: Number(s.dataset.drawn), count: Number(s.dataset.count), items: s.querySelectorAll('.pv-strip-item').length, videos: s.querySelectorAll('video').length, active: s.querySelectorAll('.pv-strip-item.active').length } : null; })()`);
    check('viewer: the filmstrip is drawn only where it can be seen, and holds no live video', Boolean(strip) && strip.count > 300 && strip.drawn < 60 && strip.items === strip.drawn && strip.videos === 0 && strip.active === 1, JSON.stringify(strip));

    await wait(500);
    await capture(win, 'viewer-stage.png');
    const complaints = await errorsIn(win);
    check('viewer: the stage reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('viewer: the stage checks ran', false, err.message);
  }
  return undefined;
}
