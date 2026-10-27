// Pictures and Image with the formats Chromium does not draw.
//
//   pictures-tiff: Pictures opens a TIFF of three pages, shows its page
//     strip, and steps through the pages by the strip's buttons, Page Down
//     and a page's own thumbnail, each page drawn as itself; and a fax
//     (CCITT Group 4) opens.
//   image-formats: Image opens a chosen page of the TIFF, steps to the next,
//     and exports BMP, ICO and TIFF that read back as the picture.
//   pictures-formats: Pictures opens QOI, PPM, Targa, PCX, Radiance HDR and a
//     cursor in turn, each drawn as itself with its size in the details, and
//     the folder's grid draws a thumbnail of each.
//   image-open: Image opens those formats, and an SVG at the size chosen in
//     the dialog it asks first.
//
// Run alone with RUTBA_VERIFY_ONLY=pictures-tiff,image-formats,pictures-formats,image-open.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { dialog } from 'electron';
import { encodeTiff } from '@rutba/imaging/encode';
import { decodePicture } from '@rutba/imaging/decode';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FAX = path.resolve(HERE, '../../../tests/fixtures/pictures/wic-bw-ccitt4.tif');

/** Three pages, each of its own size and colour: a soft gradient, a pale band and a dark disc. */
export const PAGES = [
  { width: 640, height: 420, colour: [196, 64, 48] },
  { width: 420, height: 600, colour: [40, 150, 86] },
  { width: 700, height: 360, colour: [44, 86, 200] },
];

export function pageRgba({ width, height, colour }) {
  const rgba = new Uint8Array(width * height * 4);
  const cx = width / 2;
  const cy = height / 2;
  const r = Math.min(width, height) * 0.28;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const i = (y * width + x) * 4;
      const k = 0.65 + 0.35 * (x / width);
      let [R, G, B] = colour.map((c) => Math.round(c * k + 255 * (1 - k) * 0.4));
      if (y > height * 0.08 && y < height * 0.16) [R, G, B] = [245, 240, 228];
      if ((x - cx) ** 2 + (y - cy) ** 2 < r * r) [R, G, B] = colour.map((c) => Math.round(c * 0.45));
      rgba[i] = R;
      rgba[i + 1] = G;
      rgba[i + 2] = B;
      rgba[i + 3] = 255;
    }
  }
  return rgba;
}

/** The three-page TIFF, written by the suite's own writer. */
export function threePageTiff() {
  return encodeTiff(PAGES.map((p) => ({ rgba: pageRgba(p), width: p.width, height: p.height })), { compression: 'lzw' });
}

const near = (a, b, tol = 10) => a.every((v, i) => Math.abs(v - b[i]) <= tol);
const discColour = (p) => p.colour.map((c) => Math.round(c * 0.45));

const captureOf = (wait) => async (win, name) => {
  if (!process.env.RUTBA_VERIFY_CAPTURE) return;
  win.webContents.invalidate();
  await wait(700);
  fs.mkdirSync(process.env.RUTBA_VERIFY_CAPTURE, { recursive: true });
  fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await win.webContents.capturePage()).toPNG());
};

/** The colour at the centre of what an <img> shows, read from its address (the element itself is another origin). */
const CENTRE = (selector) => `(async () => {
  const img = document.querySelector(${JSON.stringify(selector)});
  if (!img || !img.naturalWidth) return null;
  const bmp = await createImageBitmap(await (await fetch(img.src)).blob());
  const c = new OffscreenCanvas(bmp.width, bmp.height);
  const g = c.getContext('2d');
  g.drawImage(bmp, 0, 0);
  const d = g.getImageData(Math.floor(bmp.width / 2), Math.floor(bmp.height / 2), 1, 1).data;
  return { w: bmp.width, h: bmp.height, rgb: [d[0], d[1], d[2]] };
})()`;

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn
 * @param {{ dir: string }} args
 */
export async function verifyPicturesTiff({ open, check, until, wait, press, errorsIn }, { dir }) {
  const capture = captureOf(wait);
  const root = path.join(dir, 'tiff-pages');
  try {
    fs.mkdirSync(root, { recursive: true });
    const file = path.join(root, 'three pages.tif');
    fs.writeFileSync(file, threePageTiff());
    const win = await open('pictures', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const state = () => js(`(() => {
      const img = document.querySelector('.pv-image');
      return {
        w: img?.naturalWidth || 0, h: img?.naturalHeight || 0,
        count: document.querySelector('[data-role="page-count"]')?.textContent || '',
        thumbs: document.querySelectorAll('.pv-page').length,
        active: document.querySelector('.pv-page.active')?.dataset.page ?? null,
        drawn: [...document.querySelectorAll('.pv-page img')].filter((i) => i.complete && i.naturalWidth > 0).length,
      };
    })()`);
    const atPage = (n) => until(async () => {
      const s = await state();
      return s.w === PAGES[n].width && s.h === PAGES[n].height && s.count === `Page ${n + 1} of 3` && s.active === String(n);
    }, `page ${n + 1} on the stage`, 10000).then(() => true).catch(() => false);

    const first = await atPage(0);
    await until(async () => (await state()).drawn === 3, 'the page thumbnails', 10000).catch(() => {});
    const s1 = await state();
    const c1 = await js(CENTRE('.pv-image'));
    await capture(win, 'pictures-tiff-page1.png');
    check('pictures: a TIFF of three pages opens on its first page, with a strip of its three pages drawn under it',
      first && s1.thumbs === 3 && s1.drawn === 3 && c1 && near(c1.rgb, discColour(PAGES[0])),
      JSON.stringify({ ...s1, centre: c1?.rgb }));

    // The strip's Next page button.
    await js(`(() => { const b = [...document.querySelectorAll('.pv-pages .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Next page')); b?.click(); return Boolean(b); })()`);
    const second = await atPage(1);
    const c2 = await js(CENTRE('.pv-image'));
    // The details give the size of the page shown, not the file's first page.
    const sizeSaid = `${PAGES[1].width} × ${PAGES[1].height}`;
    const sized = await until(() => js(`(() => { const dd = [...document.querySelectorAll('dt')].find((d) => d.textContent === 'Dimensions')?.nextElementSibling; return (dd?.textContent || '').startsWith(${JSON.stringify(sizeSaid)}); })()`), 'the page size in the details', 5000).then(() => true).catch(() => false);
    await capture(win, 'pictures-tiff-page2.png');
    // Page Down, from the keyboard.
    win.focus();
    await js(`document.querySelector('.pv-stage')?.focus?.(), 1`);
    await press(win.webContents, 'PageDown');
    const third = await atPage(2);
    const c3 = await js(CENTRE('.pv-image'));
    // A page's own thumbnail.
    await js(`document.querySelector('.pv-page[data-page="0"]')?.click(), 1`);
    const back = await atPage(0);
    check('pictures: the next-page button, Page Down and a page\'s thumbnail each put that page on the stage, drawn as itself, its own size in the details',
      second && sized && third && back && c2 && near(c2.rgb, discColour(PAGES[1])) && c3 && near(c3.rgb, discColour(PAGES[2])),
      JSON.stringify({ second, sized, third, back, c2: c2?.rgb, c3: c3?.rgb, now: await state() }));

    // A fax: CCITT Group 4, as Windows' own encoder wrote it.
    if (fs.existsSync(FAX)) {
      const fax = path.join(root, 'fax.tif');
      fs.copyFileSync(FAX, fax);
      const fw = await open('pictures', fax);
      const fjs = (code) => fw.webContents.executeJavaScript(code);
      const faxShown = await until(() => fjs(`document.querySelector('.pv-image')?.naturalWidth === 203`), 'the fax', 10000).then(() => true).catch(() => false);
      const noStrip = await fjs(`!document.querySelector('.pv-pages')`);
      await capture(fw, 'pictures-tiff-fax.png');
      check('pictures: a fax (a TIFF in CCITT Group 4) opens, one page and no strip', faxShown && noStrip, JSON.stringify({ faxShown, noStrip }));
      const faxErrors = await errorsIn(fw);
      check('pictures: the fax reports nothing', faxErrors.length === 0, faxErrors.join(' | ') || 'nothing reported');
      fw.close();
    }

    const complaints = await errorsIn(win);
    check('pictures: the TIFF pages report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    win.close();
  } catch (err) {
    check('pictures: the TIFF page checks ran', false, err?.message || String(err));
  }
}

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn
 * @param {{ dir: string }} args
 */
export async function verifyImageFormats({ open, check, until, wait, press, errorsIn }, { dir }) {
  const capture = captureOf(wait);
  const root = path.join(dir, 'image-formats');
  const realSave = dialog.showSaveDialog;
  let saveTo = null;
  const offered = [];
  dialog.showSaveDialog = async (...args) => {
    const opts = args.find((a) => a && typeof a === 'object' && 'defaultPath' in a) || {};
    offered.push(opts);
    return saveTo ? { canceled: false, filePath: saveTo } : { canceled: true, filePath: '' };
  };
  try {
    fs.mkdirSync(root, { recursive: true });
    const file = path.join(root, 'three pages.tif');
    fs.writeFileSync(file, threePageTiff());
    // Opened as Pictures' Edit opens a page: the second.
    const win = await open('image', file, { page: '1' });
    const js = (code) => win.webContents.executeJavaScript(code);
    const result = () => js(`document.querySelector('.im-canvas-wrap')?.dataset.result || ''`);
    const count = () => js(`document.querySelector('[data-role="page-count"]')?.textContent || ''`);
    const size = (n) => `${PAGES[n].width}x${PAGES[n].height}`;
    const chosen = await until(async () => (await result()) === size(1) && (await count()) === 'Page 2 of 3', 'page 2 in Image', 10000).then(() => true).catch(() => false);
    await capture(win, 'image-tiff-page2.png');
    await js(`document.querySelector('[data-role="page-next"]')?.click(), 1`);
    const next = await until(async () => (await result()) === size(2) && (await count()) === 'Page 3 of 3', 'page 3 in Image', 10000).then(() => true).catch(() => false);
    check('image: a TIFF opens on the page asked for, and Next opens the page after it',
      chosen && next, JSON.stringify({ chosen, next, result: await result(), count: await count() }));

    // Export: BMP, ICO and TIFF, each read back as the picture it was.
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Export')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('[data-format="ico"]'))`), 'the Export tab', 4000).catch(() => {});
    await capture(win, 'image-export-formats.png');
    const written = {};
    for (const [format, ext] of [['bmp', 'bmp'], ['ico', 'ico'], ['tiff', 'tif']]) {
      await js(`document.querySelector('[data-format="${format}"]')?.click(), 1`);
      await until(() => js(`document.querySelector('[data-format="${format}"]')?.getAttribute('aria-pressed') === 'true' || document.querySelector('[data-format="${format}"]')?.classList.contains('pressed')`), `${format} chosen`, 3000).catch(() => {});
      await wait(250);
      saveTo = path.join(root, `page three.${ext}`);
      const before = offered.length;
      win.focus();
      await press(win.webContents, 'e', { modifiers: ['control'] });
      await until(() => offered.length > before && fs.existsSync(saveTo), `the ${format} file`, 10000).catch(() => {});
      await wait(400);
      let back = null;
      try {
        const bytes = fs.readFileSync(saveTo);
        const pic = decodePicture(bytes, saveTo);
        const i = (Math.floor(pic.height / 2) * pic.width + Math.floor(pic.width / 2)) * 4;
        back = { format: pic.format, w: pic.width, h: pic.height, centre: [pic.rgba[i], pic.rgba[i + 1], pic.rgba[i + 2]], pages: pic.pages, offered: path.extname(offered[before]?.defaultPath || '') };
        if (format === 'ico') back.sizes = Array.from({ length: bytes.readUInt16LE(4) }, (_, k) => bytes[6 + k * 16] || 256);
      } catch (err) {
        back = { error: err.message };
      }
      written[format] = back;
    }
    const want = PAGES[2];
    const disc = discColour(want);
    const ok = (b, w, h) => b && !b.error && b.w === w && b.h === h && near(b.centre, disc, 12);
    check('image: Export writes BMP, ICO and TIFF, offered under their own extensions, each reading back as the picture (the icon at 256 down to 16, the picture fitted in its square)',
      ok(written.bmp, want.width, want.height) && written.bmp.offered === '.bmp'
      && ok(written.tiff, want.width, want.height) && written.tiff.offered === '.tif'
      && ok(written.ico, 256, 256) && JSON.stringify(written.ico.sizes) === JSON.stringify([256, 128, 64, 48, 32, 24, 16]) && written.ico.offered === '.ico',
      JSON.stringify(written));

    const complaints = await errorsIn(win);
    check('image: the TIFF pages and the new formats report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    win.close();
  } catch (err) {
    check('image: the format checks ran', false, err?.message || String(err));
  } finally {
    dialog.showSaveDialog = realSave;
  }
}

/* ── the small formats: QOI, Netpbm, Targa, PCX, Radiance HDR, cursors; SVG at a size ── */

const SMALL = { width: 360, height: 240 };
const SMALL_COLOURS = [[200, 60, 120], [60, 140, 200], [230, 150, 30], [90, 170, 70], [140, 90, 210], [200, 200, 60]];

/** A folder of one picture in each small format, each its own colour, written by the suite's writers. */
async function smallFormats(root) {
  const { encodeQoi } = await import('../../../packages/imaging/src/codecs/qoi.js');
  const { encodePcx } = await import('../../../packages/imaging/src/codecs/pcx.js');
  const { encodeHdr } = await import('../../../packages/imaging/src/codecs/hdr.js');
  const { encodeIco } = await import('@rutba/imaging/encode');
  const { encodePng } = await import('@rutba/imaging/decode');
  const { pnmOf, tgaOf, lightOf } = await import('../../../tools/picture-sources.js');
  const pic = (k) => pageRgba({ ...SMALL, colour: SMALL_COLOURS[k] });
  const files = [
    ['1 quite ok.qoi', encodeQoi(pic(0), SMALL.width, SMALL.height), 'QOI'],
    ['2 netpbm.ppm', pnmOf(pic(1), SMALL.width, SMALL.height), 'PPM'],
    ['3 targa.tga', tgaOf(pic(2), SMALL.width, SMALL.height), 'TGA'],
    ['4 paintbrush.pcx', encodePcx(pic(3), SMALL.width, SMALL.height), 'PCX'],
    ['5 radiance.hdr', encodeHdr(lightOf(pic(4), SMALL.width, SMALL.height), SMALL.width, SMALL.height), 'HDR'],
  ];
  const cur = pageRgba({ width: 64, height: 64, colour: SMALL_COLOURS[5] });
  files.push(['6 pointer.cur', encodeIco([{ width: 64, height: 64, png: encodePng(cur, 64, 64) }], { cursor: true, hotspots: [{ x: 4, y: 4 }] }), 'CUR']);
  for (const [name, bytes] of files) fs.writeFileSync(path.join(root, name), bytes);
  return files.map(([name, , format], k) => ({ file: path.join(root, name), name, format, colour: SMALL_COLOURS[k], size: k === 5 ? [64, 64] : [SMALL.width, SMALL.height] }));
}

/** In a page: the file an address of the file host names (its path is base64). */
const NAME_OF = `const nameOf = (src) => { try { const seg = new URL(src).pathname.split('/').pop(); const b = seg.replace(/-/g, '+').replace(/_/g, '/'); return decodeURIComponent(escape(atob(b + '='.repeat((4 - (b.length % 4)) % 4)))).split(/[\\\\/]/).pop(); } catch { return ''; } };`;

const resultOf = (js) => js(`document.querySelector('.im-canvas-wrap')?.dataset.result || ''`);

/**
 * Pictures lists and opens QOI, Netpbm, Targa, PCX, Radiance HDR and cursors:
 * each drawn as itself, its details saying its size, the folder's grid
 * showing every one.
 * @param {object} h the harness: open, check, until, wait, press, errorsIn
 * @param {{ dir: string }} args
 */
export async function verifyPicturesFormats({ open, check, until, wait, press, errorsIn }, { dir }) {
  const capture = captureOf(wait);
  const root = path.join(dir, 'small-formats');
  try {
    fs.mkdirSync(root, { recursive: true });
    const files = await smallFormats(root);
    const win = await open('pictures', files[0].file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const seen = [];
    for (let k = 0; k < files.length; k++) {
      const f = files[k];
      const drawn = await until(() => js(`(() => { ${NAME_OF} const i = document.querySelector('.pv-image'); return Boolean(i && i.complete && i.naturalWidth === ${f.size[0]} && i.naturalHeight === ${f.size[1]} && nameOf(i.src) === ${JSON.stringify(f.name)}); })()`), `${f.name} drawn`, 10000).then(() => true).catch(() => false);
      const details = `(() => { const dd = (n) => [...document.querySelectorAll('dt')].find((d) => d.textContent === n)?.nextElementSibling?.textContent || ''; return { format: dd('Format'), size: dd('Dimensions') }; })()`;
      const sizeSaid = `${f.size[0]} × ${f.size[1]}`;
      await until(async () => (await js(details)).size.startsWith(sizeSaid), `${f.name} details`, 5000).catch(() => {});
      const said = await js(details);
      const c = await js(CENTRE('.pv-image'));
      // The Radiance picture is tone-mapped, so its colour is its own; the rest are their disc's colour.
      const colourOk = f.format === 'HDR' ? Boolean(c && c.rgb[2] > c.rgb[1]) : Boolean(c && near(c.rgb, f.colour.map((v) => Math.round(v * 0.45)), 12));
      seen.push({ name: f.name, drawn, colourOk, centre: c?.rgb, said });
      if (f.format === 'TGA' || f.format === 'HDR') await capture(win, `pictures-formats-${f.format.toLowerCase()}.png`);
      if (k + 1 < files.length) {
        win.focus();
        await js(`document.querySelector('.pv-stage')?.focus?.(), 1`);
        await press(win.webContents, 'Right');
      }
    }
    const allDrawn = seen.every((s) => s.drawn && s.colourOk);
    const allSaid = seen.every((s, k) => s.said?.size?.startsWith(`${files[k].size[0]} × ${files[k].size[1]}`));
    check('pictures: QOI, PPM, Targa, PCX, Radiance HDR (tone-mapped) and a cursor each open in turn, drawn as themselves, the details giving their size',
      allDrawn && allSaid, JSON.stringify(seen));
    // Back to the folder: a thumbnail for every one.
    await press(win.webContents, 'Escape');
    const want = JSON.stringify(files.map((f) => f.name));
    let thumbs = 0;
    await until(() => js(`(() => { const names = ${want}; const drawn = [...document.querySelectorAll('.pv-tile[data-name]')].filter((t) => { const i = t.querySelector('img'); return i && i.complete && i.naturalWidth > 0; }).map((t) => t.dataset.name); return names.filter((n) => drawn.includes(n)).length; })()`).then((n) => { thumbs = n; return n === files.length; }), 'the folder thumbnails', 10000).catch(() => {});
    await capture(win, 'pictures-formats-grid.png');
    check('pictures: the folder lists all six, each with its thumbnail drawn', thumbs === files.length, JSON.stringify({ thumbs }));
    const complaints = await errorsIn(win);
    check('pictures: the small formats report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    win.close();
  } catch (err) {
    check('pictures: the small-format checks ran', false, err?.message || String(err));
  }
}

const SVG = [
  '<svg xmlns="http://www.w3.org/2000/svg" width="120" height="80" viewBox="0 0 120 80">',
  '  <rect width="120" height="80" fill="#f4efe4"/>',
  '  <circle cx="60" cy="40" r="26" fill="#2c56c8"/>',
  '  <path d="M8 72 L112 8" stroke="#c43c30" stroke-width="3"/>',
  '</svg>',
].join('\n');

/**
 * Image opens the small formats, and an SVG drawn at the size chosen in the
 * dialog it asks first.
 * @param {object} h the harness: open, check, until, wait, press, errorsIn
 * @param {{ dir: string }} args
 */
export async function verifyImageOpen({ open, check, until, wait, errorsIn }, { dir }) {
  const capture = captureOf(wait);
  const root = path.join(dir, 'image-open');
  try {
    fs.mkdirSync(root, { recursive: true });
    const files = await smallFormats(root);
    const opened = {};
    for (const f of files.filter((x) => ['QOI', 'TGA', 'PCX', 'HDR', 'CUR'].includes(x.format))) {
      const win = await open('image', f.file);
      const js = (code) => win.webContents.executeJavaScript(code);
      opened[f.format] = await until(async () => (await resultOf(js)) === `${f.size[0]}x${f.size[1]}`, `${f.name} in Image`, 10000).then(() => true).catch(() => false);
      if (f.format === 'TGA') await capture(win, 'image-open-tga.png');
      const complaints = await errorsIn(win);
      if (complaints.length) opened[`${f.format} reported`] = complaints.join(' | ');
      win.close();
    }
    check('image: QOI, Targa, PCX, Radiance HDR and a cursor each open as a picture of their own size, reporting nothing',
      Object.values(opened).every((v) => v === true), JSON.stringify(opened));

    // SVG: the dialog asks the size first; 960 pixels wide draws it at 960 × 640.
    const svg = path.join(root, 'drawing.svg');
    fs.writeFileSync(svg, SVG);
    const win = await open('image', svg);
    const js = (code) => win.webContents.executeJavaScript(code);
    const asked = await until(() => js(`Boolean(document.querySelector('[data-role="resize-ok"]')) && document.body.textContent.includes('120 × 80')`), 'the size dialog', 10000).then(() => true).catch(() => false);
    await js(`(() => {
      const el = document.querySelector('[data-role="resize-width"]');
      const input = el && el.tagName === 'INPUT' ? el : el?.querySelector('input');
      if (!input) return false;
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, '960');
      input.dispatchEvent(new Event('input', { bubbles: true }));
      return true;
    })()`);
    await wait(200);
    await capture(win, 'image-svg-size.png');
    await js(`document.querySelector('[data-role="resize-ok"]')?.click(), 1`);
    const drawn = await until(async () => (await resultOf(js)) === '960x640', 'the SVG at 960', 10000).then(() => true).catch(() => false);
    await capture(win, 'image-svg-open.png');
    const result = await resultOf(js);
    check('image: an SVG asks the size to draw it at (saying its own, 120 × 80), and opens at the size chosen, 960 × 640', asked && drawn, JSON.stringify({ asked, drawn, result }));
    const complaints = await errorsIn(win);
    check('image: the small formats and the SVG report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    win.close();
  } catch (err) {
    check('image: the open checks ran', false, err?.message || String(err));
  }
}
