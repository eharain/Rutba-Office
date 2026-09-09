// One icon per kind of file, drawn here and rasterised by Electron.
//
// After "make Rutba Office the default", every document, workbook and deck
// wore the suite's one mark, and a folder of mixed files read as one kind of
// file. A person reads the icon before the name — Word's blue W, Excel's
// green X — so each family gets its own: the app's colour, the app's glyph,
// on a page. The glyphs are the ones the launcher and the ribbons draw, from
// @rutba/office-ui, so the icon on the desktop is the icon in the window.
//
//   ../../node_modules/.bin/electron build/make-file-icons.js   (from apps/desktop)
//
// Writes resources/filetypes/<family>.ico (and a .png beside it to look at)
// and resources/apps/<app>.ico, the taskbar tile each window carries.
// electron-builder.config.cjs points each association at its family page.

import { app, BrowserWindow, nativeImage } from 'electron';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const out = path.resolve(here, '..', 'resources', 'filetypes');
const appsOut = path.resolve(here, '..', 'resources', 'apps');

/** The launcher's glyphs, as drawn in packages/office-ui/src/icons.js (24-unit box). */
const GLYPHS = {
  word: ['M6 3h8l4 4v14a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z', 'M14 3v5h4', 'M8 13h7M8 17h5'],
  sheets: ['M4 5a1 1 0 0 1 1-1h14a1 1 0 0 1 1 1v14a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1z', 'M4 9h16M4 14h16M10 4v16M15 4v16'],
  slides: ['M3 5a1 1 0 0 1 1-1h16a1 1 0 0 1 1 1v9a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1z', 'M12 15v5M8 20h8'],
  pictures: ['M4 5a1 1 0 0 1 1-1h14a1 1 0 0 1 1 1v14a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1z', 'm5 16 4-4 3 3 3.5-3.5L20 16', 'M15 8.5h.01'],
  video: ['M3 7a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z', 'm16 10 5-3v10l-5-3z'],
  image: ['M4 5a1 1 0 0 1 1-1h14a1 1 0 0 1 1 1v14a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1z', 'M9 4v16M4 9h16', 'm14 15 3-3 3 3'],
  calendar: ['M4 6a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2v13a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2z', 'M4 10h16', 'M8 2v4M16 2v4', 'M8 14h3M13 14h3M8 17h3'],
  contacts: ['M5 4a1 1 0 0 1 1-1h12a1 1 0 0 1 1 1v16a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1z', 'M12 12a3 3 0 1 0 0-6 3 3 0 0 0 0 6z', 'M7 18c1-2.5 3-3.5 5-3.5s4 1 5 3.5'],
  mail: ['M3 7a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z', 'm3.5 7.5 8.5 6 8.5-6'],
};

/** The app's colour, as theme.css has it; PDF gets a red of its own. */
const FAMILIES = {
  word: { colour: '#2b5fd9', glyph: 'word' },
  sheets: { colour: '#0f9d58', glyph: 'sheets' },
  slides: { colour: '#d9534f', glyph: 'slides' },
  pictures: { colour: '#7b5cd6', glyph: 'pictures' },
  video: { colour: '#c2408f', glyph: 'video' },
  mail: { colour: '#3b7de0', glyph: 'mail' },
  calendar: { colour: '#1a9f7a', glyph: 'calendar' },
  contacts: { colour: '#5d6d7e', glyph: 'contacts' },
  pdf: { colour: '#b3261e', text: 'PDF' },
};

/**
 * A page with a folded corner, and the app's tile over its lower-left
 * corner — the arrangement Office uses, because it is the one people know:
 * at 16 pixels it is a white page with a coloured mark, and the colour is
 * the family.
 */
function svgFor({ colour, glyph, text }) {
  const mark = glyph
    ? `<g transform="translate(46 146) scale(3.4)" fill="none" stroke="#ffffff" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">${GLYPHS[glyph].map((d) => `<path d="${d}"/>`).join('')}</g>`
    : `<text x="80" y="200" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif" font-weight="700" font-size="40" fill="#ffffff">${text}</text>`;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="256" height="256" viewBox="0 0 256 256">
  <path d="M68 16h92l52 52v168a12 12 0 0 1-12 12H68a12 12 0 0 1-12-12V28a12 12 0 0 1 12-12z" fill="#ffffff" stroke="#b9bfc9" stroke-width="5"/>
  <path d="M160 16l52 52h-42a10 10 0 0 1-10-10z" fill="#e3e7ee" stroke="#b9bfc9" stroke-width="5" stroke-linejoin="round"/>
  <rect x="24" y="124" width="112" height="112" rx="22" fill="${colour}"/>
  ${mark}
</svg>`;
}

/**
 * The apps themselves, for the taskbar: each window carries its app's tile,
 * so a Word window and a Worksheets window sit under different icons and
 * each app can be pinned on its own, as Office's can. The tile is the same
 * coloured square the file-type icons wear, filling the frame.
 */
const APPS = {
  word: { colour: '#2b5fd9', glyph: 'word' },
  sheets: { colour: '#0f9d58', glyph: 'sheets' },
  slides: { colour: '#d9534f', glyph: 'slides' },
  pictures: { colour: '#7b5cd6', glyph: 'pictures' },
  image: { colour: '#e08b2b', glyph: 'image' },
  video: { colour: '#c2408f', glyph: 'video' },
  mail: { colour: '#3b7de0', glyph: 'mail' },
  calendar: { colour: '#1a9f7a', glyph: 'calendar' },
  contacts: { colour: '#5d6d7e', glyph: 'contacts' },
};

function tileSvg({ colour, glyph }) {
  const paths = GLYPHS[glyph].map((d) => `<path d="${d}"/>`).join('');
  return `<svg xmlns="http://www.w3.org/2000/svg" width="256" height="256" viewBox="0 0 256 256">
  <rect x="8" y="8" width="240" height="240" rx="52" fill="${colour}"/>
  <g transform="translate(44 44) scale(7)" fill="none" stroke="#ffffff" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">${paths}</g>
</svg>`;
}

function icoFrom(source) {
  const sizes = [256, 48, 32, 24, 16];
  const images = sizes.map((size) => ({ size, bytes: (size === 256 ? source : source.resize({ width: size, height: size, quality: 'best' })).toPNG() }));
  const header = Buffer.alloc(6);
  header.writeUInt16LE(0, 0);
  header.writeUInt16LE(1, 2);
  header.writeUInt16LE(images.length, 4);
  let offset = 6 + images.length * 16;
  const directory = Buffer.concat(
    images.map((image) => {
      const entry = Buffer.alloc(16);
      entry.writeUInt8(image.size === 256 ? 0 : image.size, 0);
      entry.writeUInt8(image.size === 256 ? 0 : image.size, 1);
      entry.writeUInt8(0, 2);
      entry.writeUInt8(0, 3);
      entry.writeUInt16LE(1, 4);
      entry.writeUInt16LE(32, 6);
      entry.writeUInt32LE(image.bytes.length, 8);
      entry.writeUInt32LE(offset, 12);
      offset += image.bytes.length;
      return entry;
    })
  );
  return Buffer.concat([header, directory, ...images.map((i) => i.bytes)]);
}

app.whenReady().then(async () => {
  fs.mkdirSync(out, { recursive: true });
  fs.mkdirSync(appsOut, { recursive: true });
  // Drawn on a canvas inside a hidden page and handed back as PNG data — not
  // captured off the screen. A transparent window's capture fails inside
  // Chromium's compositor (UnknownVizError) whenever another Electron holds
  // the GPU, and a canvas keeps the transparency the icon needs anyway.
  const win = new BrowserWindow({ width: 256, height: 256, show: false, webPreferences: { backgroundThrottling: false } });
  await win.loadURL('data:text/html;charset=utf-8,<!doctype html><meta charset="utf-8"><canvas id="c" width="256" height="256"></canvas>');
  const watchdog = setTimeout(() => {
    console.error('gave up: a family took longer than thirty seconds');
    app.exit(2);
  }, 30000);

  const jobs = [
    ...Object.entries(FAMILIES).map(([family, spec]) => [out, family, svgFor(spec)]),
    ...Object.entries(APPS).map(([key, spec]) => [appsOut, key, tileSvg(spec)]),
  ];
  for (const [dir, family, svg] of jobs) {
    const dataUrl = await win.webContents.executeJavaScript(`new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => { const c = document.getElementById('c'); const ctx = c.getContext('2d'); ctx.clearRect(0, 0, 256, 256); ctx.drawImage(img, 0, 0, 256, 256); resolve(c.toDataURL('image/png')); };
      img.onerror = () => reject(new Error('the SVG did not load'));
      img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(${JSON.stringify(svg)});
    })`);
    const png = Buffer.from(dataUrl.split(',')[1], 'base64');

    fs.writeFileSync(path.join(dir, `${family}.png`), png);
    const ico = icoFrom(nativeImage.createFromBuffer(png));
    fs.writeFileSync(path.join(dir, `${family}.ico`), ico);
    console.log(`${family.padEnd(9)} ${(ico.length / 1024).toFixed(1).padStart(6)} KB`);
  }
  clearTimeout(watchdog);
  win.destroy();
  app.exit(0);

});
