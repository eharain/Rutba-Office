// The application icon, drawn here rather than imported.
//
// Electron is already in the toolchain and already has a rasteriser, so the
// icon is an SVG in this file rendered by the app's own engine — no design tool
// in the build, nothing binary in the repository that cannot be regenerated,
// and one place to change the mark.
//
//   node build/make-icon.js
//
// electron-builder turns the 512px PNG into the .ico and .icns the installers
// need, so this writes one file.

import { app, BrowserWindow, nativeImage } from 'electron';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const resources = path.resolve(here, '..', 'resources');

/**
 * The mark: a document corner over the suite's blue, with the fold that says
 * "file" and a rule beneath that says "office". It has to read at 16px, which
 * is why it is three shapes and not a scene.
 */
const SVG = `<svg xmlns="http://www.w3.org/2000/svg" width="512" height="512" viewBox="0 0 512 512">
  <defs>
    <linearGradient id="bg" x1="0" y1="0" x2="0.35" y2="1">
      <stop offset="0" stop-color="#3f74e8"/>
      <stop offset="1" stop-color="#1f47b0"/>
    </linearGradient>
    <linearGradient id="fold" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#ffffff" stop-opacity="0.95"/>
      <stop offset="1" stop-color="#d7e2fb" stop-opacity="0.9"/>
    </linearGradient>
  </defs>
  <rect x="0" y="0" width="512" height="512" rx="112" fill="url(#bg)"/>
  <path d="M150 108h132l82 82v214a14 14 0 0 1-14 14H150a14 14 0 0 1-14-14V122a14 14 0 0 1 14-14z" fill="#ffffff" opacity="0.96"/>
  <path d="M282 108l82 82h-68a14 14 0 0 1-14-14z" fill="url(#fold)"/>
  <rect x="180" y="248" width="152" height="18" rx="9" fill="#2b5fd9" opacity="0.9"/>
  <rect x="180" y="298" width="112" height="18" rx="9" fill="#2b5fd9" opacity="0.6"/>
  <rect x="180" y="348" width="132" height="18" rx="9" fill="#2b5fd9" opacity="0.35"/>
</svg>`;

app.whenReady().then(async () => {
  fs.mkdirSync(resources, { recursive: true });

  const win = new BrowserWindow({
    width: 512,
    height: 512,
    show: false,
    frame: false,
    transparent: true,
    backgroundColor: '#00000000',
    webPreferences: { offscreen: false },
  });

  const page = `<!doctype html><meta charset="utf-8">
<style>html,body{margin:0;padding:0;background:transparent;width:512px;height:512px;overflow:hidden}</style>
${SVG}`;
  await win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(page)}`);
  await new Promise((r) => setTimeout(r, 350));

  const image = await win.webContents.capturePage();
  const png = image.toPNG();
  fs.writeFileSync(path.join(resources, 'icon.png'), png);

  // A 256px copy as well: Windows uses it for the taskbar and the installer.
  const source = nativeImage.createFromBuffer(png);
  const small = source.resize({ width: 256, height: 256, quality: 'best' });
  fs.writeFileSync(path.join(resources, 'icon-256.png'), small.toPNG());

  fs.writeFileSync(path.join(resources, 'icon.svg'), SVG);

  // The installer needs a real .ico, not a PNG with the wrong extension.
  //
  // An ICO is a tiny directory followed by its images, and since Vista those
  // images may be PNGs — so this is a header, one 16-byte entry per size, and
  // the PNG bytes Electron just produced at each of them. Every size is present
  // because Windows picks one per context: 16 in the title bar, 32 in the task
  // switcher, 256 on the desktop.
  const sizes = [256, 128, 64, 48, 32, 24, 16];
  const images = sizes.map((size) => ({
    size,
    bytes: (size === 256 ? small : source.resize({ width: size, height: size, quality: 'best' })).toPNG(),
  }));

  const header = Buffer.alloc(6);
  header.writeUInt16LE(0, 0); // reserved
  header.writeUInt16LE(1, 2); // 1 = icon
  header.writeUInt16LE(images.length, 4);

  let offset = 6 + images.length * 16;
  const directory = Buffer.concat(
    images.map((image) => {
      const entry = Buffer.alloc(16);
      entry.writeUInt8(image.size === 256 ? 0 : image.size, 0); // 0 means 256
      entry.writeUInt8(image.size === 256 ? 0 : image.size, 1);
      entry.writeUInt8(0, 2); // palette size: none, it is truecolour
      entry.writeUInt8(0, 3); // reserved
      entry.writeUInt16LE(1, 4); // colour planes
      entry.writeUInt16LE(32, 6); // bits per pixel
      entry.writeUInt32LE(image.bytes.length, 8);
      entry.writeUInt32LE(offset, 12);
      offset += image.bytes.length;
      return entry;
    })
  );

  const ico = Buffer.concat([header, directory, ...images.map((i) => i.bytes)]);
  fs.writeFileSync(path.join(resources, 'icon.ico'), ico);
  console.log(`icon.ico  ${(ico.length / 1024).toFixed(1)} KB — ${sizes.join(', ')} px`);

  console.log(`icon.png ${(png.length / 1024).toFixed(1)} KB — ${image.getSize().width}x${image.getSize().height}`);
  win.destroy();
  app.exit(0);
});
