// The system clipboard, from the main process.
//
// Electron 44 reshaped its clipboard on the W3C Clipboard API: every read and
// write is a Promise, read() and write() take ClipboardItems keyed by MIME
// type, and the narrow helpers (readHTML, readImage, writeImage, ...) are
// gone. These keep the shapes the suite reads and writes - words and a table
// together, a picture as a nativeImage - on top of that API, for the
// clipboard service in ipc.js and for the checks that hold the person's
// clipboard aside while they run.

import { clipboard, ClipboardItem, nativeImage } from 'electron';

const PICTURE_TYPES = ['image/png', 'image/jpeg'];

/** The payload of the first item that carries one of `types`, as a Buffer, or null. */
async function readFirst(types) {
  const items = await clipboard.read();
  for (const type of types) {
    const item = items.find((i) => i.types.includes(type));
    if (item) return Buffer.from(await (await item.getType(type)).arrayBuffer());
  }
  return null;
}

/** The clipboard's HTML, or '' when it holds none. */
export async function readHTML() {
  const bytes = await readFirst(['text/html']);
  return bytes ? bytes.toString('utf8') : '';
}

/** The clipboard's picture as a nativeImage, empty when it holds none. */
export async function readImage() {
  const bytes = await readFirst(PICTURE_TYPES);
  return bytes ? nativeImage.createFromBuffer(bytes) : nativeImage.createEmpty();
}

/**
 * Words, HTML and a picture in one write that replaces what was there, as
 * the old clipboard.write({ text, html, image }) did. A part left out is not
 * written; text '' is written as empty words.
 */
export async function writeParts({ text, html, image } = {}) {
  const parts = {};
  if (text != null) parts['text/plain'] = text;
  if (html) parts['text/html'] = html;
  if (image && !image.isEmpty()) parts['image/png'] = new Blob([image.toPNG()], { type: 'image/png' });
  await clipboard.write([new ClipboardItem(parts)]);
}

/** A nativeImage on the clipboard, alone. */
export function writeImage(image) {
  return writeParts({ image });
}
