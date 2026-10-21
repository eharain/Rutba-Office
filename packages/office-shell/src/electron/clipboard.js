// The clipboard, from the main process, through an adapter.
//
// The system adapter is the computer's own clipboard. A check run is given a
// stand-in (machine-fakes.js) so that check processes can run side by side
// without one's copy landing in another's paste, and without touching the
// person's clipboard; the system adapter is checked on its own, in the
// gate's machine lane (verify-machine.js).
//
// Electron 44 reshaped its clipboard on the W3C Clipboard API: every read and
// write is a Promise, read() and write() take ClipboardItems keyed by MIME
// type, and the narrow helpers (readHTML, readImage, writeImage, ...) are
// gone. The helpers below keep the shapes the suite reads and writes - words
// and a table together, a picture as a nativeImage - on top of whichever
// adapter is in use, for the clipboard service in ipc.js and for the checks.

import { clipboard as electronClipboard, ClipboardItem, nativeImage } from 'electron';
import { fakeClipboard } from './machine-fakes.js';

/** The computer's own clipboard. */
export const systemClipboard = {
  kind: 'system',
  read: () => electronClipboard.read(),
  write: (items) => electronClipboard.write(items),
  readText: () => electronClipboard.readText(),
  writeText: (text) => electronClipboard.writeText(text),
  clear: () => electronClipboard.clear(),
};

/** A clipboard of this process's own, its items Electron's ClipboardItems. */
export const memoryClipboard = () => fakeClipboard({ Item: ClipboardItem });

let current = systemClipboard;

/** Use `adapter` for every clipboard read and write from now on. */
export function useClipboard(adapter) {
  current = adapter || systemClipboard;
}

/** The adapter in use. */
export const clipboard = () => current;

const PICTURE_TYPES = ['image/png', 'image/jpeg'];

/** The payload of the first item that carries one of `types`, as a Buffer, or null. */
async function readFirst(types) {
  const items = await current.read();
  for (const type of types) {
    const item = items.find((i) => i.types.includes(type));
    if (item) return Buffer.from(await (await item.getType(type)).arrayBuffer());
  }
  return null;
}

/** The clipboard's words, or ''. */
export const readText = () => current.readText();

/** Words on the clipboard, alone. */
export const writeText = (text) => current.writeText(text ?? '');

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
  await current.write([new ClipboardItem(parts)]);
}

/** A nativeImage on the clipboard, alone. */
export function writeImage(image) {
  return writeParts({ image });
}

/** Everything on the clipboard, to put back later (a check holding the person's clipboard aside). */
export const keep = () => current.read().catch(() => []);

/** Put back what keep() answered. */
export async function putBack(items) {
  try {
    if (items?.length) await current.write(items);
    else await current.clear();
  } catch { /* the clipboard is busy: nothing more to do */ }
}
