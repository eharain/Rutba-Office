// Bytes held for a window to fetch by URL — an inline picture in a message, an
// attachment on its way to disk, a picture on a slide.
//
// Nothing ever let go of one. Every message opened held its pictures again,
// every attachment saved or opened held the file again, and the main process
// grew for as long as the suite ran. Each held blob now knows the window
// whose request made it, and goes when that window closes; one held under a
// group (the message being read, the attachment being fetched) also replaces
// that window's previous blob of the same group, which nothing draws any more.
//
// The owner is found without passing it through every service: the IPC
// layer runs each request inside `blobOwner.run(windowId, …)`, and a hold
// made anywhere under that call — however deep, across any await — reads it.

import { AsyncLocalStorage } from 'node:async_hooks';

/** The window whose request is being answered, for the blobs it is handed. */
export const blobOwner = new AsyncLocalStorage();

const blobs = new Map();
let blobSeq = 0;

/**
 * Hold bytes under a new id.
 *
 * A `group` names what the blob is for — the message being read — and its
 * `generation` which one: the next message read is a new generation, and
 * the window's blobs of an older generation in that group go. Several held
 * for one generation (a message's three pictures) keep each other.
 *
 * @param {{ group?: string, generation?: string }} [options]
 * @returns {{ id: string, size: number }}
 */
export function hold(bytes, type = 'application/octet-stream', name = '', { group = null, generation = null } = {}) {
  const owner = blobOwner.getStore() ?? null;
  if (group) {
    for (const [id, b] of blobs) if (b.owner === owner && b.group === group && b.generation !== generation) blobs.delete(id);
  }
  const id = `b${++blobSeq}`;
  // Bytes already in memory are kept as they are, not copied: a deck's
  // pictures were each held twice, once by the deck and once here.
  const buf = Buffer.isBuffer(bytes) ? bytes : bytes instanceof Uint8Array ? Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength) : Buffer.from(bytes);
  blobs.set(id, { bytes: buf, type, name, owner, group, generation });
  return { id, size: buf.length };
}

/** The held blob, or undefined once it has been let go. */
export function heldBlob(id) {
  return blobs.get(id);
}

export function release(id) {
  blobs.delete(id);
}

/** Everything a window was handed, when it closes. */
export function releaseOwner(owner) {
  for (const [id, b] of blobs) if (b.owner === owner) blobs.delete(id);
}

/** How much is held, for a check that it does not only grow. */
export function heldTotals() {
  let bytes = 0;
  for (const b of blobs.values()) bytes += b.bytes.length;
  return { count: blobs.size, bytes };
}
