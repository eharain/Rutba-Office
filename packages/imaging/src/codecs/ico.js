// Windows icons and cursors: a directory of pictures of one thing at
// several sizes, each a PNG or a DIB with its see-through mask. The pages
// are its sizes, largest (and deepest) first, so page 0 is the one to show.

import { bad, reader } from './bytes.js';
import { decodePng, isPng } from './png.js';
import { decodeDib } from './bmp.js';

export function isIco(b) {
  return b.length >= 22 && b[0] === 0 && b[1] === 0 && (b[2] === 1 || b[2] === 2) && b[3] === 0 && (b[4] | (b[5] << 8)) > 0;
}

/** The icon's entries, best first. */
function entries(b) {
  const r = reader(b, b[2] === 2 ? 'cursor' : 'icon');
  if (!isIco(b)) bad('This is not an icon or a cursor.');
  const count = r.u16(4);
  if (count > 512 || 6 + count * 16 > b.length) bad('This icon\'s list of sizes is damaged.');
  const list = [];
  for (let i = 0; i < count; i++) {
    const at = 6 + i * 16;
    const size = r.u32(at + 8);
    const offset = r.u32(at + 12);
    if (offset >= b.length || size < 8) continue;
    list.push({
      width: b[at] || 256,
      height: b[at + 1] || 256,
      bpp: b[2] === 1 ? r.u16(at + 6) : 0,
      offset,
      size: Math.min(size, b.length - offset),
    });
  }
  if (!list.length) bad('This icon holds no pictures: it is damaged.');
  list.sort((x, y) => y.width * y.height - x.width * x.height || y.bpp - x.bpp);
  return list;
}

export function icoInfo(b) {
  const list = entries(b);
  return { format: b[2] === 2 ? 'cur' : 'ico', width: list[0].width, height: list[0].height, pages: list.length };
}

export function decodeIco(b, { page = 0, limits } = {}) {
  const list = entries(b);
  const kind = b[2] === 2 ? 'cursor' : 'icon';
  if (page >= list.length) bad(list.length === 1 ? `This ${kind} has one size only; there is no size ${page + 1}.` : `This ${kind} has ${list.length} sizes; there is no size ${page + 1}.`);
  const e = list[page];
  const body = b.subarray(e.offset, e.offset + e.size);
  const pic = isPng(body) ? decodePng(body, limits) : decodeDib(body, 0, { icon: true, limits, kind });
  return { width: pic.width, height: pic.height, rgba: pic.rgba, pages: list.length };
}
