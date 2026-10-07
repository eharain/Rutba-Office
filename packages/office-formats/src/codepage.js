// 8-bit code pages, decoded the same everywhere.
//
// `new TextDecoder('windows-1252')` is not to be trusted: the encoding
// standard reads bytes 0x80-0x9F as the euro sign, the curly quotes, the
// dashes and their neighbours, and browsers do, but Node 24 — the Node in
// Electron's main process, where files are opened — reads them as the C1
// control characters with those codes. A Word document's "smart quotes" and
// dashes came out as invisible controls. So Windows-1252 is decoded here by
// table, as are the DOS code pages TextDecoder has never known; every other
// code page goes to TextDecoder.

// Windows-1252's 0x80-0x9F; the five it leaves undefined keep their own codes.
const CP1252_HIGH = [
  0x20ac, 0x81, 0x201a, 0x192, 0x201e, 0x2026, 0x2020, 0x2021, 0x2c6, 0x2030, 0x160, 0x2039, 0x152, 0x8d, 0x17d, 0x8f,
  0x90, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x2dc, 0x2122, 0x161, 0x203a, 0x153, 0x9d, 0x17e, 0x178,
];
const CP1252_BACK = new Map(CP1252_HIGH.map((c, i) => [c, 0x80 + i]));

// The upper halves of the DOS code pages Word for DOS wrote.
const DOS = {
  cp437: 'ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜ¢£¥₧ƒáíóúñÑªº¿⌐¬½¼¡«»░▒▓│┤╡╢╖╕╣║╗╝╜╛┐└┴┬├─┼╞╟╚╔╩╦╠═╬╧╨╤╥╙╘╒╓╫╪┘┌█▄▌▐▀αßΓπΣσµτΦΘΩδ∞φε∩≡±≥≤⌠⌡÷≈°∙·√ⁿ²■ ',
  cp850: 'ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜø£Ø×ƒáíóúñÑªº¿®¬½¼¡«»░▒▓│┤ÁÂÀ©╣║╗╝¢¥┐└┴┬├─┼ãÃ╚╔╩╦╠═╬¤ðÐÊËÈıÍÎÏ┘┌█▄¦Ì▀ÓßÔÒõÕµþÞÚÛÙýÝ¯´­±‗¾¶§÷¸°¨·¹³²■ ',
};

const WINDOWS_1252 = new Set(['windows-1252', 'cp1252', 'x-cp1252', 'ansi', 'latin1', 'iso-8859-1', 'us-ascii', 'ascii']);

/** Windows-1252 bytes as text. */
export function decode1252(bytes) {
  let out = '';
  for (let i = 0; i < bytes.length; i += 8192) {
    const chunk = bytes.subarray(i, i + 8192);
    const codes = new Array(chunk.length);
    for (let k = 0; k < chunk.length; k++) {
      const b = chunk[k];
      codes[k] = b >= 0x80 && b < 0xa0 ? CP1252_HIGH[b - 0x80] : b;
    }
    out += String.fromCharCode(...codes);
  }
  return out;
}

/** Text back to its Windows-1252 bytes (a character it has no byte for keeps its low byte). */
export function encode1252(text) {
  const out = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    out[i] = CP1252_BACK.get(c) ?? (c & 0xff);
  }
  return out;
}

const decoders = new Map();
/**
 * A decoder for a code page by name: { decode(bytes) → text }. Windows-1252
 * and the DOS pages by table, the rest by TextDecoder — and Windows-1252
 * for a name this runtime does not know.
 */
export function decoderFor(name) {
  const key = String(name || 'windows-1252').toLowerCase();
  if (!decoders.has(key)) {
    let d;
    if (WINDOWS_1252.has(key)) d = { decode: decode1252 };
    else if (DOS[key]) {
      const upper = DOS[key];
      d = { decode: (bytes) => { let s = ''; for (const b of bytes) s += b < 0x80 ? String.fromCharCode(b) : upper[b - 0x80]; return s; } };
    } else {
      try {
        const td = new TextDecoder(key);
        d = { decode: (bytes) => td.decode(bytes) };
      } catch {
        d = { decode: decode1252 };
      }
    }
    decoders.set(key, d);
  }
  return decoders.get(key);
}

/** Bytes in a code page as text. */
export const decodeIn = (name, bytes) => decoderFor(name).decode(bytes);
