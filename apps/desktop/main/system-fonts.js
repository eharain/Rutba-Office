// A font this computer already has, for the letters the PDF's own fonts lack.
//
// The PDF writer draws Latin text in the fourteen fonts every reader carries
// and needs a real font for anything else — Arabic, Hebrew, Greek, Cyrillic.
// None is shipped with the suite: the system's own is found, as Word finds
// one, among the faces each platform keeps that cover Arabic and Hebrew both,
// and whose licence allows embedding. Read once and kept.
import fs from 'node:fs';
import path from 'node:path';
import { readTrueType } from '@rutba/pdf/truetype';

const CANDIDATES = {
  win32: () => {
    const dir = path.join(process.env.WINDIR || 'C:\\Windows', 'Fonts');
    return ['arial.ttf', 'segoeui.ttf', 'tahoma.ttf', 'times.ttf'].map((f) => path.join(dir, f));
  },
  darwin: () => [
    '/System/Library/Fonts/Supplemental/Arial Unicode.ttf',
    '/Library/Fonts/Arial Unicode.ttf',
    '/System/Library/Fonts/Supplemental/Arial.ttf',
    '/System/Library/Fonts/Supplemental/Tahoma.ttf',
  ],
  linux: () => [
    '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
    '/usr/share/fonts/dejavu/DejaVuSans.ttf',
    '/usr/share/fonts/TTF/DejaVuSans.ttf',
    '/usr/share/fonts/truetype/freefont/FreeSans.ttf',
    '/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf',
  ],
};

let found;

/** The bytes of a system font that draws Arabic and Hebrew and may be embedded, or null. */
export function unicodeFont() {
  if (found !== undefined) return found;
  found = null;
  const list = (CANDIDATES[process.platform] || CANDIDATES.linux)();
  for (const file of list) {
    try {
      const bytes = fs.readFileSync(file);
      const font = readTrueType(bytes);
      // Beh and alef in Arabic, alef in Hebrew: a face that has those has the scripts.
      if (font.embeddable && font.glyphOf(0x0628) && font.glyphOf(0xFE91) && font.glyphOf(0x05D0)) { found = bytes; break; }
    } catch { /* not there, or not a font this reads */ }
  }
  return found;
}
