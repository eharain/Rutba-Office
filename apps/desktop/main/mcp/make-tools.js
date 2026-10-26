// The AI tools with everything they need to make and read documents: one
// construction for the two ways an AI app reaches them, started by it on
// this computer (node.js) or over the network from another (remote.js).

import fs from 'node:fs';
import path from 'node:path';
import { renderPdf } from '@rutba/doc-view/export/pdf';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { isPng } from '@rutba/ocr/png';
import { createTools } from './tools.js';
import { unicodeFont } from '../system-fonts.js';
import { createOcrService } from '../ocr.js';

export const INSTRUCTIONS = 'Rutba Office makes and reads documents (.docx), workbooks (.xlsx), presentations (.pptx) and PDFs on this computer, in the folders its user has given it, and never overwrites or deletes a file. Where the user allows each one, it also reads their mail, calendars, contacts and Data Files databases, and makes drafts, events, contacts and records; anything that leaves the computer or tells another person (a message sent, an invitation or its answer) is shown to the user in Rutba Office first and happens only when they press it. list_allowed says what is allowed. The tools are off unless the user has turned them on in Rutba Office (app menu, AI Tools); a refusal says so, and names where a setting is.';

/** The version this copy is, from the app's package. */
export function versionOf() {
  try {
    return JSON.parse(fs.readFileSync(new URL('../../package.json', import.meta.url), 'utf8')).version || '0.0.0';
  } catch {
    return '0.0.0';
  }
}

/**
 * The tools, reading the person's settings afresh for every call (`settings()`);
 * `reach` answers those for mail, calendars, address books and databases, and
 * `locked()` refuses everything while a protected profile is not open (tools.js).
 */
export function makeTools({ settings, dataDir, isAlive, reach = null, locked = null }) {
  let ocr = null;
  return createTools({
    settings,
    dataDir,
    isAlive,
    reach,
    locked,
    renderPdf,
    openDocx,
    unicodeFont,
    ocr: { recognize: (args) => (ocr ||= createOcrService()).recognize(args) },
    pictureToPng: async (file) => {
      const bytes = new Uint8Array(fs.readFileSync(file));
      if (!isPng(bytes)) throw new Error(`${path.basename(file)} is not a PNG picture. The AI tools read the words in PNG pictures; save it as a PNG first.`);
      return bytes;
    },
  });
}
