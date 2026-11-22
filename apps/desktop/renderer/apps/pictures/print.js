// Pictures, on paper: the picture on the stage, or every picture shown.
//
// One picture prints as large as the sheet takes; a file of several pages (a
// scan kept as TIFF) prints every page, a sheet each. The shown pictures
// print one, two, four, six or nine to a sheet, with their file names
// beneath when asked: a contact sheet.
//
// The print window draws the files themselves (rutba://file), so nothing is
// copied to print them, and a picture opened from the Locker stays in memory
// (main/print.js holds such a page there).

import { tn } from '@rutba/office-ui/messages';
import { pageHtml, pictureSheets } from '../../print-page.js';

/** The most pictures one print takes: past this, the dialog says so and prints nothing. */
export const MOST_PICTURES = 200;

/** One file as the pictures it prints: itself, or each of its pages. `url` gives a page's address. */
export function filePictures({ url, name = '', pages = 1, pageUrl = (u) => u } = {}) {
  const count = Math.max(1, Number(pages) || 1);
  if (count === 1) return [{ url, name }];
  return Array.from({ length: count }, (_, i) => ({ url: pageUrl(url, i), name: `${name} (${i + 1})` }));
}

/**
 * The page for these pictures, each `{ url, name }`. Throws, saying how
 * many, when there are more than one print takes.
 */
export function picturesPrintHtml({ pictures = [], perPage = 1, captions = false, title = '', options = {}, lang = 'en', dir = 'ltr' } = {}) {
  if (pictures.length > MOST_PICTURES) {
    throw new Error(tn(pictures.length, 'pictures.print.count-picture-more', { most: MOST_PICTURES }));
  }
  const sheet = pictureSheets({ pictures, perPage, captions, options });
  return pageHtml({ title, options, css: sheet.css, body: sheet.body, lang, dir });
}
