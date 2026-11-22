// The picture in Image, on paper.
//
// What prints is the picture as it stands, every layer and edit drawn, the
// same as an export would write: the window hands those bytes to the main
// process, which holds them for the print window to draw.
//
// Two sizes. Fit to the page makes the picture as large as the sheet takes,
// its shape kept. Actual size prints it as a window shows it at 100%, 96
// pixels to the inch: on one sheet when it fits, and cut across as many as
// it needs when it does not, to be laid edge to edge.

import { tn } from '@rutba/office-ui/messages';
import { pageHtml, pictureSheets, posterSheets, tilesFor, MOST_TILES } from '../../print-page.js';

/** How many sheets the picture takes at this size on this paper. */
export function imageSheets({ width, height, size = 'fit', options = {} } = {}) {
  return size === 'actual' ? tilesFor({ width, height, options }).sheets : 1;
}

/**
 * The page for the picture. `url` is where the main process holds its bytes,
 * `width` and `height` its pixels, `size` fit or actual. Throws, saying how
 * many sheets, when its own size would take more sheets than anyone means.
 */
export function imagePrintHtml({ url, width, height, name = '', size = 'fit', options = {}, lang = 'en', dir = 'ltr' } = {}) {
  if (size === 'actual') {
    const poster = posterSheets({ url, width, height, options });
    if (poster.tooMany) {
      throw new Error(tn(poster.sheets, 'image.print.actual-size-picture', { most: MOST_TILES }));
    }
    return pageHtml({ title: name, options, css: poster.css, body: poster.body, lang, dir });
  }
  const sheet = pictureSheets({ pictures: [{ url, name }], perPage: 1, options });
  return pageHtml({ title: name, options, css: sheet.css, body: sheet.body, lang, dir });
}
