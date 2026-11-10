// Which marks a search makes on a page: the boxes measured on the page's own
// glyphs, when they are sure and found the same matches as the text layer;
// otherwise none, and the window widens its own (redact.js searchQuads).

/**
 * @param {number} found how many matches the text layer found
 * @param {{ count: number, unsure: boolean, quads: number[][] } | null} measured the page's own measure
 * @returns {number[][] | null}
 */
export function exactQuads(found, measured) {
  if (!measured || measured.unsure || !found || measured.count !== found) return null;
  return measured.quads;
}
