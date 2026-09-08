/**
 * Headers and footers, as the VIEW sees them.
 *
 * These two functions lived in `@rutba/ooxml` until it turned out that importing
 * them put a format dependency in the middle of a format-neutral editor: Mail
 * asking for `@rutba/doc-view` would have dragged the whole OOXML layer in with
 * it, for two functions that do not know what OOXML is.
 *
 * The split is the same one the backend port draws everywhere else. READING a
 * header out of a file is format work and stays in the backend — `w:hdr`,
 * relationship ids, `w:fldSimple` are all OOXML's business. Deciding WHICH band
 * a given page gets, and what a page-number field says on it, is view work: it
 * depends on the page number, not on the file.
 *
 * A band is therefore plain data by the time it arrives here:
 *   { paragraphs: [{ align, text, runs: [{ text, field, bold, italic, underline }] }] }
 *
 * An HTML backend that wanted running headers would produce the same shape and
 * get the same behaviour without touching either of these.
 */

/**
 * The band that applies to a given page.
 *
 * `titlePage` is the flag that turns the `first` reference on, and it matters:
 * a file can carry a first-page header that Word never shows, and honouring it
 * anyway puts a letterhead on page one of a document that does not want one.
 *
 * @param {{default?: object, first?: object, even?: object}} bands
 * @param {number} pageNumber  1-based
 */
export function bandForPage(bands, pageNumber, { titlePage = false, evenAndOdd = false } = {}) {
  // A title page takes the first-page band and NOTHING else: when the file
  // defines none, Word leaves page 1 blank rather than reaching for the
  // default — which is how a cover page has no running header.
  if (titlePage && pageNumber === 1) return bands.first ?? null;
  if (evenAndOdd && pageNumber % 2 === 0 && bands.even) return bands.even;
  return bands.default ?? bands.even ?? bands.first ?? null;
}

/**
 * Substitute the fields we can compute; leave the rest showing their cache.
 *
 * Only the two that pagination makes knowable. Anything else keeps the result
 * the authoring tool last wrote, which is the honest fallback: a stale value the
 * author has seen beats a blank, and beats a guess.
 */
export function resolveFields(paragraphs, { page, of }) {
  return (paragraphs ?? []).map((p) => {
    const runs = p.runs.map((run) => {
      if (run.field === 'PAGE') return { ...run, text: String(page) };
      if (run.field === 'NUMPAGES') return { ...run, text: String(of) };
      return run;
    });
    return { ...p, runs, text: runs.map((r) => r.text).join('') };
  });
}

/** The fields `resolveFields` can actually compute. */
export const EVALUABLE_FIELDS = ['PAGE', 'NUMPAGES'];
