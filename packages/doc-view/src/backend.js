/**
 * The document backend port.
 *
 * `DocView` is the editing brain — caret arithmetic across runs and paragraphs,
 * insert/delete/split/merge, formatting toggles, the structural refusal. None of
 * that is specific to a file format, and Mail needs every bit of it to compose an
 * email. So the format sits behind this port and `DocView` never imports one.
 *
 *   @rutba/doc-view/backends/ooxml   Workspace — a .docx, preserving
 *   @rutba/doc-view/backends/html    Mail — an email body
 *   (Studio would add its own)
 *
 * The port is small on purpose. Everything it exposes is something the editor
 * genuinely has to do; anything a backend wants to keep private stays private.
 * `runs` carry an opaque `rPr` — whatever the backend uses to describe character
 * formatting — plus the three booleans the editor understands. The editor never
 * parses `rPr`; it only compares them for equality (to coalesce) and asks the
 * backend to toggle one. That is what lets OOXML keep a verbatim `<w:rPr>` and
 * HTML keep a class list, without the editor knowing either exists.
 *
 * REQUIRED
 *   paragraphCount()                       -> number
 *   paragraph(index)                       -> Paragraph | null
 *   setParagraphRuns(index, runs)          -> this
 *   splitParagraph(index, runIndex, offset)-> this
 *   mergeWithNext(index)                   -> this
 *   removeParagraph(index)                 -> this
 *   toggleRunFormat(rPr, tag, on)          -> rPr'      (tag is 'b' | 'i' | 'u' | 's')
 *   save()                                 -> whatever the format's artefact is
 *
 * OPTIONAL
 *   insertParagraphAfter(index, runs, opts)
 *   setContentControlText(tag, value)      named anchors, for bound templates
 *   contentControls()                      -> [{ tag, alias, text }]
 *   blocks()                               -> body in document order (see below)
 *   section()                              -> { widthPx, heightPx, margins, … }
 *   insertTable(index, rows, cols)         a fresh table after a prose paragraph
 *   appendTableRow(tableStart)             one more row — Tab in the last cell
 *   insertImage(index, spec)               a picture in its own paragraph after `index`
 *   modifiedParts()                        -> string[]  (diagnostics only)
 *
 * `blocks()` is how a table gets DRAWN — and, since D5, edited. It returns the
 * body in document order as `{kind:'paragraph', paragraphIndex}` and
 * `{kind:'table', table}`, with each editable cell paragraph in the parsed
 * table carrying its `blockIndex` into the paragraph space. A backend without
 * tables can omit it, and the view falls back to the flat paragraph list.
 *
 * `section()` is page geometry — size and margins. Omit it and the view renders
 * a continuous flow, which is what an email body wants.
 *
 * Paragraph:
 *   { index, text, style, runs: [{ rPr, text, bold, italic, underline, strike }],
 *     structural: boolean, structuralTags: string[], container: string|null }
 *
 * `container` names the table cell a paragraph lives in (null for prose; a
 * backend with no tables never sets it). Blocks share a container exactly when
 * they share a cell — the equality the view's boundary guards are built on: a
 * merge or a range delete never crosses a change of container, and Tab walks
 * to the next one.
 *
 * `structural` is the contract's most important field. It means "this paragraph
 * carries something rebuilding it would destroy" — a field code, a bookmark, a
 * content control. The editor refuses to type into those rather than flattening
 * them. A backend with no such concept simply always returns false.
 */

const REQUIRED = [
  'paragraphCount', 'paragraph', 'setParagraphRuns', 'splitParagraph',
  'mergeWithNext', 'removeParagraph', 'toggleRunFormat', 'save',
];

/** Fail loudly at construction rather than mysteriously on the first keystroke. */
export function assertBackend(backend) {
  if (!backend || typeof backend !== 'object') throw new Error('a document backend is required');
  const missing = REQUIRED.filter((m) => typeof backend[m] !== 'function');
  if (missing.length) {
    throw new Error('document backend is missing: ' + missing.join(', ')
      + ' — see packages/doc-view/src/backend.js for the port');
  }
  return backend;
}

export const supportsContentControls = (backend) =>
  typeof backend.setContentControlText === 'function' && typeof backend.contentControls === 'function';

export { REQUIRED as REQUIRED_METHODS };
