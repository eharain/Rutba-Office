# @rutba/doc-view

A word processor's editing model, minus the format.

Caret arithmetic across runs and paragraphs, insert/delete/split/merge, format
toggles, pagination, undo — none of which is specific to a file format, and all
of which Mail needs to compose an email. So the format sits behind a port and the
editor never imports one.

```js
import { DocView } from '@rutba/doc-view';
import { openDocx } from '@rutba/doc-view/backends/ooxml';   // Workspace
import { openHtml } from '@rutba/doc-view/backends/html';    // Mail

const view = openHtml('<p>Thanks for the quote.</p>');
view.setSelection({ block: 0, offset: 6 });
view.insertText(' very much');
view.render();   // { blocks, flow, pages, section, selection, history, … }
```

There is deliberately no `DocView.open(bytes)`. A format-free editor cannot know
what bytes are, and making it guess would put an import of every backend at the
top of the file. Each backend exports its own opener.

## The port

Required: `paragraphCount`, `paragraph`, `setParagraphRuns`, `splitParagraph`,
`mergeWithNext`, `removeParagraph`, `toggleRunFormat`, `save`.

Optional: `insertParagraphAfter`, `setContentControlText`, `contentControls`,
`snapshot`/`restore` (undo), `blocks` (document order, so tables draw),
`section` (page geometry), `headerFooters`.

A run carries an **opaque** `rPr` — whatever the backend uses for character
formatting — plus the three booleans the editor understands. The editor never
parses it; it compares them for equality and asks the backend to toggle one. That
is what lets OOXML keep a verbatim `<w:rPr>` and HTML keep a class list.

`structural` is the contract's most important field: "this paragraph carries
something rebuilding it would destroy" — a field code, a bookmark, a content
control. The editor refuses to type into those rather than flattening them. A
backend with no such concept always returns false.

## Two traps, both learned the hard way

**A table cell contains real `<p>` elements.** When a shell maps a DOM caret back
with `closest('p')`, a click inside a cell matches one of *those*, produces a
block index of `NaN`, and typing edits the first paragraph of the document. Mark
the editable paragraphs and match on that marker.

**A paginated paragraph renders as two elements with the same block index.** Each
must carry the character offset it starts at, or a caret on page three types on
page one.

## Rendering pages

`frame.pages` is the flow laid onto sheets, or `null` for a continuous flow —
which is what an email body wants. If you render pages, **the server owns line
breaking**: each fragment carries the lines to draw and the metrics they were
measured at. Draw at those metrics, and put each line in its own
`white-space: pre` box. Inheriting a font size from your stylesheet, or letting
the browser re-wrap, is how a page silently overflows by a line.

## Status

Plain ESM JavaScript, no build step. Depends only on `@rutba/drawing` (text
metrics) and `@rutba/editing` (undo) — both shared packages that travel with it.
The OOXML backend is the only file that imports a format layer, and
`tests/shared-seam.test.js` in Rutba-Workspace enforces that.
