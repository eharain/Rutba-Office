/**
 * @rutba/doc-view — the editable model of a document.
 *
 * `workspace.docs` had a format layer and business-data bindings but nothing
 * to type into. This is the typing.
 *
 * DOM-free, like `@rutba/sheet-view`, and for the same reason: caret arithmetic
 * across runs and paragraph boundaries is the fiddly part, and it belongs
 * somewhere it can be tested exhaustively rather than clicked at.
 *
 * FORMAT-FREE too. The editing brain sits behind a backend port, so the same
 * caret, the same selection and the same formatting toggles drive a .docx in
 * Workspace and an email body in Mail:
 *
 *   import { DocView } from '@rutba/doc-view';
 *   import { HtmlBackend } from '@rutba/doc-view/backends/html';
 *   const view = new DocView(HtmlBackend.open(emailBody));
 *
 * Importing this module does NOT pull in @rutba/ooxml — only the ooxml backend
 * does. See backend.js for the port.
 */
export {
  runsText, locate, clampPosition, samePosition, comparePositions,
  orderedRange, sliceRuns, removeRange, coalesce,
} from './positions.js';
export { DocView } from './view.js';
export { bandForPage, resolveFields, EVALUABLE_FIELDS } from './bands.js';
export { computeListLabels, formatCounter } from './lists.js';
export { assertBackend, supportsContentControls, REQUIRED_METHODS } from './backend.js';
