/**
 * @rutba/editing — the parts of "being an editor" that are not about a format.
 *
 * Sheets, documents and (soon) Mail all need the same undo semantics, and three
 * implementations of undo would drift into three different feels. Anything here
 * must be true of every editing surface; anything that is only true of one
 * belongs in that surface's package.
 */
export { History, DEFAULT_LIMIT, DEFAULT_COALESCE_MS } from './history.js';
