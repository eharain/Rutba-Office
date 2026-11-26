/**
 * What somebody types into the Name Box, read the way Excel reads it: an
 * address to go to, a name that exists to go to, or a new name for what is
 * selected. Pure text work: the view does the going and the naming.
 */

import { MAX_ROWS, MAX_COLS } from './geometry.js';

/** What a defined name may be: a letter or an underscore first, then letters, digits, dots and underscores. */
const NAME = /^[A-Za-z_][A-Za-z0-9_.]*$/;
/** A word that reads as a cell and so cannot be a name. */
const CELL_LIKE = /^[A-Za-z]{1,3}[0-9]+$/;

/** A column's letters to its 0-based index. */
function colIndexOf(letters) {
  return [...letters].reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
}

/**
 * A typed address as `{ sheet, kind, top, left, bottom, right }`, zero-based,
 * or null when the text is not one. It takes a cell (`B7`), a range
 * (`B2:D9`, either corner first), whole columns (`A:C`) and whole rows
 * (`3:7`), with or without dollar signs, and a sheet before any of them
 * (`Stock!B2`, `'May sales'!A1:C4`). `kind` is `cells`, `cols` or `rows`;
 * `sheet` is null when none was typed.
 */
export function parseAddress(text) {
  let rest = String(text ?? '').trim();
  let sheet = null;
  const quoted = /^'((?:[^']|'')+)'!(.*)$/.exec(rest);
  if (quoted) {
    sheet = quoted[1].replace(/''/g, "'");
    rest = quoted[2];
  } else {
    const bare = /^([^'!:][^!:]*)!(.*)$/.exec(rest);
    if (bare) {
      sheet = bare[1].trim();
      rest = bare[2];
    }
  }
  rest = rest.replace(/\$/g, '').replace(/\s+/g, '').toUpperCase();
  const span = (a, b) => [Math.min(a, b), Math.max(a, b)];

  const cells = /^([A-Z]{1,3})(\d{1,7})(?::([A-Z]{1,3})(\d{1,7}))?$/.exec(rest);
  if (cells) {
    const [left, right] = span(colIndexOf(cells[1]), colIndexOf(cells[3] ?? cells[1]));
    const [top, bottom] = span(Number(cells[2]) - 1, Number(cells[4] ?? cells[2]) - 1);
    if (top < 0 || bottom >= MAX_ROWS || right >= MAX_COLS) return null;
    return { sheet, kind: 'cells', top, left, bottom, right };
  }
  const cols = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(rest);
  if (cols) {
    const [left, right] = span(colIndexOf(cols[1]), colIndexOf(cols[2]));
    if (right >= MAX_COLS) return null;
    return { sheet, kind: 'cols', top: 0, left, bottom: MAX_ROWS - 1, right };
  }
  const rows = /^(\d{1,7}):(\d{1,7})$/.exec(rest);
  if (rows) {
    const [top, bottom] = span(Number(rows[1]) - 1, Number(rows[2]) - 1);
    if (top < 0 || bottom >= MAX_ROWS) return null;
    return { sheet, kind: 'rows', top, left: 0, bottom, right: MAX_COLS - 1 };
  }
  return null;
}

/**
 * What an entry in the Name Box asks for, in Excel's order: an address goes
 * there (`{ kind: 'ref', text }`); a name the workbook has, typed in any
 * case, goes to it (`{ kind: 'name', name }`, spelt as the workbook spells
 * it); a word that can be a name names what is selected
 * (`{ kind: 'define', name }`); anything else is `{ kind: 'bad', text }`.
 * Null for an empty entry. `names` is the frame's list (`[{ name }]`).
 */
export function nameBoxEntry(text, names = []) {
  const typed = String(text ?? '').trim();
  if (!typed) return null;
  if (parseAddress(typed)) return { kind: 'ref', text: typed };
  const wanted = typed.toUpperCase();
  const known = names.find((n) => String(n?.name ?? '').toUpperCase() === wanted);
  if (known) return { kind: 'name', name: known.name };
  if (!NAME.test(typed) || typed.length > 255 || CELL_LIKE.test(typed)) return { kind: 'bad', text: typed };
  return { kind: 'define', name: typed };
}
