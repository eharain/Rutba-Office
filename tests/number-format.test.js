/**
 * Settable number formats — currency, percent, date, plain decimals, General.
 *
 * A number format lives on the `<xf>` as its `numFmtId`, and the same
 * append-never-rewrite rule that governs fonts, fills and borders governs the
 * `<numFmts>` table too: every entry is addressed by id, so a custom code takes
 * the next id above the current maximum rather than overwriting one a cell
 * elsewhere points at. These tests pin that, plus the two shortcuts that keep
 * files ordinary:
 *
 *   - a code that matches a built-in (`0.00`, `0%`, `#,##0.00`) references its
 *     reserved id and adds NO `<numFmts>` entry, because every consumer already
 *     agrees on ids 0-49;
 *   - clearing a format is numFmtId 0 (General), not a custom "no format".
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { applyFormat, formatOf } from '@rutba/sheet-view/styles-write';

const book = (rows = [['Head', 'Other'], [1, 2]]) =>
  buildXlsx({ sheets: [{ name: 'S', rows }] });

const open = (bytes) => new SheetView(bytes);
const stateAt = (view, row, col) => {
  view.select(row, col);
  return view.formatState();
};

// `<numFmt ` (with a following space) counts entries only — the `<numFmts>`
// container is `<numFmts` with no space, so it is never miscounted as one.
const numFmtEntries = (xml) => [...xml.matchAll(/<numFmt\s/g)].length;
const bodyOf = (xml, tag) =>
  new RegExp('<' + tag + '\\b[^>]*>([\\s\\S]*?)</' + tag + '>').exec(xml)[1];

// ── currency, percent, date all set and survive a save+reopen ────────────────

test('a currency format is set, renders, and survives a save+reopen', () => {
  const view = open(book([['n'], [1234.5]]));
  view.select(1, 0);
  view.setFormat({ numberFormat: '$#,##0.00' });

  // End to end: the grid now shows the formatted value, not the raw number.
  assert.equal(view.displayValue(1, 0).text, '$1,234.50');
  assert.equal(view.formatState().numberFormat, '$#,##0.00');

  const reopened = open(view.save());
  assert.equal(stateAt(reopened, 1, 0).numberFormat, '$#,##0.00');
  assert.equal(reopened.displayValue(1, 0).text, '$1,234.50');
});

test('a percent format is set and survives a save+reopen', () => {
  const view = open(book([['ratio'], [0.25]]));
  view.select(1, 0);
  view.setFormat({ numberFormat: '0%' });

  assert.equal(view.displayValue(1, 0).text, '25%');
  const reopened = open(view.save());
  assert.equal(stateAt(reopened, 1, 0).numberFormat, '0%');
});

test('a date format is set and survives a save+reopen', () => {
  const view = open(book([['when'], [45000]]));
  view.select(1, 0);
  view.setFormat({ numberFormat: 'yyyy-mm-dd' });

  assert.equal(view.formatState().numberFormat, 'yyyy-mm-dd');
  const reopened = open(view.save());
  assert.equal(stateAt(reopened, 1, 0).numberFormat, 'yyyy-mm-dd');
});

test('a plain decimal format is set and reopens as itself', () => {
  const view = open(book([['n'], [3]]));
  view.select(1, 0);
  view.setFormat({ numberFormat: '0.000' });
  assert.equal(view.displayValue(1, 0).text, '3.000');
  const reopened = open(view.save());
  assert.equal(stateAt(reopened, 1, 0).numberFormat, '0.000');
});

// ── a built-in code references its reserved id, and adds no <numFmts> entry ──

test('a built-in code references its reserved id and adds no <numFmts> entry', () => {
  const view = open(book());
  view.select(1, 0);
  view.setFormat({ numberFormat: '0.00' }); // built-in numFmtId 2

  const xml = view.pkg.text('xl/styles.xml');
  assert.doesNotMatch(xml, /<numFmts/, 'a standard code needs no custom table');
  assert.match(xml, /numFmtId="2"/, 'it references the reserved id');
  assert.equal(view.formatState().numberFormat, '0.00');
});

test('the built-in percent id 9 is referenced without a custom entry', () => {
  const view = open(book());
  view.select(1, 0);
  view.setFormat({ numberFormat: '0%' });

  const xml = view.pkg.text('xl/styles.xml');
  assert.doesNotMatch(xml, /<numFmts/);
  assert.match(xml, /numFmtId="9"/);
});

// ── a custom code lands in <numFmts> at 164 and up ──────────────────────────

test('a custom code adds a <numFmts> entry starting at 164', () => {
  const view = open(book());
  view.select(1, 0);
  view.setFormat({ numberFormat: '$#,##0.00' });

  const xml = view.pkg.text('xl/styles.xml');
  assert.match(xml, /<numFmts count="1"><numFmt numFmtId="164" formatCode="\$#,##0\.00"\/><\/numFmts>/);
  assert.match(xml, /<cellXfs\b[^>]*>[\s\S]*numFmtId="164"[\s\S]*<\/cellXfs>/);
});

test('a second custom code takes the next id above the current maximum', () => {
  const view = open(book([['a', 'b']]));
  view.select(0, 0);
  view.setFormat({ numberFormat: '$#,##0.00' });   // 164
  view.select(0, 1);
  view.setFormat({ numberFormat: 'yyyy-mm-dd' });  // 165

  const xml = view.pkg.text('xl/styles.xml');
  assert.match(xml, /<numFmt numFmtId="164" formatCode="\$#,##0\.00"\/>/);
  assert.match(xml, /<numFmt numFmtId="165" formatCode="yyyy-mm-dd"\/>/);
  assert.equal(numFmtEntries(xml), 2);
});

test('a custom code with quoted literal text round-trips through escaping', () => {
  const view = open(book([['n'], [1234.5]]));
  view.select(1, 0);
  view.setFormat({ numberFormat: '#,##0.00" PKR"' });

  const xml = view.pkg.text('xl/styles.xml');
  assert.match(xml, /formatCode="#,##0\.00&quot; PKR&quot;"/, 'the quotes are escaped in the attribute');
  assert.equal(view.formatState().numberFormat, '#,##0.00" PKR"', 'and decode back on read');
  const reopened = open(view.save());
  assert.equal(stateAt(reopened, 1, 0).numberFormat, '#,##0.00" PKR"');
});

// ── the same custom code twice reuses one entry ─────────────────────────────

test('setting the same custom code twice reuses one <numFmts> entry', () => {
  const view = open(book([['a', 'b']]));
  view.select(0, 0);
  view.setFormat({ numberFormat: '$#,##0.00' });
  const afterFirst = view.pkg.text('xl/styles.xml');
  assert.equal(numFmtEntries(afterFirst), 1);

  view.select(0, 1);
  view.setFormat({ numberFormat: '$#,##0.00' });
  const afterSecond = view.pkg.text('xl/styles.xml');

  assert.equal(numFmtEntries(afterSecond), 1, 'the second currency reused the first entry');
  // Both cells end up pointing at the same style, since nothing about them differs.
  assert.equal(view._styleIndexAt('S', 0, 0), view._styleIndexAt('S', 0, 1));
});

// ── clearing back to General ─────────────────────────────────────────────────

test('clearing to General puts the cell back to numFmtId 0', () => {
  const view = open(book([['n'], [1234.5]]));
  view.select(1, 0);
  view.setFormat({ numberFormat: '$#,##0.00' });
  assert.equal(view.formatState().numberFormat, '$#,##0.00');
  assert.equal(view.displayValue(1, 0).text, '$1,234.50');

  view.setFormat({ numberFormat: null });
  assert.equal(view.formatState().numberFormat, 'General');
  // General renders the raw number again, not the currency.
  assert.equal(view.displayValue(1, 0).text, '1234.5');
});

test("the string 'General' clears the format too", () => {
  const view = open(book([['n'], [5]]));
  view.select(1, 0);
  view.setFormat({ numberFormat: '0.00' });
  view.setFormat({ numberFormat: 'General' });
  assert.equal(view.formatState().numberFormat, 'General');
});

// ── the append rule across <numFmts> and <cellXfs> ──────────────────────────

test('the numFmts and cellXfs tables stay append-only', () => {
  const view = open(book([['a', 'b'], ['c', 'd']]));
  view.select(0, 0);
  view.setFormat({ numberFormat: '$#,##0.00' });   // creates <numFmts> with 164
  const before = view.pkg.text('xl/styles.xml');

  view.select(1, 1);
  view.setFormat({ numberFormat: 'yyyy-mm-dd' });  // appends 165 and a new xf
  const after = view.pkg.text('xl/styles.xml');

  for (const tag of ['numFmts', 'cellXfs']) {
    assert.ok(bodyOf(after, tag).startsWith(bodyOf(before, tag)),
      tag + ' grew at the end; nothing already in it moved or changed');
  }
});

test('a custom format on one cell does not restyle its neighbours', () => {
  const view = open(book([['one', 'two', 'three']]));
  view.select(0, 1);
  view.setFormat({ numberFormat: '$#,##0.00' });

  const xml = view.pkg.text('xl/styles.xml');
  const at = (row, col) => formatOf(xml, view._styleIndexAt('S', row, col) ?? 0);
  assert.equal(at(0, 1).numberFormat, '$#,##0.00');
  assert.equal(at(0, 0).numberFormat, 'General', 'the cell to the left is untouched');
  assert.equal(at(0, 2).numberFormat, 'General', 'and so is the cell to the right');
});

// ── undo ─────────────────────────────────────────────────────────────────────

test('undo restores styles.xml exactly, numFmts and all', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ numberFormat: '$#,##0.00' });   // styles.xml now has a numFmts block
  const before = view.pkg.text('xl/styles.xml');

  view.select(1, 0);
  view.setFormat({ numberFormat: 'yyyy-mm-dd' });  // appends numFmt 165 and a new xf
  assert.equal(view.formatState().numberFormat, 'yyyy-mm-dd');

  view.undo();
  assert.equal(view.formatState().numberFormat, 'General', 'the cell is no longer a date');
  assert.equal(view.pkg.text('xl/styles.xml'), before,
    'the appended numFmt and xf went with the undo');
});

// ── the writer in isolation ─────────────────────────────────────────────────

test('applyFormat selects a built-in id, a custom id, and General correctly', () => {
  const styles = '<?xml version="1.0"?>'
    + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
    + '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
    + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    + '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>'
    + '</styleSheet>';

  const builtIn = applyFormat(styles, 0, { numberFormat: '#,##0.00' });
  assert.match(builtIn.xml.match(/<xf[^>]*numFmtId="4"[^>]*>?/) ? 'ok' : builtIn.xml, /ok/,
    'a built-in code takes id 4');
  assert.doesNotMatch(builtIn.xml, /<numFmts/, 'and adds no custom table');
  assert.equal(formatOf(builtIn.xml, builtIn.index).numberFormat, '#,##0.00');

  const custom = applyFormat(builtIn.xml, builtIn.index, { numberFormat: '0.0"x"' });
  assert.match(custom.xml, /<numFmt numFmtId="164" formatCode="0\.0&quot;x&quot;"\/>/);
  assert.equal(formatOf(custom.xml, custom.index).numberFormat, '0.0"x"');

  const cleared = applyFormat(custom.xml, custom.index, { numberFormat: null });
  assert.equal(formatOf(cleared.xml, cleared.index).numberFormat, 'General');
});

test('asking for a number format a cell already has is a no-op', () => {
  const view = open(book());
  view.select(0, 0);
  view.setFormat({ numberFormat: '$#,##0.00' });
  const xml = view.pkg.text('xl/styles.xml');
  const index = view._styleIndexAt('S', 0, 0);

  const again = applyFormat(xml, index, { numberFormat: '$#,##0.00' });
  assert.equal(again.changed, false, 'the same format is a no-op');
  assert.equal(again.index, index);
  assert.equal(again.xml, xml, 'and it does not grow the style table');
});
