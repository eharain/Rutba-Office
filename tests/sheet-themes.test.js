// Page Layout → Themes, Colours, Fonts and Effects for a workbook.
//
// One of the suite's themes (the same eleven Presentation offers) or its
// colours, fonts or effects rewrites `xl/theme/theme1.xml` — made, with its
// relationship and content type, when the workbook has none — and what takes
// its look from the theme follows: a cell filled or written in a theme
// colour, a table's style, a chart's series, a shape styled from the theme,
// and the default font (the Normal style's font, the theme's body face).
// Undo puts the old theme back.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { THEMES, PALETTES, FONT_PAIRS } from '@rutba/office-formats/themes';
import { THEMES as DECK_THEMES } from '@rutba/presentation';

const BOOK = () => buildXlsx({ sheets: [{ name: 'Sales', rows: [['Region', 'Units'], ['North', 12], ['South', 30], ['East', 18]] }] });
const accent1 = (view) => view.styles.theme[4];

test('the suite\'s themes are one list, shared by Presentation and Worksheets', () => {
  assert.equal(THEMES.length, 11);
  assert.equal(DECK_THEMES, THEMES, 'Presentation reads the same list');
  assert.ok(PALETTES.length > THEMES.length && FONT_PAIRS.length === 11);
});

test('a theme applied to a workbook with none makes the theme part, its relationship and type; undo takes them away', () => {
  const view = new SheetView(BOOK());
  assert.equal(view.pkg.has('xl/theme/theme1.xml'), false);
  assert.equal(view.workbookDesign().name, 'Office Theme');
  const d = view.setWorkbookTheme({ theme: 'harbour' });
  assert.deepEqual([d.name, d.builtIn, d.fonts.minor, d.colors.accent1], ['Harbour', 'harbour', 'Segoe UI', '1F6FB2']);
  const saved = SheetView.open(view.save());
  assert.ok(saved.pkg.has('xl/theme/theme1.xml'));
  assert.equal(saved.pkg.contentTypeOf('/xl/theme/theme1.xml') ?? saved.pkg.contentTypeOf('xl/theme/theme1.xml'), 'application/vnd.openxmlformats-officedocument.theme+xml');
  assert.ok(saved.pkg.rels('xl/workbook.xml').some((r) => r.Type.endsWith('/theme') && r.Target === 'theme/theme1.xml'));
  assert.match(saved.pkg.text('xl/theme/theme1.xml'), /<a:theme xmlns:a="[^"]+" name="Harbour"><a:themeElements><a:clrScheme name="Harbour">/);
  assert.equal(accent1(saved), '#1f6fb2', 'theme colours resolve from the new part');
  view.undo();
  assert.equal(view.pkg.has('xl/theme/theme1.xml'), false, 'undo takes the part away');
  assert.equal(view.workbookDesign().name, 'Office Theme');
});

test('theme-coloured cells, a table, a chart and a shape follow the theme', () => {
  // A header filled with the theme's first accent, as Excel writes one.
  const made = new SheetView(buildXlsx({ sheets: [{ name: 'Sales', rows: [['Region', 'Units'], ['North', 12], ['South', 30], ['East', 18]], styles: { 'A1:B1': { fill: '1F4E79' } } }] }));
  made.pkg.write_('xl/styles.xml', made.pkg.text('xl/styles.xml').replace('<fgColor rgb="FF1F4E79"/>', '<fgColor theme="4"/>'));
  const view = SheetView.open(made.pkg.write());
  view.setWorkbookTheme({ theme: 'rutba' });
  view.select(1, 0);
  view.formatAsTable({ style: 'TableStyleMedium2', stripes: true });
  view.select(1, 1);
  view.insertChart({ kind: 'column' });
  view.select(8, 0);
  view.insertShape({ geometry: 'rect' });
  const header = () => view.render().cells.find((c) => c.row === 0 && c.col === 0);
  const before = JSON.stringify(header().style?.fill ?? null);
  view.setWorkbookTheme({ theme: 'ember' });
  const ember = THEMES.find((t) => t.id === 'ember').palette.accent1.toLowerCase();
  const after = header();
  assert.notEqual(JSON.stringify(after.style?.fill ?? null), before, 'the header cell\'s theme fill follows');
  const frame = view.render();
  const chart = frame.drawings.find((d) => d.kind === 'chart');
  assert.ok(chart.svg.toLowerCase().includes('#' + ember), 'the chart draws in the theme\'s first accent');
  const shape = frame.drawings.find((d) => d.kind === 'shape');
  assert.ok(shape.svg.toLowerCase().includes('#' + ember), 'a shape filled with accent1 too');
  assert.equal(accent1(view), '#' + ember, 'tables paint from the theme\'s first accent');
});

test('Colours, Fonts and Effects each rewrite their one element; the default font follows the body face', () => {
  const view = new SheetView(BOOK());
  view.setWorkbookTheme({ theme: 'rutba' });
  const xml = () => view.pkg.text('xl/theme/theme1.xml');
  view.setWorkbookTheme({ colors: 'citrus' });
  assert.match(xml(), /<a:clrScheme name="Citrus">/);
  assert.match(xml(), /<a:fontScheme name="Rutba">/, 'the fonts are left alone');
  view.setWorkbookTheme({ colors: { dk1: '000000', lt1: 'FFFFFF', dk2: '222222', lt2: 'EEEEEE', accent1: 'AA0000', accent2: '00AA00', accent3: '0000AA', accent4: 'AAAA00', accent5: '00AAAA', accent6: 'AA00AA', hlink: '0000FF', folHlink: '800080' }, name: 'Brand' });
  assert.match(xml(), /<a:clrScheme name="Brand">[\s\S]*<a:accent1><a:srgbClr val="AA0000"\/><\/a:accent1>/);
  assert.throws(() => view.setWorkbookTheme({ colors: { dk1: 'zz' } }), /needs a colour/);
  view.setWorkbookTheme({ fonts: 'classic' });
  assert.match(xml(), /<a:fontScheme name="Classic"><a:majorFont><a:latin typeface="Georgia"\/>/);
  // The Normal style's font is the body face, marked as the theme's.
  const styles = view.pkg.text('xl/styles.xml');
  assert.match(/<fonts\b[^>]*>\s*<font>([\s\S]*?)<\/font>/.exec(styles)[1], /<name val="Georgia"\/>[\s\S]*<scheme val="minor"\/>/);
  assert.equal(view.styles.byStyleIndex[0]?.font?.family ?? view.render().cells.find((c) => c.row === 1 && c.col === 0)?.style?.font?.family, 'Georgia');
  view.setWorkbookTheme({ fonts: { major: 'Bahnschrift', minor: 'Verdana' }, name: 'Mine' });
  assert.match(xml(), /<a:fontScheme name="Mine">/);
  assert.match(view.pkg.text('xl/styles.xml'), /<name val="Verdana"\/>/);
  view.setWorkbookTheme({ effects: 'lifted' });
  assert.match(xml(), /<a:fmtScheme name="Lifted">/);
  assert.equal(view.workbookDesign().effects, 'lifted');
  view.undo();
  view.undo();
  assert.match(view.pkg.text('xl/styles.xml'), /<name val="Georgia"\/>/, 'undo puts the font back');
  assert.equal(view.workbookDesign().fonts.minor, 'Georgia');
});
