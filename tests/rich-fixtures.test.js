/**
 * The rich fixtures: files Office itself wrote, with the things people put
 * in files — charts, shapes, a picture, cross-sheet formulas, names, number
 * formats, merged cells, frozen panes, a table with a total row — opened
 * through the document service the windows use.
 *
 * tools/make-rich-fixtures.ps1 writes them, as OOXML and as OpenDocument
 * from one source, so the OpenDocument reader is judged against the same
 * content. Excel's own cached results sit in the .xlsx; the engine's
 * calculation is held to them below.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createDocumentService } from '../apps/desktop/main/documents.js';
import { formatValue, localeShortDate } from '@rutba/formula/numfmt';

const RICH = join(dirname(fileURLToPath(import.meta.url)), 'fixtures', 'rich');
const doc = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });

/** The active sheet's model after selecting a sheet, with the whole sheet in view. */
function sheet(id, name) {
  doc.apply({ id, ops: [{ op: 'sheet', name }, { op: 'viewport', x: 0, y: 0, width: 2400, height: 1800 }] });
  return doc.model({ id });
}
const cell = (m, ref) => m.cells.find((c) => c.ref === ref);
const text = (m, ref) => cell(m, ref)?.text;

for (const ext of ['xlsx', 'ods']) {
  const file = join(RICH, `showcase.${ext}`);
  test(`the showcase workbook (.${ext}): every sheet, every reference, every format`, { skip: !existsSync(file) && 'fixture not generated' }, () => {
    const s = doc.open({ path: file });
    try {
      // Office's OpenDocument export writes "Data types" as "Data_types".
      assert.deepEqual(s.model.sheets.map((n) => n.replace('_', ' ')), ['Sales', 'Summary', 'Charts', 'Data types', 'Table']);
      const names = Object.fromEntries((s.model.names || []).filter((n) => !n.name.startsWith('_xlnm')).map((n) => [n.name, n.ref]));
      assert.deepEqual(names, { GrandTotal: 'Sales!$F$14', Regions: 'Sales!$B$1:$E$1', SalesTotals: 'Sales!$F$2:$F$13' });

      const sales = sheet(s.id, s.model.sheets[0]);
      assert.equal(text(sales, 'F2'), '7,518', 'a row total, with the thousands separator its format asks for');
      assert.equal(text(sales, 'F14'), '142,444', 'the grand total');
      assert.equal(text(sales, 'H3'), '11.4%', 'growth as a percentage');
      assert.equal(text(sales, 'B15'), '2,377.0', 'an average to one place');
      assert.equal(text(sales, 'F15'), '11,870.3', 'the average of the totals');
      if (ext === 'xlsx') assert.deepEqual(sales.frozen, { rows: 1, cols: 1 }, 'the headings are frozen');

      const summary = sheet(s.id, s.model.sheets[1]);
      assert.equal(text(summary, 'B3'), '142,444', 'SUM across a sheet');
      assert.equal(text(summary, 'B4'), '142,444', 'a named range, which the .ods spells $$GrandTotal');
      assert.equal(text(summary, 'B5'), 'Nov', 'INDEX/MATCH across a sheet');
      assert.equal(text(summary, 'B6'), '9,230', 'VLOOKUP across a sheet');
      assert.equal(text(summary, 'B9'), '20.0%', 'a name in arithmetic, formatted as a percentage');
      assert.equal(text(summary, 'B11'), 'On target', 'IF over a name');
      assert.equal(text(summary, 'B12'), '142,444.00', 'TEXT() with a format code');
      assert.equal(text(summary, 'B13'), '45.142', 'two cells from a sheet whose name has a space');
      assert.equal(text(summary, 'B14'), '4,210.65', 'SUM over another sheet\'s table');
      assert.equal(text(summary, 'B16'), '4', 'COUNTA over a named range');
      assert.ok(cell(summary, 'A1')?.merged, 'the title is merged across four columns');
      if (ext === 'xlsx') assert.equal(summary.drawings.filter((d) => d.kind === 'chart').length, 1, 'the pie chart');

      const charts = sheet(s.id, s.model.sheets[2]);
      if (ext === 'xlsx') {
        const kinds = charts.drawings.reduce((acc, d) => ({ ...acc, [d.kind]: (acc[d.kind] || 0) + 1 }), {});
        assert.equal(kinds.chart, 4, 'four charts');
        assert.equal(kinds.image, 1, 'the picture');
        assert.ok(kinds.shape >= 14, `at least fourteen shapes, got ${kinds.shape}`);
      }

      const types = sheet(s.id, s.model.sheets[3]);
      assert.equal(text(types, 'B4'), '£1,234.50', 'currency');
      assert.equal(text(types, 'B5'), '25.6%', 'percent');
      // Excel kept `dd/mm/yyyy` as its built-in format 14, the system's short
      // date, so the .xlsx reads as this machine writes dates; the .ods
      // carries the code itself.
      assert.equal(text(types, 'B6'), ext === 'ods' ? '10/09/2026' : formatValue(46275, localeShortDate()).text, 'a date, formatted as the file says');
      assert.equal(text(types, 'B10'), 'TRUE', 'a boolean');
      assert.equal(text(types, 'B14'), '#DIV/0!', 'an error');
      assert.equal(text(types, 'B15'), '#N/A', 'NA()');
      assert.equal(text(types, 'B16'), '142,444', 'a cross-sheet reference');
      assert.ok(cell(types, 'B20')?.merged, 'a cell merged across three columns');

      const table = sheet(s.id, s.model.sheets[4]);
      assert.equal(text(table, 'D2'), '£999.96', 'Quantity × Price');
      assert.equal(text(table, 'D8'), '£4,210.65', 'the total row: SUBTOTAL(109, Orders[Amount])');
      assert.equal(text(table, 'B10'), '£4,210.65', 'a structured reference outside the table');
    } finally {
      doc.close({ id: s.id });
    }
  });
}

const deckFile = join(RICH, 'showcase.pptx');
test('the showcase deck: shapes with fills, a table, two charts drawn, a picture, notes, sections and a hidden slide', { skip: !existsSync(deckFile) && 'fixture not generated' }, () => {
  const s = doc.open({ path: deckFile });
  try {
    assert.equal(s.model.count, 8, 'eight slides');
    const slide = (i) => doc.model({ id: s.id, slide: i }).slide;
    const kinds = (sl) => sl.shapes.reduce((acc, sh) => ({ ...acc, [sh.kind]: (acc[sh.kind] || 0) + 1 }), {});

    const title = slide(0);
    assert.ok(title.shapes.some((sh) => /Rutba Office showcase/.test(JSON.stringify(sh.text || ''))), 'the title');
    assert.ok(/notes for the title slide/i.test(title.notes || ''), 'speaker notes');
    assert.ok(title.background && title.background.type !== 'solid', `a gradient background, got ${JSON.stringify(title.background)}`);

    const bullets = slide(1);
    assert.equal(kinds(bullets).picture, 1, 'the picture');
    assert.ok(/sub-sub-point/.test(JSON.stringify(bullets.shapes)), 'three levels of bullets');

    // The window's model carries geometry and text; fills and outlines
    // reach it as the slide's picture, so the picture is what is judged.
    const gallery = slide(2);
    assert.ok(gallery.shapes.length >= 15, `a gallery of shapes, got ${gallery.shapes.length}`);
    const fills = new Set(gallery.svg.match(/fill="#[0-9a-f]{6}"/gi) || []);
    assert.ok(fills.size >= 12, `twelve distinct fill colours drawn, got ${fills.size}`);
    for (const colour of ['#c00000', '#ffc000', '#70ad47', '#7030a0']) assert.ok(gallery.svg.includes(`fill="${colour}"`), `the ${colour} fill is drawn`);
    assert.ok((gallery.svg.match(/<linearGradient/g) || []).length >= 3, 'gradient fills are drawn as gradients');
    assert.ok(gallery.shapes.some((sh) => sh.geometry?.rot), 'one is rotated');
    assert.ok(gallery.shapes.some((sh) => sh.kind === 'connector'), 'the connector is read');

    const table = slide(3);
    assert.equal(kinds(table).table, 1, 'the table');
    assert.ok(table.svg.includes('7,590') && table.svg.includes('Region'), 'the table is drawn with its cells');

    const charts = slide(4);
    assert.equal(kinds(charts).chart, 2, 'two charts, read as charts');
    assert.ok(/Columns/.test(charts.svg) && /A pie/.test(charts.svg), 'both charts drawn on the slide with their titles');
    assert.ok((charts.svg.match(/<rect /g) || []).length >= 8, 'the column chart\'s bars are in the picture');

    const wordArt = slide(5);
    assert.ok(/WordArt/.test(JSON.stringify(wordArt.shapes)), 'the WordArt text');
    assert.ok(/bulleted/.test(JSON.stringify(wordArt.shapes)), 'the formatted text box');

    const closing = slide(7);
    assert.ok(/Thank you/.test(JSON.stringify(closing.shapes)), 'the closing slide');
    assert.ok(/日本語/.test(JSON.stringify(closing.shapes)), 'non-Latin text survives');
  } finally {
    doc.close({ id: s.id });
  }
});

const docFile = join(RICH, 'showcase.docx');
test('the showcase document: styles, lists, a table, a picture, floating shapes, notes, a comment, fields, a landscape section, a watermark', { skip: !existsSync(docFile) && 'fixture not generated' }, () => {
  const s = doc.open({ path: docFile });
  try {
    const m = s.model;
    const styles = m.blocks.reduce((acc, b) => ({ ...acc, [b.style || '']: (acc[b.style || ''] || 0) + 1 }), {});
    assert.equal(styles.Title, 1, 'the title');
    assert.equal(styles.Subtitle, 1, 'the subtitle');
    assert.equal(styles.Heading1, 7, 'seven level-one headings');
    assert.equal(styles.Heading2, 3, 'three level-two headings');
    assert.ok(styles.TOC1 >= 6, 'a table of contents, built by Word with entries');
    assert.ok(styles.ListParagraph >= 10, 'the bulleted, numbered and outline lists');
    assert.equal(styles.Caption, 1, 'the picture\'s caption');
    const cells = new Set(m.blocks.map((b) => b.container).filter(Boolean).map((c) => JSON.stringify(c)));
    assert.ok(cells.size >= 18, `the table's cells, got ${cells.size}`);
    assert.equal(m.footnotes.length, 1, 'the footnote');
    assert.equal(m.endnotes.length, 1, 'the endnote');
    assert.equal(m.comments.length, 1, 'the comment');
    assert.ok(m.bands && m.bands.watermark, 'the DRAFT watermark is a band of its own');
    const runs = m.blocks.flatMap((b) => b.runs || []);
    assert.ok(runs.some((r) => r.highlight), 'a highlighted run');
    assert.ok(runs.some((r) => r.strike), 'a struck run');
    assert.ok(runs.some((r) => r.smallCaps), 'small caps');
    assert.ok(runs.some((r) => r.fontName === 'Georgia' && r.fontSize === 16), 'Georgia at 16');
    assert.ok(runs.some((r) => r.vertAlign === 'superscript') && runs.some((r) => r.vertAlign === 'subscript'), 'superscript and subscript');
    assert.ok(runs.some((r) => r.link), 'the hyperlink');
    assert.ok(runs.some((r) => r.noteRef), 'a note reference in the text');
    const text = m.blocks.map((b) => b.text || '').join('\n');
    for (const needle of ['A justified, indented paragraph', 'Figure 1', 'Column text, paragraph 6', '日本語のテキスト', 'العربية', 'Ελληνικά', 'This sentence was inserted with Track Changes on']) assert.ok(text.includes(needle), `the text carries "${needle}"`);
    assert.ok(!text.includes('these words are gone'), 'deleted text is not shown in the accepted view');
    assert.ok(/landscape/.test(JSON.stringify(m)), 'the landscape section is known');
  } finally {
    doc.close({ id: s.id });
  }
});
