/**
 * The document editing model.
 *
 * Caret arithmetic across runs and paragraph boundaries is where a text editor
 * goes subtly wrong, so it is tested here rather than clicked at. The fixture is
 * the bank-confirmation letter, because the interesting cases are the ones where
 * editing meets structure it must not destroy.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { locate, sliceRuns, removeRange, coalesce } from '@rutba/doc-view';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { OoxmlPackage, Document, comparePackages, parseRuns, withToggle, hasToggle } from '@rutba/ooxml';
import { buildComplexDocument, FRAGILE_DOC_PARTS, FRAGILE_DOC_ELEMENTS } from './fixtures/complex-document.js';

const LETTER = buildComplexDocument();
/** A plain document with mixed formatting, for the text-editing cases. */
const plain = () => {
  const doc = Document.open(buildComplexDocument());
  // paragraph 5 and 6 are plain list items; use those
  return openDocx(doc.save());
};

// -------------------------------------------------------------- positions --

test('locate maps a character offset onto the right run', () => {
  const runs = [{ text: 'Dear ' }, { text: 'Sir' }, { text: ' or Madam' }];
  assert.deepEqual(locate(runs, 0), { runIndex: 0, runOffset: 0, atEnd: false });
  assert.deepEqual(locate(runs, 3), { runIndex: 0, runOffset: 3, atEnd: false });
  // at a seam, insertion biases LEFT so typing inherits what came before
  assert.deepEqual(locate(runs, 5, 'left'), { runIndex: 0, runOffset: 5, atEnd: true });
  assert.deepEqual(locate(runs, 5, 'right'), { runIndex: 1, runOffset: 0, atEnd: false });
  assert.deepEqual(locate(runs, 17, 'left'), { runIndex: 2, runOffset: 9, atEnd: true });
});

test('slicing and removing preserve run properties', () => {
  const runs = [
    { rPr: null, text: 'Dear ' },
    { rPr: '<w:rPr><w:b/></w:rPr>', text: 'Sir' },
    { rPr: null, text: ' or Madam' },
  ];
  assert.deepEqual(sliceRuns(runs, 5, 8).map((r) => r.text), ['Sir']);
  assert.equal(sliceRuns(runs, 5, 8)[0].rPr, '<w:rPr><w:b/></w:rPr>');
  assert.deepEqual(sliceRuns(runs, 3, 7).map((r) => r.text), ['r ', 'Si']);
  assert.deepEqual(removeRange(runs, 5, 8).map((r) => r.text), ['Dear ', ' or Madam']);
});

test('coalesce merges runs that carry identical properties', () => {
  const runs = [
    { rPr: null, text: 'a' }, { rPr: null, text: 'b' },
    { rPr: '<w:rPr><w:b/></w:rPr>', text: 'c' }, { rPr: null, text: 'd' }, { rPr: null, text: '' },
  ];
  const merged = coalesce(runs);
  assert.deepEqual(merged.map((r) => r.text), ['ab', 'c', 'd'], 'empty runs are dropped too');
});

// ---------------------------------------------------------------- reading --

test('a real letter opens into editable blocks', () => {
  const view = openDocx(LETTER);
  const frame = view.render();
  // 9 prose paragraphs plus the table's 13 cell paragraphs, in document order
  assert.equal(frame.blocks.length, 22);
  assert.equal(frame.blocks.filter((b) => b.container === null).length, 9);
  assert.equal(frame.blocks[0].text, 'Confirmation of balance');
  assert.equal(frame.blocks[0].style, 'Heading1');
  assert.equal(frame.blocks[0].runs[0].bold, true);
  assert.equal(frame.blocks[5].style, 'ListParagraph');
  assert.equal(frame.blocks[7].text, 'Description', 'the header row follows the prose before the table');
  assert.ok(frame.wordCount > 20);
});

test('paragraphs carrying structure are flagged, not silently editable', () => {
  const frame = openDocx(LETTER).render();
  const withField = frame.blocks.find((b) => b.text.startsWith('Ref:'));
  assert.equal(withField.structural, true);
  assert.ok(withField.structuralTags.includes('w:fldSimple'));

  const plainItem = frame.blocks.find((b) => b.text === 'Payments received');
  assert.equal(plainItem.structural, false);
});

// ---------------------------------------------------------------- editing --

test('typing inserts into the right run and inherits its formatting', () => {
  const view = plain();
  view.collapseTo({ block: 0, offset: 'Confirmation'.length });
  view.insertText(' AND AUDIT');

  const b = view.block(0);
  assert.equal(b.text, 'Confirmation AND AUDIT of balance');
  assert.equal(b.runs.length, 1, 'inserting into a uniform run does not fragment it');
  assert.equal(b.runs[0].bold, true, 'the inserted text inherited bold');
  assert.deepEqual(view.focus, { block: 0, offset: 22 });
});

test('typing after a formatted run inherits what came before the caret', () => {
  const view = plain();
  // paragraph 2: "Dear Sir or Madam," — one italic run
  view.collapseTo({ block: 5, offset: 0 });
  view.insertText('First: ');
  assert.equal(view.block(5).text, 'First: Invoices raised in the period');
  assert.equal(view.block(5).runs.length, 1);
});

test('backspace deletes a character, then merges paragraphs at offset 0', () => {
  const view = plain();
  view.collapseTo({ block: 6, offset: 3 });
  view.deleteBackward();
  // backspace removes the character BEFORE the caret: offset 3 eats the 'y'
  assert.equal(view.block(6).text, 'Paments received');
  assert.deepEqual(view.focus, { block: 6, offset: 2 });

  view.collapseTo({ block: 6, offset: 0 });
  view.deleteBackward();
  assert.equal(view.block(5).text, 'Invoices raised in the periodPaments received');
  assert.deepEqual(view.focus, { block: 5, offset: 29 }, 'the caret lands at the join');
});

test('delete forward pulls the next paragraph up at the end of one', () => {
  const view = plain();
  const before = view.blocks.length;
  view.collapseTo({ block: 5, offset: view.block(5).text.length });
  view.deleteForward();
  assert.equal(view.blocks.length, before - 1);
  assert.equal(view.block(5).text, 'Invoices raised in the periodPayments received');
});

test('Enter splits a paragraph and the new one inherits the style', () => {
  const view = plain();
  view.collapseTo({ block: 5, offset: 8 });
  view.splitParagraph();

  assert.equal(view.block(5).text, 'Invoices');
  assert.equal(view.block(6).text, ' raised in the period');
  assert.equal(view.block(6).style, 'ListParagraph', 'Enter in a list item makes another list item');
  assert.deepEqual(view.focus, { block: 6, offset: 0 });
});

test('a selection within one paragraph is replaced by typing', () => {
  const view = plain();
  view.setSelection({ block: 6, offset: 0 }, { block: 6, offset: 8 });
  view.insertText('Receipts');
  assert.equal(view.block(6).text, 'Receipts received');
  assert.ok(view.collapsed);
});

test('a selection spanning paragraphs deletes and joins them', () => {
  const view = plain();
  const before = view.blocks.length;
  view.setSelection({ block: 5, offset: 8 }, { block: 6, offset: 8 });
  view.deleteSelection();
  assert.equal(view.blocks.length, before - 1);
  assert.equal(view.block(5).text, 'Invoices received');
  assert.deepEqual(view.focus, { block: 5, offset: 8 });
});

test('typing into a structural paragraph is refused with a reason', () => {
  const view = openDocx(LETTER);
  view.collapseTo({ block: 1, offset: 4 }); // the Ref: field paragraph
  assert.equal(view.canEdit, false);
  assert.throws(() => view.insertText('x'), /w:fldSimple.*content control/s);
  assert.throws(() => view.splitParagraph(), /cannot be typed into/);
});

test('a content control can still be filled inside a structural paragraph', () => {
  const view = openDocx(LETTER);
  view.setContentControl('RUTBA_AR_TOTAL', '412,000.00');
  const controls = Object.fromEntries(view.contentControls().map((c) => [c.tag, c.text]));
  assert.equal(controls.RUTBA_AR_TOTAL, '412,000.00');
});

// ------------------------------------------------------------- formatting --

test('bold over a selection splits runs at the boundaries', () => {
  const view = plain();
  view.setSelection({ block: 6, offset: 0 }, { block: 6, offset: 8 });
  view.toggleFormat('b');

  const runs = view.block(6).runs;
  assert.equal(runs.length, 2);
  assert.deepEqual(runs.map((r) => r.text), ['Payments', ' received']);
  assert.equal(runs[0].bold, true);
  assert.equal(runs[1].bold, false);
});

test('toggling a mixed selection turns it all on, then all off', () => {
  const view = plain();
  view.setSelection({ block: 6, offset: 0 }, { block: 6, offset: 8 });
  view.toggleFormat('b');

  view.setSelection({ block: 6, offset: 0 }, { block: 6, offset: 17 });
  view.toggleFormat('b');
  assert.ok(view.block(6).runs.every((r) => r.bold), 'a mixed selection goes all-on first');

  view.toggleFormat('b');
  assert.ok(view.block(6).runs.every((r) => !r.bold), 'a fully-on selection goes off');
  assert.equal(view.block(6).runs.length, 1, 'and the runs coalesce back to one');
});

test('a toggle on a collapsed caret arms the next keystroke', () => {
  const view = plain();
  view.collapseTo({ block: 6, offset: 8 });
  view.toggleFormat('b');
  assert.equal(view.formatAtCaret().bold, true, 'the toolbar shows it armed');

  view.insertText('X');
  const runs = view.block(6).runs;
  const boldRun = runs.find((r) => r.text === 'X');
  assert.ok(boldRun, 'the typed character exists');
  assert.equal(boldRun.bold, true, 'and it came out bold');
  assert.equal(runs.find((r) => r.text === 'Payments').bold, false, 'neighbours unaffected');
});

test('formatting spanning paragraphs applies to each', () => {
  const view = plain();
  view.setSelection({ block: 5, offset: 3 }, { block: 6, offset: 4 });
  view.toggleFormat('i');
  assert.ok(view.block(5).runs.some((r) => r.italic && r.text.startsWith('oices')));
  assert.ok(view.block(6).runs.some((r) => r.italic && r.text === 'Paym'));
});

test('an rPr keeps everything we do not understand when a toggle changes', () => {
  const rPr = '<w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="32"/><w:color w:val="1F5F8B"/></w:rPr>';
  const bolded = withToggle(rPr, 'b', true);
  assert.ok(hasToggle(bolded, 'b'));
  assert.ok(bolded.includes('w:rFonts'), 'font survived');
  assert.ok(bolded.includes('w:color'), 'colour survived');
  assert.ok(bolded.includes('w:sz'), 'size survived');
  const unbolded = withToggle(bolded, 'b', false);
  assert.equal(hasToggle(unbolded, 'b'), false);
  assert.ok(unbolded.includes('w:color'));
});

// -------------------------------------------------------- caret and copy --

test('arrow keys cross paragraph boundaries', () => {
  const view = plain();
  view.collapseTo({ block: 5, offset: view.block(5).text.length });
  view.moveCaret('right');
  assert.deepEqual(view.focus, { block: 6, offset: 0 });
  view.moveCaret('left');
  assert.deepEqual(view.focus, { block: 5, offset: view.block(5).text.length });

  view.moveCaret('home');
  assert.equal(view.focus.offset, 0);
  view.moveCaret('end');
  assert.equal(view.focus.offset, view.block(5).text.length);
});

test('shift+arrow extends and the anchor stays', () => {
  const view = plain();
  view.collapseTo({ block: 6, offset: 0 });
  view.moveCaret('right', { extend: true });
  view.moveCaret('right', { extend: true });
  assert.deepEqual(view.selection.from, { block: 6, offset: 0 });
  assert.deepEqual(view.selection.to, { block: 6, offset: 2 });
  assert.equal(view.collapsed, false);
});

test('copy and paste round-trip across paragraphs', () => {
  const view = plain();
  view.setSelection({ block: 5, offset: 0 }, { block: 6, offset: 8 });
  assert.equal(view.copyText(), 'Invoices raised in the period\nPayments');

  const target = plain();
  target.collapseTo({ block: 6, offset: target.block(6).text.length });
  target.pasteText('one\ntwo');
  assert.equal(target.block(6).text, 'Payments receivedone');
  assert.equal(target.block(7).text, 'two');
});

// ------------------------------------------------------------ persistence --

test('editing rewrites only word/document.xml', () => {
  const view = openDocx(LETTER);
  view.collapseTo({ block: 5, offset: 0 });
  view.insertText('First: ');
  const out = view.save();

  const diff = comparePackages(LETTER, out);
  assert.deepEqual(diff.changed.map((c) => c.name), ['word/document.xml']);
  assert.equal(diff.removed.length, 0);
  for (const part of FRAGILE_DOC_PARTS) assert.ok(diff.identical.includes(part), part + ' disturbed');
});

test('everything structural survives an edit elsewhere in the letter', () => {
  const view = openDocx(LETTER);
  view.collapseTo({ block: 6, offset: 0 });
  view.insertText('Second: ');
  const after = OoxmlPackage.read(view.save()).text('word/document.xml');

  for (const el of FRAGILE_DOC_ELEMENTS) assert.ok(after.includes('<' + el), el + ' lost');
  assert.ok(after.includes('DOCPROPERTY'), 'the reference field survived');
  assert.ok(after.includes('w:name="Salutation"'), 'the bookmark survived');
  assert.ok(after.includes('<w:pgMar w:top="1440"'), 'the letterhead margins survived');
  assert.ok(after.includes('<w:tag w:val="RUTBA_AR_TOTAL"/>'), 'the content control survived');
});

test('a saved document reopens with the same text and formatting', () => {
  const view = openDocx(LETTER);
  view.setSelection({ block: 6, offset: 0 }, { block: 6, offset: 8 });
  view.toggleFormat('b');
  view.collapseTo({ block: 6, offset: 0 });
  view.insertText('>> ');

  const reopened = openDocx(view.save());
  assert.equal(reopened.block(6).text, '>> Payments received');
  const runs = reopened.block(6).runs;
  assert.ok(runs.some((r) => r.text.includes('Payments') && r.bold));
  assert.equal(reopened.block(0).runs[0].bold, true, 'the heading is untouched');
});

test('twenty keystrokes do not fragment the paragraph', () => {
  const view = plain();
  view.collapseTo({ block: 6, offset: 0 });
  for (const ch of 'abcdefghijklmnopqrst') view.insertText(ch);
  assert.equal(view.block(6).text, 'abcdefghijklmnopqrstPayments received');
  assert.equal(view.block(6).runs.length, 1, 'coalescing keeps it one run, not twenty');

  const reopened = openDocx(view.save());
  assert.equal(parseRuns(reopened.block(6).xml).length, 1);
});


// ---------------------------------------------------------------------------
// Tables and the page — drawn in place and, since D5, edited in place.
// ---------------------------------------------------------------------------

test('the flow puts a table back where the author wrote it', () => {
  const view = openDocx(LETTER);
  const frame = view.render();

  // The paragraph list is the caret's address space, and since D5 it reaches
  // into the table: cell paragraphs are blocks, each naming its cell. What
  // keeps a caret safe is no longer exclusion but the container discipline —
  // the boundary guards are tested below.
  const opening = frame.blocks.find((b) => b.text === 'Opening balance');
  assert.match(opening.container, /^t\d+:r1:c0$/);
  assert.equal(frame.blocks[0].container, null, 'prose has no container');

  const kinds = frame.flow.map((e) => e.kind);
  assert.ok(kinds.includes('table'), 'the flow carries the table');
  const tableAt = kinds.indexOf('table');
  assert.ok(tableAt > 0 && tableAt < kinds.length - 1, 'with prose on both sides');

  // every paragraph entry points at a real PROSE paragraph, in order; the
  // cell paragraphs are addressed through the table entry's blockIndex, not
  // duplicated as flow entries
  const paras = frame.flow.filter((e) => e.kind === 'paragraph').map((e) => e.paragraphIndex);
  assert.deepEqual(paras, [...paras].sort((a, b) => a - b));
  assert.deepEqual(paras, frame.blocks.filter((b) => b.container === null).map((b) => b.index));
});

test('a table reports its structure rather than resolving it', () => {
  const { table } = openDocx(LETTER).render().flow.find((e) => e.kind === 'table');

  assert.equal(table.rowCount, 4);
  assert.equal(table.columnCount, 3);
  assert.equal(table.widthPct, 100);              // 5000 fiftieths of a percent
  assert.deepEqual(table.columns.map(Math.round), [301, 150, 150]); // 4508 twips = 300.5px

  // a header row repeats across pages and reads differently; say so
  assert.equal(table.rows[0].header, true);
  assert.equal(table.rows[1].header, false);
  assert.equal(table.rows[0].cells[0].shading, '#d9e2f3');
  assert.equal(table.rows[0].cells[2].blocks[0].align, 'right');

  // merges are REPORTED, not resolved — only the renderer knows how many
  // continuation rows followed, so it computes the rowspan
  assert.equal(table.rows[1].cells[0].vMerge, 'restart');
  assert.equal(table.rows[2].cells[0].vMerge, 'continue');
  assert.equal(table.rows[3].cells[0].gridSpan, 2);

  // insideH is what actually draws a TableGrid's rules
  assert.equal(table.borders.top.style, 'single');
  assert.equal(table.borders.insideH.colour, '#d9d9d9');
});

test('a nested table stays nested and keeps its siblings in order', () => {
  const { table } = openDocx(LETTER).render().flow.find((e) => e.kind === 'table');
  const cell = table.rows[2].cells[1];

  // paragraph, nested table, paragraph — document order, not grouped by kind
  assert.deepEqual(cell.blocks.map((b) => b.kind), ['paragraph', 'table', 'paragraph']);
  assert.equal(cell.blocks[0].text, 'INV-1002');
  assert.equal(cell.blocks[2].text, 'see note 4');
  assert.equal(cell.blocks[1].table.rowCount, 1);
  assert.equal(cell.blocks[1].table.rows[0].cells[0].text, 'part-paid');

  // and the inner table's row did NOT leak into the outer one
  assert.equal(table.rowCount, 4);
});

test('the page comes from the document, not from us', () => {
  const section = openDocx(LETTER).render().section;
  assert.equal(Math.round(section.widthPx), 794);    // A4: 11906 twips
  assert.equal(Math.round(section.heightPx), 1123);
  assert.equal(section.orientation, 'portrait');
  assert.equal(Math.round(section.margins.top), 96); // 1440 twips = 1 inch
  assert.equal(Math.round(section.margins.header), 47);
  assert.equal(Math.round(section.contentWidthPx), 602);
  assert.equal(section.declared, true);
});

test('editing prose leaves the table byte-identical', () => {
  const view = openDocx(LETTER);
  const before = view.render().flow.find((e) => e.kind === 'table').table;

  const target = view.render().blocks.findIndex((b) => !b.structural && b.text.length > 4);
  view.setSelection({ block: target, offset: 0 });
  view.insertText('NOTE: ');
  const saved = view.save();

  const after = openDocx(saved).render();
  assert.ok(after.blocks[target].text.startsWith('NOTE: '));
  assert.deepEqual(after.flow.find((e) => e.kind === 'table').table, before);

  // and the XML says the same thing: the table element was never rewritten
  const tableXml = (buf) => /<w:tbl>[\s\S]*<\/w:tbl>/.exec(OoxmlPackage.read(buf).text('word/document.xml'))[0];
  assert.equal(tableXml(saved), tableXml(LETTER));
});

test('a backend with no tables still renders — the flow degrades to paragraphs', async () => {
  const { openHtml } = await import('@rutba/doc-view/backends/html');
  const frame = openHtml('<p>Thanks for the quote.</p><p>Approved.</p>').render();

  assert.deepEqual(frame.flow.map((e) => e.kind), ['paragraph', 'paragraph']);
  assert.deepEqual(frame.flow.map((e) => e.paragraphIndex), [0, 1]);
  // an email body has no page, and saying so is better than inventing A4
  assert.equal(frame.section, null);
});


test('a caret the client could not locate never becomes paragraph zero', () => {
  // Found in the browser the day tables became visible: a <p> inside a table
  // cell was matched as a top-level paragraph, its block index came through as
  // NaN, and typing edited the FIRST paragraph of the document instead. A lost
  // caret must clamp to a defined position, never silently pick one.
  const view = openDocx(LETTER);
  const first = view.render().blocks[0].text;

  view.setSelection({ block: NaN, offset: NaN });
  const at = view.render().selection.focus;
  assert.equal(at.block, 0);
  assert.equal(at.offset, 0);
  assert.ok(Number.isInteger(at.block) && Number.isInteger(at.offset), 'no NaN escapes');

  // and the same for a wildly out-of-range index
  view.setSelection({ block: 9999, offset: 9999 });
  const far = view.render().selection.focus;
  assert.equal(far.block, view.render().blocks.length - 1);
  assert.equal(far.offset, view.render().blocks.at(-1).text.length);

  assert.equal(view.render().blocks[0].text, first, 'clamping alone edits nothing');
});
