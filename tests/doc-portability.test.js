/**
 * The proof that the editor is shareable.
 *
 * `DocView` is the editing brain. If it is genuinely format-free, the same
 * caret arithmetic, the same selection semantics and the same formatting toggles
 * must drive an email body as well as they drive a `.docx` — with no branch
 * anywhere in the editor that says "if this is HTML".
 *
 * So this file runs the same sequence of edits against BOTH backends and asserts
 * the same resulting text. If Mail ever needs a behaviour Workspace does not
 * have, that is a signal to extend the port, not to fork the editor.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { DocView, assertBackend, REQUIRED_METHODS, supportsContentControls } from '@rutba/doc-view';
import { HtmlBackend, openHtml, sanitise } from '@rutba/doc-view/backends/html';
import { OoxmlBackend, openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildComplexDocument } from './fixtures/complex-document.js';
import { buildDocx } from '@rutba/ooxml';

// ------------------------------------------------------------------ port --

test('the port refuses an incomplete backend at construction', () => {
  assert.throws(() => new DocView({}), /missing: paragraphCount/);
  assert.throws(() => new DocView(null), /a document backend is required/);
  assert.throws(() => assertBackend({ paragraphCount: () => 0 }), /missing: paragraph,/);
  assert.ok(REQUIRED_METHODS.includes('toggleRunFormat'));
});

test('both shipped backends satisfy the port', () => {
  assert.doesNotThrow(() => assertBackend(HtmlBackend.open('<p>hello</p>')));
  assert.doesNotThrow(() => assertBackend(OoxmlBackend.open(buildComplexDocument())));
});

test('content controls are optional, and the editor says so rather than crashing', () => {
  const html = openHtml('<p>hello</p>');
  assert.equal(html.supportsContentControls, false);
  assert.deepEqual(html.contentControls(), []);
  assert.throws(() => html.setContentControl('X', 'y'), /no content controls/);

  const docx = openDocx(buildComplexDocument());
  assert.equal(docx.supportsContentControls, true);
  assert.ok(docx.contentControls().length > 0);
});

// ------------------------------------------------------------- HTML side --

test('an email body parses into blocks and runs', () => {
  const view = openHtml('<h1>Quarterly update</h1><p>Revenue is <b>up 14%</b> this quarter.</p><li>North</li>');
  assert.equal(view.blocks.length, 3);
  assert.equal(view.block(0).text, 'Quarterly update');
  assert.equal(view.block(0).style, 'Heading1');
  assert.equal(view.block(1).text, 'Revenue is up 14% this quarter.');
  assert.deepEqual(view.block(1).runs.map((r) => r.bold), [false, true, false]);
  assert.equal(view.block(2).style, 'ListParagraph');
});

test('nested inline tags accumulate', () => {
  const view = openHtml('<p>plain <b>bold <i>both</i></b></p>');
  const runs = view.block(0).runs;
  const both = runs.find((r) => r.text === 'both');
  assert.equal(both.bold, true);
  assert.equal(both.italic, true);
});

test('an email body is untrusted input and is sanitised', () => {
  const nasty = '<p onclick="steal()">hi</p><script>steal()</script><p><a href="javascript:go()">x</a></p>';
  const clean = sanitise(nasty);
  assert.ok(!clean.includes('<script'));
  assert.ok(!clean.includes('onclick'));
  assert.ok(!clean.includes('javascript:'));

  const view = openHtml(nasty);
  const out = view.save();
  assert.ok(!out.includes('script'));
  assert.ok(!out.includes('onclick'));
});

test('HTML output is conservative and escaped', () => {
  const view = openHtml('<p>a &amp; b</p>');
  view.collapseTo({ block: 0, offset: 5 });
  view.insertText(' <tag>');
  const out = view.save();
  assert.match(out, /^<p>/);
  assert.ok(out.includes('&amp;'), 'the ampersand stays escaped');
  assert.ok(out.includes('&lt;tag&gt;'), 'typed angle brackets are escaped, not injected');
  assert.ok(!out.includes('<tag>'));
});

test('an email carries a plain-text alternative', () => {
  const backend = HtmlBackend.open('<h1>Title</h1><p>Body <b>text</b></p>');
  assert.equal(backend.toPlainText(), 'Title\nBody text');
});

// ------------------------------------------------ the same edits, both ways --

/** One script of edits, run against whatever backend is handed in. */
function runEditScript(view) {
  view.collapseTo({ block: 1, offset: 0 });
  view.insertText('Note: ');

  // bold a word in the middle
  view.setSelection({ block: 1, offset: 6 }, { block: 1, offset: 13 });
  view.toggleFormat('b');

  // split, then merge back — the pair must be a no-op on the text
  view.collapseTo({ block: 1, offset: 6 });
  view.splitParagraph();
  view.collapseTo({ block: 2, offset: 0 });
  view.deleteBackward();

  // type at the end with an armed format
  view.collapseTo({ block: 1, offset: view.block(1).text.length });
  view.toggleFormat('i');
  view.insertText(' (italic tail)');

  return {
    text: view.blocks.map((b) => b.text),
    boldRuns: view.blocks.flatMap((b) => b.runs.filter((r) => r.bold).map((r) => r.text)),
    italicRuns: view.blocks.flatMap((b) => b.runs.filter((r) => r.italic).map((r) => r.text)),
    caret: { ...view.focus },
  };
}

test('the identical edit script gives identical results on both backends', () => {
  // The SAME two paragraphs, in two formats. No index shifting, no special
  // casing — if the editor is format-free, the script cannot tell them apart.
  const html = openHtml('<p>Confirmation of balance</p><p>Payments received</p>');
  const docx = openDocx(buildDocx({ paragraphs: ['Confirmation of balance', 'Payments received'] }));

  assert.deepEqual(
    html.blocks.map((b) => b.text),
    docx.blocks.map((b) => b.text),
    'the two documents start identical',
  );

  const htmlResult = runEditScript(html);
  const docxResult = runEditScript(docx);

  assert.deepEqual(htmlResult.text, docxResult.text, 'the same edits produce the same text');
  assert.deepEqual(htmlResult.boldRuns, docxResult.boldRuns, 'the same runs came out bold');
  assert.deepEqual(htmlResult.italicRuns, docxResult.italicRuns, 'and the same ones italic');
  assert.deepEqual(htmlResult.caret, docxResult.caret, 'the caret ended in the same place');

  // and the result is what a person would expect, not merely self-consistent
  assert.deepEqual(htmlResult.text, [
    'Confirmation of balance',
    'Note: Payments received (italic tail)',
  ]);
  assert.deepEqual(htmlResult.boldRuns, ['Payment']);
  assert.deepEqual(htmlResult.italicRuns, [' (italic tail)']);
});

test('split then merge is a no-op on the text, in both formats', () => {
  for (const view of [
    openHtml('<p>one two three</p>'),
    openDocx(buildDocx({ paragraphs: ['one two three'] })),
  ]) {
    const before = view.block(0).text;
    view.collapseTo({ block: 0, offset: 3 });
    view.splitParagraph();
    view.collapseTo({ block: 1, offset: 0 });
    view.deleteBackward();
    assert.equal(view.block(0).text, before);
  }
});

test('coalescing works on both, because rPr is compared not parsed', () => {
  const html = openHtml('<p>abc</p>');
  html.collapseTo({ block: 0, offset: 0 });
  for (const ch of 'xyz') html.insertText(ch);
  assert.equal(html.block(0).text, 'xyzabc');
  assert.equal(html.block(0).runs.length, 1, 'three keystrokes, one run');

  const docx = openDocx(buildComplexDocument());
  docx.collapseTo({ block: 6, offset: 0 });
  for (const ch of 'xyz') docx.insertText(ch);
  assert.equal(docx.block(6).runs.length, 1);
});

test('an email composed with the editor round-trips through its own parser', () => {
  const view = openHtml('<p>Dear customer,</p><p>Your balance is due.</p>');
  view.setSelection({ block: 1, offset: 5 }, { block: 1, offset: 12 });
  view.toggleFormat('b');
  const html = view.save();
  assert.match(html, /<b>balance<\/b>/);

  const reopened = openHtml(html);
  assert.equal(reopened.block(1).text, 'Your balance is due.');
  assert.ok(reopened.block(1).runs.find((r) => r.text === 'balance').bold);
  assert.equal(reopened.save(), html, 'a second round trip is stable');
});

test('the html backend never reports a paragraph as structural', () => {
  const view = openHtml('<p>anything at all</p>');
  assert.equal(view.canEdit, true);
  assert.equal(view.block(0).structural, false);
  assert.doesNotThrow(() => view.insertText('x'));
});

test('supportsContentControls is a function of the backend, not the view', () => {
  assert.equal(supportsContentControls(HtmlBackend.open('<p>x</p>')), false);
  assert.equal(supportsContentControls(OoxmlBackend.open(buildComplexDocument())), true);
});
