// Tracked changes — the engine's half.
//
// Review → Track Changes turns recording on: typed text becomes `w:ins`
// (an id, an author, a date), a deletion becomes `w:del` with `w:delText`
// rather than disappearing — unless it is the SAME author's own still-open
// insertion, which Word simply un-inserts. Accept keeps an insertion's
// words and drops a deletion's; Reject is the reverse. A paragraph carrying
// either stays editable, the way a field or a bookmark now does.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';

const fixture = () => openDocx(buildDocx({
  styles: true,
  paragraphs: [
    { text: 'First paragraph text' },
    { text: 'Second paragraph text' },
    { text: 'Third paragraph text' },
  ],
}));

/** Recording on, author "Kim", caret at the end of the first paragraph. */
const recording = () => {
  const view = fixture();
  view.setTrackChanges(true, 'Kim');
  view.setSelection({ block: 0, offset: view.blocks[0].text.length });
  return view;
};

test('Track Changes writes trackRevisions in settings.xml, and reads it back', () => {
  const view = recording();
  assert.equal(view.doc.doc.trackRevisions(), true);
  assert.equal(view.recording, true);

  const reopened = openDocx(view.save());
  assert.equal(reopened.doc.doc.trackRevisions(), true);
  assert.equal(reopened.recording, true, 'a file that has it opens recording');
});

test('turning Track Changes off removes the setting again', () => {
  const view = recording();
  view.setTrackChanges(false);
  assert.equal(view.doc.doc.trackRevisions(), false);
  assert.equal(openDocx(view.save()).doc.doc.trackRevisions(), false);
});

test('typed text while recording is written as w:ins, with an id, author and date', () => {
  const view = recording();
  view.insertText('X');

  const run = view.blocks[0].runs.find((r) => r.ins);
  assert.ok(run, 'the typed text carries ins');
  assert.equal(run.text, 'X');
  assert.equal(run.ins.author, 'Kim');
  assert.ok(run.ins.date);
  assert.ok(view.blocks[0].text.endsWith('X'), 'the text is part of the paragraph, caret and all');

  const pXml = view.doc.doc.editParagraph(0).xml;
  assert.match(pXml, /<w:ins w:id="\d+" w:author="Kim" w:date="[^"]+"><w:r><w:t xml:space="preserve">X<\/w:t><\/w:r><\/w:ins>/);
});

test('typed text with recording off is plain, as before', () => {
  const view = fixture();
  view.setSelection({ block: 0, offset: view.blocks[0].text.length });
  view.insertText('X');
  assert.equal(view.blocks[0].runs.some((r) => r.ins), false);
});

test('deleting existing words while recording marks them w:del with delText, rather than removing them', () => {
  const view = recording();
  // Delete "text" off the end of "First paragraph text".
  view.setSelection({ block: 0, offset: 16 }, { block: 0, offset: 20 });
  view.deleteSelection();

  // The words are still THERE for review, but not part of the live text.
  assert.equal(view.blocks[0].text, 'First paragraph ');
  const del = view.blocks[0].runs.find((r) => r.del);
  assert.ok(del, 'a del run is present');
  assert.equal(del.del.text, 'text');
  assert.equal(del.del.author, 'Kim');
  assert.equal(del.text, '', 'a deletion is zero characters wide in the editable text');

  const pXml = view.doc.doc.editParagraph(0).xml;
  assert.match(pXml, /<w:del w:id="\d+" w:author="Kim" w:date="[^"]+"><w:r><w:delText xml:space="preserve">text<\/w:delText><\/w:r><\/w:del>/);
});

test('deleting a word inserted earlier in the SAME pending insertion just removes it — no w:del', () => {
  const view = recording();
  view.insertText('ABC');
  assert.equal(view.blocks[0].text, 'First paragraph textABC');

  view.setSelection({ block: 0, offset: view.blocks[0].text.length - 1 }, { block: 0, offset: view.blocks[0].text.length });
  view.deleteSelection();

  assert.equal(view.blocks[0].text, 'First paragraph textAB');
  assert.equal(view.blocks[0].runs.some((r) => r.del), false, 'un-inserted, not marked deleted');
  const ins = view.blocks[0].runs.find((r) => r.ins);
  assert.equal(ins.text, 'AB');
});

test('Accept this change keeps an insertion\'s words and drops a deletion\'s', () => {
  const view = recording();
  view.insertText('X');
  view.setSelection({ block: 1, offset: 0 }, { block: 1, offset: 7 });
  view.deleteSelection();
  assert.equal(view.blocks[1].text, 'paragraph text');

  view.setSelection({ block: 0, offset: 0 });
  view.acceptChanges({});
  assert.equal(view.blocks[0].text, 'First paragraph textX');
  assert.equal(view.blocks[0].runs.some((r) => r.ins || r.del), false);
  // The second paragraph's deletion is untouched — only "this change" (block 0) was accepted.
  assert.ok(view.blocks[1].runs.some((r) => r.del));

  view.setSelection({ block: 1, offset: 0 });
  view.acceptChanges({});
  assert.equal(view.blocks[1].text, 'paragraph text');
  assert.equal(view.blocks[1].runs.some((r) => r.del), false);
});

test('Reject this change removes an insertion\'s words and brings a deletion\'s back', () => {
  const view = recording();
  view.insertText('X');
  view.setSelection({ block: 1, offset: 0 }, { block: 1, offset: 7 });
  view.deleteSelection();

  view.setSelection({ block: 0, offset: 0 });
  view.rejectChanges({});
  assert.equal(view.blocks[0].text, 'First paragraph text');
  assert.equal(view.blocks[0].runs.some((r) => r.ins), false);

  view.setSelection({ block: 1, offset: 0 });
  view.rejectChanges({});
  assert.equal(view.blocks[1].text, 'Second paragraph text');
  assert.equal(view.blocks[1].runs.some((r) => r.del), false);
});

test('Accept All and Reject All sweep every paragraph at once', () => {
  const view = recording();
  view.insertText('X');
  view.setSelection({ block: 1, offset: 0 });
  view.insertText('Y');
  view.setSelection({ block: 2, offset: 0 }, { block: 2, offset: 6 });
  view.deleteSelection();

  view.acceptChanges({ all: true });
  assert.equal(view.blocks[0].text, 'First paragraph textX');
  assert.equal(view.blocks[1].text, 'YSecond paragraph text');
  assert.equal(view.blocks[2].text, 'paragraph text');
  assert.equal(view.blocks.some((b) => b.runs.some((r) => r.ins || r.del)), false);
});

test('a paragraph carrying a tracked change stays editable — typing elsewhere in it works', () => {
  const view = recording();
  view.setSelection({ block: 1, offset: 0 }, { block: 1, offset: 7 });
  view.deleteSelection();
  assert.equal(view.blocks[1].structural, false, 'a tracked change no longer locks the paragraph');

  view.setSelection({ block: 1, offset: view.blocks[1].text.length });
  view.insertText('!');
  assert.equal(view.blocks[1].text, 'paragraph text!');
  assert.ok(view.blocks[1].runs.some((r) => r.del), 'the earlier deletion survived the later edit');
});

test('tracked changes survive a save and reopen, and can still be accepted there', () => {
  const view = recording();
  view.insertText('X');
  view.setSelection({ block: 1, offset: 0 }, { block: 1, offset: 7 });
  view.deleteSelection();

  const reopened = openDocx(view.save());
  assert.ok(reopened.blocks[0].runs.some((r) => r.ins?.text !== undefined || r.ins), 'insertion round-tripped');
  assert.equal(reopened.blocks[0].text, 'First paragraph textX');
  assert.ok(reopened.blocks[1].runs.some((r) => r.del));
  assert.equal(reopened.blocks[1].text, 'paragraph text');

  reopened.acceptChanges({ all: true });
  assert.equal(reopened.blocks.some((b) => b.runs.some((r) => r.ins || r.del)), false);
});

test('a Word-authored w:ins/w:del fixture is read, shown, and accepted', () => {
  const view = fixture();
  const engine = view.doc.doc;
  const p = engine.editParagraph(0);
  const handBuilt = p.open + (p.pPr || '') +
    '<w:r><w:t xml:space="preserve">First paragraph </w:t></w:r>' +
    '<w:ins w:id="9" w:author="Priya" w:date="2026-01-01T00:00:00Z"><w:r><w:t xml:space="preserve">brand new </w:t></w:r></w:ins>' +
    '<w:del w:id="10" w:author="Priya" w:date="2026-01-01T00:00:00Z"><w:r><w:delText xml:space="preserve">text</w:delText></w:r></w:del>' +
    '</w:p>';
  engine._spliceBody(p.start, p.end, handBuilt);

  const reopened = openDocx(view.save());
  assert.equal(reopened.blocks[0].text, 'First paragraph brand new ');
  const ins = reopened.blocks[0].runs.find((r) => r.ins);
  const del = reopened.blocks[0].runs.find((r) => r.del);
  assert.equal(ins.ins.author, 'Priya');
  assert.equal(del.del.text, 'text');
  assert.equal(reopened.blocks[0].structural, false);

  reopened.setSelection({ block: 0, offset: 0 });
  reopened.acceptChanges({ all: true });
  assert.equal(reopened.blocks[0].text, 'First paragraph brand new ');
  assert.equal(reopened.blocks[0].runs.some((r) => r.ins || r.del), false);
});

test('undo after a tracked insertion takes it off again, and after a tracked deletion restores the words', () => {
  const view = recording();
  view.insertText('X');
  assert.ok(view.blocks[0].runs.some((r) => r.ins));
  view.undo();
  assert.equal(view.blocks[0].text, 'First paragraph text');
  assert.equal(view.blocks[0].runs.some((r) => r.ins), false);

  view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 5 });
  view.deleteSelection();
  assert.ok(view.blocks[0].runs.some((r) => r.del));
  view.undo();
  assert.equal(view.blocks[0].text, 'First paragraph text');
  assert.equal(view.blocks[0].runs.some((r) => r.del), false);
});
