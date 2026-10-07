// Tracked changes of formatting, and of paragraph marks — the engine's half.
//
// Word records a change of formatting as a `w:rPrChange` inside the run's
// `w:rPr`, holding the properties before; a change to a paragraph's own
// formatting as a `w:pPrChange`; and a paragraph mark put in or taken out as
// a self-closing `w:ins`/`w:del` in the mark's `w:rPr`. Each nests markup the
// engine's patterns used to cut short: a paragraph holding one was written
// back broken after a keystroke, or after Accept. Now they are read whole,
// recorded while Track Changes is on, and accepted or rejected as Word does.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx, OoxmlPackage } from '@rutba/ooxml';

const DATE = 'w:date="2026-01-01T00:00:00Z"';

/** A document whose first paragraph is `para` (raw XML), the second "Second paragraph". */
const withParagraph = (para) => {
  const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: [{ text: 'PLACEHOLDER' }, { text: 'Second paragraph' }] }));
  const xml = pkg.text('word/document.xml').replace(/<w:p\b(?:(?!<\/w:p>)[\s\S])*?PLACEHOLDER[\s\S]*?<\/w:p>/, () => para);
  pkg.write_('word/document.xml', xml);
  return openDocx(pkg.write());
};

const bodyOf = (view) => /<w:body>([\s\S]*)<\/w:body>/.exec(OoxmlPackage.read(view.save()).text('word/document.xml'))[1];

/** Every element closed in order — what a reader that refuses broken XML wants. */
const wellFormed = (xml) => {
  const stack = [];
  for (const m of xml.matchAll(/<(\/?)([\w:]+)[^>]*?(\/?)>/g)) {
    if (m[3] === '/') continue;
    if (m[1] === '/') { if (stack.pop() !== m[2]) return false; } else stack.push(m[2]);
  }
  return stack.length === 0;
};

const FORMATTED = `<w:p><w:r><w:t xml:space="preserve">Plain </w:t></w:r><w:r><w:rPr><w:b/><w:rPrChange w:id="7" w:author="Ann" ${DATE}><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>words</w:t></w:r><w:r><w:t xml:space="preserve"> here</w:t></w:r></w:p>`;

test('a run carrying Word\'s change of formatting reads as it looks now, and survives a keystroke whole', () => {
  const view = withParagraph(FORMATTED);
  const run = view.blocks[0].runs.find((r) => r.text === 'words');
  assert.equal(run.bold, true);
  assert.equal(run.italic, false, 'italic is how it looked before, not now');
  const frame = view.render({ pages: false }).blocks[0];
  const shown = frame.runs.find((r) => r.text === 'words');
  assert.deepEqual({ author: shown.formatChange.author, bold: shown.formatChange.was.bold, italic: shown.formatChange.was.italic }, { author: 'Ann', bold: false, italic: true });
  assert.equal(frame.tracked.formatted, 1);
  view.setSelection({ block: 0, offset: 0 });
  view.insertText('Some ');
  const xml = bodyOf(view);
  assert.ok(wellFormed(xml), 'the file is written whole');
  assert.match(xml, /<w:rPr><w:b\/><w:rPrChange w:id="7" w:author="Ann" w:date="2026-01-01T00:00:00Z"><w:rPr><w:i\/><\/w:rPr><\/w:rPrChange><\/w:rPr><w:t xml:space="preserve">words<\/w:t>/);
});

test('formatting while recording is written as a w:rPrChange holding how the words looked before', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'First paragraph text' }] }));
  view.setTrackChanges(true, 'Kim');
  view.setSelection({ block: 0, offset: 6 }, { block: 0, offset: 15 });
  view.toggleFormat('b');
  let xml = bodyOf(view);
  assert.match(xml, /<w:r><w:rPr><w:b\/><w:rPrChange w:id="\d+" w:author="Kim" w:date="[^"]+"><w:rPr><\/w:rPr><\/w:rPrChange><\/w:rPr><w:t xml:space="preserve">paragraph<\/w:t><\/w:r>/);
  assert.equal(view.render({ pages: false }).blocks[0].tracked.formatted, 1);
  // A second change is still from how the words first were.
  view.setSelection({ block: 0, offset: 6 }, { block: 0, offset: 15 });
  view.setRunFormat({ fontSize: 16 });
  xml = bodyOf(view);
  assert.match(xml, /<w:rPr><w:b\/><w:sz w:val="32"\/><w:rPrChange [^>]*><w:rPr><\/w:rPr><\/w:rPrChange><\/w:rPr>/);
  // Formatted back to how they were: no change left to record.
  view.setSelection({ block: 0, offset: 6 }, { block: 0, offset: 15 });
  view.setRunFormat({ fontSize: null });
  view.setSelection({ block: 0, offset: 6 }, { block: 0, offset: 15 });
  view.toggleFormat('b');
  xml = bodyOf(view);
  assert.ok(!/rPrChange/.test(xml), 'back as it was');
  assert.ok(!view.render({ pages: false }).blocks[0].tracked);
});

test('Accept keeps the new formatting and Reject puts the old back; Clear Formatting is recorded too', () => {
  const make = () => {
    const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Bold words', bold: true }] }));
    view.setTrackChanges(true, 'Kim');
    view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 4 });
    view.clearFormat();
    return view;
  };
  const cleared = make();
  assert.match(bodyOf(cleared), /<w:rPr><w:rPrChange [^>]*><w:rPr><w:b\/><\/w:rPr><\/w:rPrChange><\/w:rPr><w:t xml:space="preserve">Bold<\/w:t>/);
  assert.equal(cleared.blocks[0].runs[0].bold, false);

  const accepted = make();
  accepted.setSelection({ block: 0, offset: 0 });
  accepted.acceptChanges();
  assert.ok(!/rPrChange/.test(bodyOf(accepted)));
  assert.equal(accepted.blocks[0].runs.find((r) => r.text.startsWith('Bold'))?.bold, false);

  const rejected = make();
  rejected.setSelection({ block: 0, offset: 0 });
  rejected.rejectChanges();
  const xml = bodyOf(rejected);
  assert.ok(!/rPrChange/.test(xml) && wellFormed(xml));
  assert.ok(rejected.blocks[0].runs.filter((r) => r.text).every((r) => r.bold), 'bold again throughout');
});

test('words still in one\'s own pending insertion are formatted as part of it, not as a change', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Start' }] }));
  view.setTrackChanges(true, 'Kim');
  view.setSelection({ block: 0, offset: 5 });
  view.insertText(' added');
  view.setSelection({ block: 0, offset: 6 }, { block: 0, offset: 11 });
  view.toggleFormat('i');
  const xml = bodyOf(view);
  assert.ok(!/rPrChange/.test(xml));
  assert.match(xml, /<w:ins [^>]*><w:r><w:rPr><w:i\/><\/w:rPr><w:t xml:space="preserve">added<\/w:t><\/w:r><\/w:ins>/);
});

const MARK_IN = `<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:ins w:id="1" w:author="Ann" ${DATE}/></w:rPr></w:pPr><w:r><w:t xml:space="preserve">Kept </w:t></w:r><w:ins w:id="2" w:author="Ann" ${DATE}><w:r><w:t>new</w:t></w:r></w:ins></w:p>`;

test('a paragraph mark Word put in is not mistaken for inserted words, and Accept takes the mark\'s record away', () => {
  const view = withParagraph(MARK_IN);
  assert.deepEqual(view.blocks[0].runs.map((r) => [r.text, Boolean(r.ins)]), [['Kept ', false], ['new', true]]);
  assert.equal(view.render({ pages: false }).blocks[0].tracked.mark, 'inserted');
  view.setSelection({ block: 0, offset: 0 });
  view.acceptChanges();
  const xml = bodyOf(view);
  assert.ok(wellFormed(xml), 'written whole');
  assert.match(xml, /^<w:p><w:pPr><w:jc w:val="center"\/><\/w:pPr><w:r><w:t xml:space="preserve">Kept <\/w:t><\/w:r><w:r><w:t>new<\/w:t><\/w:r><\/w:p>/);
  assert.equal(view.blocks.length, 2);
});

test('Reject on an inserted paragraph mark joins the paragraph to the next, which keeps its own properties', () => {
  const view = withParagraph(MARK_IN);
  view.setSelection({ block: 0, offset: 0 });
  view.rejectChanges();
  const xml = bodyOf(view);
  assert.ok(wellFormed(xml));
  assert.equal(view.blocks.length, 1);
  assert.equal(view.blocks[0].text, 'Kept Second paragraph');
  assert.ok(!/<w:jc w:val="center"\/>/.test(xml), 'the joined paragraph has the next one\'s properties, as the mark held them');
});

test('Accept on a deleted paragraph mark joins the two paragraphs; Reject keeps them apart', () => {
  const para = `<w:p><w:pPr><w:rPr><w:del w:id="3" w:author="Ann" ${DATE}/></w:rPr></w:pPr><w:r><w:t xml:space="preserve">Joined </w:t></w:r></w:p>`;
  const accepted = withParagraph(para);
  accepted.acceptChanges({ all: true });
  assert.equal(accepted.blocks.length, 1);
  assert.equal(accepted.blocks[0].text, 'Joined Second paragraph');
  const rejected = withParagraph(para);
  rejected.rejectChanges({ all: true });
  assert.equal(rejected.blocks.length, 2);
  assert.ok(!/<w:del\b/.test(bodyOf(rejected)) && wellFormed(bodyOf(rejected)));
});

test('a change to a paragraph\'s own formatting is kept by Accept and put back by Reject, whole', () => {
  const para = `<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="4" w:author="Ann" ${DATE}><w:pPr><w:jc w:val="left"/><w:ind w:left="720"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>Moved</w:t></w:r></w:p>`;
  const view = withParagraph(para);
  assert.equal(view.doc.getParagraphProps(0).align, 'center', 'read as it is now');
  assert.equal(view.render({ pages: false }).blocks[0].tracked.formatted, 1);
  view.setSelection({ block: 0, offset: 5 });
  view.insertText('!');
  assert.ok(wellFormed(bodyOf(view)), 'a keystroke leaves the change whole');
  const rejected = withParagraph(para);
  rejected.rejectChanges({ all: true });
  assert.match(bodyOf(rejected), /^<w:p><w:pPr><w:jc w:val="left"\/><w:ind w:left="720"\/><\/w:pPr><w:r><w:t>Moved<\/w:t><\/w:r><\/w:p>/);
  const accepted = withParagraph(para);
  accepted.acceptChanges({ all: true });
  assert.match(bodyOf(accepted), /^<w:p><w:pPr><w:jc w:val="center"\/><\/w:pPr><w:r><w:t>Moved<\/w:t><\/w:r><\/w:p>/);
});

const deletions = (view) => (OoxmlPackage.read(view.save()).text('word/document.xml').match(/<w:del\b(?![^>]*\/>)/g) || []).length;
const shape = (view, i = 0) => view.blocks[i].runs.map((r) => (r.del ? '-' + r.del.text : r.text)).filter(Boolean);

test('typing beside a tracked deletion writes it once, where it stands — it used to come back at the end with every keystroke', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Keep these words safe' }] }));
  view.setTrackChanges(true, 'Ann');
  view.setSelection({ block: 0, offset: 5 }, { block: 0, offset: 11 });
  view.deleteSelection();
  view.setTrackChanges(false);
  view.setSelection({ block: 0, offset: 0 });
  view.insertText('X');
  view.insertText('Y');
  assert.equal(deletions(view), 1);
  assert.deepEqual(shape(view), ['XYKeep ', '-these ', 'words safe']);
  const again = openDocx(view.save());
  again.setSelection({ block: 0, offset: 0 });
  again.insertText('Z');
  assert.equal(deletions(again), 1, 'and once again after reopening');
});

test('formatting across somebody\'s tracked deletion leaves it where it was', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Keep these words safe' }] }));
  view.setTrackChanges(true, 'Ann');
  view.setSelection({ block: 0, offset: 5 }, { block: 0, offset: 11 });
  view.deleteSelection();
  view.setTrackChanges(false);
  view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: view.blocks[0].text.length });
  view.toggleFormat('b');
  view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 4 });
  view.setRunFormat({ fontSize: 14 });
  assert.deepEqual(shape(view), ['Keep', ' ', '-these ', 'words safe']);
  assert.equal(deletions(view), 1);
});

test('a paragraph holding a tracked move stays as Word wrote it, and Accept or Reject resolves the move', () => {
  const para = `<w:p><w:moveFromRangeStart w:id="5" w:author="Ann" ${DATE} w:name="move1"/><w:moveFrom w:id="6" w:author="Ann" ${DATE}><w:r><w:t xml:space="preserve">Moved words. </w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:id="5"/><w:r><w:t>Staying words.</w:t></w:r></w:p>`;
  const view = withParagraph(para);
  assert.equal(view.blocks[0].structural, true, 'read-only until the move is resolved');
  const accepted = withParagraph(para);
  accepted.acceptChanges({ all: true });
  assert.equal(accepted.blocks[0].text, 'Staying words.');
  assert.ok(wellFormed(bodyOf(accepted)) && !/w:move/.test(bodyOf(accepted)));
  const rejected = withParagraph(para);
  rejected.rejectChanges({ all: true });
  assert.equal(rejected.blocks[0].text, 'Moved words. Staying words.');
  assert.equal(rejected.blocks[0].structural, false);
});
