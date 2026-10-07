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

test('paragraph formatting while recording is written as a w:pPrChange; Reject puts the paragraph back, Accept keeps it', () => {
  const make = () => {
    const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Centre me' }, { text: 'Leave me' }] }));
    view.setTrackChanges(true, 'Kim');
    view.setSelection({ block: 0, offset: 2 });
    view.setParagraphFormat({ align: 'center' });
    return view;
  };
  const view = make();
  assert.match(bodyOf(view), /^<w:p><w:pPr><w:jc w:val="center"\/><w:pPrChange w:id="\d+" w:author="Kim" w:date="[^"]+"><w:pPr><\/w:pPr><\/w:pPrChange><\/w:pPr>/);
  assert.equal(view.doc.getParagraphProps(0).align, 'center');
  assert.equal(view.render({ pages: false }).blocks[0].tracked.formatted, 1);
  // Formatted again, the change is still from how it first was; back to that, none.
  view.setParagraphFormat({ align: 'right' });
  assert.match(bodyOf(view), /<w:jc w:val="right"\/><w:pPrChange [^>]*><w:pPr><\/w:pPr><\/w:pPrChange>/);
  view.setParagraphFormat({ align: null });
  assert.ok(!/pPrChange/.test(bodyOf(view)));
  const rejected = make();
  rejected.rejectChanges({ all: true });
  assert.equal(rejected.doc.getParagraphProps(0).align, null);
  const accepted = make();
  accepted.acceptChanges({ all: true });
  assert.equal(accepted.doc.getParagraphProps(0).align, 'center');
  assert.ok(!/pPrChange/.test(bodyOf(accepted)));
});

test('Replace All while recording takes the words found out and puts the replacement in after them, each change its own id', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'The cat sat by the cat flap.' }] }));
  view.setTrackChanges(true, 'Kim');
  assert.equal(view.replaceAll('cat', 'dog'), 2);
  assert.deepEqual(shape(view), ['The ', '-cat', 'dog', ' sat by the ', '-cat', 'dog', ' flap.']);
  assert.equal(view.blocks[0].text, 'The dog sat by the dog flap.');
  const ids = [...bodyOf(view).matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map((m) => m[1]);
  assert.equal(ids.length, 4);
  assert.equal(new Set(ids).size, 4, 'no two changes share an id');
  view.rejectChanges({ all: true });
  assert.equal(view.blocks[0].text, 'The cat sat by the cat flap.');
});

test('Sort while recording moves each paragraph\'s words as a deletion and an insertion, its formatting as a change, and Reject All restores the order', () => {
  const make = () => {
    const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Cherry' }, { text: 'apple', style: 'Heading1' }, { text: 'Banana' }] }));
    view.setTrackChanges(true, 'Kim');
    view.setSelection({ block: 0, offset: 0 }, { block: 2, offset: 6 });
    view.sortParagraphs();
    return view;
  };
  const view = make();
  assert.deepEqual(view.blocks.map((b) => b.text), ['apple', 'Banana', 'Cherry']);
  assert.deepEqual(shape(view, 0), ['-Cherry', 'apple']);
  assert.equal(view.blocks[0].style, 'Heading1', 'the heading\'s formatting went with its words');
  assert.match(bodyOf(view), /^<w:p><w:pPr><w:pStyle w:val="Heading1"\/><w:pPrChange [^>]*><w:pPr><\/w:pPr><\/w:pPrChange><\/w:pPr>/);
  const rejected = make();
  rejected.rejectChanges({ all: true });
  assert.deepEqual(rejected.blocks.map((b) => [b.text, b.style ?? null]), [['Cherry', null], ['apple', 'Heading1'], ['Banana', null]]);
  const accepted = make();
  accepted.acceptChanges({ all: true });
  assert.deepEqual(accepted.blocks.map((b) => [b.text, b.style ?? null]), [['apple', 'Heading1'], ['Banana', null], ['Cherry', null]]);
  assert.ok(wellFormed(bodyOf(accepted)) && !/<w:(?:ins|del|pPrChange)\b/.test(bodyOf(accepted)));
});

const texts = (view) => view.blocks.map((b) => b.text);

test('Enter while recording records the new paragraph mark as put in; Reject joins the halves again', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'One two' }, { text: 'Three' }] }));
  view.setTrackChanges(true, 'Kim');
  view.setSelection({ block: 0, offset: 3 });
  view.splitParagraph();
  assert.deepEqual(texts(view), ['One', ' two', 'Three']);
  assert.equal(view.doc.paragraphMark(0).ins.author, 'Kim');
  assert.equal(view.doc.paragraphMark(1).ins, null, 'the paragraph\'s own mark goes on with the second half');
  assert.equal(view.render({ pages: false }).blocks[0].tracked.mark, 'inserted');
  assert.match(bodyOf(view), /^<w:p><w:pPr><w:rPr><w:ins w:id="\d+" w:author="Kim" w:date="[^"]+"\/><\/w:rPr><\/w:pPr>/);
  view.rejectChanges({ all: true });
  assert.deepEqual(texts(view), ['One two', 'Three']);
});

test('Backspace and Delete across a paragraph mark while recording record it as taken out; Accept joins, Reject keeps apart', () => {
  const make = (how) => {
    const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'First' }, { text: 'Second' }, { text: 'Third' }] }));
    view.setTrackChanges(true, 'Kim');
    if (how === 'back') { view.setSelection({ block: 1, offset: 0 }); view.deleteBackward(); }
    else { view.setSelection({ block: 0, offset: 5 }); view.deleteForward(); }
    return view;
  };
  for (const how of ['back', 'forward']) {
    const view = make(how);
    assert.deepEqual(texts(view), ['First', 'Second', 'Third'], how + ': still apart');
    assert.equal(view.doc.paragraphMark(0).del.author, 'Kim', how);
    assert.equal(view.render({ pages: false }).blocks[0].tracked.mark, 'deleted');
    const accepted = make(how);
    accepted.acceptChanges({ all: true });
    assert.deepEqual(texts(accepted), ['FirstSecond', 'Third'], how + ': joined by Accept');
    const rejected = make(how);
    rejected.rejectChanges({ all: true });
    assert.deepEqual(texts(rejected), ['First', 'Second', 'Third'], how + ': kept apart by Reject');
    assert.ok(!/<w:del\b/.test(bodyOf(rejected)));
  }
  // Backspace leaves the caret before the mark, so a second press takes a letter.
  const view = make('back');
  assert.deepEqual(view.selection.focus ?? view.focus, { block: 0, offset: 5 });
});

test('deleting across paragraphs while recording strikes the words through and records each mark between as taken out', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Alpha beta' }, { text: 'Gamma' }, { text: 'Delta epsilon' }] }));
  view.setTrackChanges(true, 'Kim');
  view.setSelection({ block: 0, offset: 6 }, { block: 2, offset: 6 });
  view.deleteSelection();
  assert.deepEqual(texts(view), ['Alpha ', '', 'epsilon'], 'apart, their words struck through');
  assert.deepEqual(shape(view, 0), ['Alpha ', '-beta']);
  assert.ok(view.doc.paragraphMark(0).del && view.doc.paragraphMark(1).del && !view.doc.paragraphMark(2).del);
  view.acceptChanges({ all: true });
  assert.deepEqual(texts(view), ['Alpha epsilon']);
  const again = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Alpha beta' }, { text: 'Gamma' }, { text: 'Delta epsilon' }] }));
  again.setTrackChanges(true, 'Kim');
  again.setSelection({ block: 0, offset: 6 }, { block: 2, offset: 6 });
  again.deleteSelection();
  again.rejectChanges({ all: true });
  assert.deepEqual(texts(again), ['Alpha beta', 'Gamma', 'Delta epsilon']);
});

test('one\'s own new paragraph taken out again simply goes, and a paste of several lines records each new mark', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'One two' }, { text: 'Three' }] }));
  view.setTrackChanges(true, 'Kim');
  view.setSelection({ block: 0, offset: 3 });
  view.splitParagraph();
  view.setSelection({ block: 1, offset: 0 });
  view.deleteBackward();
  assert.deepEqual(texts(view), ['One two', 'Three']);
  assert.ok(!view.render({ pages: false }).blocks[0].tracked, 'nothing left to review');
  view.setSelection({ block: 1, offset: 5 });
  view.pasteText(' and\nfour\nfive');
  assert.deepEqual(texts(view), ['One two', 'Three and', 'four', 'five']);
  assert.ok(view.doc.paragraphMark(1).ins && view.doc.paragraphMark(2).ins && !view.doc.paragraphMark(3).ins);
  view.rejectChanges({ all: true });
  assert.deepEqual(texts(view), ['One two', 'Three']);
});
