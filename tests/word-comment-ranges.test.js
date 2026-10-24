/**
 * Comments anchored to the words they are about (2.18.0, A3.5).
 *
 * In 2.17.1 a new comment hung on the end of its paragraph, with no words
 * marked, and a comment Word had anchored to some words made its paragraph
 * read-only. Now a comment over a selection writes Word's range marks
 * (commentRangeStart and commentRangeEnd) round the selected words, the
 * paragraph stays editable, and the marks stay round the words through later
 * typing, Enter and Backspace.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { OoxmlPackage } from '@rutba/ooxml';
import { parseRuns, renderRuns } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';

/** The words between a comment's marks, read off the saved body. */
function rangeText(xml, id) {
  const start = xml.indexOf('<w:commentRangeStart w:id="' + id + '"/>');
  const end = xml.indexOf('<w:commentRangeEnd w:id="' + id + '"/>');
  if (start < 0 || end < 0 || end < start) return null;
  return [...xml.slice(start, end).matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<\/w:p>/g)].map((m) => (m[1] === undefined ? '\n' : m[1])).join('');
}

const body = (view) => view.doc.doc.xml;

test('a comment over selected words marks those words, and is signed by its author', () => {
  const view = openDocx(buildDocx({ paragraphs: ['Count the till before you lock the door.'] }));
  view.setSelection({ block: 0, offset: 10 }, { block: 0, offset: 14 });
  view.addComment('Both tills?', { author: 'Sam Patel' });
  const xml = body(view);
  assert.equal(rangeText(xml, '1'), 'till');
  assert.match(xml, /<w:commentRangeEnd w:id="1"\/>.*<w:commentReference w:id="1"\/>/s, 'the reference after the range');
  const c = view.doc.comments()[0];
  assert.equal(c.author, 'Sam Patel');
  assert.equal(c.text, 'Both tills?');
  assert.equal(c.blockIndex, 0);
  // The page marks the words.
  const frame = view.render({ pages: false });
  const marked = frame.blocks[0].runs.filter((r) => r.comments?.includes('1')).map((r) => r.text).join('');
  assert.equal(marked, 'till');
  assert.equal(frame.blocks[0].structural, false, 'the paragraph stays editable');
});

test('a caret in a word comments on that word', () => {
  const view = openDocx(buildDocx({ paragraphs: ['Lock the door.'] }));
  view.collapseTo({ block: 0, offset: 10 });
  view.addComment('Front or back?', { author: 'Sam' });
  assert.equal(rangeText(body(view), '1'), 'door');
});

test('typing, Enter and Backspace keep the marks round the commented words', () => {
  const view = openDocx(buildDocx({ paragraphs: ['Count the till before you lock the door.'] }));
  view.setSelection({ block: 0, offset: 10 }, { block: 0, offset: 14 });
  view.addComment('Both tills?', { author: 'Sam' });
  // Typing before the range moves it along.
  view.collapseTo({ block: 0, offset: 0 });
  view.insertText('Please ');
  assert.equal(rangeText(body(view), '1'), 'till');
  // Typing inside it grows it.
  view.collapseTo({ block: 0, offset: 20 });
  view.insertText('s');
  assert.equal(rangeText(body(view), '1'), 'tilsl');
  // Enter in the middle of the words: the range runs across the two paragraphs.
  view.collapseTo({ block: 0, offset: 19 });
  view.splitParagraph();
  assert.equal(rangeText(body(view), '1'), 'ti\nlsl');
  assert.equal((body(view).match(/<w:commentRangeStart w:id="1"\/>/g) || []).length, 1, 'one start');
  assert.equal((body(view).match(/<w:commentRangeEnd w:id="1"\/>/g) || []).length, 1, 'one end');
  // Typing in the second half keeps them so.
  view.collapseTo({ block: 1, offset: 1 });
  view.insertText('!');
  assert.equal((body(view).match(/<w:commentRangeEnd w:id="1"\/>/g) || []).length, 1);
  // Backspace joins them again: one range, as it was.
  view.collapseTo({ block: 1, offset: 0 });
  view.deleteBackward();
  assert.equal(rangeText(body(view), '1'), 'til!sl');
  assert.equal((body(view).match(/<w:commentRange(?:Start|End) w:id="1"\/>/g) || []).length, 2);
  // Saved and opened again: Word's form, the paragraph editable.
  const again = openDocx(view.save());
  assert.equal(again.render({ pages: false }).blocks[0].structural, false);
  assert.equal(rangeText(OoxmlPackage.read(view.save()).text('word/document.xml'), '1'), 'til!sl');
});

test('a selection over two paragraphs is one range from the first word to the last', () => {
  const view = openDocx(buildDocx({ paragraphs: ['Opening jobs', 'Unlock the door', 'Turn the till on'] }));
  view.setSelection({ block: 0, offset: 8 }, { block: 2, offset: 4 });
  view.addComment('Who does these?', { author: 'Sam' });
  const xml = body(view);
  assert.equal(rangeText(xml, '1'), 'jobs\nUnlock the door\nTurn');
  // Each paragraph still takes typing, and the marks stay single.
  view.collapseTo({ block: 2, offset: 16 });
  view.insertText(' now');
  view.collapseTo({ block: 0, offset: 0 });
  view.insertText('The ');
  assert.equal(rangeText(body(view), '1'), 'jobs\nUnlock the door\nTurn');
  assert.equal((body(view).match(/<w:commentRange(?:Start|End) w:id="1"\/>/g) || []).length, 2);
});

test('a file with Word\'s range marks opens editable, and editing keeps the marks', () => {
  const docx = buildDocx({ paragraphs: ['Words'] });
  const pkg = OoxmlPackage.read(docx);
  const doc = pkg.text('word/document.xml').replace(/<w:p>[\s\S]*?<\/w:p>/, '<w:p><w:r><w:t xml:space="preserve">Check </w:t></w:r><w:commentRangeStart w:id="0"/><w:r><w:t>this</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r><w:r><w:t xml:space="preserve"> line.</w:t></w:r></w:p>');
  pkg.write_('word/document.xml', doc);
  const view = openDocx(pkg.write());
  assert.equal(view.render({ pages: false }).blocks[0].structural, false);
  view.collapseTo({ block: 0, offset: 0 });
  view.insertText('Please ');
  assert.equal(view.block(0).text, 'Please Check this line.');
  assert.equal(rangeText(body(view), '0'), 'this');
  assert.equal((body(view).match(/<w:commentReference w:id="0"\/>/g) || []).length, 1, 'the reference once');
});

test('one undo takes the comment and its marks away', () => {
  const view = openDocx(buildDocx({ paragraphs: ['Lock the door.'] }));
  const before = body(view);
  view.setSelection({ block: 0, offset: 0 }, { block: 0, offset: 4 });
  view.addComment('Twice', { author: 'Sam' });
  view.undo();
  assert.equal(body(view), before);
  assert.equal(view.doc.comments().length, 0);
});

test('runs copied out of a range do not take the marks with them', () => {
  const runs = parseRuns('<w:p><w:commentRangeStart w:id="4"/><w:r><w:t>here</w:t></w:r><w:commentRangeEnd w:id="4"/></w:p>');
  assert.deepEqual(runs[0].comments, ['4']);
  assert.doesNotMatch(renderRuns(runs), /commentRange/, 'without the paragraph\'s own plan, no marks');
});
