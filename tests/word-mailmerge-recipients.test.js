// Mail merge: Edit Recipient List's ticks and sort are kept in the document
// — in a part of the suite's own, which Word passes over — and come back
// when the letter is opened again and its list re-attached, each left-out
// record found again by its values even after the list has moved its rows.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml/build';
import { parseDelimited, sourceFromRows } from '@rutba/ooxml/mailmerge';
import { OoxmlPackage } from '@rutba/ooxml';

const CSV = 'First Name,City\r\nJoshua,London\r\nCynthia,Leeds\r\nAmira,Paris\r\n';
const MOVED = 'First Name,City\r\nAmira,Paris\r\nJoshua,London\r\nCynthia,Leeds\r\n';
const source = (text) => sourceFromRows(parseDelimited(text), { kind: 'csv', name: 'list.csv', path: 'list.csv' });

function letter() {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Dear ' }] }));
  view.startMailMerge('formLetters');
  view.attachMergeSource(source(CSV));
  view.setMergeRecipients({ excluded: [1], sort: { field: 'City', descending: false } });
  return view;
}

test('the ticks and the sort are kept in a part of the suite\'s own and come back with the list', () => {
  const bytes = letter().save();
  const pkg = OoxmlPackage.read(bytes);
  assert.ok(pkg.has('customXml/rutbaMergeRecipients.xml'));
  const again = openDocx(bytes);
  again.attachMergeSource(source(CSV), { restore: true });
  assert.deepEqual(again.merge.excluded, [1], 'Cynthia still left out');
  assert.deepEqual(again.merge.sort, { field: 'City', descending: false });
  assert.deepEqual(again.mergeOrder(), [0, 2], 'London, Paris');
});

test('a list that has moved its rows still leaves out the same person', () => {
  const again = openDocx(letter().save());
  again.attachMergeSource(source(MOVED), { restore: true });
  assert.deepEqual(again.merge.excluded, [2], 'Cynthia, now third');
});

test('every recipient ticked again and no sort: the part says so and nothing is restored', () => {
  const view = letter();
  view.setMergeRecipients({ excluded: [], sort: null });
  const again = openDocx(view.save());
  again.attachMergeSource(source(CSV), { restore: true });
  assert.deepEqual(again.merge.excluded, []);
  assert.equal(again.merge.sort, null);
});
