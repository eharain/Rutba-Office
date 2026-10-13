// The PDF window's list of pages: each operation a new list, so undo and
// redo are the lists before and after (renderer/apps/pdf/pages.js).

import test from 'node:test';
import assert from 'node:assert/strict';
import {
  pagesOf, rotatePages, deletePages, insertPages, movePages, rangeKeys, rangeText, chosenKeys, shownSize, splitGroups,
} from '../apps/desktop/renderer/apps/pdf/pages.js';

const info = { pages: [{ width: 595, height: 842, rotate: 0 }, { width: 595, height: 842, rotate: 90 }, { width: 612, height: 792, rotate: 0 }, { width: 595, height: 842, rotate: 0 }] };
const list = () => pagesOf('a', info);
const order = (pages) => pages.map((p) => p.page);

test('a file\'s pages become entries with keys of their own and the turn they have', () => {
  const pages = list();
  assert.deepEqual(order(pages), [0, 1, 2, 3]);
  assert.equal(new Set(pages.map((p) => p.key)).size, 4);
  assert.deepEqual(pages.map((p) => p.rotate), [0, 90, 0, 0]);
});

test('moving pages: one to the front, several together, to the end, and a move that changes nothing', () => {
  const pages = list();
  assert.deepEqual(order(movePages(pages, [pages[2].key], 0)), [2, 0, 1, 3]);
  assert.deepEqual(order(movePages(pages, [pages[0].key, pages[2].key], 4)), [1, 3, 0, 2]);
  assert.deepEqual(order(movePages(pages, [pages[3].key, pages[1].key], 1)), [0, 1, 3, 2]);
  assert.equal(movePages(pages, [pages[1].key], 1), pages, 'dropped where it was: the same list');
  assert.equal(movePages(pages, [pages[1].key], 2), pages, 'just after itself: the same list');
});

test('rotating turns the pages named, and keeps a turn a quarter at a time', () => {
  const pages = list();
  const turned = rotatePages(pages, [pages[1].key, pages[2].key], 90);
  assert.deepEqual(turned.map((p) => p.rotate), [0, 180, 90, 0]);
  assert.deepEqual(rotatePages(turned, [pages[2].key], -90).map((p) => p.rotate), [0, 180, 0, 0]);
  assert.deepEqual(rotatePages(pages, [pages[0].key], -90).map((p) => p.rotate)[0], 270);
  assert.equal(pages[1].rotate, 90, 'the list it came from is left as it was');
});

test('deleting and inserting', () => {
  const pages = list();
  assert.deepEqual(order(deletePages(pages, [pages[0].key, pages[3].key])), [1, 2]);
  const more = pagesOf('b', { pages: [{ width: 100, height: 100 }, { width: 100, height: 100 }] });
  const ins = insertPages(pages, 1, more);
  assert.deepEqual(ins.map((p) => `${p.source}${p.page}`), ['a0', 'b0', 'b1', 'a1', 'a2', 'a3']);
  assert.deepEqual(insertPages(pages, 99, more).slice(-2).map((p) => p.source), ['b', 'b']);
});

test('the selection: a run picked with Shift, and the page in view when nothing is picked', () => {
  const pages = list();
  assert.deepEqual(rangeKeys(pages, pages[3].key, pages[1].key), [pages[1].key, pages[2].key, pages[3].key]);
  assert.deepEqual(chosenKeys(pages, [pages[3].key, pages[0].key], pages[1].key), [pages[0].key, pages[3].key], 'in the list\'s order');
  assert.deepEqual(chosenKeys(pages, [], pages[2].key), [pages[2].key]);
  assert.deepEqual(chosenKeys(pages, [], 'gone'), []);
});

test('page numbers written short, a turned page\'s size, and where a split falls', () => {
  assert.equal(rangeText([5, 1, 2, 3, 8, 9]), '1-3, 5, 8-9');
  assert.equal(rangeText([4]), '4');
  assert.deepEqual(shownSize({ width: 595, height: 842 }, 90), { width: 842, height: 595 });
  assert.deepEqual(shownSize({ width: 595, height: 842 }, 180), { width: 595, height: 842 });
  assert.deepEqual(splitGroups(5, { every: 2 }), [[0, 1], [2, 3], [4]]);
  assert.deepEqual(splitGroups(6, { starts: [4, 2] }), [[0, 1], [2, 3], [4, 5]]);
  assert.deepEqual(splitGroups(3, { starts: [0] }), [[0, 1, 2]], 'a cut before the first page cuts nothing');
});
