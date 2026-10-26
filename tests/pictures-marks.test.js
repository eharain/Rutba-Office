// Pictures: ratings and tags in the grid's order and filter, and a search
// through the folders inside one, by name or by tag.
import test from 'node:test';
import assert from 'node:assert/strict';
import { arrange, searchTree, SORTS, RATED } from '../apps/desktop/renderer/apps/pictures/library.js';

const e = (name, extra = {}) => ({ name, path: `C:/P/${name}`, dir: false, mtime: 0, size: 0, ...extra });
const entries = [e('a.jpg'), e('b.jpg'), e('c.png'), e('d.mp4')];
const marks = { 'C:/P/b.jpg': { rating: 5, tags: ['Holiday'] }, 'C:/P/c.png': { rating: 3 }, 'C:/P/d.mp4': { tags: ['family'] } };

test('the filter finds a tag as well as a name, the rating filter keeps so many stars or more, and Highest rated sorts by them', () => {
  assert.deepEqual(arrange(entries, { query: 'holi', marks }).map((f) => f.name), ['b.jpg']);
  assert.deepEqual(arrange(entries, { query: 'FAM', marks }).map((f) => f.name), ['d.mp4']);
  assert.deepEqual(arrange(entries, { rated: 3, marks }).map((f) => f.name), ['b.jpg', 'c.png']);
  assert.deepEqual(arrange(entries, { sort: 'rating', marks }).map((f) => f.name), ['b.jpg', 'c.png', 'a.jpg', 'd.mp4']);
  assert.ok(SORTS.some((s) => s.id === 'rating'));
  assert.equal(RATED.length, 6);
  assert.deepEqual(arrange(entries, {}).map((f) => f.name), ['a.jpg', 'b.jpg', 'c.png', 'd.mp4'], 'no marks, as before');
});

test('Subfolders searches the folders inside, nearest first, by name or tag, as deep and as many as it is allowed', async () => {
  const tree = {
    'C:/P': [e('top.jpg'), { name: 'Trips', path: 'C:/P/Trips', dir: true }],
    'C:/P/Trips': [{ name: 'beach.jpg', path: 'C:/P/Trips/beach.jpg', dir: false }, { name: 'Deep', path: 'C:/P/Trips/Deep', dir: true }],
    'C:/P/Trips/Deep': [{ name: 'beach-2.jpg', path: 'C:/P/Trips/Deep/beach-2.jpg', dir: false }, { name: 'x.jpg', path: 'C:/P/Trips/Deep/x.jpg', dir: false }],
  };
  const list = async (dir) => tree[dir] || [];
  assert.deepEqual((await searchTree(list, 'C:/P', { query: 'beach' })).map((f) => f.name), ['beach.jpg', 'beach-2.jpg']);
  assert.deepEqual((await searchTree(list, 'C:/P', { query: 'sunny', marks: { 'C:/P/Trips/Deep/x.jpg': { tags: ['Sunny'] } } })).map((f) => f.name), ['x.jpg']);
  assert.deepEqual((await searchTree(list, 'C:/P', { query: 'beach', depth: 1 })).map((f) => f.name), ['beach.jpg']);
  assert.deepEqual((await searchTree(list, 'C:/P', { query: 'beach', limit: 1 })).map((f) => f.name), ['beach.jpg']);
  assert.deepEqual(await searchTree(list, 'C:/P', { query: '  ' }), []);
});

test('ratings and tags follow a file renamed or moved: known by its size and time, taken over only by one file that matches alone', async () => {
  const { followMarks, fingerprintOf } = await import('../apps/desktop/renderer/apps/pictures/library.js');
  const row = (p, size, mtime) => ({ name: p.split('/').pop(), path: p, dir: false, size, mtime });
  assert.equal(fingerprintOf(row('C:/P/a.jpg', 1200, 1700000000123.4)), '1200:1700000000123');
  assert.equal(fingerprintOf({ path: 'C:/P', dir: true }), null);
  const kept = { 'C:/P/beach.jpg': { rating: 4, tags: ['Holiday'], fp: '5000:111' }, 'C:/P/old.jpg': { rating: 2 } };
  // Renamed in the same folder: the new name takes the marks.
  let r = followMarks(kept, [row('C:/P/Beach day.jpg', 5000, 111), row('C:/P/old.jpg', 900, 222)]);
  assert.deepEqual(r.moves, [{ from: 'C:/P/beach.jpg', to: 'C:/P/Beach day.jpg' }]);
  assert.deepEqual(r.fill, { 'C:/P/old.jpg': '900:222' }, 'a mark kept before fingerprints gets one');
  // Moved to another folder, listed there.
  r = followMarks(kept, [row('D:/Trips/beach.jpg', 5000, 111)]);
  assert.deepEqual(r.moves, [{ from: 'C:/P/beach.jpg', to: 'D:/Trips/beach.jpg' }]);
  // Still listed where it was: nothing to follow.
  assert.deepEqual(followMarks(kept, [row('C:/P/beach.jpg', 5000, 111), row('C:/P/copy.jpg', 5000, 111)]).moves, []);
  // Two files listed with the same fingerprint: neither is guessed at.
  assert.deepEqual(followMarks(kept, [row('D:/a.jpg', 5000, 111), row('D:/b.jpg', 5000, 111)]).moves, []);
  // A file with marks of its own keeps them.
  assert.deepEqual(followMarks({ ...kept, 'D:/x.jpg': { rating: 1, fp: '5000:111' } }, [row('D:/x.jpg', 5000, 111)]).moves, []);
  // Edited since (a new time): its fingerprint is brought up to date.
  assert.deepEqual(followMarks(kept, [row('C:/P/beach.jpg', 5100, 999)]).fill, { 'C:/P/beach.jpg': '5100:999' });
});
