// The recent-files list's own bookkeeping, apart from Electron and the disk.
import test from 'node:test';
import assert from 'node:assert/strict';
import { removeEntry, renameEntry } from '../packages/office-shell/src/recent-store.js';

const list = () => [
  { path: '/docs/report.docx', name: 'report.docx', app: 'word', at: 3 },
  { path: '/docs/sales.xlsx', name: 'sales.xlsx', app: 'sheets', at: 2 },
];

test('removeEntry drops the matching entry and leaves the rest', () => {
  const out = removeEntry(list(), '/docs/report.docx');
  assert.deepEqual(out.map((r) => r.path), ['/docs/sales.xlsx']);
});

test('removeEntry on an unknown path is a no-op', () => {
  const before = list();
  const out = removeEntry(before, '/docs/nowhere.docx');
  assert.deepEqual(out, before);
});

test('renameEntry updates the path and the name, keeping everything else', () => {
  const out = renameEntry(list(), '/docs/report.docx', '/docs/quarterly.docx');
  assert.deepEqual(out[0], { path: '/docs/quarterly.docx', name: 'quarterly.docx', app: 'word', at: 3 });
  assert.equal(out[1].path, '/docs/sales.xlsx'); // untouched
});

test('renameEntry on an unknown path is a no-op', () => {
  const before = list();
  const out = renameEntry(before, '/docs/nowhere.docx', '/docs/elsewhere.docx');
  assert.deepEqual(out, before);
});
