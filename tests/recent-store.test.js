// The recent-files list's own bookkeeping, apart from Electron and the disk.
import test from 'node:test';
import assert from 'node:assert/strict';
import { addEntry, removeEntry, renameEntry, presentEntries } from '../packages/office-shell/src/recent-store.js';

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

test('one file is one row on Windows, whichever slashes or case the path came with', () => {
  const win = { platform: 'win32' };
  let rows = addEntry([], { path: 'C:\\Docs\\Report.docx', name: 'Report.docx', at: 1 }, win);
  rows = addEntry(rows, { path: 'C:/docs/report.docx', name: 'report.docx', at: 2 }, win);
  assert.deepEqual(rows.map((r) => r.at), [2]);
});

test('paths that differ only in case are two files where the system says so', () => {
  const posix = { platform: 'linux' };
  const rows = addEntry(addEntry([], { path: '/docs/a.txt', at: 1 }, posix), { path: '/docs/A.txt', at: 2 }, posix);
  assert.equal(rows.length, 2);
});

test('the newest entry goes first and the list stops at its limit', () => {
  let rows = [];
  for (let i = 0; i < 5; i++) rows = addEntry(rows, { path: `/f${i}`, at: i }, { most: 3, platform: 'linux' });
  assert.deepEqual(rows.map((r) => r.at), [4, 3, 2]);
});

test('renameEntry on an unknown path is a no-op', () => {
  const before = list();
  const out = renameEntry(before, '/docs/nowhere.docx', '/docs/elsewhere.docx');
  assert.deepEqual(out, before);
});

test('the Recent list waits on no file for long: a drive that does not answer is listed, a file that is gone is not', async () => {
  const list = [{ path: 'C:/here.docx' }, { path: 'Z:/lost/drive.xlsx' }, { path: 'C:/gone.pptx' }];
  const exists = (p) => (p.startsWith('Z:') ? new Promise(() => {}) : Promise.resolve(p !== 'C:/gone.pptx'));
  const started = Date.now();
  const shown = await presentEntries(list, { exists, timeoutMs: 80 });
  assert.ok(Date.now() - started < 1000, 'answered without waiting on the lost drive');
  assert.deepEqual(shown.map((r) => r.path), ['C:/here.docx', 'Z:/lost/drive.xlsx']);
  assert.equal((await presentEntries(Array.from({ length: 50 }, (_, i) => ({ path: `f${i}` })), { exists: () => true, limit: 30 })).length, 30);
});
