// A query whose source a workbook names on a network share is not read on
// Refresh: reading a share hands it this computer's sign-in, and the path
// came from the file. A local source refreshes as ever.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-query-share-'));

const workbookWith = (source) => {
  const view = SheetView.open(buildXlsx({ sheets: [{ name: 'Sheet1', rows: [] }] }), { viewportWidth: 600, viewportHeight: 300 });
  view.addQuery({ name: 'Stock', source: { kind: 'csv', path: source }, steps: [{ kind: 'promoteHeaders' }], read: () => 'Item,Qty\nPens,4\n' });
  const file = path.join(dir, `book-${Math.random().toString(36).slice(2)}.xlsx`);
  fs.writeFileSync(file, view.serialize());
  return file;
};

test('Refresh does not read a source on a network share that only the workbook names', () => {
  const service = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
  const opened = service.open({ path: workbookWith('\\\\files.example\\share\\stock.csv'), kind: 'sheet' });
  assert.throws(() => service.apply({ id: opened.id, ops: [{ op: 'refreshQueries' }] }), /network share/);
});

test('a local source refreshes as ever', () => {
  const csv = path.join(dir, 'stock.csv');
  fs.writeFileSync(csv, 'Item,Qty\nPens,9\n');
  const service = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
  const opened = service.open({ path: workbookWith(csv), kind: 'sheet' });
  service.apply({ id: opened.id, ops: [{ op: 'refreshQueries' }] });
});
