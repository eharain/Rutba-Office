// Opening off the main process: a workbook and a document open from a
// package whose parts were inflated asynchronously exactly as from their
// bytes, and save byte for byte the same.
import test from 'node:test';
import assert from 'node:assert/strict';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { Workbook, Document } from '@rutba/ooxml';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

test('a workbook opens from a package read asynchronously as it does from its bytes', async () => {
  const bytes = buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['a', 1], ['b', '=B1*2']] }] });
  const pkg = await OoxmlPackage.readAsync(bytes);
  const fromPkg = Workbook.open(pkg);
  assert.equal(fromPkg.pkg, pkg, 'the package is used as it is, not read again');
  const view = SheetView.open(await OoxmlPackage.readAsync(bytes), { viewportWidth: 600, viewportHeight: 300 });
  assert.ok(Buffer.from(view.serialize()).equals(Buffer.from(SheetView.open(bytes, { viewportWidth: 600, viewportHeight: 300 }).serialize())), 'saves the same');
});

test('a document opens from a package read asynchronously as it does from its bytes', async () => {
  const bytes = buildDocx({ paragraphs: [{ text: 'One' }, { text: 'Two' }] });
  const pkg = await OoxmlPackage.readAsync(bytes);
  assert.equal(Document.open(pkg).pkg, pkg);
  const view = openDocx(await OoxmlPackage.readAsync(bytes));
  assert.deepEqual(view.render().blocks.map((b) => b.text), openDocx(bytes).render().blocks.map((b) => b.text));
});

test('the service opens a file waiting on it as it opens one at once, and refuses a missing one the same way', async () => {
  const fs = await import('node:fs');
  const os = await import('node:os');
  const path = await import('node:path');
  const { createDocumentService } = await import('../apps/desktop/main/documents.js');
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-open-async-'));
  const file = path.join(dir, 'book.xlsx');
  fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['a', 1], ['b', '=B1*2']] }] }));
  const service = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
  const now = service.open({ path: file, kind: 'sheet' });
  const waited = service.openAsync({ path: file, kind: 'sheet' });
  assert.ok(waited instanceof Promise, 'it waits');
  const later = await waited;
  assert.equal(later.kind, now.kind);
  assert.deepEqual(later.model.sheets, now.model.sheets);
  const missing = path.join(dir, 'gone.xlsx');
  let a = null;
  let b = null;
  try { service.open({ path: missing }); } catch (err) { a = err.message; }
  try { await service.openAsync({ path: missing }); } catch (err) { b = err.message; }
  assert.ok(a && a === b, `${a} | ${b}`);
});
