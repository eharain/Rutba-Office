// Old spreadsheets and databases from public test collections (fetched by
// tools/fetch-old-samples.mjs into .corpus/old, which git ignores): each
// opens as Worksheets reads it or is refused in words, never a crash; each
// that opens imports into Data Files; and the dBASE tables match the summary
// of fields and records their source keeps beside them. Skipped when the
// samples have not been fetched.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import MDBReader from 'mdb-reader';
import { legacySheetKind, readLegacySheets } from '../packages/office-formats/src/legacy-sheets.js';
import { Base, importSheets } from '../packages/database/src/index.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OLD = path.join(ROOT, '.corpus', 'old');
const fetched = fs.existsSync(OLD);
const skip = !fetched && 'run node tools/fetch-old-samples.mjs first';

const files = fetched
  ? fs.readdirSync(OLD).flatMap((src) => fs.readdirSync(path.join(OLD, src)).filter((n) => !/\.(md|txt|dbt|fpt|yml)$/i.test(n)).map((n) => path.join(OLD, src, n)))
  : [];

/** A file read as Worksheets reads it, its memo file beside it when there is one. */
function read(file) {
  const bytes = new Uint8Array(fs.readFileSync(file));
  const ext = path.extname(file).toLowerCase();
  const kind = legacySheetKind(bytes, ext);
  if (!kind) return { kind: null };
  let memo = null;
  let memoKind = null;
  if (kind === 'dbf') {
    for (const m of ['dbt', 'fpt']) {
      const p = file.replace(/\.dbf$/i, `.${m}`);
      if (fs.existsSync(p)) { memo = new Uint8Array(fs.readFileSync(p)); memoKind = m; }
    }
  }
  return { kind, model: readLegacySheets(bytes, kind, { name: path.parse(file).name, memo, memoKind, Reader: MDBReader }) };
}

test('every old sample opens or is refused in words, and every one that opens imports into Data Files', { skip }, () => {
  let opened = 0;
  for (const file of files) {
    let got;
    try {
      got = read(file);
    } catch (err) {
      assert.ok(err instanceof Error && err.message.length > 10, `${path.basename(file)} failed without saying why: ${err}`);
      continue;
    }
    if (!got.kind) continue;
    opened++;
    const db = path.join(os.tmpdir(), `old-sample-${process.pid}-${opened}.rbase`);
    const base = Base.open(db);
    try {
      importSheets(base, got.model);
    } finally {
      base.close();
      fs.rmSync(db, { force: true });
    }
  }
  assert.ok(opened >= 30, `${opened} of ${files.length} opened`);
});

test('the dBASE and FoxPro tables have the fields and the records their summaries give', { skip }, () => {
  const dir = path.join(OLD, 'infused-dbf');
  let compared = 0;
  for (const summary of fs.readdirSync(dir).filter((n) => n.endsWith('_summary.txt'))) {
    const table = path.join(dir, summary.replace('_summary.txt', '.dbf'));
    if (!fs.existsSync(table)) continue;
    const text = fs.readFileSync(path.join(dir, summary), 'latin1');
    const records = Number(/Records:\s*(\d+)/.exec(text)?.[1]);
    const fields = [...text.split(/-{10,}/)[1]?.matchAll(/^(\S+)\s+[A-Z+@]\s+\d+/gm) || []].map((m) => m[1]);
    const { model } = read(table);
    const rows = model.sheets[0].rows;
    const header = rows[0].map((c) => c?.v ?? c?.value ?? c);
    const deleted = model.notes.find((n) => n.code === 'deleted')?.count || 0;
    assert.deepEqual(header, fields, `${summary}: the fields`);
    assert.equal(rows.length - 1 + deleted, records, `${summary}: the records`);
    compared++;
  }
  assert.ok(compared >= 8, `${compared} tables compared`);
});
