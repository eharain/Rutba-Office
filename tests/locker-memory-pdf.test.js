// A locked PDF opened in memory: read from the table, and a save of its pages
// sealed through it, never written to the disk (2.26.0, D-2026-11-13-20).
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { tempDir } from './temp-dirs.js';
import { createPdfService } from '../apps/desktop/main/pdf.js';
import { createMemoryFiles, setMemoryFiles } from '../apps/desktop/main/memory-files.js';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const LETTER = fs.readFileSync(path.join(HERE, 'corpus-public', 'pdf', 'letter.pdf'));

after(() => setMemoryFiles(null));

test('a PDF opened in memory reads from the table, and saving its pages seals through it', async () => {
  const home = tempDir('rl-mempdf-');
  const table = createMemoryFiles({ root: path.join(home, 'mem') });
  setMemoryFiles(table);
  const sealed = [];
  const p = table.add({ name: 'Letter.pdf', bytes: LETTER, save: async (b) => { sealed.push(Buffer.from(b)); } });
  const pdf = createPdfService();
  const opened = await pdf.open({ path: p });
  assert.ok(opened.pages?.length || opened.pageCount, 'it opened');
  const pages = Array.from({ length: opened.pages?.length || opened.pageCount }, (_, i) => ({ source: opened.id, page: i }));
  const out = await pdf.assemble({ pages, path: p });
  assert.equal(sealed.length, 1);
  assert.equal(out.size, sealed[0].length);
  assert.equal(sealed[0].subarray(0, 5).toString('latin1'), '%PDF-');
  assert.deepEqual(fs.readdirSync(path.join(home, 'mem')).map((n) => fs.statSync(path.join(home, 'mem', n)).isFile()), [true], 'only the marker is there');
  const looked = await pdf.inspect({ path: p });
  assert.ok(looked.pages.length >= 1);
});

test('a PDF whose Locker has locked since reads as gone', async () => {
  const home = tempDir('rl-mempdf-');
  const table = createMemoryFiles({ root: path.join(home, 'mem') });
  setMemoryFiles(table);
  const p = table.add({ name: 'Letter.pdf', bytes: LETTER, save: async () => {} });
  await table.clear();
  const pdf = createPdfService();
  await assert.rejects(pdf.open({ path: p }), (err) => err.code === 'ENOENT');
});

test('pages from a PDF in the Locker are saved only back into it', async () => {
  const home = tempDir('rl-mempdf-');
  const table = createMemoryFiles({ root: path.join(home, 'mem') });
  setMemoryFiles(table);
  const p = table.add({ name: 'Letter.pdf', bytes: LETTER, save: async () => {} });
  const pdf = createPdfService();
  const opened = await pdf.open({ path: p });
  const pages = [{ source: opened.id, page: 0 }];
  const out = path.join(home, 'out', 'Letter.pdf');
  const said = /Letter\.pdf is in the Locker, so it is saved only back into it/;
  await assert.rejects(pdf.assemble({ pages, path: out }), said);
  await assert.rejects(pdf.redact({ pages, path: out }), said);
  assert.ok(!fs.existsSync(path.join(home, 'out')));
});
