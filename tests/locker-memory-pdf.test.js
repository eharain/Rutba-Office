// A locked PDF opened in memory: read from the table, and a save of its pages
// sealed through it, never written to the disk (2.26.0, D-2026-11-13-20).
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { tempDir } from './temp-dirs.js';
import { createPdfService } from '../apps/desktop/main/pdf.js';
import { register } from '../apps/desktop/main/pdf-pack.js';
import { createMemoryFiles, setMemoryFiles, memoryFiles } from '../apps/desktop/main/memory-files.js';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const LETTER = fs.readFileSync(path.join(HERE, 'corpus-public', 'pdf', 'letter.pdf'));

after(() => setMemoryFiles(null));

test('a PDF opened in memory reads from the table, and saving its pages seals through it', async () => {
  const home = tempDir('rl-mempdf-');
  const table = createMemoryFiles({ root: path.join(home, 'mem') });
  setMemoryFiles(table);
  const sealed = [];
  const p = table.add({ name: 'Letter.pdf', bytes: LETTER, save: async (b) => { sealed.push(Buffer.from(b)); } });
  // Made as the core makes it, through the pack's register and the context it is handed.
  const { pdf } = register({ memoryFiles });
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
  const pdf = createPdfService({ memoryFiles });
  await assert.rejects(pdf.open({ path: p }), (err) => err.code === 'ENOENT');
});

test('pages from a PDF in the Locker are saved only back into it', async () => {
  const home = tempDir('rl-mempdf-');
  const table = createMemoryFiles({ root: path.join(home, 'mem') });
  setMemoryFiles(table);
  const p = table.add({ name: 'Letter.pdf', bytes: LETTER, save: async () => {} });
  const pdf = createPdfService({ memoryFiles });
  const opened = await pdf.open({ path: p });
  const pages = [{ source: opened.id, page: 0 }];
  const out = path.join(home, 'out', 'Letter.pdf');
  const said = /Letter\.pdf is in the Locker, so it is saved only back into it/;
  await assert.rejects(pdf.assemble({ pages, path: out }), said);
  await assert.rejects(pdf.redact({ pages, path: out }), said);
  assert.ok(!fs.existsSync(path.join(home, 'out')));
});

test('the PDF pack reaches the in-memory table through the context the core hands it, not by importing the core', () => {
  // A pack that imports a file staying in the core cannot move into its pack, and the release build stops (14 Nov 2026).
  const src = fs.readFileSync(path.join(HERE, '..', 'apps', 'desktop', 'main', 'pdf.js'), 'utf8');
  assert.doesNotMatch(src, /from '\.\/memory-files\.js'/);
  // And the core hands it: main.js puts the table's getter in the context registerPacks hands the packs.
  assert.match(fs.readFileSync(path.join(HERE, '..', 'apps', 'desktop', 'main', 'main.js'), 'utf8'), /await registerPacks\(plan, \{[^}]*\bmemoryFiles,[^}]*\}, \{ services/);
});
