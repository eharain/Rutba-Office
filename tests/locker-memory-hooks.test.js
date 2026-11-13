// Every place that reads or writes a document's bytes by path asks the table
// of files opened in memory first (2.26.0, D-2026-11-13-20). The shell's code
// runs only inside Electron, so these read its source: a new file call that
// takes a path and does not ask the table fails here, before it can write a
// locked file out in the clear.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const read = (rel) => fs.readFileSync(path.join(ROOT, rel), 'utf8').replace(/\r\n/g, '\n');

/** The body of each `name: async? ({ path: p ... }) =>` entry inside `impl.<ns> = { ... };`. */
function entriesOf(src, ns) {
  const start = src.indexOf(`impl.${ns} = {`);
  assert.ok(start >= 0, `impl.${ns} is there`);
  const end = src.indexOf('\n  };\n', start);
  const body = src.slice(start, end);
  const out = new Map();
  const re = /\n {4}(\w+): (?:async )?\(([^)]*)\)/g;
  const marks = [...body.matchAll(re)];
  marks.forEach((m, i) => {
    const to = i + 1 < marks.length ? marks[i + 1].index : body.length;
    out.set(m[1], { args: m[2], body: body.slice(m.index, to) });
  });
  return out;
}

// Calls that only ever touch the scratch files made in this run, never a person's file.
const SCRATCH_ONLY = new Set(['temp', 'append', 'patch', 'dropTemp']);

test('every shell file call that takes a path asks the table first', () => {
  const src = read('packages/office-shell/src/electron/ipc.js');
  const fsCalls = entriesOf(src, 'fs');
  assert.ok(fsCalls.size > 10);
  const missed = [];
  for (const [name, { args, body }] of fsCalls) {
    if (SCRATCH_ONLY.has(name)) continue;
    if (!/path: p\b|\bto\b/.test(args)) continue;
    const asks = name === 'keepTemp' ? /mem\(to\)/.test(body) : /mem\(p\)|mem\(path\.join/.test(body);
    if (!asks) missed.push(name);
  }
  assert.deepEqual(missed, [], `these do not ask the table: ${missed.join(', ')}`);
  for (const name of ['read', 'readHead', 'readText', 'write', 'writeText', 'stat', 'exists', 'list', 'keepTemp']) assert.ok(fsCalls.has(name), name);
});

test('the shell keeps a file opened in memory out of the recent list, the thumbnail cache, the folder view and other programs', () => {
  const src = read('packages/office-shell/src/electron/ipc.js');
  assert.match(src, /addRecent: \(p\) => \(mem\(p\?\.path \?\? p\) \? undefined : stores\.recent\.add\(p\)\)/);
  assert.match(src, /put: async \(\{ path: p, bytes \}\) => \(thumbnailer && !mem\(p\)/);
  const app = entriesOf(src, 'shell');
  assert.match(app.get('showInFolder').body, /mem\(p\)/);
  assert.match(app.get('openPath').body, /if \(mem\(p\)\) throw notHere\(p\)/);
});

test('rutba://file and rutba://thumb answer a file opened in memory from the table, never from the disk', () => {
  const src = read('packages/office-shell/src/electron/protocol.js');
  const file = src.slice(src.indexOf("if (host === 'file')"), src.indexOf("if (host === 'thumb')"));
  const held = file.indexOf('held?.isUnder(target)');
  assert.ok(held > 0 && held < file.indexOf('fs.existsSync(target)'), 'the table is asked before the disk');
  assert.ok(held < file.indexOf('pictureResponse('), 'and before anything is decoded through a path');
  const thumb = src.slice(src.indexOf("if (host === 'thumb')"), src.indexOf("if (host === 'blob')"));
  assert.ok(thumb.indexOf('memoryFiles()?.isUnder(target)') < thumb.indexOf('thumbnailer.respond'));
});

test('the shell is handed the table by the desktop app', () => {
  assert.match(read('packages/office-shell/src/electron/main.js'), /installProtocol\(\{[^}]*memoryFiles \}\)/);
  assert.match(read('packages/office-shell/src/electron/main.js'), /buildImplementations\(\{[^}]*memoryFiles \}\)/);
  assert.match(read('apps/desktop/main/main.js'), /\n {2}memoryFiles,\n/);
});

test('documents.js and pdf.js read and write a document only through the hooked helpers', () => {
  const doc = read('apps/desktop/main/documents.js');
  // The open runners read through readDocument; the two write helpers ask inMemory first.
  assert.doesNotMatch(doc, /step\.value\.read \? (await )?fs\./);
  for (const fn of ['function writeWhole(target, data, encoding) {\n    if (inMemory(target))', 'async function writeWholeAsync(target, data, encoding) {\n    if (inMemory(target))']) {
    assert.ok(doc.includes(fn), fn);
  }
  const pdf = read('apps/desktop/main/pdf.js');
  // Every finished PDF goes through writeOut, which asks the table.
  assert.equal((pdf.match(/await writeWhole\(/g) || []).length, 1, 'writeWhole is called only from writeOut');
  assert.doesNotMatch(pdf, /await fsp\.readFile\(file\)/);
});
