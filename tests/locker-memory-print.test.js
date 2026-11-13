// Print and Export as PDF of a document from the Locker (2.26.0): a PDF of it
// is written only back into the Locker, so Export as PDF to anywhere else is
// refused in words, and a print job is put in the table of files opened in
// memory and loaded from there, never through a file in the temp folder.
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { tempDir } from './temp-dirs.js';
import { createPrintService } from '../apps/desktop/main/print.js';
import { createMemoryFiles, setMemoryFiles } from '../apps/desktop/main/memory-files.js';

after(() => setMemoryFiles(null));

/** A stand-in for Electron's BrowserWindow: what the hidden window was asked to load, and what the table held then. */
function fakeElectron(table) {
  const loads = [];
  class BrowserWindow {
    async loadURL(url) {
      const held = Buffer.from(url.replace('rutba://file/', ''), 'base64url').toString('utf8');
      loads.push({ url, held, there: table?.has(held) ? (await table.read(held)).toString() : null });
    }
    async loadFile(file) {
      loads.push({ file });
    }
    isDestroyed() { return false; }
    destroy() {}
    get webContents() {
      return { print: (_o, cb) => cb(true), printToPDF: async () => Buffer.from('%PDF-printed') };
    }
  }
  return { loads, electron: { BrowserWindow } };
}

const lockerOnly = () => { throw Object.assign(new Error('Tender.docx is in the Locker, so it is saved only back into it.'), { code: 'LOCKER_ONLY' }); };

test('Export as PDF of a document from the Locker is refused before anything is laid out or written', async () => {
  const dir = tempDir('rl-print-');
  const target = path.join(dir, 'Tender.pdf');
  let laidOut = false;
  const docs = { staysInLocker: lockerOnly, printSource: () => { laidOut = true; return { pdf: new Uint8Array([1]) }; } };
  await assert.rejects(createPrintService({ docs }).pdf({ id: 'x', path: target }), (err) => err.code === 'LOCKER_ONLY');
  assert.equal(laidOut, false);
  assert.equal(fs.existsSync(target), false);
});

test('a document not in the Locker exports and prints as before', async () => {
  const dir = tempDir('rl-print-');
  const target = path.join(dir, 'Plain.pdf');
  const { loads, electron } = fakeElectron(null);
  const docs = { staysInLocker: () => {}, printSource: () => ({ name: 'Plain.docx', locked: false, pdf: new Uint8Array(Buffer.from('%PDF-plain')) }) };
  const print = createPrintService({ docs, electron });
  await print.pdf({ id: 'x', path: target });
  assert.equal(fs.readFileSync(target, 'utf8'), '%PDF-plain');
  assert.deepEqual(await print.document({ id: 'x' }), { ok: true, reason: null });
  assert.equal(loads.length, 1);
  assert.ok(loads[0].file, 'through its temp file, as before');
});

test('a document from the Locker prints from memory: the job is in the table while it prints, and gone after', async () => {
  const home = tempDir('rl-print-');
  const table = createMemoryFiles({ root: path.join(home, 'locker', 'mem') });
  setMemoryFiles(table);
  const { loads, electron } = fakeElectron(table);
  for (const source of [
    { name: 'Tender.docx', locked: true, pdf: new Uint8Array(Buffer.from('%PDF-LOCKED-JOB')) },
    { name: 'Budget.xlsx', locked: true, html: '<p>LOCKED-SHEET</p>' },
  ]) {
    const print = createPrintService({ docs: { printSource: () => source }, electron });
    assert.deepEqual(await print.document({ id: 'x' }), { ok: true, reason: null });
  }
  assert.equal(loads.length, 2);
  assert.ok(loads.every((l) => !l.file && l.url.startsWith('rutba://file/')), 'loaded over rutba://file, never from a file');
  assert.equal(loads[0].there, '%PDF-LOCKED-JOB');
  assert.equal(path.basename(loads[1].held), 'job.html');
  assert.equal(loads[1].there, '<p>LOCKED-SHEET</p>');
  assert.equal(table.entries().length, 0, 'each job is let go when it ends');
});
