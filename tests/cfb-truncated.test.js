// A compound file (an old .doc, .xls or .ppt) cut short inside its 512-byte
// header was read past its end and refused with a DataView range error; the
// fuzzer found it at seed 4244. It is refused in a sentence now.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { CompoundFile } from '../packages/office-formats/src/cfb.js';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const fixture = path.join(path.dirname(fileURLToPath(import.meta.url)), 'fixtures', 'binary', 'locked.xls');

test('a compound file cut inside its header is refused in a sentence, never with a range error', () => {
  const whole = fs.readFileSync(fixture);
  for (const length of [8, 30, 60, 100, 300, 511]) {
    assert.throws(() => new CompoundFile(whole.subarray(0, length)), (err) => err.name === 'CfbError' && /cut short/.test(err.message), `cut to ${length} bytes`);
  }
  assert.ok(new CompoundFile(whole), 'the whole file still opens');
});

test('opening one says it could not be read, from the reader\'s own sentence', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-cfb-'));
  const file = path.join(dir, 'cut.xls');
  fs.writeFileSync(file, fs.readFileSync(fixture).subarray(0, 300));
  const service = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  assert.throws(() => service.open({ path: file }), (err) => {
    let e = err;
    while (e) { if (e.name === 'RangeError') return false; e = e.cause; }
    return /could not be read/.test(err.message);
  });
});
