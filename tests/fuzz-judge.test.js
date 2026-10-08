// The fuzzer's judge: a sentence is a refusal, but a sentence made from a
// RangeError, a stack overflow or a TypeError is a fault, since the parser
// did not choose it. And a damaged archive pointing past its own end is
// refused in a sentence of the reader's own, not read into a RangeError.
import test from 'node:test';
import assert from 'node:assert/strict';
import { judge } from '../tools/fuzz-open.js';
import { readZip } from '../packages/ooxml/src/zip.js';
import { buildDocx } from '@rutba/ooxml/build';

test('the judge looks through the sentence to what it was made from', () => {
  assert.equal(judge(new Error('letter.docx is not a complete .docx file.')).outcome, 'refused');
  assert.equal(judge(new Error('letter.docx could not be read. (bad)', { cause: new RangeError('Offset is outside the bounds of the DataView') })).outcome, 'FAULT');
  assert.equal(judge(new Error('x could not be read. (Maximum call stack size exceeded)')).outcome, 'FAULT');
  assert.equal(judge(new TypeError("Cannot read properties of undefined (reading 'x')")).outcome, 'FAULT');
  assert.equal(judge(new Error('???')).outcome, 'FAULT', 'not a sentence');
});

test('an archive whose records point past its end is refused by the reader in a sentence', () => {
  const good = Buffer.from(buildDocx({ paragraphs: [{ text: 'One' }] }));
  // The central directory said to start beyond the file.
  const eocd = good.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  const bad = Buffer.from(good);
  bad.writeUInt32LE(bad.length + 1000, eocd + 16);
  assert.throws(() => readZip(bad), (err) => err.name === 'Error' && /not a complete zip archive/.test(err.message));
  // An entry's local header moved past the end.
  const cd = good.readUInt32LE(eocd + 16);
  const bad2 = Buffer.from(good);
  bad2.writeUInt32LE(0x7ffffff0, cd + 42);
  assert.throws(() => readZip(bad2), (err) => err.name === 'Error' && /not a complete zip archive/.test(err.message));
});
