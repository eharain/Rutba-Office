// Documents held to Word (tools/word-check.mjs built the files there, made the
// same edits, and opened what we write).
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { EDITS, TRACKED, applyEdit, applyTracked, shapeOf } from '../tools/word-check.mjs';

const bytes = fs.readFileSync(new URL('./fixtures/word-edits.docx', import.meta.url));
const tracked = fs.readFileSync(new URL('./fixtures/word-tracked.docx', import.meta.url));
const answers = JSON.parse(fs.readFileSync(new URL('./fixtures/word-edits.json', import.meta.url), 'utf8'));

test('the file Word wrote reads the same here', () => {
  assert.deepEqual(shapeOf(bytes), answers.base);
});

EDITS.forEach((edit, i) => {
  test(`as Word: ${edit.name}`, () => {
    assert.deepEqual(shapeOf(applyEdit(bytes, edit)), answers.edits[i].word);
  });
});

TRACKED.forEach((edit, i) => {
  test(`as Word: ${edit.name}`, () => {
    assert.deepEqual(shapeOf(applyTracked(tracked, edit)), answers.tracked[i].word);
  });
});
