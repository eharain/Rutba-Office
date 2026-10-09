// No English left where a translation cannot reach it: every word a window
// shows goes through the message catalogue (tools/find-english.mjs reads the
// window files for the shapes words take). tests/messages.test.js holds the
// catalogue's list, messages.json, to the code.
import test from 'node:test';
import assert from 'node:assert/strict';
import { scan, findEnglish } from '../tools/find-english.mjs';

test('every word a window shows goes through the catalogue', () => {
  const left = scan().filter((f) => f.findings.length);
  const listed = left.flatMap((f) => f.findings.slice(0, 5).map((x) => `${f.file}:${x.line} ${x.kind} ${x.text}`));
  assert.equal(left.length, 0, `English left in ${left.length} files:\n${listed.join('\n')}`);
});

test('the finder sees words in the shapes they take, and leaves code and marked names alone', () => {
  const found = (src) => findEnglish(src).map((f) => f.text);
  assert.deepEqual(found('<Button label="Save as" />'), ['Save as']);
  assert.deepEqual(found("items.push({ label: 'Open recent', run })"), ['Open recent']);
  assert.deepEqual(found("toast('Saved the file.')"), ['Saved the file.']);
  assert.deepEqual(found('<span>Line spacing</span>'), ['Line spacing']);
  assert.deepEqual(found("<Button label={t('Save as')} />"), []);
  assert.deepEqual(found('<b>{x > 1 ? a : b}</b>'), []);
  assert.deepEqual(found('<option value="Calibri">Calibri</option> {/* words-ok: a font */}'), []);
  assert.deepEqual(found('className="rw-btn tall"'), []);
});
