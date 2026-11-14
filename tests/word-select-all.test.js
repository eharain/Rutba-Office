// Documents: Ctrl+A is the engine's Select All, as the ribbon's is. Left to
// the browser, Ctrl+A showed the whole page selected while the engine kept
// the old caret: a colour, Bold or a size chosen next went to the caret, and
// the page then drew the caret, so the selection vanished too.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { selectAllKey } from '../apps/desktop/renderer/apps/word/caret.js';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';
import { asEnglish } from './label-english.js';

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const on = (target) => ({ closest: (sel) => (sel.split(',').map((s) => s.trim()).includes(target) ? {} : null) });

test('Ctrl+A and Cmd+A are Select All; with Shift or Alt, or a bare A, they are not', () => {
  assert.equal(selectAllKey({ ctrlKey: true, key: 'a', code: 'KeyA' }), true);
  assert.equal(selectAllKey({ ctrlKey: true, key: 'A', code: 'KeyA' }), true, 'Caps Lock on');
  assert.equal(selectAllKey({ metaKey: true, key: 'a', code: 'KeyA' }), true, 'a Mac');
  assert.equal(selectAllKey({ ctrlKey: true, key: 'ش', code: 'KeyA' }), true, 'an Arabic or Urdu layout: the key, not its letter');
  assert.equal(selectAllKey({ ctrlKey: true, shiftKey: true, key: 'A', code: 'KeyA' }), false);
  assert.equal(selectAllKey({ ctrlKey: true, altKey: true, key: 'a', code: 'KeyA' }), false);
  assert.equal(selectAllKey({ key: 'a', code: 'KeyA' }), false);
  assert.equal(selectAllKey({ ctrlKey: true, key: 'b', code: 'KeyB' }), false);
  assert.equal(selectAllKey(null), false);
});

test('inside a form field\'s own box Ctrl+A is the field\'s', () => {
  assert.equal(selectAllKey({ ctrlKey: true, key: 'a', code: 'KeyA', target: on('.wd-cc') }), false);
  assert.equal(selectAllKey({ ctrlKey: true, key: 'a', code: 'KeyA', target: on('input') }), false);
  assert.equal(selectAllKey({ ctrlKey: true, key: 'a', code: 'KeyA', target: on('.wd-block') }), true, 'anywhere on the page, a table cell included');
});

test('the page sends Ctrl+A to the engine\'s Select All, the op the ribbon\'s Select All sends, and keeps it from the browser', () => {
  const word = asEnglish(fs.readFileSync(path.join(root, 'apps/desktop/renderer/apps/word.js'), 'utf8'));
  const key = word.indexOf('if (selectAllKey(e)');
  assert.ok(key > 0, 'the page\'s key handler asks selectAllKey');
  const body = word.slice(key, word.indexOf('return;', key));
  assert.match(body, /e\.preventDefault\(\)/, 'the browser does not select on its own');
  assert.match(body, /apply\(\{ op: 'selectAll' \}\)/, 'the engine selects');
  const ribbon = asEnglish(fs.readFileSync(path.join(root, 'apps/desktop/renderer/apps/word/ribbon.js'), 'utf8'));
  assert.match(ribbon, /title=\{t\('Select all \(Ctrl\+A\)'\)\} onClick=\{\(\) => dispatch\(\{ op: 'selectAll' \}\)\}/, 'the ribbon sends the same op');
});

test('after Select All a colour, Bold and a size go on every paragraph, table cells included, and the selection stays', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'One' }, { table: { rows: [['cell', 'other']] } }, { text: 'Last' }] }));
  view.setSelection({ block: 1, offset: 2 }, { block: 1, offset: 2 });
  view.selectAll();
  const range = { from: { ...view.selection.from }, to: { ...view.selection.to } };
  view.setRunFormat({ fontColour: 'FF0000' });
  view.toggleFormat('b');
  view.setRunFormat({ fontSize: 16 });
  for (const b of view.blocks) {
    for (const r of b.runs.filter((x) => x.text)) {
      const p = view.doc.readRunProps(r.rPr);
      assert.equal(p.fontColour, 'FF0000', `${r.text}: red`);
      assert.equal(p.fontSize, 16, `${r.text}: 16 pt`);
      assert.match(r.rPr, /<w:b\/>/, `${r.text}: bold`);
    }
  }
  assert.equal(view.collapsed, false);
  assert.deepEqual({ from: { ...view.selection.from }, to: { ...view.selection.to } }, range, 'the same selection after all three');
});
