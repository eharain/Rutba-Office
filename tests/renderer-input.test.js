// Shell defects that no unit test would catch, pinned by reading the source.
//
// Both of these were shipped in this repository, both looked like the product
// simply not working, and both are one line:
//
// 1. React's `onBeforeInput` is not the `beforeinput` event. It is a synthetic
//    event react-dom assembles from `keypress`, `textInput` and
//    `compositionend` — a polyfill older than the standard — and it carries no
//    `inputType`. The document editor cancelled the browser's own insertion and
//    then asked what to do about `undefined`, which is nothing: typing did
//    absolutely nothing, silently. Backspace never even arrived, because
//    `keypress` does not fire for it.
//
// 2. `position: sticky` keeps a box in normal flow. Column headers that are
//    block-level and carry an inline `left` therefore each start a new line
//    further right than the last, and the header row renders as a staircase
//    descending across the sheet.
//
// The first of these was introduced here by deleting the very test that had
// been written to prevent it, which is the argument for this file existing in
// the same repository as the code it guards.

import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync, statSync } from 'node:fs';
import { join, dirname, relative } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const RENDERER = join(ROOT, 'apps', 'desktop', 'renderer');

function sources(dir) {
  const out = [];
  const walk = (d) => {
    for (const entry of readdirSync(d)) {
      if (entry === 'node_modules' || entry === 'build') continue;
      const path = join(d, entry);
      if (statSync(path).isDirectory()) walk(path);
      else if (entry.endsWith('.js')) out.push(path);
    }
  };
  walk(dir);
  return out;
}

const FILES = sources(RENDERER).map((path) => ({ path, name: relative(ROOT, path), text: readFileSync(path, 'utf8') }));

test('the renderer has sources to check', () => {
  assert.ok(FILES.length >= 8, `found ${FILES.length} renderer sources`);
});

test('no editor relies on React onBeforeInput — it carries no inputType', () => {
  const offenders = FILES.filter((f) => /onBeforeInput\s*=\s*\{/.test(f.text)).map((f) => f.name);
  assert.deepEqual(
    offenders,
    [],
    'React\'s onBeforeInput is a polyfill with no inputType and no backspace. ' +
      'Attach a native listener: el.addEventListener("beforeinput", …).'
  );
});

test('every editor that cancels beforeinput also listens for it natively', () => {
  const cancels = FILES.filter((f) => /inputType/.test(f.text));
  assert.ok(cancels.length > 0, 'at least one surface reads inputType');
  for (const file of cancels) {
    assert.match(
      file.text,
      /addEventListener\(\s*['"]beforeinput['"]/,
      `${file.name} reads inputType but never attaches a native beforeinput listener`
    );
  }
});

test('a cancelled beforeinput is always turned into an operation', () => {
  // Cancelling the browser's edit and then not replacing it is the failure
  // mode; every handler that calls preventDefault must dispatch something.
  const word = FILES.find((f) => f.name.endsWith('word.js'));
  assert.ok(word, 'the document editor is present');
  // From the handler's definition to where it is attached — bounded by two
  // things that can only appear in that order.
  const from = word.text.indexOf('const handleBeforeInput');
  const to = word.text.indexOf("addEventListener('beforeinput'", from);
  assert.ok(from >= 0 && to > from, 'the handler is defined before it is attached');
  const handler = word.text.slice(from, to);
  assert.match(handler, /e\.preventDefault\(\)/, 'the browser edit is cancelled');
  assert.match(handler, /apply\(/, 'and something is applied in its place');
  for (const inputType of ['insertText', 'insertParagraph', 'deleteContentBackward', 'deleteContentForward']) {
    assert.match(handler, new RegExp(`'${inputType}'`), `${inputType} is handled`);
  }
});

test('the sheet grid does not lay its headers out with sticky block elements', () => {
  // The staircase: sticky keeps a box in flow, so block-level headers with an
  // inline `left` each start a new line further across than the last.
  const sheets = FILES.find((f) => f.name.endsWith('sheets.js'));
  assert.ok(sheets, 'the sheet is present');
  assert.match(sheets.text, /\.sh-canvas \{ display: grid; \}/, 'the grid panes are laid out with CSS grid');
  assert.match(sheets.text, /grid-row: 1; grid-column: 2;/, 'the column headers have a track of their own');
  assert.match(sheets.text, /grid-row: 2; grid-column: 1;/, 'the row headers have a track of their own');
});

test('the cell editor keeps its own text while it is being typed', () => {
  // Sending every keystroke to the engine and reading the value back made the
  // field lag the keyboard; fast typing dropped and reordered characters.
  const sheets = FILES.find((f) => f.name.endsWith('sheets.js'));
  assert.match(sheets.text, /const \[draft, setDraft\] = useState\(null\)/, 'there is a local draft');
  assert.doesNotMatch(
    sheets.text,
    /onChange=\{\(e\) => dispatch\(\{ op: 'updateDraft'/,
    'the editor no longer round-trips per character'
  );
  assert.match(sheets.text, /startingRef/, 'starting an edit is decided synchronously');
});

test('typing into a document sends a patch, not the whole document', () => {
  const service = readFileSync(join(ROOT, 'apps', 'desktop', 'main', 'documents.js'), 'utf8');
  assert.match(service, /function docDelta/, 'the service can answer with a difference');
  assert.match(service, /session\.lastBlocks/, 'and remembers what the window already has');

  const word = FILES.find((f) => f.name.endsWith('word.js'));
  assert.match(word.text, /next\.patch/, 'the editor applies a patch when it gets one');
  assert.match(word.text, /React\.memo/, 'and unchanged paragraphs do not re-render');
});
