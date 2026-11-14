// The checks find a control by the key its label's English names (stage 4 of
// the meaning keys, D-2026-11-14-16): with an English copy the window words a
// label its own way and a button's tip is its complete version, so a check
// that matched the words or the tip missed it (the integrator's causes 2 and
// 3 of 14 Nov). apps/desktop/main/verify-keys.js gives the predicate each
// finder tries first; its words are the fallback for words that name no label.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { keyFor, keyParts, labelKeyed, keySelector } from '../apps/desktop/main/verify-keys.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const MAIN = path.join(ROOT, 'apps', 'desktop', 'main');

/** An element as the window has it, for the predicate: its dataset alone. */
const el = (dataset) => ({ dataset });

test("a label's English gives its key, and the predicate holds for an element carrying it as its label or as its tip", () => {
  assert.equal(keyFor('Multiple Pages'), 'word.ribbon.multiple');
  assert.deepEqual(keyParts('Multiple Pages'), { scope: 'word.ribbon', name: 'multiple' });
  const is = eval(labelKeyed('Multiple Pages'));
  assert.equal(is(el({ scope: 'word.ribbon', key: 'multiple' })), true, 'its label');
  assert.equal(is(el({ scope: 'word', key: 'multiple' })), false, 'another scope');
  assert.equal(is(el({ scope: 'word.ribbon', key: 'multiple-pages' })), false, 'another name');
  assert.equal(is(el({})), false);
  assert.equal(is(null), false);
  // Named by its tip: Change, whose tip is "Change, this one" in its own words.
  const english = Object.entries(JSON.parse(fs.readFileSync(path.join(ROOT, 'packages/office-ui/src/catalogues/labels.json'), 'utf8')).labels).find(([k]) => k === 'ui.proofing.change-this-one')[1].english;
  const tip = eval(labelKeyed(english));
  assert.equal(tip(el({ scope: 'common', key: 'change', tipScope: 'ui.proofing', tipKey: 'change-this-one' })), true, 'its tip');
  assert.equal(tip(el({ scope: 'common', key: 'change' })), false);
  assert.equal(keySelector('Multiple Pages'), '[data-scope="word.ribbon"][data-key="multiple"], [data-tip-scope="word.ribbon"][data-tip-key="multiple"]');
});

test("the start of one label's words names that label; the start of several names none", () => {
  const labels = JSON.parse(fs.readFileSync(path.join(ROOT, 'packages/office-ui/src/catalogues/labels.json'), 'utf8')).labels;
  const citations = labels['word.references-toa.update-table-citations'].english;
  // The TOA check names its Update Table by the start of its tip.
  const start = citations.slice(0, citations.indexOf(' the citations') + ' the citations'.length);
  assert.notEqual(start, citations);
  assert.equal(keyFor(start), 'word.references-toa.update-table-citations');
  assert.equal(keyFor('Update Table'), 'word.update-table', 'a label of its own words first');
  assert.equal(keyFor('Grou'), null, 'Group, Group by and more begin so');
});

test('words that name no label give a predicate that holds for none, so a check falls back to its words', () => {
  for (const words of ['Report.docx', 'Quarterly Review', '', null, undefined]) {
    assert.equal(keyFor(words), null);
    assert.equal(keyParts(words), null);
    assert.equal(eval(labelKeyed(words))(el({ scope: 'common', key: 'save' })), false);
    assert.equal(keySelector(words), null);
  }
});

test('a key the predicate names is written as data, never as code', () => {
  // The scope and name are JSON strings in the source a window runs.
  assert.match(labelKeyed('Multiple Pages'), /n\.dataset\.scope === "word\.ribbon" && n\.dataset\.key === "multiple"/);
});

test("every check that finds a control by its words tries the label's key first, and imports the predicate", () => {
  const scripts = fs.readdirSync(MAIN).filter((f) => /^verify-.*\.js$/.test(f) && f !== 'verify-keys.js');
  const byWords = /\[\.\.\.document\.querySelectorAll\('[^']*'\)\]\.find\(\((\w)\) => (?:\1\.textContent\.trim\(\) === \$\{JSON\.stringify\(\w+\)\}|[^;\n]*?startsWith\(\$\{JSON\.stringify\(\w+\)\}\))/;
  const left = [];
  const unimported = [];
  for (const f of scripts) {
    const text = fs.readFileSync(path.join(MAIN, f), 'utf8');
    if (byWords.test(text)) left.push(f);
    if (text.includes('${labelKeyed(') && !/import \{ labelKeyed \} from '\.\/verify-keys\.js';/.test(text)) unimported.push(f);
  }
  assert.deepEqual(left, [], 'a finder by words alone');
  assert.deepEqual(unimported, [], 'uses labelKeyed without importing it');
  // No script has a name of its own that would hide the predicate.
  const hidden = scripts.filter((f) => /\b(?:const|let|var|function)\s+labelKeyed\b/.test(fs.readFileSync(path.join(MAIN, f), 'utf8')));
  assert.deepEqual(hidden, []);
});
