// The colour palette's words are translated in a file of their own
// (apps/desktop/renderer/apps/colour-translations.js): every word the palette
// shows is there or in the shared catalogues, in every language the suite
// ships, with its places kept, and written in the language's own script.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { COLOUR_WORDS } from '../apps/desktop/renderer/apps/colour-translations.js';
import ur from '../packages/office-ui/src/catalogues/ur.js';
import ar from '../packages/office-ui/src/catalogues/ar.js';
import hi from '../packages/office-ui/src/catalogues/hi.js';
import { englishOf } from './label-english.js';

const SHARED = { ur, ar, hi };
const SCRIPTS = { ur: /[؀-ۿ]/, ar: /[؀-ۿ]/, hi: /[ऀ-ॿ]/ };
/** The files that take their words from colour-words.js. */
const FILES = ['colour/palette.js'];
const unescape = (s) => s.replace(/\\(n|.)/g, (m, c) => (c === 'n' ? '\n' : c));
const places = (s) => [...String(s).matchAll(/\{(\w+)\}/g)].map((m) => m[1]).sort();

const used = () => {
  const words = new Set();
  for (const f of FILES) {
    const text = fs.readFileSync(new URL(`../apps/desktop/renderer/apps/${f}`, import.meta.url), 'utf8');
    assert.match(text, /from '\.\.\/colour-words\.js'/, `${f} takes its words from colour-words.js`);
    assert.doesNotMatch(text, /import\s*\{[^}]*\bt\b[^}]*\}\s*from\s*'@rutba\/office-ui'/, `${f} takes no t() from the shared catalogue`);
    for (const m of text.matchAll(/(?<![\w.$])t\(\s*'((?:[^'\\\n]|\\.)*)'/g)) words.add(englishOf(unescape(m[1])));
  }
  return words;
};

test('every word the palette shows is translated, and nothing is kept that it no longer shows', () => {
  const words = used();
  assert.ok(words.size >= 10, 'the words were read');
  for (const [tag, list] of Object.entries(COLOUR_WORDS)) {
    assert.deepEqual([...words].filter((w) => !list[w] && !SHARED[tag][w]), [], `${tag}: not translated`);
    assert.deepEqual(Object.keys(list).filter((w) => !words.has(w)), [], `${tag}: no longer used`);
    assert.deepEqual(Object.keys(list).filter((w) => SHARED[tag][w] !== undefined), [], `${tag}: already in the shared catalogue`);
  }
});

test('the words the palette shares with the suite come from the shared catalogues, in every language', () => {
  const text = fs.readFileSync(new URL('../apps/desktop/renderer/apps/colour/names.js', import.meta.url), 'utf8');
  // Taken from the shared catalogue, so the list of messages keeps them.
  assert.match(text, /import \{ t \} from '@rutba\/office-ui'/);
  const shared = [...text.matchAll(/(?<![\w.$])t\(\s*'((?:[^'\\\n]|\\.)*)'/g)].map((m) => unescape(m[1]));
  assert.ok(shared.length > 30, 'the words were read');
  for (const [tag, list] of Object.entries(SHARED)) assert.deepEqual(shared.filter((w) => !list[w]), [], `${tag}: not in the shared catalogue`);
});

test('each translation keeps its places and is in its own script, in every language the suite ships', () => {
  assert.deepEqual(Object.keys(COLOUR_WORDS).sort(), Object.keys(SHARED).sort());
  for (const [tag, list] of Object.entries(COLOUR_WORDS)) {
    for (const [english, v] of Object.entries(list)) {
      assert.equal(typeof v, 'string', `${tag}: "${english}"`);
      assert.deepEqual(places(v), places(english), `${tag}: "${english}" keeps its places`);
      assert.match(v, SCRIPTS[tag], `${tag}: "${english}" is written in ${tag}`);
    }
  }
});
