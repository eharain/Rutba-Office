// Talk's words from Office 2.25 on are translated in a file of their own
// (apps/desktop/renderer/apps/talk-words.js): every one Talk shows is there or
// in the shared catalogues, in every language the suite ships, with every
// plural form, its places kept, and written in the language's own script.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { TALK_WORDS } from '../apps/desktop/renderer/apps/talk-translations.js';
import ur from '../packages/office-ui/src/catalogues/ur.js';
import ar from '../packages/office-ui/src/catalogues/ar.js';
import hi from '../packages/office-ui/src/catalogues/hi.js';

const SHARED = { ur, ar, hi };
const SCRIPTS = { ur: /[\u0600-\u06FF]/, ar: /[\u0600-\u06FF]/, hi: /[\u0900-\u097F]/ };
/** The files that take their words from talk-words.js. */
const FILES = ['talk/invite-email.js', 'talk/connect.js', 'talk/address.js'];
const STR = String.raw`'((?:[^'\\\n]|\\.)*)'`;
const unescape = (s) => s.replace(/\\(n|.)/g, (m, c) => (c === 'n' ? '\n' : c));
const places = (s) => [...String(s).matchAll(/\{(\w+)\}/g)].map((m) => m[1]).sort();

const used = () => {
  const words = new Map();
  for (const f of FILES) {
    const text = fs.readFileSync(new URL(`../apps/desktop/renderer/apps/${f}`, import.meta.url), 'utf8');
    assert.match(text, /from '\.\.\/talk-words\.js'/, `${f} takes its words from talk-words.js`);
    assert.doesNotMatch(text, /import\s*\{[^}]*\bt\b[^}]*\}\s*from\s*'@rutba\/office-ui'/, `${f} takes no t() from the shared catalogue`);
    for (const m of text.matchAll(new RegExp(String.raw`(?<![\w.$])t\(\s*${STR}`, 'g'))) words.set(unescape(m[1]), { plural: null });
    for (const m of text.matchAll(new RegExp(String.raw`(?<![\w.$])tn\([^']*?,\s*${STR}\s*,\s*${STR}`, 'g'))) words.set(unescape(m[1]), { plural: unescape(m[2]) });
  }
  return words;
};

test('every word Talk shows from its own words is translated, and nothing is kept that it no longer shows', () => {
  const words = used();
  assert.ok(words.size > 20, 'the words were read');
  for (const [tag, list] of Object.entries(TALK_WORDS)) {
    assert.deepEqual([...words.keys()].filter((w) => !list[w] && !SHARED[tag][w]), [], `${tag}: not translated`);
    assert.deepEqual(Object.keys(list).filter((w) => !words.has(w)), [], `${tag}: no longer used`);
    assert.deepEqual(Object.keys(list).filter((w) => SHARED[tag][w] !== undefined), [], `${tag}: already in the shared catalogue`);
  }
});

test('each translation has every plural form its language needs, keeps its places, and is in its own script', () => {
  const words = used();
  assert.deepEqual(Object.keys(TALK_WORDS).sort(), Object.keys(SHARED).sort(), 'every language the suite ships');
  for (const [tag, list] of Object.entries(TALK_WORDS)) {
    const forms = new Intl.PluralRules(tag).resolvedOptions().pluralCategories;
    for (const [english, v] of Object.entries(list)) {
      const plural = words.get(english)?.plural;
      if (plural) {
        assert.equal(typeof v, 'object', `${tag}: "${english}" is counted`);
        for (const f of forms) assert.ok(typeof v[f] === 'string' && v[f].trim(), `${tag}: "${english}" has its ${f} form`);
        for (const f of forms) {
          const want = places(f === 'one' ? english : plural).filter((p) => p !== 'count');
          assert.deepEqual(places(v[f]).filter((p) => p !== 'count'), want, `${tag}: "${english}" (${f}) keeps its places`);
          assert.match(v[f], SCRIPTS[tag], `${tag}: "${english}" (${f}) is written in ${tag}`);
        }
      } else {
        assert.equal(typeof v, 'string', `${tag}: "${english}"`);
        assert.deepEqual(places(v), places(english), `${tag}: "${english}" keeps its places`);
        assert.match(v, SCRIPTS[tag], `${tag}: "${english}" is written in ${tag}`);
      }
    }
  }
});
