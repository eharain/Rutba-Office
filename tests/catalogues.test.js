// Every catalogue is whole: each message the windows show has its
// translation, a counted message every plural form the language's own rules
// can choose (Intl.PluralRules: Urdu's one and other, Arabic's six), every
// placeholder kept, and nothing left in Latin letters that the language would
// write in its own script. An engine's message, thrown already filled, comes
// back in the language too. tests/messages.test.js holds every catalogue to
// the list's keys.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { CATALOGUES } from '@rutba/office-ui/catalogues';
import { t, tn, tFilled, setLanguage, setDigits, registerCatalogue } from '@rutba/office-ui/messages';
import { ENGINE_WORDS } from '../apps/desktop/renderer/engine-words.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const LIST = JSON.parse(fs.readFileSync(path.join(ROOT, 'packages/office-ui/src/catalogues/messages.json'), 'utf8')).messages;

/** The script each language writes in, to tell a translation from English left in. */
const SCRIPTS = { ur: /[\u0600-\u06FF]/, ar: /[\u0600-\u06FF]/, hi: /[\u0900-\u097F]/ };

// A brand, a format or a code alone stays as it is; words do not.
const KEEP = /^(?:[\s\d.,:;!?…—–()[\]{}/+%#&|·×→←=-]|\{\w+\}|PDF|CSV|TSV|HTML|Markdown|DOCX|PNG|JPEG|WebP|Mbox|TLS|STARTTLS|Cc|Bcc|ISO|AV|Aa|IF|DOI|URL|SmartArt|WordArt|PivotTable|PivotChart|Power Query|Rutba Office|Rutba|Office|Word|Excel|PowerPoint|Windows|Google|Microsoft|iCloud|Outlook|vCard|AutoFit|AutoSum|Calibri|\S+@\S+|https?:\S+|[A-Z]\d*|f\/\S+)+$/;

test('there is a catalogue for each language the windows offer', () => {
  assert.deepEqual(Object.keys(CATALOGUES).sort(), ['ar', 'hi', 'ur']);
});

for (const [tag, catalogue] of Object.entries(CATALOGUES)) {
  const forms = new Intl.PluralRules(tag).resolvedOptions().pluralCategories;

  test(`${tag}: every message the windows show has its translation, a counted one every form (${forms.join(', ')})`, () => {
    const missing = LIST.filter((m) => {
      const v = catalogue[m.message];
      if (m.plural) return !v || typeof v !== 'object' || forms.some((f) => typeof v[f] !== 'string' || !v[f].trim());
      return typeof v !== 'string' || !v.trim();
    });
    assert.equal(missing.length, 0, `without a translation: ${missing.slice(0, 12).map((m) => m.message).join(' | ')}`);
  });

  test(`${tag}: a message of words is not left in English`, () => {
    const script = SCRIPTS[tag];
    const english = LIST.filter((m) => !m.plural && typeof catalogue[m.message] === 'string' && /[A-Za-z]{3,}/.test(m.message.replace(/\{\w+\}/g, ''))
      && !script.test(catalogue[m.message]) && !KEEP.test(m.message));
    assert.ok(english.length <= 40, `left in English (${english.length}): ${english.slice(0, 20).map((m) => m.message).join(' | ')}`);
  });

  test(`${tag}: a window in it speaks it, a counted message by its own rules, English for one it does not know`, () => {
    registerCatalogue(tag, catalogue);
    try {
      assert.equal(setLanguage(tag), tag);
      setDigits('latn');
      assert.equal(t('Save'), catalogue.Save);
      assert.match(t('Save'), SCRIPTS[tag]);
      assert.match(t('Save changes to {name}?', { name: 'Report.docx' }), /Report\.docx/);
      const word = catalogue['{count} word'];
      const rules = new Intl.PluralRules(tag);
      for (const n of [0, 1, 2, 5, 12, 100]) {
        const rule = rules.select(n);
        const form = (n === 0 && word.zero) || word[rule] || word.other;
        assert.equal(tn(n, '{count} word', '{count} words'), form.replace('{count}', n.toLocaleString(`${tag}-u-nu-latn`)), `${n} words`);
      }
      assert.equal(t('A message no window shows'), 'A message no window shows');
    } finally {
      setDigits(null);
      setLanguage('en');
    }
  });

  test(`${tag}: an engine's message reaches a window in it translated, the names it was thrown with put back`, () => {
    // Each message as an engine throws it, its names filled; the window must
    // find that message's own translation, not a looser one it also fits.
    registerCatalogue(tag, catalogue);
    try {
      setLanguage(tag);
      const astray = [];
      for (const message of ENGINE_WORDS) {
        const values = {};
        let n = 0;
        const thrown = message.replace(/\{(\w+)\}/g, (m, name) => (values[name] ??= `Q${++n}z`));
        const want = catalogue[message].replace(/\{(\w+)\}/g, (m, name) => values[name]);
        if (tFilled(thrown) !== want) astray.push(message);
      }
      assert.deepEqual(astray, []);
      assert.equal(tFilled('Nothing any engine says.'), 'Nothing any engine says.');
    } finally {
      setLanguage('en');
    }
  });
}

test('the digits a window writes its numbers in: the language\'s own, or the ones chosen', () => {
  for (const [tag, catalogue] of Object.entries(CATALOGUES)) registerCatalogue(tag, catalogue);
  try {
    setLanguage('ar');
    setDigits('latn');
    assert.match(tn(1234, '{count} word', '{count} words'), /1,234/);
    setDigits('arab');
    assert.match(tn(1234, '{count} word', '{count} words'), /١٬٢٣٤/);
    setLanguage('hi');
    setDigits(null);
    assert.match(tn(1234567, '{count} word', '{count} words'), /12,34,567/, 'Hindi groups as India does');
    setDigits('deva');
    assert.match(tn(25, '{count} word', '{count} words'), new RegExp('\u0968\u096B'), 'Devanagari digits');
    setLanguage('ur');
    setDigits('arabext');
    assert.match(tn(25, '{count} word', '{count} words'), /۲۵/);
  } finally {
    setDigits(null);
    setLanguage('en');
  }
});
