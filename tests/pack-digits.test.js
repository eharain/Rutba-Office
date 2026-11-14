// The digits a window writes its numbers in, in Arabic, Urdu and Hindi, with
// those languages' words taken from their packs as a window takes them
// (readLanguagePack, tableOf, registerPack): the language's own digits, or
// the ones chosen, and Hindi's numbers grouped as India groups them. Office
// carries no words of these languages since D-2026-11-14-16, so the test
// reads the packs the release ships when RUTBA_PACKS names them
// (<packages>/languages/packs), and otherwise packs written here with the
// counted label's forms as those packs give them, which the RUTBA_PACKS run
// holds to the real packs. Every plural form and nothing left in English, per
// language, are the packages repository's checks (languages/tools/check-packs.mjs).
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { tn, setLanguage, setDigits, registerPack } from '@rutba/office-ui/messages';
import { codeLabels, keyOf, hashOf, tableOf, readLanguagePack } from '../packages/office-shell/src/electron/language-packs.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const labels = JSON.parse(fs.readFileSync(path.join(ROOT, 'packages/office-ui/src/catalogues/labels.json'), 'utf8'));
const keys = codeLabels(labels);
const hashes = keys.map((m) => (typeof m.source === 'string' && /^[0-9a-f]{12}$/.test(m.source) ? m.source : hashOf(keyOf(m))));
const COUNTED = 'common.count-word';
const source = hashes[keys.findIndex((m) => m.key === COUNTED)];

/** The counted label's forms, as lang-ar, lang-ur and lang-hi 2.26.0 give them. */
const WRITTEN = {
  ar: { zero: 'لا توجد كلمات', one: 'كلمة واحدة', two: 'كلمتان', few: '{count} كلمات', many: '{count} كلمة', other: '{count} كلمة' },
  ur: { one: '{count} لفظ', other: '{count} الفاظ' },
  hi: { one: '{count} शब्द', other: '{count} शब्द' },
};

/** Each language's pack entries: the release's packs, or the ones written here. */
function entriesOf(tag) {
  if (process.env.RUTBA_PACKS) {
    const pack = readLanguagePack(path.join(process.env.RUTBA_PACKS, `lang-${tag}`));
    assert.ok(pack, `lang-${tag} in ${process.env.RUTBA_PACKS}`);
    return pack.entries;
  }
  return { [`pack.${COUNTED}`]: { forms: WRITTEN[tag], source } };
}

test('the counted label is the one the packs word, and the forms written here are the packs\' own', (t) => {
  assert.ok(source, `${COUNTED} is a label`);
  if (!process.env.RUTBA_PACKS) return t.diagnostic('RUTBA_PACKS names no packs: the forms written here stand for them');
  for (const tag of Object.keys(WRITTEN)) {
    const entry = Object.values(readLanguagePack(path.join(process.env.RUTBA_PACKS, `lang-${tag}`)).entries).find((e) => e?.source === source);
    assert.deepEqual(entry?.forms, WRITTEN[tag], `lang-${tag}'s forms for ${COUNTED}: copy them here`);
  }
});

test('the digits a window writes its numbers in: the language\'s own, or the ones chosen, and Hindi grouped as India does', () => {
  for (const tag of Object.keys(WRITTEN)) {
    const { table, used } = tableOf(keys, hashes, entriesOf(tag));
    assert.ok(used >= 1 && table[COUNTED], `lang-${tag} words ${COUNTED}`);
    registerPack(tag, table);
  }
  const counted = (n, tag) => {
    const forms = WRITTEN[tag];
    const rule = new Intl.PluralRules(tag).select(n);
    return (n === 0 && forms.zero) || forms[rule] || forms.other;
  };
  try {
    setLanguage('ar');
    setDigits('latn');
    assert.equal(tn(1234, COUNTED), counted(1234, 'ar').replace('{count}', '1,234'));
    setDigits('arab');
    assert.match(tn(1234, COUNTED), /١٬٢٣٤/);
    assert.equal(tn(2, COUNTED), 'كلمتان', 'Arabic\'s dual, from the pack');
    setLanguage('hi');
    setDigits(null);
    assert.match(tn(1234567, COUNTED), /12,34,567/, 'Hindi groups as India does');
    setDigits('deva');
    assert.match(tn(25, COUNTED), new RegExp('२५'), 'Devanagari digits');
    assert.match(tn(25, COUNTED), /शब्द/);
    setLanguage('ur');
    setDigits('arabext');
    assert.match(tn(25, COUNTED), /۲۵/);
    assert.equal(tn(1, COUNTED), '۱ لفظ');
  } finally {
    setDigits(null);
    setLanguage('en');
  }
});
