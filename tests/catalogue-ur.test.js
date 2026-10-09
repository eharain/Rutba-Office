// The Urdu catalogue is whole: every message the windows show has its Urdu,
// a counted message both of Urdu's forms (one and other), every placeholder
// kept, and nothing left in Latin letters that Urdu would write in its own.
// tests/messages.test.js holds every catalogue to the list's keys.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { CATALOGUES } from '@rutba/office-ui/catalogues';
import { t, tn, setLanguage, registerCatalogue } from '@rutba/office-ui/messages';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const LIST = JSON.parse(fs.readFileSync(path.join(ROOT, 'packages/office-ui/src/catalogues/messages.json'), 'utf8')).messages;
const UR = CATALOGUES.ur;

test('every message the windows show has its Urdu, a counted one both forms', () => {
  assert.ok(UR, 'an Urdu catalogue');
  const missing = LIST.filter((m) => {
    const v = UR[m.message];
    if (m.plural) return !v || typeof v !== 'object' || !v.one || !v.other;
    return typeof v !== 'string' || !v.trim();
  });
  assert.equal(missing.length, 0, `without Urdu: ${missing.slice(0, 12).map((m) => m.message).join(' | ')}`);
});

test('the Urdu is Urdu: a message of words is not left in English', () => {
  // A brand, a format or a code alone stays as it is; words do not.
  const keep = /^(?:[\s\d.,:;!?…—–()[\]{}/+%#&|·×→←=-]|\{\w+\}|PDF|CSV|TSV|HTML|Markdown|DOCX|PNG|JPEG|WebP|Mbox|TLS|STARTTLS|Cc|Bcc|ISO|AV|Aa|IF|SmartArt|WordArt|PivotTable|PivotChart|Power Query|Rutba Office|Rutba|Office|Word|Excel|PowerPoint|Windows|Google|Microsoft|iCloud|Outlook|vCard|AutoFit|AutoSum|Calibri|\S+@\S+|https?:\S+|[A-Z]\d*|f\/\S+)+$/;
  const english = LIST.filter((m) => !m.plural && typeof UR[m.message] === 'string' && /[A-Za-z]{3,}/.test(m.message.replace(/\{\w+\}/g, ''))
    && !/[\u0600-\u06FF]/.test(UR[m.message]) && !keep.test(m.message));
  assert.ok(english.length <= 40, `left in English (${english.length}): ${english.slice(0, 20).map((m) => m.message).join(' | ')}`);
});

test('an Urdu window speaks Urdu: a message, a counted one by Urdu\'s rules, and English for one it does not know', () => {
  registerCatalogue('ur', UR);
  try {
    assert.equal(setLanguage('ur-PK'), 'ur', 'Pakistan\'s Urdu finds the Urdu catalogue');
    assert.equal(t('Save'), UR.Save);
    assert.match(t('Save'), /[\u0600-\u06FF]/);
    assert.match(t('Save changes to {name}?', { name: 'Report.docx' }), /Report\.docx/);
    const one = tn(1, '{count} word', '{count} words');
    const many = tn(5, '{count} word', '{count} words');
    assert.equal(one, UR['{count} word'].one.replace('{count}', '1'));
    assert.equal(many, UR['{count} word'].other.replace('{count}', '5'));
    assert.equal(t('A message no window shows'), 'A message no window shows');
  } finally {
    setLanguage('en');
  }
});
