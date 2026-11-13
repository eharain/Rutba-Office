// A label's short and complete versions (L3): what versionsOf finds for a text
// t() or tn() gave, filled with the same names and in the same plural form;
// a language's own versions only, lang-en's only for a label shown in English
// for want of a translation, and none for a text two labels give alike.
import test from 'node:test';
import assert from 'node:assert/strict';
import { t, tn, setLanguage, registerCatalogue, setFallback, setVersions, versionsOf } from '@rutba/office-ui/messages';

const english = {
  'Save as': { short: 'Save as', complete: 'Save a copy under another name' },
  'Paragraph Spacing': { short: 'Spacing' },
  'Saved {name}': { short: 'Saved', complete: 'Saved {name} to this computer' },
  '{count} file': { short: { one: '1 f', other: '{count} f' } },
};

test('in a French window a label has its French versions, never English ones', () => {
  registerCatalogue('fr', {
    'Paragraph Spacing': 'Espacement des paragraphes',
    'Saved {name}': '{name} enregistré',
    '{count} file': { one: '{count} fichier', other: '{count} fichiers' },
  });
  setLanguage('fr');
  setFallback(null);
  setVersions({ 'Saved {name}': { short: '{name} enreg.', complete: '{name} enregistré sur cet ordinateur' }, '{count} file': { short: { one: '{count} fich.', other: '{count} fich.' } } }, english);
  try {
    // Its own short, with the same names filled.
    assert.equal(t('Saved {name}', { name: 'Budget' }), 'Budget enregistré');
    assert.deepEqual(versionsOf('Budget enregistré'), { short: 'Budget enreg.', complete: 'Budget enregistré sur cet ordinateur' });
    // Translated, with no French short: the French medium alone, never an English short.
    assert.equal(t('Paragraph Spacing'), 'Espacement des paragraphes');
    assert.equal(versionsOf('Espacement des paragraphes'), null);
    // Shown in English for want of a translation: lang-en's versions.
    assert.equal(t('Save as'), 'Save as');
    assert.deepEqual(versionsOf('Save as'), { short: 'Save as', complete: 'Save a copy under another name' });
    // A counted label's short follows the medium's form.
    assert.equal(tn(3, '{count} file', '{count} files'), '3 fichiers');
    assert.deepEqual(versionsOf('3 fichiers'), { short: '3 fich.' });
    // Not given lately: nothing.
    assert.equal(versionsOf('Jamais vu'), null);
  } finally {
    setVersions(null, null);
    setLanguage('en');
  }
});

test('two labels that read alike but have different versions give none', () => {
  registerCatalogue('fr', { Close: 'Fermer', 'Close the file': 'Fermer' });
  setLanguage('fr');
  setVersions({ Close: { short: 'Ferm.' }, 'Close the file': { short: 'Fich.' } }, null);
  try {
    t('Close');
    assert.deepEqual(versionsOf('Fermer'), { short: 'Ferm.' });
    t('Close the file');
    assert.equal(versionsOf('Fermer'), null, 'a wrong short is worse than none');
    t('Close');
    assert.equal(versionsOf('Fermer'), null, 'and stays so');
  } finally {
    setVersions(null, null);
    setLanguage('en');
  }
});

test('two labels that read alike with the same versions keep them', () => {
  registerCatalogue('fr', { Close: 'Fermer', 'Close the window': 'Fermer' });
  setLanguage('fr');
  setVersions({ Close: { short: 'Ferm.' }, 'Close the window': { short: 'Ferm.' } }, null);
  try {
    t('Close');
    t('Close the window');
    assert.deepEqual(versionsOf('Fermer'), { short: 'Ferm.' });
  } finally {
    setVersions(null, null);
    setLanguage('en');
  }
});

test('an English window has lang-en\'s versions, counted by English\'s rules', () => {
  setLanguage('en');
  setVersions(english, english);
  try {
    assert.equal(t('Saved {name}', { name: 'Notes' }), 'Saved Notes');
    assert.deepEqual(versionsOf('Saved Notes'), { short: 'Saved', complete: 'Saved Notes to this computer' });
    assert.equal(tn(1, '{count} file', '{count} files'), '1 file');
    assert.deepEqual(versionsOf('1 file'), { short: '1 f' });
    assert.equal(tn(4, '{count} file', '{count} files'), '4 files');
    assert.deepEqual(versionsOf('4 files'), { short: '4 f' });
  } finally {
    setVersions(null, null);
  }
});

test('with no versions nothing is remembered, and the list of texts stays bounded', () => {
  setLanguage('en');
  setVersions(null, null);
  t('Saved {name}', { name: 'x' });
  assert.equal(versionsOf('Saved x'), null);
  setVersions(english, english);
  try {
    for (let i = 0; i < 5000; i++) t('Saved {name}', { name: `n${i}` });
    assert.equal(versionsOf('Saved n0'), null, 'the oldest have gone');
    assert.deepEqual(versionsOf('Saved n4999'), { short: 'Saved', complete: 'Saved n4999 to this computer' });
  } finally {
    setVersions(null, null);
  }
});
