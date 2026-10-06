// What the apps are called, and what must not move when they are renamed.
//
// Rutba Word became Documents and Presentation became Presentations in
// 1.29.0: plurals, like Worksheets, Contacts and Pictures. A name is what a
// person reads; the key under it is what the operating system remembers —
// the taskbar identity a pin is filed under, the app that owns `.docx` in the
// installer, the window a double-clicked file opens. The names change here;
// the keys may not.

import test from 'node:test';
import assert from 'node:assert/strict';
import { APPS, APP_EXTENSIONS, fileAssociations } from '@rutba/office-formats/registry';

test('the apps carry their names: Documents and Presentations, plurals beside Worksheets', () => {
  const names = Object.fromEntries(Object.entries(APPS).map(([key, a]) => [key, a.short]));
  assert.deepEqual(names, {
    mail: 'Mail',
    word: 'Documents',
    sheets: 'Worksheets',
    slides: 'Presentations',
    pictures: 'Pictures',
    image: 'Image',
    video: 'Video',
    calendar: 'Calendar',
    contacts: 'Contacts',
  });
  for (const [key, app] of Object.entries(APPS)) {
    assert.equal(app.name, `Rutba ${app.short}`, `${key}'s full name is the brand and its short name`);
    assert.equal(app.key, key, `${key} keeps its key`);
  }
});

test('a rename leaves every key where people\'s pins and defaults expect it', () => {
  // The keys the taskbar identities (co.techstyle.rutba.office.<key>), the
  // tiles (resources/apps/<key>.ico) and the file associations are filed
  // under. Renaming one would orphan every pin and default already made.
  assert.ok(APPS.word && APPS.slides, 'the keys are still word and slides');
  const owner = Object.fromEntries(fileAssociations().map((a) => [a.ext, a.app]));
  assert.equal(owner.docx, 'word');
  assert.equal(owner.pptx, 'slides');
  assert.equal(owner.xlsx, 'sheets');
  for (const key of Object.keys(APP_EXTENSIONS)) assert.ok(APPS[key], `${key}, which owns extensions, is still an app key`);
});

test('a new file is offered in the singular', () => {
  // The launcher's New button reads its noun: "New document", not "New
  // documents", now that the app's name is a plural.
  assert.equal(APPS.word.noun, 'document');
  assert.equal(APPS.sheets.noun, 'workbook');
  assert.equal(APPS.slides.noun, 'presentation');
});
