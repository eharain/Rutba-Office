// The gate by app: a changed file names the apps whose window checks it can
// break, and what every app shares names them all.
import test from 'node:test';
import assert from 'node:assert/strict';
import { appsForPath, appsForChanges, VERIFY_APPS } from '../tools/gate.js';

test('gate by app: an app\'s own files, its checks and its service name that app', () => {
  assert.deepEqual(appsForPath('apps/desktop/renderer/apps/mail.js'), ['mail']);
  assert.deepEqual(appsForPath('apps/desktop/renderer/apps/mail/reader.js'), ['mail']);
  assert.deepEqual(appsForPath('apps/desktop/renderer/apps/video/player.js'), ['video']);
  assert.deepEqual(appsForPath('apps/desktop/renderer/apps/recorder/program.js'), ['recorder']);
  assert.deepEqual(appsForPath('apps/desktop/main/verify-deck-zoom.js'), ['slides']);
  assert.deepEqual(appsForPath('apps/desktop/main/verify-sheet-pointing.js'), ['sheets']);
  assert.deepEqual(appsForPath('apps/desktop/main/mail-archive.js'), ['mail']);
  assert.deepEqual(appsForPath('apps/desktop/main/live.js'), ['recorder']);
});

test('gate by app: an engine names the apps built on it; notes and engine tests name none', () => {
  assert.deepEqual(appsForPath('packages/ooxml/src/docx.js'), ['word', 'sheets', 'slides']);
  assert.deepEqual(appsForPath('packages/mailbox/src/store.js'), ['mail']);
  assert.deepEqual(appsForPath('packages/studio/src/compositor.js'), ['video', 'recorder']);
  assert.deepEqual(appsForPath('docs/releases/v2.7.0.md'), []);
  assert.deepEqual(appsForPath('README.md'), []);
  assert.deepEqual(appsForPath('tests/gate-apps.test.js'), []);
});

test('gate by app: the shell, the shared packages and the main process name every app', () => {
  assert.deepEqual(appsForPath('apps/desktop/renderer/apps/home.js'), ['shell']);
  assert.deepEqual(appsForPath('packages/office-ui/src/index.js'), ['shell']);
  assert.deepEqual(appsForPath('apps/desktop/main/main.js'), ['shell']);
  assert.deepEqual(appsForChanges(['packages/office-ui/src/index.js', 'apps/desktop/renderer/apps/mail.js']), VERIFY_APPS);
  assert.deepEqual(appsForChanges(['apps/desktop/renderer/apps/mail/reader.js', 'packages/ooxml/src/x.js']), ['word', 'sheets', 'slides', 'mail']);
  assert.deepEqual(appsForChanges(['docs/a.md']), []);
});
