// Uninstalling and the Locker's safety copies (owner, 12 Nov 2026): the key
// and the copies go only if the person ticks a box, and the page says they stay
// otherwise. The box starts unticked, an update never ticks it, and only the
// safety folders of each profile are removed, never anything else.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { SAFETY_DIR } from '../apps/desktop/main/extras/locker/safety.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const nsh = fs.readFileSync(path.join(here, '..', 'apps', 'desktop', 'resources', 'installer.nsh'), 'utf8').replace(/\r\n/g, '\n');
const config = fs.readFileSync(path.join(here, '..', 'apps', 'desktop', 'electron-builder.config.cjs'), 'utf8');

const between = (text, start, end) => {
  const a = text.indexOf(start);
  assert.ok(a >= 0, `${start} is in installer.nsh`);
  const b = text.indexOf(end, a + start.length);
  assert.ok(b > a, `${end} follows ${start}`);
  return text.slice(a, b + end.length);
};

test('the uninstaller keeps the profile, and with it the copies, unless asked', () => {
  assert.match(config, /deleteAppDataOnUninstall:\s*false/);
  const top = /!ifdef BUILD_UNINSTALLER\n\s*!define MUI_COMPONENTSPAGE_TEXT_TOP "([^"]+)"\n!endif/.exec(nsh);
  assert.ok(top, 'the page that has the box says what stays');
  assert.match(top[1], /safety copies .* kept too, with their key, unless you tick the box/);
  assert.ok(![...top[1]].some((c) => c.codePointAt(0) > 0x7e), 'plain characters only in what a person reads: no long dash');
});

test('the electron-builder in use shows the page with the box, and the uninstall itself cannot be unticked', (t) => {
  // The page is not ours to add: electron-builder's template adds it when customUnInstallSection is defined.
  const templates = path.join(here, '..', 'node_modules', 'app-builder-lib', 'templates', 'nsis');
  if (!fs.existsSync(path.join(templates, 'installer.nsi'))) {
    t.skip('electron-builder is not installed here');
    return;
  }
  const nsi = fs.readFileSync(path.join(templates, 'installer.nsi'), 'utf8').replace(/\r\n/g, '\n');
  assert.match(nsi, /!ifdef BUILD_UNINSTALLER\n\s*!ifmacrodef customUnInstallSection\n(?:\s*!define [^\n]*\n)*\s*!insertmacro MUI_UNPAGE_COMPONENTS/, 'a components page, in the uninstaller only');
  const un = fs.readFileSync(path.join(templates, 'uninstaller.nsh'), 'utf8').replace(/\r\n/g, '\n');
  assert.match(un, /Section "un\.\$\{UNINSTALL_SECTION_NAME\}"\n\s*SectionIn RO/, 'the uninstall itself is read-only on that page');
  assert.ok(un.indexOf('!insertmacro customUnInstallSection') > un.indexOf('Section "un.${UNINSTALL_SECTION_NAME}"'), 'ours comes after it');
});

test('the box starts unticked, an update never removes anything, and only the safety folders go', () => {
  const section = between(nsh, '!macro customUnInstallSection', '!macroend');
  const box = /Section \/o "un\.([^"]+)" rutbaUnSafety/.exec(section);
  assert.ok(box, 'a section of its own, unticked (/o)');
  assert.match(box[1], /safety copies and their key/);
  assert.match(section, /\$\{IfNot\} \$\{isUpdated\}/, 'an update (silent, --updated) skips it');
  const removes = [...section.matchAll(/RMDir[^\n]*/g)].map((m) => m[0]);
  assert.deepEqual(removes, [], 'nothing removed directly in the section');
  assert.match(section, /rutbaRemoveSafetyIn "\$APPDATA\\Rutba Office"/, 'the first profile');
  assert.match(section, /Profiles\\\$R1/, 'and every other profile');
  const helper = between(nsh, '!macro rutbaRemoveSafetyIn DIR', '!macroend');
  const rm = [...helper.matchAll(/RMDir[^\n]*/g)].map((m) => m[0].trim());
  assert.deepEqual(rm, [`RMDir /r "\${DIR}\\locker\\${SAFETY_DIR}"`], 'the one folder removed is the profile\'s locker\\safety');
  assert.match(section, /Section "-un\.rutbaSafetyKept"[\s\S]*SectionIsSelected\} \$\{rutbaUnSafety\}[\s\S]*DetailPrint "The Locker's safety copies, if any, are kept/, 'unticked, the uninstaller says they were kept');
});
