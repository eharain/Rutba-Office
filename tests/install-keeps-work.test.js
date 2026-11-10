// An update or an install in the way of work: Cancel leaves the copy, Restart
// and update asks about unsaved work first. The installer script is read as
// text here; the built installer is run by the release checks.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { whatStopsRestart } from '../apps/desktop/main/update-error.js';

const nsh = fs.readFileSync(new URL('../apps/desktop/resources/installer.nsh', import.meta.url), 'utf8').replace(/\r\n/g, '\n');
const between = (from, to) => nsh.slice(nsh.indexOf(from), nsh.indexOf(to, nsh.indexOf(from)));

test('Restart and update waits for unsaved work, and names it', () => {
  assert.deepEqual(whatStopsRestart({ ready: true, windows: [{ dirty: true, name: 'Budget.xlsx' }, { dirty: false, name: 'Saved.docx' }] }).unsaved, ['Budget.xlsx']);
  assert.equal(whatStopsRestart({ ready: true, windows: [{ dirty: true, name: 'A' }] }).installed, false);
  assert.equal(whatStopsRestart({ ready: true, windows: [{ dirty: false }] }), null, 'nothing unsaved, nothing in the way');
  assert.equal(whatStopsRestart({ ready: false, windows: [] }).installed, false);
});

test('with pages, the old copy stays until the last page is passed', () => {
  const init = between('!macro customInit', '!macroend');
  assert.match(init, /\$\{If\} \$\{Silent\}\s+\$\{If\} \$hasPerMachineInstallation != "1"\s+\$\{OrIf\} \$\{UAC_IsAdmin\}\s+!insertmacro rutbaRemoveInstalledCopy HKCU currentuser\s+!insertmacro rutbaRemoveInstalledCopy HKLM allusers\s+\$\{EndIf\}/);
  const pages = between('!macro customPageAfterChangeDir', 'Var rutbaAppsShown');
  assert.ok(pages.indexOf('rutbaRemoveOldCopies') > pages.indexOf('rutbaAppsPageShow'), 'after the last page');
  assert.match(between('Function rutbaRemoveOldCopies', 'FunctionEnd'), /rutbaRemoveInstalledCopy HKCU[\s\S]*rutbaRemoveInstalledCopy HKLM[\s\S]*Abort/);
});

test('Cancel in the Retry box ends the install, from a page as well', () => {
  const removal = between('!macro rutbaRemoveInstalledCopy', '!macroend');
  assert.match(removal, /IDRETRY \+2\n\s+Quit\n/);
  assert.doesNotMatch(removal, /^\s*Abort\s*$/m, 'Abort on a page only moves to the next one');
});

test('a window that may be asking about saving is not ended while a person is there', () => {
  const removal = between('!macro rutbaRemoveInstalledCopy', '!macroend');
  const forced = removal.indexOf('/F /T');
  const guard = removal.indexOf('${IfNot} ${Silent}');
  assert.ok(guard > 0 && guard < forced, 'the question comes before the force');
  assert.match(removal, /\$\{If\} \$R6 != "0"\s+DetailPrint "Closing what is left/);
});

test('an update shows no pickers', () => {
  for (const fn of ['Function rutbaPacksPageShow', 'Function rutbaAppsPageShow']) {
    assert.match(between(fn, 'FunctionEnd'), /^Function \w+\n\s+\$\{If\} \$\{isUpdated\}\n\s+Abort/);
  }
});

test('the release build carries the Google client or stops', () => {
  const workflow = fs.readFileSync(new URL('../.github/workflows/release.yml', import.meta.url), 'utf8');
  assert.match(workflow, /RUTBA_GOOGLE_CLIENT: \$\{\{ secrets\.RUTBA_GOOGLE_CLIENT \}\}/);
  assert.match(workflow, /-z "\$RUTBA_GOOGLE_CLIENT"[\s\S]*exit 1/);
  assert.match(workflow, /RUTBA_GOOGLE_CLIENT_FILE="\$RUNNER_TEMP\/google-client\.json" npm run dist:win/);
});
