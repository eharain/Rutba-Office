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
  // In .onInit only a silent install removes anything, and only through the one removal (rutbaRemoveCopiesNow).
  const init = between('!macro customInit', '!macroend');
  assert.match(init, /\$\{If\} \$\{Silent\}\s+\$\{If\} \$hasPerMachineInstallation != "1"\s+\$\{OrIf\} \$\{rutbaIsAdmin\}\s+StrCpy \$rutbaKept "user"\s+\$\{If\} \$hasPerMachineInstallation == "1"\s+StrCpy \$rutbaKept "all"\s+\$\{EndIf\}\s+Call rutbaRemoveCopiesNow\s+\$\{EndIf\}\s+\$\{EndIf\}/);
  assert.equal((init.match(/rutbaRemoveCopiesNow|rutbaRemoveCopies\b|rutbaRemoveInstalledCopy|rutbaClearLeftovers|rutbaSweep/g) ?? []).length, 1, 'nothing else in .onInit removes a copy');
  // What .onInit runs before it, and every macro those insert, only finds the copies and writes where they are.
  for (const name of ['rutbaRecordFolders', 'rutbaSilentScope', 'rutbaFindCopies', 'rutbaScopeFolder', 'rutbaDefaultFolder', 'rutbaWorkingScope', 'rutbaCopyWorks', 'rutbaKeepFolder']) {
    const body = between(`!macro ${name}`, '!macroend');
    assert.ok(body.length > 20, `installer.nsh has ${name}`);
    assert.doesNotMatch(body, /\b(Delete|RMDir|Rename|ExecWait|ExecShellWait|nsExec|DeleteRegKey|DeleteRegValue)\b|rutbaClearLeftovers|rutbaRemove|rutbaSweep|rutbaProcessesIn/, `${name} removes nothing`);
  }
  // The removal itself, both kinds in either order, is one function, called from .onInit as above and from the
  // moment after the last page, and from nowhere else.
  assert.match(between('Function rutbaRemoveCopiesNow', 'FunctionEnd'), /!insertmacro rutbaRemoveCopies\n/);
  assert.match(between('!macro rutbaRemoveCopies\n', '!macroend'), /rutbaRemoveInstalledCopy HKCU[\s\S]*rutbaRemoveInstalledCopy HKLM[\s\S]*\$\{Else\}[\s\S]*rutbaRemoveInstalledCopy HKLM[\s\S]*rutbaRemoveInstalledCopy HKCU/);
  assert.equal((nsh.match(/Call rutbaRemoveCopiesNow/g) ?? []).length, 2, 'called from .onInit and after the last page only');
  assert.equal((nsh.match(/!insertmacro rutbaRemoveCopies\n/g) ?? []).length, 1, 'and inserted nowhere else');
  const pages = between('!macro customPageAfterChangeDir', 'Var rutbaAppsShown');
  assert.ok(pages.indexOf('rutbaRemoveOldCopies') > pages.indexOf('rutbaAppsPageShow'), 'after the last page');
  assert.ok(pages.indexOf('rutbaRemoveOldCopies') > pages.indexOf('rutbaLanguagesPageShow'), 'after the Languages page too');
  assert.match(between('Function rutbaRemoveOldCopies', 'FunctionEnd'), /Call rutbaRemoveCopiesNow\s+Abort/);
});

test('Cancel in the Retry box ends the install, from a page as well', () => {
  const removal = between('!macro rutbaRemoveInstalledCopy', '!macroend');
  assert.match(removal, /IDRETRY \+2\n\s+Quit\n/);
  assert.doesNotMatch(removal, /^\s*Abort\s*$/m, 'Abort on a page only moves to the next one');
});

test('a window that may be asking about saving is not ended while a person is there', () => {
  // Only the end mode forces (taskkill /F), and only two places insert it: the removal's retry loop and
  // rutbaFreeProgram. Each is checked below.
  const processes = between('!macro rutbaProcessesIn', '!macroend');
  assert.equal((processes.match(/\/F /g) ?? []).length, 1, 'one forced kill in rutbaProcessesIn');
  assert.match(processes, /!else if "\$\{MODE\}" == "end"\n[^\n]*\/F \/T/, 'and it is the end mode');
  assert.equal((nsh.match(/rutbaProcessesIn "[^"]*" end /g) ?? []).length, 2, 'the end mode is used in two places only');

  // The retry loop: with a person here, what is left is ended only when no window from that path is up
  // (or the check could not run); otherwise the Retry box asks them.
  const removal = between('!macro rutbaRemoveInstalledCopy', '!macroend');
  assert.equal((removal.match(/rutbaProcessesIn "\$R4" end /g) ?? []).length, 1);
  assert.match(removal, /\$\{ElseIf\} \$R7 == 10\n\s+StrCpy \$R6 "0"\n\s+\$\{IfNot\} \$\{Silent\}\n\s+!insertmacro rutbaProcessesIn "\$R4" windows \$R6\n\s+\$\{EndIf\}\n\s+\$\{If\} \$R6 == "0"\n\s+\$\{OrIf\} \$R6 == "error"\n\s+DetailPrint "Closing what is left of Rutba Office"\n\s+!insertmacro rutbaProcessesIn "\$R4" end \$R6\n\s+\$\{EndIf\}\n\s+\$\{EndIf\}\n\s+\$\{If\} \$R7 >= 15\n\s+MessageBox MB_RETRYCANCEL\|MB_ICONEXCLAMATION "[^"\n]*" \/SD IDCANCEL IDRETRY \+2\n/);

  // rutbaFreeProgram, as strictly: after the grace, a window up is never ended, with a person here or not.
  // The grace loop leaves only by these four ways: nothing runs; no window is up (or the check could not run);
  // silent, with a window up (left, and logged); Cancel (left). Retry starts the grace again.
  const free = between('!macro rutbaFreeProgram', '!macroend');
  const grace = between('!macro rutbaFreeProgram', '    ${Loop}');
  assert.equal((grace.match(/\$\{ExitDo\}/g) ?? []).length, 4, 'four ways out of the grace');
  assert.equal((grace.match(/\$0 >= 10/g) ?? []).length, 1, 'the grace ends in one place');
  const after = grace.match(/IntOp \$0 \$0 \+ 1\n\s+!insertmacro rutbaProcessesIn "\$\{DIR\}" count \$R3\n\s+\$\{If\} \$R3 == "0"\n\s+\$\{ExitDo\}\n\s+\$\{EndIf\}\n\s+\$\{If\} \$0 >= 10\n(?:\s+;[^\n]*\n)*\s+!insertmacro rutbaProcessesIn "\$\{DIR\}" windows \$R3\n\s+\$\{If\} \$R3 == "0"\n\s+\$\{OrIf\} \$R3 == "error"\n\s+\$\{ExitDo\}\n\s+\$\{EndIf\}\n\s+\$\{If\} \$\{Silent\}\n\s+DetailPrint "[^"\n]*"\n\s+StrCpy \$R3 "window"\n\s+\$\{ExitDo\}\n\s+\$\{EndIf\}\n([\s\S]*?)\$\{If\} \$\{Cmd\} `MessageBox MB_RETRYCANCEL\|MB_ICONEXCLAMATION "[^"\n]*" \/SD IDCANCEL IDRETRY`\n\s+StrCpy \$0 0\n\s+\$\{Else\}\n\s+StrCpy \$rutbaCancelled "1"\n\s+StrCpy \$R3 "window"\n\s+\$\{ExitDo\}\n\s+\$\{EndIf\}\n\s+\$\{EndIf\}\n$/);
  assert.ok(after, 'after the grace: the window check, then silent leaves it, else the Retry box');
  assert.doesNotMatch(after[1], /ExitDo|rutbaProcessesIn|Delete|Rename|"window"/, 'nothing between the silent case and the box ends or leaves');
  // Left means left: neither the end-and-delete loop nor the move runs, and the only end is inside that guard.
  assert.equal((free.match(/rutbaProcessesIn "\$\{DIR\}" end /g) ?? []).length, 1);
  assert.match(free, /\$\{Loop\}\n\s+Pop \$0\n\s+\$\{EndIf\}\n\s+\$\{If\} \$R3 != "window"\n\s+Push \$0\n\s+StrCpy \$0 0\n\s+\$\{Do\}\n\s+IntOp \$0 \$0 \+ 1\n\s+!insertmacro rutbaProcessesIn "\$\{DIR\}" end \$R3\n/);
  assert.match(free, /\$\{If\} \$R3 != "window"\n\s+\$\{AndIf\} \$\{FileExists\} "\$\{DIR\}\\\$\{APP_EXECUTABLE_FILENAME\}"\n\s+System::Call 'kernel32::GetTickCount\(\) i \.R3'\n\s+ClearErrors\n\s+Rename /);
  assert.equal((free.match(/\bRename\b|\bDelete\b/g) ?? []).length, 2, 'one delete, in the guarded loop, and one move');
  // What is left is then the caller's: rutbaClearLeftovers touches nothing else while the program is there.
  assert.match(between('!macro rutbaClearLeftovers', '!macroend'), /!insertmacro rutbaFreeProgram \$\{SCOPE\} "\$\{DIR\}"\n\s+\$\{EndIf\}\n\s+\$\{IfNot\} \$\{FileExists\} "\$\{DIR\}\\\$\{APP_EXECUTABLE_FILENAME\}"\n/);
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
