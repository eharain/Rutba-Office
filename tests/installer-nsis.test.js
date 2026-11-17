// The installer's own NSIS, built and run (apps/desktop/resources/installer.nsh):
// the Languages page's slots, read with WordFind from a list that starts and
// ends with a comma. And the update has no elevated step for add-ons: every
// copy keeps them in the person's own folder (D-2026-11-13-35).
//
// The macros are taken from installer.nsh as they are, built into a small
// silent program with the makensis electron-builder uses, and run. Skipped
// where there is no makensis (not Windows, or electron-builder never ran).
//
// The helpers are in installer-nsis-kit.js. The probes of the scope and the
// removal are in installer-nsis-scope.test.js, and the copy that cannot be
// removed in installer-nsis-left-whole.test.js: three files, side by side.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawn, spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { ROOT, NSH, makensis, MAKENSIS, skip, macro, build, run, runAsync, write, EXE, APP, testHead, scopeVars, present, scopeMacros, removalMacros, tempDir, copyIn, regDrop, regGet, holdOpen, let_go, layouts, pageProbe } from './installer-nsis-kit.js';

test('the Languages page fills its slots from the first pack found, and has none when none was', { skip }, () => {
  const slots = (list) => [
    `  StrCpy $rutbaLangFound "${list}"`,
    '  !insertmacro rutbaLangSlot 1', '  !insertmacro rutbaLangSlot 2', '  !insertmacro rutbaLangSlot 3',
    write(`${list} 1`, '$rutbaLang1'), write(`${list} 2`, '$rutbaLang2'), write(`${list} 3`, '$rutbaLang3'),
  ].join('\r\n');
  const { lines } = run(['rutbaLangSlot'], ['rutbaLangFound', 'rutbaLang1', 'rutbaLang2', 'rutbaLang3'], [slots(',lang-fr,lang-de,'), slots(','), slots(',lang-ur,')].join('\r\n'));
  assert.deepEqual(lines, [
    ',lang-fr,lang-de, 1=[lang-fr]', ',lang-fr,lang-de, 2=[lang-de]', ',lang-fr,lang-de, 3=[]',
    ', 1=[]', ', 2=[]', ', 3=[]',
    ',lang-ur, 1=[lang-ur]', ',lang-ur, 2=[]', ',lang-ur, 3=[]',
  ]);
});

test('the installer runs no step for add-ons, and the uninstaller asks only about the folder of the person removing it', () => {
  assert.doesNotMatch(NSH, /addons-apply-staged|addons-staging|ExecBounded|PROGRAMDATA/);
  assert.match(NSH, /ReadEnvStr \$R0 LOCALAPPDATA\s+StrCpy \$R0 "\$R0\\Rutba Office\\addons"/);
});

test('no macro is defined inside another, in any of the installer files (NSIS refuses it, and no installer builds)', () => {
  // bac18fb1 put the Languages page's slot macros inside customPageAfterChangeDir: the macro test above still
  // found them, makensis said "can't define a macro inside a macro!", and no Windows installer built (14 Nov 2026).
  // Read as text, so it holds with no makensis on the computer.
  for (const name of ['installer.nsh', 'installer-words.nsh', 'installer-languages.nsh']) {
    const lines = fs.readFileSync(path.join(ROOT, 'apps', 'desktop', 'resources', name), 'utf8').split(/\r?\n/);
    let depth = 0;
    lines.forEach((line, i) => {
      const code = line.replace(/^\s*[;#].*$/, '').trim();
      if (/^!macro\s/i.test(code)) {
        depth += 1;
        assert.equal(depth, 1, `${name}:${i + 1} defines a macro inside another`);
      } else if (/^!macroend\b/i.test(code)) {
        depth -= 1;
        assert.ok(depth >= 0, `${name}:${i + 1} ends a macro that was never begun`);
      }
    });
    assert.equal(depth, 0, `${name} leaves a macro open`);
  }
});

test('the foot of every installer and uninstaller page names the company, Tech Style Ltd, from the build', () => {
  // The owner, relayed by the PM: "in installer kindly add company name tech style".
  assert.match(NSH, /!macro customHeader\s+!ifdef COMPANY_NAME\s+BrandingText "\$\{PRODUCT_NAME\} \$\{VERSION\} - \$\{COMPANY_NAME\}"\s+!endif\s+!macroend/);
  assert.equal(JSON.parse(fs.readFileSync(path.join(ROOT, 'apps', 'desktop', 'package.json'), 'utf8')).author?.name, 'Tech Style Ltd');
});

test('the kind replaced is removed before the kind kept, in either order of kinds', () => {
  const body = NSH.match(/!macro rutbaRemoveCopies\b[\s\S]*?!macroend/)?.[0] || '';
  assert.ok(body, 'installer.nsh has rutbaRemoveCopies');
  const [all, user] = body.split('${Else}');
  const order = (text) => [...text.matchAll(/rutbaRemoveInstalledCopy (HK\w+) \w+ (other|own)/g)].map((m) => `${m[1]} ${m[2]}`);
  assert.deepEqual(order(all), ['HKCU other', 'HKLM own'], 'all users kept: the per-user copy first');
  assert.deepEqual(order(user), ['HKLM other', 'HKCU own'], 'per-user kept: the all-users copy first');
  assert.ok(all.indexOf('rutbaSweepKind user') < all.indexOf('HKLM allusers own'), 'and what is left of it, before the kind kept');
  assert.ok(user.indexOf('rutbaSweepKind machine') < user.indexOf('HKCU currentuser own'));
  // Neither the old copy's own removal nor the leftovers' ever raises an administrator prompt in a silent install.
  for (const m of ['rutbaClearLeftovers', 'rutbaRemoveInstalledCopy']) {
    const text = NSH.match(new RegExp(`!macro ${m}\\b[\\s\\S]*?!macroend`))?.[0] || '';
    for (const at of [...text.matchAll(/ExecShellWait "runas"/g)].map((x) => x.index)) {
      const before = text.slice(0, at);
      assert.ok(before.lastIndexOf('${Else}') > before.lastIndexOf('${If} ${Silent}') && before.lastIndexOf('${If} ${Silent}') >= 0, `${m}: a runas outside the branch a silent install never takes`);
    }
  }
});

test('an old uninstaller that finished but left its program held open: what runs from that folder is ended, nothing else, and the folder goes', { skip }, async () => {
  const id = `held-${process.pid}`;
  const root = tempDir('rutba-held-');
  const children = [];
  try {
    const tools = path.join(root, 'tools');
    fs.mkdirSync(tools);
    // A program that runs until it is ended: the copy's own, held open, and the same program in another folder.
    // Fifteen minutes: longer than any run of this test (168 s beside a full gate on 16 Nov, when two minutes ran out and the
    // other folder's program ended by itself); each is ended in the finally.
    const sleeper = build(tools, 'sleeper', ['Section', '  Sleep 900000', 'SectionEnd']);
    const copy = copyIn(path.join(root, 'copy', APP));
    const other = copyIn(path.join(root, 'other', APP));
    fs.copyFileSync(sleeper, path.join(copy, EXE));
    fs.copyFileSync(sleeper, path.join(other, EXE));
    // Its uninstaller, as an update runs it: removes all it can, its entry last, and says it is done.
    const uninstaller = build(tools, 'uninstaller', [
      'Section',
      `  RMDir /r "${copy}\\resources"`, `  RMDir /r "${copy}\\locales"`, `  Delete "${copy}\\ffmpeg.dll"`, `  Delete "${copy}\\${EXE}"`,
      `  DeleteRegKey HKCU "Software\\RutbaTest\\${id}\\Install"`, `  DeleteRegKey HKCU "Software\\RutbaTest\\${id}\\Uninstall"`,
      '  SetErrorLevel 0', 'SectionEnd',
    ]);
    fs.copyFileSync(uninstaller, path.join(copy, 'Uninstall RutbaTestApp.exe'));
    for (const dir of [copy, other]) {
      const child = spawn(path.join(dir, EXE), [], { stdio: 'ignore' });
      child.ended = new Promise((resolve) => child.on('exit', () => resolve(Date.now())));
      children.push(child);
    }
    await new Promise((r) => setTimeout(r, 1500));
    const { status, lines, started } = await runAsync(removalMacros(), scopeVars, [
      `  WriteRegStr HKCU "\${UNINSTALL_REGISTRY_KEY}" "UninstallString" '"${copy}\\Uninstall RutbaTestApp.exe" /currentuser'`,
      `  WriteRegStr HKCU "\${INSTALL_REGISTRY_KEY}" "InstallLocation" "${copy}"`,
      '  !insertmacro rutbaRemoveInstalledCopy HKCU currentuser own',
      write('carried on', 'yes'),
    ].join('\r\n'), { head: testHead(id) });
    assert.equal(status, 0, 'the install carries on rather than stopping with the program left');
    assert.deepEqual(lines, ['carried on=[yes]']);
    assert.equal(fs.existsSync(copy), false, 'no husk: the program and its folder are gone');
    const ended = await Promise.race([children[0].ended, new Promise((r) => setTimeout(() => r(false), 5000))]);
    assert.ok(ended, 'what ran from the copy was ended');
    assert.ok(ended - started >= 10000, `it was given its grace before it was ended, not ${ended - started} ms`);
    // The grace, as written (a time alone cannot tell it: PowerShell's own start takes seconds on some computers):
    // asked to close after five seconds, ended after ten (when no window is up), and only then.
    const free = NSH.match(/!macro rutbaFreeProgram\b[\s\S]*?!macroend/)?.[0] || '';
    assert.match(free, /\$\{If\} \$0 == 5\s+DetailPrint "Asking what runs from \$\{DIR\} to close"\s+!insertmacro rutbaProcessesIn "\$\{DIR\}" close/);
    assert.match(free, /\$\{If\} \$0 >= 10\s+(?:;[^\n]*\s+)?!insertmacro rutbaProcessesIn "\$\{DIR\}" windows \$R3\s+\$\{If\} \$R3 == "0"\s+\$\{OrIf\} \$R3 == "error"\s+\$\{ExitDo\}/);
    assert.ok(free.indexOf('" close ') < free.indexOf('" end '), 'asked to close before it is ended');
    assert.equal(children[1].exitCode, null, 'the same program in another folder is still running');
    assert.equal(fs.existsSync(path.join(other, EXE)), true);
    assert.equal(regGet(`${id}\\Uninstall`, 'UninstallString'), null, 'the entry is gone with the copy');
  } finally {
    await let_go(children);
    regDrop(id);
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test('a removal knows every file the build puts in the install folder, and no folder but resources and locales', () => {
  const unpacked = path.join(ROOT, 'apps', 'desktop', 'release', 'win-unpacked');
  const list = NSH.match(/!macro rutbaOurFiles\b[\s\S]*?!macroend/)?.[0] || '';
  const clear = NSH.match(/!macro rutbaClearLeftovers\b[\s\S]*?!macroend/)?.[0] || '';
  assert.match(list, /APP_EXECUTABLE_FILENAME/);
  assert.match(list, /UNINSTALL_FILENAME/);
  assert.match(clear, /RMDir \/r "\$\{DIR\}\\resources"/);
  assert.match(clear, /RMDir \/r "\$\{DIR\}\\locales"/);
  if (!fs.existsSync(unpacked)) return;
  for (const e of fs.readdirSync(unpacked, { withFileTypes: true })) {
    if (e.isDirectory()) assert.ok(['resources', 'locales'].includes(e.name), `the build has a folder ${e.name} a removal would leave`);
    else if (e.name !== 'Rutba Office.exe') assert.ok(list.includes(`"${e.name}"`), `rutbaOurFiles does not name ${e.name}, which the build puts in the install folder`);
  }
});
