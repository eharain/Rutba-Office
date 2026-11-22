// "Start with Windows" does not outlive the copy it starts (the icon by the
// clock, D-2026-11-22-23). Office asks Windows to start it at sign-in with a
// value a profile under the person's Run key; a copy that is removed takes
// its own with it, and an update leaves them. The installer's own NSIS
// (rutbaForgetStartup, called from customUnInstall), built and run against
// stand-in keys under HKCU\Software\RutbaTest, so no real entry is touched.
// Skipped where there is no makensis.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { skip, run, write, build, macro, tempDir, regDrop, regGet, NSH, ROOT } from './installer-nsis-kit.js';

// The name Windows keeps the start under: Electron's default is the application's id.
const ID = 'co.techstyle.rutba.office';
const HERE = 'C:\\Programs\\Rutba Office';
const exe = (dir, args = '--tray') => `"${dir}\\Rutba Office.exe" ${args}`;

test('a copy removed takes its own "Start with Windows" entries with it, and leaves every other one', { skip }, () => {
  const id = `startup-${process.pid}`;
  const RUN = `Software\\RutbaTest\\${id}\\Run`;
  const APPROVED = `Software\\RutbaTest\\${id}\\Approved`;
  const set = (key, name, value) => `  WriteRegStr HKCU "${key}" "${name}" '${value}'`;
  const entries = [
    ['First', '"C:\\Elsewhere\\first.exe"', 'stays'],
    [ID, exe(HERE), 'goes'],
    [`${ID}.p2`, exe(HERE, '--tray --profile=p2'), 'goes'],
    // Another copy of Rutba Office, in another folder: its own to keep.
    [`${ID}.p3`, exe('D:\\Other\\Rutba Office', '--tray --profile=p3'), 'stays'],
    // A name that only begins alike, and another program in this folder: not ours.
    [`${ID}r`, exe(HERE), 'stays'],
    ['Helper', exe(HERE, '--helper'), 'stays'],
    // Windows compares neither names nor folders by their case.
    [`${ID.toUpperCase()}.P4`, exe(HERE.toLowerCase(), '--tray --profile=p4'), 'goes'],
    ['Last', '"C:\\Elsewhere\\last.exe"', 'stays'],
  ];
  regDrop(id);
  try {
    const ran = run(['rutbaForgetStartup'], [], [
      ...entries.map(([name, value]) => set(RUN, name, value)),
      ...entries.map(([name]) => set(APPROVED, name, 'allowed')),
      `  !insertmacro rutbaForgetStartup "${RUN}" "${APPROVED}" "${ID}" "${HERE}"`,
      // Run again: nothing of ours is left, and nothing else is taken.
      `  !insertmacro rutbaForgetStartup "${RUN}" "${APPROVED}" "${ID}" "${HERE}"`,
      '  StrCpy $R0 "kept"',
      write('registers', '$R0'),
    ].join('\r\n'));
    assert.equal(ran.status, 0);
    assert.deepEqual(ran.lines, ['registers=[kept]']);
    for (const [name, value, what] of entries) {
      const run_ = regGet(`${id}\\Run`, name);
      const approved = regGet(`${id}\\Approved`, name);
      if (what === 'goes') assert.deepEqual([run_, approved], [null, null], `${name} goes, with Windows' note of it`);
      else assert.deepEqual([run_, approved], [value, 'allowed'], `${name} stays`);
    }
  } finally {
    regDrop(id);
  }
});

test('a Run key that is not there, or holds nothing of ours, is left as it is', { skip }, () => {
  const id = `startup-none-${process.pid}`;
  const RUN = `Software\\RutbaTest\\${id}\\Run`;
  regDrop(id);
  try {
    const none = run(['rutbaForgetStartup'], [], [`  !insertmacro rutbaForgetStartup "${RUN}" "${RUN}-approved" "${ID}" "${HERE}"`, write('done', 'yes')].join('\r\n'));
    assert.deepEqual([none.status, none.lines], [0, ['done=[yes]']]);
    const other = run(['rutbaForgetStartup'], [], [`  WriteRegStr HKCU "${RUN}" "Other" 'other.exe'`, `  !insertmacro rutbaForgetStartup "${RUN}" "${RUN}-approved" "${ID}" "${HERE}"`, write('done', 'yes')].join('\r\n'));
    assert.deepEqual([other.status, other.lines], [0, ['done=[yes]']]);
    assert.equal(regGet(`${id}\\Run`, 'Other'), 'other.exe');
  } finally {
    regDrop(id);
  }
});

test('it builds inside an uninstaller, where customUnInstall puts it', { skip }, () => {
  const dir = tempDir('rutba-startup-');
  try {
    // Built only: an uninstaller's section takes the same tools under their own names there.
    const built = build(dir, 'unprobe', [
      '!include LogicLib.nsh', '!include WordFunc.nsh', '!include FileFunc.nsh', '!include x64.nsh',
      macro('rutbaForgetStartup'),
      'Section', '  WriteUninstaller "$EXEDIR\\un.exe"', 'SectionEnd',
      'Section "un.forget"', `  !insertmacro rutbaForgetStartup "Software\\RutbaTest\\never\\Run" "Software\\RutbaTest\\never\\Approved" "${ID}" "$INSTDIR"`, 'SectionEnd',
    ]);
    assert.ok(fs.existsSync(built));
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('the uninstaller forgets them only on a real removal, under the id Office gives Windows', () => {
  // An update runs the old copy's uninstaller too: the setting is the person's, and stays.
  const call = /\$\{IfNot\} \$\{isUpdated\}\r?\n\s*!insertmacro rutbaForgetStartup "([^"]+)" "([^"]+)" "([^"]+)" "\$INSTDIR"\r?\n\s*\$\{EndIf\}/.exec(NSH);
  assert.ok(call, 'customUnInstall calls it, and not on an update');
  assert.equal(call[1], 'Software\\Microsoft\\Windows\\CurrentVersion\\Run');
  assert.equal(call[2], 'Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\StartupApproved\\Run');
  assert.equal(call[3], ID);
  const uninstall = NSH.match(/!macro customUnInstall\b[\s\S]*?!macroend/)[0];
  assert.ok(uninstall.includes(call[0]), 'in the uninstaller itself');
  // The same id the application gives Windows, and the name a second profile's start is kept under.
  const shell = fs.readFileSync(path.join(ROOT, 'packages', 'office-shell', 'src', 'electron', 'main.js'), 'utf8');
  assert.ok(shell.includes(`const appUserModelId = '${ID}';`));
  const entry = fs.readFileSync(path.join(ROOT, 'apps', 'desktop', 'main', 'main.js'), 'utf8');
  assert.ok(entry.includes('startName: PROFILE?.isDefault === false ? `' + ID + '.${PROFILE.id}` : null'));
});
