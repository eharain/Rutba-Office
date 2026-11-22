// An install that put nothing on the computer does not end as a success (the
// owner, 22 Nov 2026: "The installer must end with an error when it installs
// nothing"). After 2.26.1 an update run without elevation, over what was left
// of a copy for all users, ended with 0 and no Office registered anywhere.
// The installer's own NSIS (rutbaInstallLanded, the last step of
// customInstall), built and run against the stand-ins in
// installer-nsis-kit.js. Skipped where there is no makensis.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { skip, macro, build, EXE, APP, testHead, tempDir, copyIn, regDrop, regGet } from './installer-nsis-kit.js';

const UNINSTALLER = 'Uninstall RutbaTestApp.exe';
const INCLUDES = ['!include LogicLib.nsh', '!include WordFunc.nsh', '!include FileFunc.nsh', '!include x64.nsh'];
// The probe says "landed" only when the step let the install go on to its end.
const SAY = ['  FileOpen $9 "$EXEDIR\\out.txt" w', '  FileWrite $9 "landed"', '  FileClose $9'];

/** The probe run as an install into `folder` (/D= is how NSIS is told the folder): how it ended, and whether it got to its end. */
function install(probe, folder) {
  const out = path.join(path.dirname(probe), 'out.txt');
  fs.rmSync(out, { force: true });
  const ran = spawnSync(probe, [`/D=${folder}`], { timeout: 90_000 });
  return { status: ran.status, landed: fs.existsSync(out) };
}
const regSet = (id, name, value) => assert.equal(spawnSync('reg', ['add', `HKCU\\Software\\RutbaTest\\${id}\\Uninstall`, '/v', name, '/t', 'REG_SZ', '/d', value, '/f'], { stdio: 'ignore' }).status, 0);
/** A whole copy in `dir`, with its uninstaller. */
const whole = (dir) => {
  copyIn(dir);
  fs.writeFileSync(path.join(dir, UNINSTALLER), 'stand-in');
  return dir;
};
/** A program that carries `version`, as the build stamps the real one. */
const stamped = (dir, name, version) => build(dir, name, [`VIProductVersion "${version}.0"`, `VIAddVersionKey "FileVersion" "${version}"`, `VIAddVersionKey "ProductName" "Rutba Test"`, 'Section', 'SectionEnd']);

test('an install whose program, uninstaller or entry is not there ends with an error, and a whole one ends well', { skip }, () => {
  const root = tempDir('rutba-landed-');
  const id = `landed-${process.pid}`;
  try {
    const probe = build(root, 'probe', [
      ...INCLUDES, ...testHead(id),
      macro('rutbaCopyWorks'), macro('rutbaInstallLanded'),
      'Section', '  !insertmacro rutbaInstallLanded', ...SAY, 'SectionEnd',
    ]);
    const entry = (version, folder) => {
      regDrop(id);
      if (version != null) regSet(id, 'DisplayVersion', version);
      if (folder != null) regSet(id, 'InstallLocation', folder);
    };
    const good = whole(path.join(root, 'good', APP));

    // The whole install: the program, its resources and uninstaller, and an entry that names this version here.
    entry('9.9.9', good);
    assert.deepEqual(install(probe, good), { status: 0, landed: true }, 'a whole install');

    // Nothing written at all, as into a folder this person may not write to: no folder, or an empty one.
    assert.deepEqual(install(probe, path.join(root, 'never', APP)), { status: 2, landed: false }, 'no folder');
    const empty = path.join(root, 'empty', APP);
    fs.mkdirSync(empty, { recursive: true });
    entry('9.9.9', empty);
    assert.deepEqual(install(probe, empty), { status: 2, landed: false }, 'an empty folder');

    // What an earlier removal left, and nothing added to it: the program alone.
    const husk = copyIn(path.join(root, 'husk', APP), { asar: false });
    entry('9.9.9', husk);
    assert.deepEqual(install(probe, husk), { status: 2, landed: false }, 'only the program');

    // The files without the uninstaller.
    const noUninstaller = copyIn(path.join(root, 'no-uninstaller', APP));
    entry('9.9.9', noUninstaller);
    assert.deepEqual(install(probe, noUninstaller), { status: 2, landed: false }, 'no uninstaller');

    // The files there, and Windows not told: no entry, the entry of the copy before, or one that names another folder.
    entry(null, null);
    assert.deepEqual(install(probe, good), { status: 2, landed: false }, 'no entry');
    entry('9.9.8', good);
    assert.deepEqual(install(probe, good), { status: 2, landed: false }, 'the entry of the version before');
    entry('9.9.9', path.join(root, 'elsewhere', APP));
    assert.deepEqual(install(probe, good), { status: 2, landed: false }, 'an entry for another folder');
    entry('9.9.9', null);
    assert.deepEqual(install(probe, good), { status: 2, landed: false }, 'an entry with no folder');
  } finally {
    regDrop(id);
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test('a program of another version left in the folder, as one a helper held open is left, ends the install with an error', { skip }, () => {
  const root = tempDir('rutba-landed-');
  const id = `landed-version-${process.pid}`;
  try {
    const probe = build(root, 'probe', [
      ...INCLUDES, ...testHead(id),
      macro('rutbaCopyWorks'), macro('rutbaInstallLanded'),
      'Section', '  !insertmacro rutbaInstallLanded', ...SAY, 'SectionEnd',
    ]);
    const dir = whole(path.join(root, 'copy', APP));
    regSet(id, 'DisplayVersion', '9.9.9');
    regSet(id, 'InstallLocation', dir);
    // Everything else as a whole install leaves it; only the program differs.
    fs.copyFileSync(stamped(root, 'older', '9.9.8'), path.join(dir, EXE));
    assert.deepEqual(install(probe, dir), { status: 2, landed: false }, 'the program of the version before');
    fs.copyFileSync(stamped(root, 'this', '9.9.9'), path.join(dir, EXE));
    assert.deepEqual(install(probe, dir), { status: 0, landed: true }, 'the program of this version');
  } finally {
    regDrop(id);
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test('the installer runs the check itself, after it has written its entry: customInstall as it is', { skip }, () => {
  const root = tempDir('rutba-landed-');
  const id = `landed-custom-${process.pid}`;
  try {
    // customInstall with its own steps: the entry it writes (rutbaEntryTrue) is the installer's; only the app
    // packs, which are another file's, are left out.
    const probe = build(root, 'probe', [
      ...INCLUDES, ...testHead(id),
      ...['installMode', 'rutbaAppsShown', 'rutbaAppsOff', 'rutbaAppsOn', 'rutbaAddOns', 'rutbaLangShown', 'rutbaLangChosen'].map((v) => `Var ${v}`),
      '!macro rutbaWritePacks', '!macroend',
      macro('rutbaCopyWorks'), macro('rutbaForgetOtherScope'), macro('rutbaEntryTrue'), macro('rutbaInstallLanded'), macro('customInstall'),
      'Section', '  StrCpy $installMode "CurrentUser"', '  !insertmacro customInstall', ...SAY, 'SectionEnd',
    ]);
    const good = whole(path.join(root, 'good', APP));
    assert.deepEqual(install(probe, good), { status: 0, landed: true }, 'a whole install');
    assert.equal(regGet(`${id}\\Uninstall`, 'DisplayVersion'), '9.9.9');
    assert.equal(regGet(`${id}\\Uninstall`, 'InstallLocation'), good);

    // The 2.26.1 run: the entry written (or not), and no files beside the program left from before.
    const husk = copyIn(path.join(root, 'husk', APP), { asar: false });
    assert.deepEqual(install(probe, husk), { status: 2, landed: false }, 'only the program');
    assert.deepEqual(install(probe, path.join(root, 'never', APP)), { status: 2, landed: false }, 'nothing written');
  } finally {
    regDrop(id);
    fs.rmSync(root, { recursive: true, force: true });
  }
});
