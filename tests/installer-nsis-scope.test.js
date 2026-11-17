// The installer's own NSIS, built and run (apps/desktop/resources/installer.nsh):
// the choice of scope and the removal of the copy there was, against the
// stand-ins in installer-nsis-kit.js. Skipped where there is no makensis.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawn, spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { ROOT, NSH, makensis, MAKENSIS, skip, macro, build, run, runAsync, write, EXE, APP, testHead, scopeVars, present, scopeMacros, removalMacros, tempDir, copyIn, regDrop, regGet, holdOpen, let_go, layouts, pageProbe } from './installer-nsis-kit.js';

test('the install-mode page opens on the kind of copy there is, in its folder, whatever way it was installed', { skip }, () => {
  const root = tempDir('rutba-scope-');
  try {
    for (const [i, c] of layouts(root).entries()) {
      const { status, lines } = pageProbe(c, i, root, ['  !insertmacro customInstallMode']);
      assert.equal(status, 0);
      const [mode, folder] = c.want;
      assert.deepEqual(lines, [`mode=[${mode} ${folder ? path.join(root, folder) : ''}]`], c.name);
    }
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test('leaving the page keeps the folder found, as does /D= the folder it names', { skip }, () => {
  const root = tempDir('rutba-leave-');
  try {
    const custom = { name: 'a custom per-user folder, its install key lost', entry: copyIn(path.join(root, 'custom', APP)) };
    // The page's Leave sets the folder again from the install key, else the default folder: what .onInit wrote
    // (rutbaRecordFolders) is what keeps the custom folder there.
    const left = pageProbe(custom, 0, root, ['  !insertmacro rutbaRecordFolders', '  !insertmacro customInstallMode', '  !insertmacro setInstallModePerUser']);
    assert.equal(left.status, 0);
    assert.deepEqual(left.lines, [`mode=[CurrentUser ${custom.entry}]`], 'the custom folder after Leave');
    // NSIS takes /D= off the command line it gives the installer: the folder it names is still the one used.
    const elsewhere = path.join(root, 'elsewhere', APP);
    const told = pageProbe(custom, 1, root, ['  !insertmacro customInstallMode'], [`/D=${elsewhere}`]);
    assert.equal(told.status, 0);
    assert.deepEqual(told.lines, [`mode=[CurrentUser ${elsewhere}]`], '/D= over the folder found');
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test('a silent update keeps the kind of copy there is, elevating only for all users and never only for a husk; /allusers and /currentuser still decide', { skip }, () => {
  const root = tempDir('rutba-scope-');
  try {
    const works = copyIn(path.join(root, 'works'));
    const husk = copyIn(path.join(root, 'husk'), { asar: false });
    const empty = path.join(root, 'empty');
    fs.mkdirSync(empty);
    // As initMultiUser leaves it: per-user with both registered (or with only the per-user one), the install
    // section elevating and installing for all users when $hasPerMachineInstallation is 1.
    const rule = (label, machine, user, hasMachine) => [
      `  StrCpy $perMachineInstallationFolder "${machine}"`,
      `  StrCpy $perUserInstallationFolder "${user}"`,
      '  StrCpy $installMode "CurrentUser"',
      `  StrCpy $hasPerMachineInstallation "${hasMachine}"`,
      '  !insertmacro rutbaSilentScope',
      write(label, '$hasPerMachineInstallation $installMode'),
    ].join('\r\n');
    const body = [
      rule('all works', works, empty, '1'),
      rule('user works', empty, works, '1'),
      rule('both', works, works, '1'),
      // The all-users registration lost, its husk in Program Files, and nothing else.
      rule('all-users husk, unregistered', '', '', '0'),
    ].join('\r\n');
    const macros = [...scopeMacros(), 'rutbaSilentScope'];
    const head = (opts) => testHead('silent', { machine: husk, user: empty, ...opts });
    const plain = run(macros, scopeVars, body, { head: head(), args: ['/S', '--updated'] });
    assert.equal(plain.status, 0);
    // Not elevated, only a husk left for all users: per-user, with no prompt; the husk is left, and said.
    assert.deepEqual(plain.lines, ['all works=[1 CurrentUser]', 'user works=[0 CurrentUser]', 'both=[1 CurrentUser]', 'all-users husk, unregistered=[0 CurrentUser]']);
    // Elevated already: all users, in its folder, set here for the removal to know where the install goes.
    const admin = run(macros, scopeVars, body, { head: head({ admin: true }), args: ['/S', '--updated'] });
    assert.deepEqual(admin.lines, ['all works=[1 all]', 'user works=[0 CurrentUser]', 'both=[1 all]', 'all-users husk, unregistered=[1 all]']);
    // The elevated inner instance is all users, always.
    const inner = run(macros, scopeVars, body, { head: head({ inner: true, admin: true }), args: ['/S', '--updated'] });
    assert.deepEqual(inner.lines, ['all works=[1 all]', 'user works=[1 all]', 'both=[1 all]', 'all-users husk, unregistered=[1 all]']);
    for (const given of ['/allusers', '/currentuser']) {
      const told = run(macros, scopeVars, body, { head: head(), args: ['/S', given] });
      assert.deepEqual(told.lines, ['all works=[1 CurrentUser]', 'user works=[1 CurrentUser]', 'both=[1 CurrentUser]', 'all-users husk, unregistered=[0 CurrentUser]'], `${given} is left to initMultiUser`);
    }
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test('a copy no entry names is removed, only our files and never the one installed over; a husk that cannot go is left and the install goes on, a whole copy stops it', { skip }, async () => {
  const root = tempDir('rutba-sweep-');
  const held = [];
  try {
    const sweep = (machine, user, { admin = false, instdir = path.join(root, 'here', APP) } = {}) => run(removalMacros(), scopeVars, [
      `  StrCpy $INSTDIR "${instdir}"`,
      '  !insertmacro rutbaSweepKind machine',
      '  !insertmacro rutbaSweepKind user',
      write('swept', 'yes'),
    ].join('\r\n'), { head: testHead(`sweep-${process.pid}`, { machine, user, admin }) });
    // A husk for all users that cannot be removed (held as Program Files holds it), and one for this person beside a file of someone else's.
    const m = copyIn(path.join(root, 'machine', APP), { asar: false });
    const u = copyIn(path.join(root, 'user', APP), { asar: false, more: ['keep-me.txt'] });
    held.push(await holdOpen(path.join(m, EXE)));
    const both = sweep(m, u);
    assert.equal(both.status, 0, 'a husk left does not stop the install');
    assert.deepEqual(both.lines, ['swept=[yes]']);
    assert.equal(fs.existsSync(path.join(u, EXE)), false, 'the per-user husk is gone');
    assert.equal(fs.existsSync(path.join(u, 'keep-me.txt')), true, 'with a file of someone else\'s left');
    assert.equal(fs.existsSync(path.join(m, EXE)), true, 'the held one stays');
    assert.equal(fs.existsSync(path.join(m, 'ffmpeg.dll')), true, 'and nothing else of it was touched');
    await let_go(held);
    assert.equal(sweep(m, path.join(root, 'none'), { admin: true }).status, 0);
    assert.equal(fs.existsSync(m), false, 'once it can be, it is removed, its folder with it');
    // A whole copy with only its install key (no entry): the kind replaced, removed rather than left beside the new one.
    const whole = copyIn(path.join(root, 'whole', APP));
    assert.equal(sweep(path.join(root, 'none'), whole).status, 0);
    assert.equal(fs.existsSync(whole), false, 'a whole copy no entry names is removed');
    // The folder this install goes to is installed over, not removed.
    const here = copyIn(path.join(root, 'here', APP));
    assert.equal(sweep(here, here, { admin: true, instdir: here }).status, 0);
    assert.equal(fs.existsSync(path.join(here, 'resources', 'app.asar')), true, 'the copy installed over is left');
    // A whole copy that cannot be removed: the install stops, and the copy is as it was.
    const stuck = copyIn(path.join(root, 'stuck', APP));
    held.push(await holdOpen(path.join(stuck, EXE)));
    const stopped = sweep(stuck, path.join(root, 'none'));
    assert.equal(stopped.status, 2, 'it stops rather than leave two');
    assert.equal(fs.existsSync(path.join(stuck, 'resources', 'app.asar')), true, 'with the copy whole');
    assert.equal(fs.existsSync(path.join(stuck, 'ffmpeg.dll')), true);
  } finally {
    await let_go(held);
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test('an entry whose uninstaller is gone is forgotten only after what is left of its copy is removed, and only our files', { skip }, () => {
  const id = `gone-${process.pid}`;
  const root = tempDir('rutba-gone-');
  try {
    // A stale per-user entry with InstallLocation empty: the folder is found from UninstallString.
    const husk = path.join(root, 'husk', APP);
    fs.mkdirSync(path.join(husk, 'locales'), { recursive: true });
    fs.writeFileSync(path.join(husk, EXE), 'stand-in');
    fs.writeFileSync(path.join(husk, 'locales', 'en-US.pak'), 'x');
    // A folder holding our program and a file of someone else's: the program goes, their file and the folder stay.
    const shared = copyIn(path.join(root, 'shared'), { asar: false, more: ['keep-me.txt'] });
    // A folder that is not our install (no program of ours): nothing in it is touched.
    const strange = path.join(root, 'strange');
    fs.mkdirSync(path.join(strange, 'resources'), { recursive: true });
    fs.writeFileSync(path.join(strange, 'resources', 'app.asar'), 'theirs');
    fs.writeFileSync(path.join(strange, 'ffmpeg.dll'), 'theirs');
    // An entry naming a folder of anyone's (D:\Tools), our program in it: resources and locales are not removed whole.
    const tools = copyIn(path.join(root, 'Tools'));
    fs.writeFileSync(path.join(tools, 'resources', 'theirs.txt'), 'theirs');
    fs.writeFileSync(path.join(tools, 'locales', 'theirs.dat'), 'theirs');
    const entry = (dir, location) => [
      `  WriteRegStr HKCU "\${UNINSTALL_REGISTRY_KEY}" "UninstallString" '"${dir}\\Uninstall RutbaTestApp.exe" /currentuser'`,
      `  WriteRegStr HKCU "\${UNINSTALL_REGISTRY_KEY}" "InstallLocation" "${location}"`,
      `  WriteRegStr HKCU "\${INSTALL_REGISTRY_KEY}" "InstallLocation" "${dir}"`,
      '  !insertmacro rutbaRemoveInstalledCopy HKCU currentuser other',
      '  ReadRegStr $0 HKCU "${UNINSTALL_REGISTRY_KEY}" "UninstallString"',
      '  ReadRegStr $1 HKCU "${INSTALL_REGISTRY_KEY}" "InstallLocation"',
      write(path.basename(path.dirname(dir)) + '/' + path.basename(dir), 'entry $0 install $1'),
    ].join('\r\n');
    const { status, lines } = run(removalMacros(), scopeVars, [entry(husk, ''), entry(shared, shared), entry(strange, ''), entry(tools, tools)].join('\r\n'), { head: testHead(id) });
    assert.equal(status, 0, `the probe stopped (${status}): ${lines.join(' | ')}`);
    assert.equal(fs.existsSync(husk), false, 'the husk and its folder are gone');
    assert.equal(fs.existsSync(path.join(shared, EXE)), false, 'our program is gone from the shared folder');
    assert.equal(fs.existsSync(path.join(shared, 'ffmpeg.dll')), false, 'and our files with it');
    assert.equal(fs.existsSync(path.join(shared, 'keep-me.txt')), true, 'a file of someone else\'s stays');
    assert.equal(fs.existsSync(path.join(strange, 'resources', 'app.asar')), true, 'a folder without our program is not touched');
    assert.equal(fs.existsSync(path.join(strange, 'ffmpeg.dll')), true);
    assert.equal(fs.existsSync(path.join(tools, EXE)), false, 'our program is gone from a folder not named as ours');
    assert.equal(fs.existsSync(path.join(tools, 'resources', 'app.asar')), false, 'with what is ours in its resources');
    assert.equal(fs.existsSync(path.join(tools, 'resources', 'theirs.txt')), true, 'but its resources folder is not removed whole');
    assert.equal(fs.existsSync(path.join(tools, 'locales', 'theirs.dat')), true, 'nor its locales');
    // Every entry forgotten, and where the scope said it was installed with it: no working copy is left there.
    assert.deepEqual(lines, ['husk/RutbaTest=[entry  install ]', `${path.basename(root)}/shared=[entry  install ]`, `${path.basename(root)}/strange=[entry  install ]`, `${path.basename(root)}/Tools=[entry  install ]`]);
  } finally {
    regDrop(id);
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test('after the install, the entry says the version installed and names its folder, and the other scope keeps none', { skip }, () => {
  const id = `entry-${process.pid}`;
  const root = tempDir('rutba-entry-');
  try {
    const { status, lines } = run(present(['rutbaForgetOtherScope', 'rutbaEntryTrue']), scopeVars, [
      `  StrCpy $INSTDIR "${root}"`,
      '  StrCpy $installMode "CurrentUser"',
      // As a stale entry looks: an older version, no InstallLocation.
      '  WriteRegStr HKCU "${UNINSTALL_REGISTRY_KEY}" "DisplayVersion" "2.18.0"',
      '  !insertmacro rutbaEntryTrue',
      '  ReadRegStr $0 HKCU "${UNINSTALL_REGISTRY_KEY}" "DisplayVersion"',
      '  ReadRegStr $1 HKCU "${UNINSTALL_REGISTRY_KEY}" "InstallLocation"',
      write('entry', '$0 $1'),
      // The other scope's stale entry, forgotten (here in HKCU, as for an install for all users).
      '  WriteRegStr HKCU "${INSTALL_REGISTRY_KEY}" "InstallLocation" "C:\\gone"',
      '  !insertmacro rutbaForgetOtherScope HKCU',
      '  ReadRegStr $0 HKCU "${UNINSTALL_REGISTRY_KEY}" "DisplayVersion"',
      '  ReadRegStr $1 HKCU "${INSTALL_REGISTRY_KEY}" "InstallLocation"',
      write('other', '$0 $1'),
    ].join('\r\n'), { head: testHead(id) });
    assert.equal(status, 0);
    assert.deepEqual(lines, [`entry=[9.9.9 ${root}]`, 'other=[ ]']);
  } finally {
    regDrop(id);
    fs.rmSync(root, { recursive: true, force: true });
  }
});

