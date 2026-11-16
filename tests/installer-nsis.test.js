// The installer's own NSIS, built and run (apps/desktop/resources/installer.nsh):
// the Languages page's slots, read with WordFind from a list that starts and
// ends with a comma. And the update has no elevated step for add-ons: every
// copy keeps them in the person's own folder (D-2026-11-13-35).
//
// The macros are taken from installer.nsh as they are, built into a small
// silent program with the makensis electron-builder uses, and run. Skipped
// where there is no makensis (not Windows, or electron-builder never ran).

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawn, spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
// RUTBA_TEST_INSTALLER_NSH puts another installer.nsh to these tests (an older one, to see a test fail without its change).
const NSH = fs.readFileSync(process.env.RUTBA_TEST_INSTALLER_NSH || path.join(ROOT, 'apps', 'desktop', 'resources', 'installer.nsh'), 'utf8');

function makensis() {
  if (process.platform !== 'win32' || !process.env.LOCALAPPDATA) return null;
  const cache = path.join(process.env.LOCALAPPDATA, 'electron-builder', 'Cache');
  try {
    for (const d of fs.readdirSync(cache).filter((n) => n.startsWith('nsis-')).sort().reverse()) {
      for (const sub of fs.readdirSync(path.join(cache, d))) {
        const exe = path.join(cache, d, sub, 'makensis.exe');
        if (fs.existsSync(exe)) return exe;
      }
    }
  } catch {
    // no cache
  }
  return null;
}
const MAKENSIS = makensis();
const skip = !MAKENSIS && 'no makensis here (electron-builder has not run on this computer)';

const macro = (name) => {
  const m = NSH.match(new RegExp(`!macro ${name}\\b[\\s\\S]*?!macroend`));
  assert.ok(m, `installer.nsh has ${name}`);
  return m[0];
};

/** A small silent program built from `lines` (NSIS source) into `dir` as `name`. */
function build(dir, name, lines) {
  fs.writeFileSync(path.join(dir, `${name}.nsi`), ['Unicode true', `OutFile "${name}.exe"`, 'SilentInstall silent', 'RequestExecutionLevel user', ...lines, ''].join('\r\n'));
  const built = spawnSync(MAKENSIS, ['-V2', `${name}.nsi`], { cwd: dir, encoding: 'utf8' });
  assert.equal(built.status, 0, built.stdout + built.stderr);
  return path.join(dir, `${name}.exe`);
}

/**
 * Builds `body` (inside a Section) with the macros named, runs it, and answers the lines it wrote.
 * `head` is more source before the macros (stubs, defines); `args` the probe's command line.
 */
function run(macros, vars, body, { head = [], args = [] } = {}) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-nsis-'));
  try {
    const exe = build(dir, 'probe', [
      '!include LogicLib.nsh', '!include WordFunc.nsh', '!include FileFunc.nsh', '!include x64.nsh',
      ...head,
      ...vars.map((v) => `Var ${v}`),
      ...macros.map(macro),
      'Section', '  FileOpen $9 "$EXEDIR\\out.txt" w', body, '  FileClose $9', 'SectionEnd',
    ]);
    const started = Date.now();
    const ran = spawnSync(exe, args, { cwd: dir, timeout: 90_000 });
    const out = path.join(dir, 'out.txt');
    return { status: ran.status, lines: fs.existsSync(out) ? fs.readFileSync(out, 'utf8').split(/\r?\n/).filter(Boolean) : [], ms: Date.now() - started };
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

/** As run, but not blocking: what else runs meanwhile can be timed against it. */
async function runAsync(macros, vars, body, { head = [], args = [] } = {}) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-nsis-'));
  try {
    const exe = build(dir, 'probe', [
      '!include LogicLib.nsh', '!include WordFunc.nsh', '!include FileFunc.nsh', '!include x64.nsh',
      ...head,
      ...vars.map((v) => `Var ${v}`),
      ...macros.map(macro),
      'Section', '  FileOpen $9 "$EXEDIR\\out.txt" w', body, '  FileClose $9', 'SectionEnd',
    ]);
    const started = Date.now();
    const status = await new Promise((resolve) => spawn(exe, args, { cwd: dir, stdio: 'ignore' }).on('exit', (code) => resolve(code)));
    const out = path.join(dir, 'out.txt');
    return { status, started, lines: fs.existsSync(out) ? fs.readFileSync(out, 'utf8').split(/\r?\n/).filter(Boolean) : [] };
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

const write = (label, v) => `  FileWrite $9 "${label}=[${v}]$\\r$\\n"`;

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

// The removal and the choice of scope (the owner's PC, 2.26.0: an all-users
// copy replaced per-user, its program left alone in Program Files, and every
// pin made from it starting that husk). Run against stand-ins: a program
// named RutbaTestApp.exe in folders under the temporary folder, and the
// entries under HKCU\Software\RutbaTest\<id>, so no Rutba Office, no real
// entry and no administrator prompt is ever touched. A probe is never an
// administrator here, and never the elevated inner instance; it is silent,
// as an update is.
const EXE = 'RutbaTestApp.exe';
// The name an install folder ends in (the directory page puts it on the end of any folder chosen).
const APP = 'RutbaTest';
/**
 * The defines and stand-ins a probe of the removal and the scope is built with. `machine` and `user` stand in for
 * the default folders of each scope (Program Files and the person's Programs folder, which a test cannot make);
 * `admin` and `inner` say whether the probe plays an elevated installer, or the elevated inner instance.
 */
const testHead = (id, { machine = null, user = null, admin = false, inner = false } = {}) => [
  `!define APP_EXECUTABLE_FILENAME "${EXE}"`,
  `!define APP_FILENAME "${APP}"`,
  '!define UNINSTALL_FILENAME "Uninstall RutbaTestApp.exe"',
  '!define PRODUCT_NAME "Rutba Test"',
  '!define VERSION "9.9.9"',
  `!define UNINSTALL_REGISTRY_KEY "Software\\RutbaTest\\${id}\\Uninstall"`,
  `!define INSTALL_REGISTRY_KEY "Software\\RutbaTest\\${id}\\Install"`,
  ...(machine ? [`!define RUTBA_MACHINE_FOLDER "${machine}"`] : []),
  ...(user ? [`!define RUTBA_USER_FOLDER "${user}"`] : []),
  '!macro _rutbaTestNo _a _b _t _f',
  '  StrCmp "0" "1" `${_t}` `${_f}`',
  '!macroend',
  '!macro _rutbaTestYes _a _b _t _f',
  '  StrCmp "1" "1" `${_t}` `${_f}`',
  '!macroend',
  `!define UAC_IsAdmin \`"" ${admin ? 'rutbaTestYes' : 'rutbaTestNo'} ""\``,
  `!define UAC_IsInnerInstance \`"" ${inner ? 'rutbaTestYes' : 'rutbaTestNo'} ""\``,
  // multiUser.nsh's GetDParameter: /D= from the whole command line, which NSIS does not take it off.
  'Var rutbaTestD',
  '!macro GetDParameter outVar',
  '  Push $R5', '  Push $R6', '  Push $R7', '  Push $R8',
  "  System::Call 'kernel32::GetCommandLineW() t .R8'",
  '  StrCpy $R7 ""', '  StrCpy $R6 0',
  '  ${Do}',
  '    StrCpy $R5 $R8 3 $R6',
  '    ${If} $R5 == ""', '      ${ExitDo}', '    ${EndIf}',
  '    ${If} $R5 == "/D="', '      IntOp $R6 $R6 + 3', '      StrCpy $R7 $R8 "" $R6', '      ${ExitDo}', '    ${EndIf}',
  '    IntOp $R6 $R6 + 1',
  '  ${Loop}',
  '  StrCpy $rutbaTestD $R7',
  '  Pop $R8', '  Pop $R7', '  Pop $R6', '  Pop $R5',
  '  StrCpy ${outVar} $rutbaTestD',
  '!macroend',
  // What multiUser.nsh's two macros do, as the page's Leave and the inner instance run them: the mode, and the
  // folder from the scope's install key (read again here for this person; all users' is the variable, which a test
  // cannot write in HKLM), else the default folder; /D= over both.
  '!macro setInstallModePerAllUsers',
  '  StrCpy $installMode all',
  '  ${If} $perMachineInstallationFolder != ""',
  '    StrCpy $INSTDIR "$perMachineInstallationFolder"',
  '  ${Else}',
  `    StrCpy $INSTDIR "${machine || ''}"`,
  '  ${EndIf}',
  '  !insertmacro GetDParameter $R0',
  '  ${If} $R0 != ""', '    StrCpy $INSTDIR $R0', '  ${EndIf}',
  '!macroend',
  '!macro setInstallModePerUser',
  '  StrCpy $installMode CurrentUser',
  '  ReadRegStr $perUserInstallationFolder HKCU "${INSTALL_REGISTRY_KEY}" InstallLocation',
  '  ${If} $perUserInstallationFolder != ""',
  '    StrCpy $INSTDIR "$perUserInstallationFolder"',
  '  ${Else}',
  `    StrCpy $INSTDIR "${user || ''}"`,
  '  ${EndIf}',
  '  !insertmacro GetDParameter $R0',
  '  ${If} $R0 != ""', '    StrCpy $INSTDIR $R0', '  ${EndIf}',
  '!macroend',
];
const scopeVars = ['installMode', 'perMachineInstallationFolder', 'perUserInstallationFolder', 'hasPerMachineInstallation', 'rutbaMachineDir', 'rutbaUserDir', 'rutbaKept', 'rutbaOldVersion', 'rutbaOldPublisher', 'rutbaCancelled'];
// Macros a newer installer.nsh has and an older one may not: the probe is built with those it has, so a test
// runs against the installer.nsh before a change too, and fails on what it does, or on the name it lacks.
const present = (names) => names.filter((n) => new RegExp(`!macro ${n}\\b`).test(NSH));
const scopeMacros = () => present(['rutbaCopyWorks', 'rutbaDefaultFolder', 'rutbaScopeFolder', 'rutbaFindCopies', 'rutbaRecordFolders', 'rutbaWorkingScope', 'rutbaKeepFolder']);
const removalMacros = () => present(['rutbaCopyWorks', 'rutbaSamePath', 'rutbaDefaultFolder', 'rutbaScopeFolder', 'rutbaOurFile', 'rutbaOurFiles', 'rutbaOurFolder', 'rutbaProcessesIn', 'rutbaEndProcessesIn', 'rutbaFreeProgram', 'rutbaRegViewFlag', 'rutbaClearLeftovers', 'rutbaRemoveInstalledCopy', 'rutbaSweepOne', 'rutbaSweepKind', 'rutbaRemoveCopies']);

/** A folder under the temporary folder, by its long name (the path a process's image is reported by). */
const tempDir = (prefix) => fs.realpathSync.native(fs.mkdtempSync(path.join(os.tmpdir(), prefix)));
/** A copy in `dir`: the program, and with `asar` the archive beside it. */
function copyIn(dir, { asar = true, more = [] } = {}) {
  fs.mkdirSync(path.join(dir, 'resources'), { recursive: true });
  fs.mkdirSync(path.join(dir, 'locales'), { recursive: true });
  fs.writeFileSync(path.join(dir, EXE), 'stand-in');
  fs.writeFileSync(path.join(dir, 'locales', 'en-US.pak'), 'x');
  fs.writeFileSync(path.join(dir, 'ffmpeg.dll'), 'x');
  if (asar) fs.writeFileSync(path.join(dir, 'resources', 'app.asar'), 'x');
  for (const f of more) fs.writeFileSync(path.join(dir, f), 'theirs');
  return dir;
}
const regDrop = (id) => spawnSync('reg', ['delete', `HKCU\\Software\\RutbaTest\\${id}`, '/f'], { stdio: 'ignore' });
const regGet = (key, name) => {
  const r = spawnSync('reg', ['query', `HKCU\\Software\\RutbaTest\\${key}`, '/v', name], { encoding: 'utf8' });
  if (r.status !== 0) return null;
  return /REG_SZ\s+(.*)$/m.exec(r.stdout)?.[1]?.trim() ?? '';
};
/**
 * A process that keeps `file` open, sharing it for reading only, as Program Files keeps a program from a person
 * who is not an administrator: it can be neither deleted nor moved, and no process runs from it to end.
 */
function holdOpen(file) {
  const child = spawn('powershell', ['-NoProfile', '-NonInteractive', '-Command', "$f = [IO.File]::Open($env:RUTBA_TEST_HOLD, 'Open', 'Read', 'Read'); [Console]::Out.WriteLine('held'); [Console]::Out.Flush(); Start-Sleep 180"], { env: { ...process.env, RUTBA_TEST_HOLD: file }, stdio: ['ignore', 'pipe', 'ignore'] });
  return new Promise((resolve, reject) => {
    child.stdout.on('data', () => resolve(child));
    child.on('exit', () => reject(new Error(`could not hold ${file}`)));
  });
}
const let_go = async (children) => {
  for (const c of children) if (c.exitCode == null) c.kill();
  await new Promise((r) => setTimeout(r, 700));
};

// The owner's ruling, 16 Nov (D-2026-11-16-03): "the installer should detect these changes and variations".
// Each case: what is on the computer, and the kind and folder the install keeps. `machine` and `user` are each
// scope's InstallLocation as initMultiUser read it; `entry` an HKCU uninstall entry with no install key (the per-user
// copy's registration half lost, in a custom folder); `defaults` what each scope's default folder holds.
const layouts = (root) => {
  const works = (name) => copyIn(path.join(root, name));
  const husk = (name) => copyIn(path.join(root, name), { asar: false });
  const nothing = (name) => {
    fs.mkdirSync(path.join(root, name), { recursive: true });
    return path.join(root, name);
  };
  return [
    { name: 'per-user over per-user', user: works('u1'), want: ['CurrentUser', 'u1'] },
    { name: 'all users over all users', machine: works('m1'), want: ['all', 'm1'] },
    // The owner's case: 2.25.0 in Program Files, and a stale per-user registration 2.18.0 left (its folder gone or a husk).
    { name: 'per-user over all users (the owner\'s)', machine: works('m2'), user: nothing('u2'), want: ['all', 'm2'] },
    { name: 'per-user over all users, the per-user folder a husk', machine: works('m3'), user: husk('u3'), want: ['all', 'm3'] },
    { name: 'all users over per-user', machine: nothing('m4'), user: works('u4'), want: ['CurrentUser', 'u4'] },
    { name: 'a custom per-user folder, its install key lost', entry: works('custom-u5'), want: ['CurrentUser', 'custom-u5'] },
    { name: 'a custom all-users folder', machine: works('custom-m6'), want: ['all', 'custom-m6'] },
    // Both kinds whole at once: all users, which other people on this computer may use too.
    { name: 'both kinds at once', machine: works('m7'), user: works('u7'), want: ['all', 'm7'] },
    // A husk an earlier removal left, its entries gone: the kind it was, in its folder, which the install makes whole.
    { name: 'an all-users husk with no entry', defaults: { machine: husk('dm8') }, want: ['all', 'dm8'] },
    { name: 'a per-user husk with no entry', defaults: { user: husk('du9') }, want: ['CurrentUser', 'du9'] },
    // A working copy beside a husk of the other kind: the working one.
    { name: 'a per-user copy beside an all-users husk', user: works('u10'), defaults: { machine: husk('dm10') }, want: ['CurrentUser', 'u10'] },
    // Two husks: the kind Windows still lists.
    { name: 'two husks, the per-user one still registered', user: husk('u12'), defaults: { machine: husk('dm12') }, want: ['CurrentUser', 'u12'] },
    { name: 'nothing anywhere: electron-builder decides', defaults: { machine: nothing('dm11'), user: nothing('du11') }, want: ['untouched', ''] },
  ];
};

/** The page probe for a layout: as initMultiUser leaves it, then `steps`. */
function pageProbe(c, i, root, steps, args = []) {
  const id = `scope-${process.pid}-${i}`;
  try {
    const body = [
      `  StrCpy $perMachineInstallationFolder "${c.machine || ''}"`,
      `  StrCpy $perUserInstallationFolder "${c.user || ''}"`,
      ...(c.user ? [`  WriteRegStr HKCU "\${INSTALL_REGISTRY_KEY}" "InstallLocation" "${c.user}"`] : []),
      ...(c.entry ? [`  WriteRegStr HKCU "\${UNINSTALL_REGISTRY_KEY}" "UninstallString" '"${c.entry}\\Uninstall RutbaTestApp.exe" /currentuser'`] : []),
      '  StrCpy $installMode "untouched"',
      '  StrCpy $INSTDIR ""',
      ...steps,
      write('mode', '$installMode $INSTDIR'),
    ].join('\r\n');
    const empty = path.join(root, `none-${i}`);
    fs.mkdirSync(empty, { recursive: true });
    return run([...scopeMacros(), 'customInstallMode'], scopeVars, body, {
      head: testHead(id, { machine: c.defaults?.machine || empty, user: c.defaults?.user || empty }),
      args,
    });
  } finally {
    regDrop(id);
  }
}

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

test('a copy whose program cannot be removed is left whole, with its entry: the kind kept stops the install, a husk of the kind replaced does not', { skip }, async () => {
  const id = `stuck-${process.pid}`;
  const root = tempDir('rutba-stuck-');
  const held = [];
  try {
    const entry = (dir) => [
      `  WriteRegStr HKCU "\${UNINSTALL_REGISTRY_KEY}" "UninstallString" '"${dir}\\Uninstall RutbaTestApp.exe" /currentuser'`,
      `  WriteRegStr HKCU "\${UNINSTALL_REGISTRY_KEY}" "InstallLocation" "${dir}"`,
    ];
    // The kind kept, whole, its uninstaller gone and its program held: nothing of it is touched, and the install stops.
    const whole = copyIn(path.join(root, 'whole', APP));
    held.push(await holdOpen(path.join(whole, EXE)));
    const own = run(removalMacros(), scopeVars, [...entry(whole), '  !insertmacro rutbaRemoveInstalledCopy HKCU currentuser own', write('carried on', 'yes')].join('\r\n'), { head: testHead(id) });
    assert.equal(own.status, 2, 'the install stops');
    assert.equal(fs.existsSync(path.join(whole, 'resources', 'app.asar')), true, 'the copy is whole: its resources were not removed without its program');
    assert.equal(fs.existsSync(path.join(whole, 'ffmpeg.dll')), true);
    assert.equal(fs.existsSync(path.join(whole, 'locales', 'en-US.pak')), true);
    assert.notEqual(regGet(`${id}\\Uninstall`, 'UninstallString'), null, 'and Windows still lists it');
    regDrop(id);
    // The kind replaced first, when all that is left of it is a husk that cannot go: it is left, with its entry,
    // and the install goes on to the kind kept (here nothing), rather than stop with no copy that works.
    const husk = copyIn(path.join(root, 'husk', APP), { asar: false });
    held.push(await holdOpen(path.join(husk, EXE)));
    const other = run(removalMacros(), scopeVars, [...entry(husk), '  StrCpy $rutbaKept "all"', `  StrCpy $INSTDIR "${path.join(root, 'kept', APP)}"`, '  !insertmacro rutbaRemoveCopies', write('carried on', 'yes')].join('\r\n'), { head: testHead(id, { machine: path.join(root, 'none'), user: path.join(root, 'none') }) });
    assert.equal(other.status, 0, `the install goes on (${other.status})`);
    assert.deepEqual(other.lines, ['carried on=[yes]']);
    assert.equal(fs.existsSync(path.join(husk, 'ffmpeg.dll')), true, 'nothing of it was touched');
    assert.notEqual(regGet(`${id}\\Uninstall`, 'UninstallString'), null, 'its entry stays while its program does');
  } finally {
    await let_go(held);
    regDrop(id);
    fs.rmSync(root, { recursive: true, force: true });
  }
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
    const sleeper = build(tools, 'sleeper', ['Section', '  Sleep 120000', 'SectionEnd']);
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
