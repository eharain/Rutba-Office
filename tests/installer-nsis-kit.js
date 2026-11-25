// What the installer tests share (installer-nsis*.test.js): the macros taken
// from installer.nsh as they are, built into a small silent program with the
// makensis electron-builder uses, and the stand-ins the removal and the scope
// are probed against. Not a test file: the tests are in three files so they
// run side by side (the one file took 264 seconds in the 2.26.1 gate).
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawn, spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

export const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
// RUTBA_TEST_INSTALLER_NSH puts another installer.nsh to these tests (an older one, to see a test fail without its change).
export const NSH = fs.readFileSync(process.env.RUTBA_TEST_INSTALLER_NSH || path.join(ROOT, 'apps', 'desktop', 'resources', 'installer.nsh'), 'utf8');

export function makensis() {
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
export const MAKENSIS = makensis();
export const skip = !MAKENSIS && 'no makensis here (electron-builder has not run on this computer)';

export const macro = (name) => {
  const m = NSH.match(new RegExp(`!macro ${name}\\b[\\s\\S]*?!macroend`));
  assert.ok(m, `installer.nsh has ${name}`);
  return m[0];
};

/** A small silent program built from `lines` (NSIS source) into `dir` as `name`. */
export function build(dir, name, lines) {
  fs.writeFileSync(path.join(dir, `${name}.nsi`), ['Unicode true', `OutFile "${name}.exe"`, 'SilentInstall silent', 'RequestExecutionLevel user', ...lines, ''].join('\r\n'));
  const built = spawnSync(MAKENSIS, ['-V2', `${name}.nsi`], { cwd: dir, encoding: 'utf8' });
  assert.equal(built.status, 0, built.stdout + built.stderr);
  return path.join(dir, `${name}.exe`);
}

/**
 * Builds `body` (inside a Section) with the macros named, runs it, and answers the lines it wrote.
 * `head` is more source before the macros (stubs, defines); `args` the probe's command line.
 */
export function run(macros, vars, body, { head = [], args = [] } = {}) {
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
export async function runAsync(macros, vars, body, { head = [], args = [] } = {}) {
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

export const write = (label, v) => `  FileWrite $9 "${label}=[${v}]$\\r$\\n"`;

// The removal and the choice of scope (the owner's PC, 2.26.0: an all-users
// copy replaced per-user, its program left alone in Program Files, and every
// pin made from it starting that husk). Run against stand-ins: a program
// named RutbaTestApp.exe in folders under the temporary folder, and the
// entries under HKCU\Software\RutbaTest\<id>, so no Rutba Office, no real
// entry and no administrator prompt is ever touched. A probe is never an
// administrator here, and never the elevated inner instance; it is silent,
// as an update is.
export const EXE = 'RutbaTestApp.exe';
// The name an install folder ends in (the directory page puts it on the end of any folder chosen).
export const APP = 'RutbaTest';
/**
 * The defines and stand-ins a probe of the removal and the scope is built with. `machine` and `user` stand in for
 * the default folders of each scope (Program Files and the person's Programs folder, which a test cannot make);
 * `admin` and `inner` say whether the probe plays an elevated installer, or the elevated inner instance.
 */
export const testHead = (id, { machine = null, user = null, admin = false, inner = false } = {}) => [
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
  // installer.nsh's own test (the plugin and Windows both asked, rutbaAskAdmin): a probe plays the answer.
  `!define rutbaIsAdmin \`"" ${admin ? 'rutbaTestYes' : 'rutbaTestNo'} ""\``,
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
export const scopeVars = ['installMode', 'perMachineInstallationFolder', 'perUserInstallationFolder', 'hasPerMachineInstallation', 'rutbaMachineDir', 'rutbaUserDir', 'rutbaKept', 'rutbaOldVersion', 'rutbaOldPublisher', 'rutbaCancelled'];
// Macros a newer installer.nsh has and an older one may not: the probe is built with those it has, so a test
// runs against the installer.nsh before a change too, and fails on what it does, or on the name it lacks.
export const present = (names) => names.filter((n) => new RegExp(`!macro ${n}\\b`).test(NSH));
export const scopeMacros = () => present(['rutbaCopyWorks', 'rutbaDefaultFolder', 'rutbaScopeFolder', 'rutbaFindCopies', 'rutbaRecordFolders', 'rutbaWorkingScope', 'rutbaKeepFolder']);
export const removalMacros = () => present(['rutbaCopyWorks', 'rutbaSamePath', 'rutbaDefaultFolder', 'rutbaScopeFolder', 'rutbaOurFile', 'rutbaOurFiles', 'rutbaOurFolder', 'rutbaProcessesIn', 'rutbaEndProcessesIn', 'rutbaFreeProgram', 'rutbaRegViewFlag', 'rutbaClearLeftovers', 'rutbaRemoveInstalledCopy', 'rutbaSweepOne', 'rutbaSweepKind', 'rutbaRemoveCopies']);

/** A folder under the temporary folder, by its long name (the path a process's image is reported by). */
export const tempDir = (prefix) => fs.realpathSync.native(fs.mkdtempSync(path.join(os.tmpdir(), prefix)));
/** A copy in `dir`: the program, and with `asar` the archive beside it. */
export function copyIn(dir, { asar = true, more = [] } = {}) {
  fs.mkdirSync(path.join(dir, 'resources'), { recursive: true });
  fs.mkdirSync(path.join(dir, 'locales'), { recursive: true });
  fs.writeFileSync(path.join(dir, EXE), 'stand-in');
  fs.writeFileSync(path.join(dir, 'locales', 'en-US.pak'), 'x');
  fs.writeFileSync(path.join(dir, 'ffmpeg.dll'), 'x');
  if (asar) fs.writeFileSync(path.join(dir, 'resources', 'app.asar'), 'x');
  for (const f of more) fs.writeFileSync(path.join(dir, f), 'theirs');
  return dir;
}
export const regDrop = (id) => spawnSync('reg', ['delete', `HKCU\\Software\\RutbaTest\\${id}`, '/f'], { stdio: 'ignore' });
export const regGet = (key, name) => {
  const r = spawnSync('reg', ['query', `HKCU\\Software\\RutbaTest\\${key}`, '/v', name], { encoding: 'utf8' });
  if (r.status !== 0) return null;
  return /REG_SZ\s+(.*)$/m.exec(r.stdout)?.[1]?.trim() ?? '';
};
/**
 * A process that keeps `file` open, sharing it for reading only, as Program Files keeps a program from a person
 * who is not an administrator: it can be neither deleted nor moved, and no process runs from it to end.
 */
export function holdOpen(file) {
  const child = spawn('powershell', ['-NoProfile', '-NonInteractive', '-Command', "$f = [IO.File]::Open($env:RUTBA_TEST_HOLD, 'Open', 'Read', 'Read'); [Console]::Out.WriteLine('held'); [Console]::Out.Flush(); Start-Sleep 180"], { env: { ...process.env, RUTBA_TEST_HOLD: file }, stdio: ['ignore', 'pipe', 'ignore'] });
  return new Promise((resolve, reject) => {
    child.stdout.on('data', () => resolve(child));
    child.on('exit', () => reject(new Error(`could not hold ${file}`)));
  });
}
export const let_go = async (children) => {
  for (const c of children) if (c.exitCode == null) c.kill();
  await new Promise((r) => setTimeout(r, 700));
};

// The owner's ruling, 16 Nov (D-2026-11-16-03): "the installer should detect these changes and variations".
// Each case: what is on the computer, and the kind and folder the install keeps. `machine` and `user` are each
// scope's InstallLocation as initMultiUser read it; `entry` an HKCU uninstall entry with no install key (the per-user
// copy's registration half lost, in a custom folder); `defaults` what each scope's default folder holds.
export const layouts = (root) => {
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
export function pageProbe(c, i, root, steps, args = []) {
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
