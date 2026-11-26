// What the program wrote for its Windows notifications does not outlive the
// copy it starts (B18). To have a notification's press reach it, the program
// writes a key for itself under the person's Software\Classes\CLSID: the
// key's own value is "Electron Notification Activator", and LocalServer32
// under it names the program file. A copy that is removed takes its own with
// it, and an update leaves them. The installer's own NSIS
// (rutbaForgetToastActivators, called from customUnInstall), built and run
// against stand-in keys under HKCU\Software\RutbaTest, so no real key is
// touched. Skipped where there is no makensis.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { skip, run, write, build, macro, tempDir, regDrop, NSH, ROOT } from './installer-nsis-kit.js';

const MARK = 'Electron Notification Activator';
const HERE = 'C:\\Programs\\Rutba Office\\Rutba Office.exe';

/** A stand-in key's own value: null where the key is not there, '' where it has none. */
const own = (key) => {
  const r = spawnSync('reg', ['query', `HKCU\\Software\\RutbaTest\\${key}`, '/ve'], { encoding: 'utf8' });
  if (r.status !== 0) return null;
  return /REG_SZ[ \t]+(.*)$/m.exec(r.stdout)?.[1]?.trim() ?? '';
};

test('a copy removed takes the keys its notifications were answered on, and leaves every other one', { skip }, () => {
  const id = `toasts-${process.pid}`;
  const CLASSES = `Software\\RutbaTest\\${id}\\CLSID`;
  // [the key, its own value, the program it names (null: none), what becomes of it]
  const keys = [
    ['{00000000-0000-4000-8000-000000000001}', 'Some other class', 'C:\\Elsewhere\\other.exe', 'stays'],
    // The id every installed copy uses from 2.28.0 on (app-identity.js).
    ['{74583220-6F39-4387-A371-FE35BACE5663}', MARK, HERE, 'goes'],
    // One drawn by chance by 2.27.0, and the one after it: two in a row both go.
    ['{B245C59E-8FB0-48A9-962B-5EFC8EE03F65}', MARK, HERE, 'goes'],
    ['{B245C59E-8FB0-48A9-962B-5EFC8EE03F66}', MARK, HERE, 'goes'],
    // Another copy of Rutba Office, in another folder, and another program made the same way: theirs to keep.
    ['{C0000000-0000-4000-8000-000000000001}', MARK, 'D:\\Other\\Rutba Office\\Rutba Office.exe', 'stays'],
    ['{C0000000-0000-4000-8000-000000000002}', MARK, 'C:\\Programs\\Another App\\Another App.exe', 'stays'],
    // A class of another make that happens to start this program: not written by it.
    ['{D0000000-0000-4000-8000-000000000001}', 'Rutba Office document', HERE, 'stays'],
    // The mark with no program named: nothing says it is ours.
    ['{D0000000-0000-4000-8000-000000000002}', MARK, null, 'stays'],
    // Windows compares no folder by its case.
    ['{E0000000-0000-4000-8000-000000000001}', MARK, HERE.toLowerCase(), 'goes'],
    ['{F0000000-0000-4000-8000-000000000001}', 'Last', 'C:\\Elsewhere\\last.exe', 'stays'],
  ];
  regDrop(id);
  try {
    const ran = run(['rutbaForgetToastActivators'], [], [
      ...keys.flatMap(([key, value, program]) => [
        `  WriteRegStr HKCU "${CLASSES}\\${key}" "" "${value}"`,
        `  WriteRegDWORD HKCU "${CLASSES}\\${key}" "CustomActivator" 1`,
        ...(program ? [`  WriteRegStr HKCU "${CLASSES}\\${key}\\LocalServer32" "" "${program}"`] : []),
      ]),
      '  StrCpy $R0 "kept"',
      '  StrCpy $R1 "kept"',
      '  StrCpy $R2 "kept"',
      `  !insertmacro rutbaForgetToastActivators "${CLASSES}" "${HERE}"`,
      // Run again: nothing of ours is left, and nothing else is taken.
      `  !insertmacro rutbaForgetToastActivators "${CLASSES}" "${HERE}"`,
      write('registers', '$R0 $R1 $R2'),
    ].join('\r\n'));
    assert.equal(ran.status, 0);
    assert.deepEqual(ran.lines, ['registers=[kept kept kept]']);
    for (const [key, value, program, what] of keys) {
      const now = [own(`${id}\\CLSID\\${key}`), own(`${id}\\CLSID\\${key}\\LocalServer32`)];
      if (what === 'goes') assert.deepEqual(now, [null, null], `${key} goes, with the program it names`);
      else assert.deepEqual(now, [value, program], `${key} stays`);
    }
  } finally {
    regDrop(id);
  }
});

test('a key that is not there, or holds nothing of ours, is left as it is', { skip }, () => {
  const id = `toasts-none-${process.pid}`;
  const CLASSES = `Software\\RutbaTest\\${id}\\CLSID`;
  regDrop(id);
  try {
    const none = run(['rutbaForgetToastActivators'], [], [`  !insertmacro rutbaForgetToastActivators "${CLASSES}" "${HERE}"`, write('done', 'yes')].join('\r\n'));
    assert.deepEqual([none.status, none.lines], [0, ['done=[yes]']]);
    assert.equal(own(`${id}\\CLSID`), null, 'and it is not made');
    const other = run(['rutbaForgetToastActivators'], [], [
      `  WriteRegStr HKCU "${CLASSES}\\{A0000000-0000-4000-8000-000000000001}" "" "Theirs"`,
      `  !insertmacro rutbaForgetToastActivators "${CLASSES}" "${HERE}"`,
      write('done', 'yes'),
    ].join('\r\n'));
    assert.deepEqual([other.status, other.lines], [0, ['done=[yes]']]);
    assert.equal(own(`${id}\\CLSID\\{A0000000-0000-4000-8000-000000000001}`), 'Theirs');
  } finally {
    regDrop(id);
  }
});

test('it builds inside an uninstaller, where customUnInstall puts it', { skip }, () => {
  const dir = tempDir('rutba-toasts-');
  try {
    // Built only: an uninstaller's section takes the same tools under their own names there.
    const built = build(dir, 'unprobe', [
      '!include LogicLib.nsh', '!include WordFunc.nsh', '!include FileFunc.nsh', '!include x64.nsh',
      macro('rutbaForgetToastActivators'),
      'Section', '  WriteUninstaller "$EXEDIR\\un.exe"', 'SectionEnd',
      'Section "un.forget"', '  !insertmacro rutbaForgetToastActivators "Software\\RutbaTest\\never\\CLSID" "$INSTDIR\\RutbaTestApp.exe"', 'SectionEnd',
    ]);
    assert.ok(fs.existsSync(built));
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('the uninstaller forgets them only on a real removal, where the program writes them, by the program it removes', () => {
  // An update runs the old copy's uninstaller too: the key is still the new copy's, and stays.
  const call = /\$\{IfNot\} \$\{isUpdated\}\r?\n(?:(?!\$\{EndIf\})[\s\S])*?!insertmacro rutbaForgetToastActivators "([^"]+)" "([^"]+)"\r?\n\s*\$\{EndIf\}/.exec(NSH);
  assert.ok(call, 'customUnInstall calls it, and not on an update');
  assert.equal(call[1], 'Software\\Classes\\CLSID');
  assert.equal(call[2], '$INSTDIR\\${APP_EXECUTABLE_FILENAME}');
  const uninstall = NSH.match(/!macro customUnInstall\b[\s\S]*?!macroend/)[0];
  assert.ok(uninstall.includes(call[0]), 'in the uninstaller itself');
  // Only the person's own keys are read or removed, and only by what they hold.
  const body = macro('rutbaForgetToastActivators');
  assert.ok(!/HKLM|HKCR|SHCTX/.test(body), 'never the computer\'s keys');
  assert.ok(body.includes(`== "${MARK}"`), 'by the mark the program writes');
  assert.ok(body.includes('== "${EXE}"'), 'and by the program named');
  // The id an installed copy asks for is one the uninstaller's search would find: it is written with the same mark.
  const identity = fs.readFileSync(path.join(ROOT, 'packages', 'office-shell', 'src', 'electron', 'app-identity.js'), 'utf8');
  assert.match(identity, /export const TOAST_ACTIVATOR = '\{[0-9A-F-]{36}\}';/);
});
