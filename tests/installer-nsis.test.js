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
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const NSH = fs.readFileSync(path.join(ROOT, 'apps', 'desktop', 'resources', 'installer.nsh'), 'utf8');

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

/** Builds `body` (inside a Section) with the macros named, runs it, and answers the lines it wrote. */
function run(macros, vars, body) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-nsis-'));
  try {
    const script = [
      'Unicode true', '!include LogicLib.nsh', '!include WordFunc.nsh',
      'OutFile "probe.exe"', 'SilentInstall silent', 'RequestExecutionLevel user',
      ...vars.map((v) => `Var ${v}`),
      ...macros.map(macro),
      'Section', '  FileOpen $9 "$EXEDIR\\out.txt" w', body, '  FileClose $9', 'SectionEnd', '',
    ].join('\r\n');
    fs.writeFileSync(path.join(dir, 'probe.nsi'), script);
    const built = spawnSync(MAKENSIS, ['-V2', 'probe.nsi'], { cwd: dir, encoding: 'utf8' });
    assert.equal(built.status, 0, built.stdout + built.stderr);
    const started = Date.now();
    const ran = spawnSync(path.join(dir, 'probe.exe'), [], { cwd: dir, timeout: 60_000 });
    assert.equal(ran.status, 0);
    return { lines: fs.readFileSync(path.join(dir, 'out.txt'), 'utf8').split(/\r?\n/).filter(Boolean), ms: Date.now() - started };
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
