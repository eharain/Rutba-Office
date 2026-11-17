// The installer's own NSIS, built and run (apps/desktop/resources/installer.nsh):
// a copy whose program cannot be removed, against the stand-ins in
// installer-nsis-kit.js. Skipped where there is no makensis.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawn, spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { ROOT, NSH, makensis, MAKENSIS, skip, macro, build, run, runAsync, write, EXE, APP, testHead, scopeVars, present, scopeMacros, removalMacros, tempDir, copyIn, regDrop, regGet, holdOpen, let_go, layouts, pageProbe } from './installer-nsis-kit.js';

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

