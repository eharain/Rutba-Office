// What the second review of the installer, updates and packs found.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { elevatedScript } from '../apps/desktop/main/packs-service.js';
import { keepAnonymous } from '../apps/desktop/main/update-signature.js';
import { runsWhenOpened } from '../packages/office-shell/src/runs.js';

const require = createRequire(new URL('../apps/desktop/package.json', import.meta.url));

const win = process.platform === 'win32';

function parsed(script) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-ps-'));
  const file = path.join(dir, 'script.ps1');
  fs.writeFileSync(file, '\uFEFF' + script, 'utf8');
  const probe = `$e=$null;$t=$null;$a=[System.Management.Automation.Language.Parser]::ParseFile('${file.replace(/'/g, "''")}',[ref]$t,[ref]$e);`
    + `$s=$a.FindAll({param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.StringConstantType -eq 'SingleQuoted'},$true)|ForEach-Object{$_.Value};`
    + `[Console]::OutputEncoding=[Text.Encoding]::UTF8;@{errors=@($e).Count;strings=@($s);commands=@($a.FindAll({param($n) $n -is [System.Management.Automation.Language.CommandAst]},$true)|ForEach-Object{$_.GetCommandName()})}|ConvertTo-Json -Compress`;
  const r = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', probe], { encoding: 'utf8' });
  fs.rmSync(dir, { recursive: true, force: true });
  return JSON.parse(r.stdout);
}

for (const quote of ["'", '\u2018', '\u2019', '\u201A', '\u201B']) {
  test(`a folder with ${quote} in its name stays one string in the elevated script`, { skip: !win }, () => {
    const dir = `C:\\Users\\O${quote}Neil; Remove-Item C:\\ -Recurse #`;
    const exe = `${dir}\\Rutba Office.exe`;
    const out = parsed(elevatedScript({ pid: 1, dir, exe, removals: ['mail'] }));
    assert.equal(out.errors, 0);
    assert.ok(out.strings.some((s) => s === path.win32.join(dir, 'mail.pack.json')), 'the name comes through whole');
    assert.ok(out.strings.some((s) => s === `"${exe}"`));
    assert.equal(out.commands.filter((c) => c === 'Remove-Item').length, 3, 'nothing was added by the folder name');
  });
}

test('the elevated script starts Office through the shell, never itself', () => {
  const script = elevatedScript({ pid: 1, dir: 'C:\\p', exe: 'C:\\p\\Rutba Office.exe' });
  assert.doesNotMatch(script, /Start-Process -FilePath 'C:/);
  assert.match(script, /explorer\.exe/);
});

test('the real updater sends one word where it would send a lasting id', () => {
  const { NsisUpdater } = require('electron-updater');
  const app = {
    version: '1.0.0', name: 'x', isPackaged: true, appUpdateConfigPath: path.join(os.tmpdir(), 'none.yml'),
    userDataPath: os.tmpdir(), baseCachePath: os.tmpdir(), onQuit: () => {}, whenReady: async () => {}, quit: () => {}, relaunch: () => {},
  };
  const updater = new NsisUpdater(undefined, app);
  assert.equal(updater.computeFinalHeaders({ 'x-user-staging-id': 'a-lasting-id' })['x-user-staging-id'], 'a-lasting-id', 'what it does unless told');
  keepAnonymous(updater);
  assert.equal(updater.computeFinalHeaders({ 'x-user-staging-id': 'a-lasting-id' })['x-user-staging-id'], 'rutba-office');
});

test('keepAnonymous leaves headers already set', () => {
  const updater = { requestHeaders: { Accept: 'x' } };
  keepAnonymous(updater);
  assert.equal(updater.requestHeaders.Accept, 'x');
});

for (const ext of ['ade', 'adp', 'mda', 'mdb', 'mde', 'mdt', 'mdw', 'wsb', 'cer', 'bas', 'htc', 'ksh', 'msh1', 'msh2', 'ps2xml', 'psc2', 'pyc', 'pyz', 'vbp', 'vsmacros']) {
  test(`.${ext} warns before it opens`, () => assert.equal(runsWhenOpened(`a.${ext}`), true));
}

test('the upgrade proof runs when a release is published, which is when the new pair exists', () => {
  const yml = fs.readFileSync(new URL('../.github/workflows/upgrade.yml', import.meta.url), 'utf8');
  assert.match(yml, /^ {2}release:\r?\n {4}types: \[published\]/m);
  assert.doesNotMatch(yml, /workflow_run/);
});

test('removing an all-users copy waits for the elevated pass of a silent update, and keeps where it was', () => {
  const nsh = fs.readFileSync(new URL('../apps/desktop/resources/installer.nsh', import.meta.url), 'utf8').replace(/\r\n/g, '\n');
  const init = nsh.slice(nsh.indexOf('!macro customInit'), nsh.indexOf('!macroend', nsh.indexOf('!macro customInit')));
  assert.ok(init.indexOf('${OrIf} ${UAC_IsAdmin}') < init.indexOf('rutbaRemoveInstalledCopy'));
  assert.equal((init.match(/Push \$R8/g) ?? []).length, 1);
  assert.equal((init.match(/Pop \$R8/g) ?? []).length, 1);
  const copy = nsh.slice(nsh.indexOf('!macro rutbaRemoveInstalledCopy'), nsh.indexOf('!macroend'));
  assert.ok(copy.indexOf('ReadRegStr $R8') < copy.indexOf("ExecWait '"));
  assert.ok(copy.indexOf('WriteRegStr') > copy.indexOf("ExecWait '"));
});
