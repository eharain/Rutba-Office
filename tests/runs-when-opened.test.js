// What runs when it is opened: the extension as Windows takes it, so a name
// that ends in a space, a dot or a stream's name is still a program.
import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { runsWhenOpened, openedExtension, markFromInternet, FROM_INTERNET } from '@rutba/office-shell/runs';
import { isListed } from '../packages/office-shell/src/electron/packs.js';
import { elevationCommand, elevatedLaunch } from '../apps/desktop/main/packs-service.js';

test('programs are known by the extension Windows opens them under', () => {
  for (const name of ['setup.exe', 'C:\\Temp\\run.EXE', 'run.exe ', 'run.exe.', 'run.exe::$DATA', 'report.pdf.js', '/tmp/x.sh', 'disk.iso', 'link.lnk']) {
    assert.equal(runsWhenOpened(name), true, name);
  }
  for (const name of ['letter.docx', 'photo.jpg', 'README', 'C:\\setup.exe\\notes.txt', 'archive.zip', '']) {
    assert.equal(runsWhenOpened(name), false, name);
  }
  assert.equal(openedExtension('run.exe::$DATA'), 'exe');
  assert.equal(openedExtension('noext'), '');
});

test('the kinds mail is used to deliver lately are programs too', () => {
  for (const name of ['a.msix', 'a.appinstaller', 'a.library-ms', 'a.settingcontent-ms', 'a.py', 'a.xll', 'a.ps1xml', 'a.diagcab', 'a.search-ms']) assert.equal(runsWhenOpened(name), true, name);
});

test('an attachment is marked as from the internet, as a browser marks a download', { skip: process.platform !== 'win32' }, async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-zone-'));
  const file = path.join(dir, 'report.pdf');
  fs.writeFileSync(file, 'x');
  assert.equal(await markFromInternet(file, { write: fs.promises.writeFile, platform: process.platform }), true);
  assert.equal(fs.readFileSync(`${file}:Zone.Identifier`, 'utf8'), FROM_INTERNET);
  assert.equal(fs.readFileSync(file, 'utf8'), 'x', 'the file itself is untouched');
});

test('the mark is only written where it can be, and a failure is not the person\'s problem', async () => {
  const none = async () => { throw new Error('should not be called'); };
  assert.equal(await markFromInternet('C:\\a.pdf', { write: none, platform: 'linux' }), false);
  assert.equal(await markFromInternet('C:\\a.pdf:evil', { write: none, platform: 'win32' }), false);
  assert.equal(await markFromInternet('relative.pdf', { write: none, platform: 'win32' }), false);
  assert.equal(await markFromInternet('C:\\a.pdf', { write: async () => { throw new Error('no streams here'); }, platform: 'win32' }), false);
});

test('every way Mail hands an attachment over marks it', () => {
  const source = fs.readFileSync(path.join(import.meta.dirname, '..', 'apps', 'desktop', 'renderer', 'apps', 'mail.js'), 'utf8');
  assert.equal(source.match(/fromInternet: true/g).length, 3);
});

test('a path named after what every object answers to is not listed', () => {
  assert.equal(isListed({ 'a.js': {} }, 'a.js'), true);
  for (const p of ['constructor', 'toString', '__proto__']) assert.equal(isListed({ 'a.js': {} }, p), false, p);
});

const encodedRun = (b64) => spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', b64], { encoding: 'utf8' });

test('the elevated step runs only the script that was written', { skip: process.platform !== 'win32', timeout: 60000 }, () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "rutba-it's-elevate-"));
  const out = path.join(dir, 'ran.txt');
  const script = path.join(dir, 'apply.ps1');
  const text = `Set-Content -LiteralPath '${out.replace(/'/g, "''")}' -Value 'caf\u00e9'\r\n`;
  fs.writeFileSync(script, text);
  const hash = crypto.createHash('sha256').update(text, 'utf8').digest('hex');

  fs.appendFileSync(script, "Set-Content -LiteralPath 'nope' -Value 'changed'\r\n");
  assert.equal(encodedRun(elevationCommand(script, hash)).status, 77, 'a changed script is not run');
  assert.equal(fs.existsSync(out), false);

  fs.writeFileSync(script, text);
  assert.equal(encodedRun(elevationCommand(script, hash)).status, 0);
  assert.match(fs.readFileSync(out, 'latin1'), /caf/);
});

test('the elevated launch asks Windows first, then runs only the script that was written', { skip: process.platform !== 'win32', timeout: 60000 }, () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "rutba-it's-launch-"));
  const out = path.join(dir, 'ran.txt');
  const script = path.join(dir, 'apply.ps1');
  const text = `Set-Content -LiteralPath '${out.replace(/'/g, "''")}' -Value 'launched'\r\n`;
  fs.writeFileSync(script, text);
  const [program, args] = elevatedLaunch(elevationCommand(script, crypto.createHash('sha256').update(text, 'utf8').digest('hex')));
  assert.equal(program, 'powershell.exe');
  assert.deepEqual(args.slice(0, 2), ['-NoProfile', '-Command']);
  // Start-Process stood in for by one that says what it was asked and runs it here, without the prompt.
  const standIn = 'function Start-Process { param([string]$FilePath, [string]$Verb, [string]$WindowStyle, [string[]]$ArgumentList) Write-Output "verb=$Verb file=$FilePath"; & $FilePath @ArgumentList; exit $LASTEXITCODE }; ';
  const run = () => spawnSync('powershell.exe', ['-NoProfile', '-Command', standIn + args[2]], { encoding: 'utf8' });
  const first = run();
  assert.match(first.stdout, /verb=RunAs file=powershell\.exe/);
  assert.equal(first.status, 0, first.stderr);
  assert.equal(fs.readFileSync(out, 'latin1').trim(), 'launched');
  fs.rmSync(out);
  fs.appendFileSync(script, "Set-Content -LiteralPath 'nope' -Value 'changed'\r\n");
  assert.equal(run().status, 77, 'a script changed after it was written is not run');
  assert.equal(fs.existsSync(out), false);
});
