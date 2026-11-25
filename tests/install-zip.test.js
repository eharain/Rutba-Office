// The command-line install, as a zip beside the Windows installer (apps/desktop/build/write-install-zip.js).
//
// From 2.28.0 the build makes rutba-office-<version>-install.zip from the installer it has just built,
// and it goes up as a file of the release, for office.rutba.io/download to link. The zip holds one
// script, which fetches that installer, stops unless it has the SHA-256 the script names, and starts
// it. cmd runs the script on a person's computer, so what goes into it is held to its shape here, and
// a release is not signed unless its zip is the one its own installer gives (tests/update-signature.test.js).
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { installScript, installZip, installZipProblems, installZipName, installScriptName, installerName, installerUrl, sha256Of, writeInstallZip } from '../apps/desktop/build/write-install-zip.js';
import { readZip } from '../packages/ooxml/src/zip.js';
import { releaseAssets } from '../tools/publish-release.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const read = (rel) => fs.readFileSync(path.join(ROOT, rel), 'utf8');
const VERSION = '2.28.0';
const INSTALLER = Buffer.from('the installer of 2.28.0, a stand-in for a hundred megabytes');
const SHA = sha256Of(INSTALLER);

test('the script is the template with its four values, plain ASCII, with the line ends cmd reads', () => {
  const script = installScript({ version: VERSION, sha256: SHA });
  const text = script.toString('latin1');
  assert.ok(text.includes(`set "VERSION=${VERSION}"\r\n`));
  assert.ok(text.includes(`set "FILE=Rutba-Office-${VERSION}-win-x64.exe"\r\n`));
  assert.ok(text.includes(`set "URL=https://github.com/eharain/Rutba-Office/releases/download/v${VERSION}/Rutba-Office-${VERSION}-win-x64.exe"\r\n`));
  assert.ok(text.includes(`set "SHA256=${SHA}"\r\n`));
  assert.doesNotMatch(text, /\{\{|\}\}/, 'no field left unfilled');
  assert.doesNotMatch(text, /[^\t\r\n\x20-\x7e]/, 'plain ASCII');
  assert.doesNotMatch(text.replace(/\r\n/g, ''), /[\r\n]/, 'every line ends CRLF, and none ends any other way');
  assert.ok(text.startsWith('@echo off\r\n'));
  // What the script does is what the template says: it stops unless the file has the SHA-256 named.
  assert.match(text, /if \/i not "%GOT%"=="%SHA256%" goto :bad_file/);
  assert.match(text, /:bad_file\r\necho The file does not match its published SHA-256, so it was not run\./);
});

test('a value that is not its shape never reaches the script', () => {
  for (const version of ['2.28', '2.28.0-beta', '2.28.0 & calc', '2.28.0"', '', 'v2.28.0', '2.28.0\n']) {
    assert.throws(() => installScript({ version, sha256: SHA }), /is not a version/, JSON.stringify(version));
  }
  for (const sha256 of ['', SHA.toUpperCase(), SHA.slice(1), `${SHA}0`, `${SHA.slice(0, 63)}g`, `${SHA.slice(0, 60)}&x&y`]) {
    assert.throws(() => installScript({ version: VERSION, sha256 }), /SHA-256 is not 64 lower-case hex/, sha256);
  }
  // The name and the address are made from the version alone.
  assert.equal(installerName(VERSION), 'Rutba-Office-2.28.0-win-x64.exe');
  assert.equal(installerUrl(VERSION), 'https://github.com/eharain/Rutba-Office/releases/download/v2.28.0/Rutba-Office-2.28.0-win-x64.exe');
});

test('a template that names another field, lacks one of the four, or is not plain ASCII is refused', () => {
  const good = 'set "V={{VERSION}}"\nset "F={{FILE}}"\nset "U={{URL}}"\nset "S={{SHA256}}"\n';
  assert.equal(installScript({ version: VERSION, sha256: SHA, template: good }).toString('latin1').split('\r\n').length, 5);
  assert.throws(() => installScript({ version: VERSION, sha256: SHA, template: `${good}echo {{OTHER}}\n` }), /names a field this build does not fill: \{\{OTHER\}\}/);
  assert.throws(() => installScript({ version: VERSION, sha256: SHA, template: good.replace('set "S={{SHA256}}"\n', '') }), /has no \{\{SHA256\}\}/);
  assert.throws(() => installScript({ version: VERSION, sha256: SHA, template: `${good}rem caf\u00e9\n` }), /plain ASCII: it has U\+00E9/);
  assert.throws(() => installScript({ version: VERSION, sha256: SHA, template: `${good}rem {{VERSION\n` }), /left unfilled/);
});

test('the zip holds the one script, and the same installer gives the same zip byte for byte', () => {
  const zip = installZip({ version: VERSION, sha256: SHA });
  const { entries } = readZip(zip);
  assert.deepEqual(entries.map((e) => e.name), [installScriptName(VERSION)]);
  assert.equal(installScriptName(VERSION), 'Install-Rutba-Office-2.28.0.cmd');
  assert.ok(Buffer.from(entries[0].data).equals(installScript({ version: VERSION, sha256: SHA })));
  assert.ok(zip.equals(installZip({ version: VERSION, sha256: SHA })), 'made again, the same bytes');
  assert.equal(installZipName(VERSION), 'rutba-office-2.28.0-install.zip');
});

test('a zip is held to the installer it is for', () => {
  const zip = installZip({ version: VERSION, sha256: SHA });
  assert.deepEqual(installZipProblems({ zip, installer: INSTALLER, version: VERSION }), []);
  // A zip made for another installer, or for another version, would accept another file.
  assert.match(installZipProblems({ zip, installer: Buffer.from('another build'), version: VERSION })[0], /is not the one Rutba-Office-2\.28\.0-win-x64\.exe gives/);
  assert.match(installZipProblems({ zip: installZip({ version: '2.27.0', sha256: SHA }), installer: INSTALLER, version: VERSION })[0], /holds Install-Rutba-Office-2\.27\.0\.cmd, not the one file Install-Rutba-Office-2\.28\.0\.cmd/);
  assert.match(installZipProblems({ zip: Buffer.from('not a zip at all'), installer: INSTALLER, version: VERSION })[0], /cannot be read as a zip/);
});

test('the build writes it beside the installer, and stops when there is no installer to make it from', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-install-zip-'));
  try {
    assert.throws(() => writeInstallZip(dir, VERSION), /Rutba-Office-2\.28\.0-win-x64\.exe is not in/);
    fs.writeFileSync(path.join(dir, installerName(VERSION)), INSTALLER);
    const { out, sha256 } = writeInstallZip(dir, VERSION);
    assert.equal(path.basename(out), 'rutba-office-2.28.0-install.zip');
    assert.equal(sha256, SHA);
    assert.deepEqual(installZipProblems({ zip: fs.readFileSync(out), installer: INSTALLER, version: VERSION }), []);
    // It is one of the files that go up with the release.
    assert.ok(releaseAssets(fs.readdirSync(dir), `v${VERSION}`).includes('rutba-office-2.28.0-install.zip'));
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
  // The Windows build makes it last, and both workflows build with that script.
  assert.match(read('apps/desktop/build/dist.js'), /if \(platform === 'win'\) step\(\[path\.join\(APP, 'build', 'write-install-zip\.js'\)\]\);/);
  assert.match(read('.github/workflows/release.yml'), /npm run dist:win/);
  assert.match(read('.github/workflows/candidate.yml'), /npm run dist:win/);
  // The installer packs only build/out of the build folder: the template is not in it.
  assert.doesNotMatch(read('apps/desktop/electron-builder.config.cjs'), /install\.cmd\.template/);
});
