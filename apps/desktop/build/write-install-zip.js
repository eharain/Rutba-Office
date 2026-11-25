// The command-line install, as a zip beside the Windows installer (2.28.0).
//
//   node build/write-install-zip.js [release folder]
//
// office.rutba.io/download lists the lines that fetch and start the installer
// from Command Prompt, for a person whose browser download Windows will not
// run, and offers the same steps as a script inside a zip (owner, 24 Nov
// 2026). The site made that zip itself. From 2.28.0 the build makes it, from
// the installer it has just built, and it goes up as a file of the release: the
// site links a release file and makes none of its own.
//
// The zip, rutba-office-<version>-install.zip, holds one file,
// Install-Rutba-Office-<version>.cmd, filled from build/install.cmd.template
// with four values: the version, the installer's name, its address in the
// release on GitHub, and its SHA-256. The script downloads that installer,
// stops if the file is not the one with that SHA-256, starts it, and removes
// the download. It changes no Windows setting.
//
// Every value is held to its shape before it goes in, because cmd runs the
// script: the version is digits and two dots, the SHA-256 is 64 lower-case hex
// characters, and the name and the address are made here from the version. The
// script is plain ASCII with CRLF line ends (cmd reads labels wrongly with
// LF); the template is LF on disk.
//
// The same installer gives the same zip, byte for byte: the entry's date is
// fixed. So what is signed and published can be held to it: the zip in a draft
// must be the one that draft's installer gives (tools/publish-release.js,
// installZipProblems), or nothing is signed.

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { ZipEntry, writeZip, readZip } from '../../../packages/ooxml/src/zip.js';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const TEMPLATE = path.join(HERE, 'install.cmd.template');

export const installerName = (version) => `Rutba-Office-${version}-win-x64.exe`;
export const installZipName = (version) => `rutba-office-${version}-install.zip`;
export const installScriptName = (version) => `Install-Rutba-Office-${version}.cmd`;
export const installerUrl = (version) => `https://github.com/eharain/Rutba-Office/releases/download/v${version}/${installerName(version)}`;

/** A file's SHA-256 as the script and certutil write it: 64 lower-case hex characters. */
export const sha256Of = (bytes) => crypto.createHash('sha256').update(bytes).digest('hex');

/**
 * The install script for a version and its installer's SHA-256: the template
 * filled, as the bytes cmd reads. Throws on a value that is not its shape, a
 * field the template names that is not one of the four, or a character that is
 * not plain ASCII.
 */
export function installScript({ version, sha256, template = fs.readFileSync(TEMPLATE, 'utf8') }) {
  if (!/^\d+\.\d+\.\d+$/.test(String(version))) throw new Error(`"${version}" is not a version: digits and two dots`);
  if (!/^[0-9a-f]{64}$/.test(String(sha256))) throw new Error('the SHA-256 is not 64 lower-case hex characters');
  const values = { VERSION: String(version), FILE: installerName(version), URL: installerUrl(version), SHA256: String(sha256) };
  const text = String(template).replace(/\r\n/g, '\n').replace(/\{\{([^{}]*)\}\}/g, (all, key) => {
    if (!Object.hasOwn(values, key)) throw new Error(`the install script's template names a field this build does not fill: ${all}`);
    return values[key];
  });
  if (/\{\{|\}\}/.test(text)) throw new Error('the install script still has a field left unfilled');
  for (const key of Object.keys(values)) if (!String(template).includes(`{{${key}}}`)) throw new Error(`the install script's template has no {{${key}}}`);
  const odd = /[^\t\n\x20-\x7e]/.exec(text);
  if (odd) throw new Error(`the install script must be plain ASCII: it has U+${odd[0].codePointAt(0).toString(16).toUpperCase().padStart(4, '0')}`);
  return Buffer.from(text.replace(/\n/g, '\r\n'), 'ascii');
}

/** The zip that holds the one script, for a version and its installer's SHA-256. */
export function installZip({ version, sha256, template }) {
  const entry = new ZipEntry({ name: installScriptName(version), method: 8, crc: 0, compressedSize: 0, uncompressedSize: 0, compressed: Buffer.alloc(0), dosTime: 0, dosDate: 0x21, flags: 0, externalAttrs: 0, comment: Buffer.alloc(0) });
  entry.data = installScript({ version, sha256, template });
  return writeZip([entry]);
}

/**
 * What is wrong between an install zip and the installer it is for, as
 * sentences; empty when the zip is the one that installer gives.
 *
 * @param {{ zip: Buffer, installer: Buffer, version: string, template?: string }} given
 */
export function installZipProblems({ zip, installer, version, template }) {
  const name = installZipName(version);
  let entries;
  try {
    ({ entries } = readZip(zip));
  } catch (err) {
    return [`${name} cannot be read as a zip (${err.message})`];
  }
  if (entries.length !== 1 || entries[0].name !== installScriptName(version)) {
    return [`${name} holds ${entries.map((e) => e.name).join(', ') || 'nothing'}, not the one file ${installScriptName(version)}`];
  }
  let script;
  try {
    script = entries[0].data;
  } catch (err) {
    return [`the script in ${name} cannot be read (${err.message})`];
  }
  const want = installScript({ version, sha256: sha256Of(installer), template });
  if (!Buffer.from(script).equals(want)) return [`the script in ${name} is not the one ${installerName(version)} gives: it would fetch or accept another file`];
  return [];
}

/** Write the zip beside the installer in `dir`; answers its path. Throws when the installer is not there. */
export function writeInstallZip(dir, version) {
  const installer = path.join(dir, installerName(version));
  if (!fs.existsSync(installer)) throw new Error(`${installerName(version)} is not in ${dir}: the install zip is made from the installer`);
  const sha256 = sha256Of(fs.readFileSync(installer));
  const out = path.join(dir, installZipName(version));
  fs.writeFileSync(out, installZip({ version, sha256 }));
  return { out, sha256 };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const dir = path.resolve(process.argv[2] || path.join(HERE, '..', 'release'));
  const { version } = JSON.parse(fs.readFileSync(path.join(HERE, '..', 'package.json'), 'utf8'));
  try {
    const { out, sha256 } = writeInstallZip(dir, version);
    console.log(`${path.basename(out)}: ${fs.statSync(out).size} bytes, for ${installerName(version)} with SHA-256 ${sha256}`);
  } catch (err) {
    console.error(err.message);
    process.exit(1);
  }
}
