// The languages this release no longer carries built in, written into its
// update feeds (release/latest*.yml) after electron-builder has made them:
//
//   node build/write-left-core.js
//
// An installed Office reads an update's feed before anything else of it. With
// leftCore there, it holds the update for a person whose language the update
// leaves out of the core, even when the packages feed for that version cannot
// be had (D-2026-11-13-36); with no feed and the language still built in, it
// holds nothing. The list is the add-ons contract's languages.leftCore, which
// tools/publish-release.js checks again against the feed it uploads.
// Signing (publish-release.js --sign) signs the installers and packs, never
// this file, so what is written here is what installed copies read.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const APP = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const CONTRACT = path.join(APP, '..', '..', 'packages', 'office-shell', 'src', 'addons-contract.json');

/** The contract's leftCore: the language tags this release leaves out of the core. */
export function contractLeftCore(file = CONTRACT) {
  const list = JSON.parse(fs.readFileSync(file, 'utf8'))?.languages?.leftCore;
  if (!Array.isArray(list) || !list.every((t) => typeof t === 'string' && /^[a-z]{2,3}$/.test(t))) throw new Error(`${file}: languages.leftCore is not a list of language tags`);
  return [...new Set(list)].sort();
}

/** The feed's text with one `leftCore:` line, the given list, in place of any it had; its line endings kept. */
export function withLeftCore(text, list) {
  const eol = /\r\n/.test(text) ? '\r\n' : '\n';
  const lines = String(text).replace(/\r\n/g, '\n').split('\n').filter((l) => !/^leftCore:/.test(l));
  while (lines.length && lines[lines.length - 1] === '') lines.pop();
  lines.push(`leftCore: [${list.map((t) => `'${t}'`).join(', ')}]`);
  return lines.join(eol) + eol;
}

function main() {
  const dir = path.join(APP, 'release');
  const feeds = fs.existsSync(dir) ? fs.readdirSync(dir).filter((n) => /^latest.*\.yml$/.test(n)) : [];
  if (!feeds.length) {
    console.error(`write-left-core: no update feed in ${dir}`);
    process.exit(1);
  }
  const list = contractLeftCore();
  for (const name of feeds) {
    const file = path.join(dir, name);
    fs.writeFileSync(file, withLeftCore(fs.readFileSync(file, 'utf8'), list));
    console.log(`${name}: leftCore [${list.join(', ')}]`);
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) main();
