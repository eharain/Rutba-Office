// Takes a released lang-en into Office as its English copy (C15).
//
//   node tools/take-lang-en.mjs <release folder> [--from <url>] [--keys <file>] [--unchecked]
//
// The release folder holds lang-en.pack.json and lang-en/, as the packages
// repository's tools/sign.mjs leaves them (languages/dist/v<version>/). Both
// are copied byte for byte to apps/desktop/resources/languages/, with
// lang-en.source.json saying which release they are and the hash of every
// file, and the copy is shipped beside the archive (electron-builder.config.cjs).
//
// Before it copies, it checks what it takes as Office will: the envelope's
// signature under the packages keys (addons-contract.json, or --keys, a JSON
// list of public keys), the kind and format, and every file against the
// manifest. --unchecked takes a copy without the signature, for a build that
// is not a release; the release build refuses such a copy
// (tools/check-english-copy.mjs).
//
// It also says how many of the labels the code uses (catalogues/messages.json)
// the copy has: when the copy is at this Office's version, the gate wants all.
// The copy is committed by itself, as the release steps say (RELEASING.md).

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { usable, contractFor } from '../packages/office-shell/src/electron/addons-compat.js';
import { checkFolder, sha512Of } from '../packages/office-shell/src/electron/addons.js';
import { keyOf, hashOf, LANGUAGE_FORMATS } from '../packages/office-shell/src/electron/language-packs.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const COPY_DIR = path.join(ROOT, 'apps', 'desktop', 'resources', 'languages');
const CONTRACT = JSON.parse(fs.readFileSync(path.join(ROOT, 'packages', 'office-shell', 'src', 'addons-contract.json'), 'utf8'));
const KEYS_LIST = path.join(ROOT, 'packages', 'office-ui', 'src', 'catalogues', 'messages.json');

/** Every file under a folder, relative, with forward slashes. */
function filesUnder(dir, rel = '', out = []) {
  for (const e of fs.readdirSync(path.join(dir, rel), { withFileTypes: true })) {
    const p = rel ? `${rel}/${e.name}` : e.name;
    if (e.isDirectory()) filesUnder(dir, p, out);
    else out.push(p);
  }
  return out.sort();
}

/**
 * Checks a release folder's lang-en and copies it in.
 * @param {object} o
 * @param {string} o.from          the release folder
 * @param {string} [o.to]          where the copy goes (resources/languages)
 * @param {string} [o.source]      where the release was taken from, for lang-en.source.json
 * @param {string[]} [o.keys]      the packages keys to check the signature with
 * @param {boolean} [o.unchecked]  take it with no signature check, when there are no keys yet
 * @param {object[]} [o.labels]    the labels the code uses
 * @param {string} [o.office]      this Office's version
 * @returns {{ id, version, revision, files: number, labels: { used: number, found: number }, checked: boolean }}
 */
export function takeLangEn({ from, to = COPY_DIR, source = null, keys = CONTRACT.keys?.packages || [], unchecked = false, labels = null, office = null }) {
  const envelopeFile = path.join(from, 'lang-en.pack.json');
  const folder = path.join(from, 'lang-en');
  if (!fs.existsSync(envelopeFile)) throw new Error(`There is no lang-en.pack.json in ${from}.`);
  const envelope = JSON.parse(fs.readFileSync(envelopeFile, 'utf8'));
  const m = envelope?.manifest;
  if (m?.id !== 'lang-en' || m.kind !== 'language') throw new Error('That envelope is not lang-en\'s.');
  if (!LANGUAGE_FORMATS.includes(m.format)) throw new Error(`lang-en is in format ${m.format}, which this Office does not read.`);

  let checked = false;
  if (keys.length) {
    const verdict = usable(envelope, { contract: contractFor(CONTRACT, { office: office || m.version }), office: office || m.version, keys: { packages: keys } });
    if (!verdict.usable) throw new Error(verdict.reason);
    checked = true;
  } else if (!unchecked) {
    throw new Error('There are no packages keys to check lang-en with yet. Take it with --unchecked for a build that is not a release.');
  }
  const refused = checkFolder(folder, m, { hashAll: true });
  if (refused) throw new Error(refused.reason);

  // What the code uses that the copy has: by the key's hash, as Office looks it up.
  const list = labels || JSON.parse(fs.readFileSync(KEYS_LIST, 'utf8')).messages;
  const entries = JSON.parse(fs.readFileSync(path.join(folder, 'messages.json'), 'utf8'));
  const sources = new Set(Object.values(entries).map((e) => e?.source));
  const found = list.filter((k) => sources.has(hashOf(keyOf(k)))).length;

  // Byte for byte: the old copy goes first, so nothing of it is left.
  fs.rmSync(to, { recursive: true, force: true });
  fs.mkdirSync(path.join(to, 'lang-en'), { recursive: true });
  fs.copyFileSync(envelopeFile, path.join(to, 'lang-en.pack.json'));
  const names = filesUnder(folder);
  for (const p of names) {
    const target = path.join(to, 'lang-en', ...p.split('/'));
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.copyFileSync(path.join(folder, ...p.split('/')), target);
  }
  const record = {
    id: 'lang-en',
    version: m.version,
    revision: m.revision || 0,
    from: source,
    signatureChecked: checked,
    envelope: { sha512: sha512Of(path.join(to, 'lang-en.pack.json')) },
    files: Object.fromEntries(names.map((p) => [p, { sha512: sha512Of(path.join(to, 'lang-en', ...p.split('/'))) }])),
  };
  fs.writeFileSync(path.join(to, 'lang-en.source.json'), `${JSON.stringify(record, null, 2)}\n`);
  return { id: 'lang-en', version: m.version, revision: record.revision, files: names.length, labels: { used: list.length, found }, checked };
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const args = process.argv.slice(2);
  const flag = (name) => {
    const i = args.indexOf(name);
    if (i < 0) return null;
    const v = args[i + 1];
    args.splice(i, 2);
    return v;
  };
  const source = flag('--from');
  const keysFile = flag('--keys');
  const unchecked = args.includes('--unchecked');
  const from = args.filter((a) => a !== '--unchecked')[0];
  if (!from) {
    console.error('node tools/take-lang-en.mjs <release folder> [--from <url>] [--keys <file>] [--unchecked]');
    process.exit(2);
  }
  try {
    const keys = keysFile ? JSON.parse(fs.readFileSync(keysFile, 'utf8')) : undefined;
    const office = JSON.parse(fs.readFileSync(path.join(ROOT, 'apps', 'desktop', 'package.json'), 'utf8')).version;
    const r = takeLangEn({ from: path.resolve(from), source, keys, unchecked, office });
    console.log(`lang-en ${r.version} r${r.revision}: ${r.files} files taken, ${r.labels.found} of the ${r.labels.used} labels the code uses${r.checked ? '' : '. Its signature was not checked'}.`);
    if (r.version === office && r.labels.found !== r.labels.used) console.log(`It is at this Office's version (${office}) and lacks ${r.labels.used - r.labels.found} labels: the gate will refuse it until a lang-en built for this Office is taken.`);
    console.log(`Commit ${path.relative(ROOT, COPY_DIR)} by itself.`);
  } catch (err) {
    console.error(err.message);
    process.exit(1);
  }
}
