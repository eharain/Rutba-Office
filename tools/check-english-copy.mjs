// The English copy a release ships (C15; private design 1.4), checked before
// it is built (apps/desktop/build/dist.js runs this after the contract's keys).
//
// A release must carry the copy (resources/languages/lang-en.pack.json and
// lang-en/), taken with tools/take-lang-en.mjs from a released lang-en:
//   - its envelope holds under the contract's packages keys (a copy taken
//     --unchecked, or signed with any other key, is refused);
//   - every file is as the manifest and lang-en.source.json say;
//   - at this Office's version it matches the labels the code uses exactly:
//     none missing and none stale. A copy for an earlier Office (the
//     fallback when lang-en at this version was not ready) passes, with its
//     gaps reported.
//
//   node tools/check-english-copy.mjs

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { usable, contractFor } from '../packages/office-shell/src/electron/addons-compat.js';
import { checkFolder, sha512Of } from '../packages/office-shell/src/electron/addons.js';
import { keyOf, hashOf } from '../packages/office-shell/src/electron/language-packs.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const COPY_DIR = path.join(ROOT, 'apps', 'desktop', 'resources', 'languages');
const read = (file) => JSON.parse(fs.readFileSync(file, 'utf8'));

/**
 * @param {object} o
 * @param {string} [o.dir]        the copy's folder (resources/languages)
 * @param {string} o.office       this Office's version
 * @param {object} o.contract     the add-ons contract the release ships
 * @param {object[]} o.labels     the labels the code uses (catalogues/messages.json)
 * @returns {{ faults: string[], report: string|null, version: string|null }}
 */
export function englishCopyFaults({ dir = COPY_DIR, office, contract, labels }) {
  const faults = [];
  const envelopeFile = path.join(dir, 'lang-en.pack.json');
  if (!fs.existsSync(envelopeFile)) return { faults: ['there is no English copy: take the released lang-en with tools/take-lang-en.mjs (RELEASING.md)'], report: null, version: null };
  let envelope;
  let record;
  try {
    envelope = read(envelopeFile);
    record = read(path.join(dir, 'lang-en.source.json'));
  } catch (err) {
    return { faults: [`the English copy cannot be read (${err?.message || err})`], report: null, version: null };
  }
  const m = envelope?.manifest || {};
  if (m.id !== 'lang-en' || m.kind !== 'language') faults.push('the English copy is not lang-en');
  if (record.signatureChecked !== true) faults.push('the English copy was taken without its signature checked (--unchecked): take it again with the packages key');
  const keys = contract?.keys?.packages || [];
  const verdict = usable(envelope, { contract: contractFor(contract, { office }), office, keys: { packages: keys } });
  if (!verdict.usable && verdict.code !== 'version') faults.push(`the English copy is not usable: ${verdict.reason}`);
  const disk = checkFolder(path.join(dir, 'lang-en'), m, { hashAll: true });
  if (disk) faults.push(`the English copy: ${disk.reason}`);
  if (record.version !== m.version || (record.revision || 0) !== (m.revision || 0)) faults.push('lang-en.source.json names another lang-en than the copy');
  if (record.envelope?.sha512 !== sha512Of(envelopeFile)) faults.push('lang-en.source.json names another envelope than the copy');
  if (disk) return { faults, report: null, version: m.version || null };

  const entries = read(path.join(dir, 'lang-en', 'messages.json'));
  const sources = new Set(Object.values(entries).map((e) => e?.source));
  const wanted = new Set(labels.map((k) => hashOf(keyOf(k))));
  const uncovered = labels.filter((k) => !sources.has(hashOf(keyOf(k)))).map((k) => k.message);
  const stale = Object.entries(entries).filter(([, e]) => !wanted.has(e?.source)).map(([id]) => id);
  let report = null;
  if (m.version === office) {
    if (uncovered.length) faults.push(`lang-en ${office} lacks ${uncovered.length} labels the code uses (first: ${uncovered[0]}): take a lang-en built from this Office`);
    if (stale.length) faults.push(`lang-en ${office} has ${stale.length} entries the code no longer uses (first: ${stale[0]}): take a lang-en built from this Office`);
  } else {
    report = `the copy is lang-en ${m.version} r${m.revision || 0}, for an earlier Office: ${uncovered.length} labels uncovered, ${stale.length} stale`;
  }
  return { faults, report, version: m.version || null };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const office = read(path.join(ROOT, 'apps', 'desktop', 'package.json')).version;
  const contract = read(path.join(ROOT, 'packages', 'office-shell', 'src', 'addons-contract.json'));
  const labels = read(path.join(ROOT, 'packages', 'office-ui', 'src', 'catalogues', 'messages.json')).messages;
  const { faults, report, version } = englishCopyFaults({ office, contract, labels });
  if (faults.length) {
    for (const f of faults) console.error(f);
    process.exit(1);
  }
  console.log(report || `the English copy is lang-en ${version}, as this Office's labels`);
}
