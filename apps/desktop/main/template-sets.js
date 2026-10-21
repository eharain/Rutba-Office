// Newer template sets, fetched when Office checks for updates (the owner,
// 2026-10-21: "the templates should have min version check so that the
// outdate software or templates can easily be upgraded to new set").
//
// The templates are their own repository, Rutba-Office-Templates, and each
// published set is a release there: a bundle, and a short description of it
// (the set's version, the Office it needs, each part's, the bundle's size
// and SHA-256) signed with a key that never leaves the publisher's machine.
// Office holds the public key. A set is taken only when its description's
// signature is good, it is newer than the set in use, and this Office is
// new enough for it; then its bundle is downloaded, its size and digest
// checked, and its files written under Office's own data folder, to be used
// from the next start. A set that needs a newer Office is reported instead,
// so the app's update is offered. Nothing unsigned is ever used: the
// samples are code, since each fills its records.

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { compareVersions, usableBy } from '@rutba/templates/versions';

export const SETS_REPO = 'eharain/Rutba-Office-Templates';

/** The keys a set may be signed with (a new one is added here before the old one is retired). */
export const PUBLIC_KEYS = [
  '-----BEGIN PUBLIC KEY-----\nMCowBQYDK2VwAyEABHQE0YTtEfhajNbh7I4g56a/kc27OYSN1Nx5gP8kSmI=\n-----END PUBLIC KEY-----\n',
];

/** A signed description's set, or throws: { description, signature } checked against `keys`. */
export function verifyDescription(signed, keys = PUBLIC_KEYS) {
  const { description, signature } = signed || {};
  if (typeof description !== 'string' || typeof signature !== 'string') throw new Error('The template set has no signed description.');
  const good = keys.some((k) => {
    try { return crypto.verify(null, Buffer.from(description), crypto.createPublicKey(k), Buffer.from(signature, 'base64')); } catch { return false; }
  });
  if (!good) throw new Error('The template set is not signed by Rutba: it was not used.');
  const d = JSON.parse(description);
  if (!/^\d+\.\d+\.\d+$/.test(d.set || '') || !d.bundle?.sha256 || !d.bundle?.name) throw new Error('The template set\'s description is not one Office reads.');
  return d;
}

/** A bundle's files written under `dir`, each path kept inside it. Answers how many. */
export function writeBundle(dir, bundle) {
  const files = bundle?.files || {};
  const root = path.resolve(dir);
  let n = 0;
  for (const [rel, b64] of Object.entries(files)) {
    const to = path.resolve(root, rel);
    // A name that climbs out of the set's folder, or is absolute, is refused.
    if (!to.startsWith(root + path.sep) || path.isAbsolute(rel) || rel.includes('\0')) throw new Error(`The template set names a file outside itself: ${rel}.`);
    fs.mkdirSync(path.dirname(to), { recursive: true });
    fs.writeFileSync(to, Buffer.from(b64, 'base64'));
    n++;
  }
  return n;
}

/**
 * The service. `root` is where sets are kept (Office's data folder,
 * templates/); `bundled` is the manifest of the set Office was built with;
 * `office` its version. `fetchJson(url)` and `fetchText(url)` are fetch,
 * stood in by the checks.
 */
export function createTemplateSets({ root, bundled, office, fetchJson, fetchText, keys = PUBLIC_KEYS, repo = SETS_REPO, api = 'https://api.github.com' }) {
  const currentFile = path.join(root, 'current.json');
  let state = { phase: 'idle', set: null, needs: null, error: null };

  /** The set to use: a downloaded one newer than the bundled set that this Office can use, else null for the bundled. */
  function inUse() {
    try {
      const c = JSON.parse(fs.readFileSync(currentFile, 'utf8'));
      const dir = path.join(root, c.set);
      const m = JSON.parse(fs.readFileSync(path.join(dir, 'manifest.json'), 'utf8'));
      if (m.set !== c.set || compareVersions(m.set, bundled.set) <= 0 || !usableBy(office, m.minOffice)) return null;
      return { set: m.set, dir, manifest: m };
    } catch {
      return null;
    }
  }

  /** The set in use's version, downloaded or bundled. */
  const version = () => inUse()?.set || bundled.set;

  async function check() {
    state = { ...state, phase: 'checking', error: null };
    try {
      const release = await fetchJson(`${api}/repos/${repo}/releases/latest`);
      const asset = (name) => (release.assets || []).find((a) => a.name === name);
      const desc = asset('templates-set.json');
      if (!desc) { state = { phase: 'current', set: version(), needs: null, error: null }; return state; }
      const d = verifyDescription(JSON.parse(await fetchText(desc.browser_download_url)), keys);
      if (compareVersions(d.set, version()) <= 0) { state = { phase: 'current', set: version(), needs: null, error: null }; return state; }
      if (!usableBy(office, d.minOffice)) { state = { phase: 'needs-office', set: d.set, needs: d.minOffice, error: null }; return state; }
      const b = asset(d.bundle.name);
      if (!b) throw new Error('The template set\'s bundle is missing from its release.');
      const text = await fetchText(b.browser_download_url);
      if (Buffer.byteLength(text) !== d.bundle.size || crypto.createHash('sha256').update(text).digest('hex') !== d.bundle.sha256) throw new Error('The template set\'s bundle is not the one its description signs: it was not used.');
      const bundle = JSON.parse(text);
      const dir = path.join(root, d.set);
      const staging = `${dir}.part`;
      fs.rmSync(staging, { recursive: true, force: true });
      writeBundle(staging, bundle);
      fs.rmSync(dir, { recursive: true, force: true });
      fs.renameSync(staging, dir);
      fs.writeFileSync(currentFile, JSON.stringify({ set: d.set, at: new Date().toISOString() }));
      state = { phase: 'ready', set: d.set, needs: null, error: null };
    } catch (err) {
      state = { ...state, phase: 'error', error: err.message || String(err) };
    }
    return state;
  }

  return { inUse, version, check, state: () => state };
}
