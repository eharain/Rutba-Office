// What a database made from one of Data Files' templates or samples keeps of
// where it came from (owner, 2026-10-22, D-2026-10-22-03 and -04): which one,
// from which template set, and a fingerprint of each design item (query,
// form, report, print template) and each table's fields as they were made.
// The fingerprints are what let a newer design be offered later without
// touching anything the person has changed since (upgrade.js).

import crypto from 'node:crypto';
import path from 'node:path';
import { readManifest, MANIFEST_PATH } from '@rutba/templates/manifest';

/** The design items a database keeps, each read as name and stored design. */
export const DESIGN_KINDS = {
  queries: (b) => b.queries().map((x) => [x.name, x.design]),
  forms: (b) => b.forms().map((x) => [x.name, x.layout]),
  reports: (b) => b.reports().map((x) => [x.name, x.layout]),
  documents: (b) => b.documents().map((x) => [x.name, x.layout]),
};

/** A design item's fingerprint: the same stored design, the same print. */
export const printOf = (value) => crypto.createHash('sha256').update(JSON.stringify(value)).digest('hex').slice(0, 20);

/** Every design item's fingerprint, and each table's fields. */
export function designPrint(base) {
  const out = { tables: Object.fromEntries(base.tables().map((t) => [t.name, base.fields(t.name).map((f) => f.name)])) };
  for (const [kind, read] of Object.entries(DESIGN_KINDS)) out[kind] = Object.fromEntries(read(base).map(([n, v]) => [n, printOf(v)]));
  return out;
}

/** The template set in use: the one Office downloaded, or the one it was built with. */
export function activeSet() {
  const dir = process.env.RUTBA_TEMPLATES_DIR;
  try { return readManifest(dir ? path.join(dir, 'manifest.json') : MANIFEST_PATH).set; } catch { return null; }
}

/** Remember that `base` was just made from template or sample `id`. */
export function recordOrigin(base, { kind, id }) {
  base.setMeta('origin', JSON.stringify({ kind, id, set: activeSet(), made: new Date().toISOString() }));
  base.setMeta('origin.design', JSON.stringify(designPrint(base)));
}

/** Where `base` came from ({ kind: 'sample' | 'template', id, set, made }), or null for a database of one's own. */
export function originOf(base) {
  try {
    const o = JSON.parse(base.meta('origin') || 'null');
    return o && (o.kind === 'sample' || o.kind === 'template') && o.id ? o : null;
  } catch {
    return null;
  }
}

/** The fingerprints kept when it was made (or last brought up to date). */
export function originPrint(base) {
  try { return JSON.parse(base.meta('origin.design') || 'null'); } catch { return null; }
}
