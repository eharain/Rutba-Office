/**
 * Fidelity measurement.
 *
 * "Fidelity is critical" is only a requirement if it is a NUMBER. This module
 * turns it into one: open a file, do something to it, write it back, and report
 * exactly which parts changed and which survived byte-identically.
 *
 * Two things it is careful about:
 *
 * 1. It compares the PACKAGE, not the bytes of the archive. Two zips can differ
 *    in compression while carrying identical parts; that is not a fidelity loss
 *    and reporting it as one would train everyone to ignore the report.
 *
 * 2. It reports loss per part with a reason. "97% preserved" is useless if the
 *    3% is the chart. The report names names.
 *
 * The intended use is a corpus: point `fidelity-report` at a directory of real
 * customer files — the ones that actually go to banks and government — and get a
 * per-file, per-part verdict. Synthetic tests prove the mechanism; only real
 * files prove the claim.
 */
import { createHash } from 'node:crypto';
import { OoxmlPackage } from './package.js';

const digest = (buf) => createHash('sha256').update(buf).digest('hex');

/** Whitespace-insensitive XML comparison, for parts we DID rewrite. */
export function xmlEquivalent(a, b) {
  const norm = (s) =>
    String(s)
      .replace(/<\?xml[^>]*\?>/g, '')
      .replace(/>\s+</g, '><')
      .replace(/\s+/g, ' ')
      .trim();
  return norm(a) === norm(b);
}

/**
 * Compare two packages part by part.
 *
 * @returns {{
 *   total: number, identical: string[], changed: object[],
 *   added: string[], removed: string[], score: number
 * }}
 */
export function comparePackages(beforeBuf, afterBuf) {
  const before = OoxmlPackage.read(beforeBuf);
  const after = OoxmlPackage.read(afterBuf);

  const beforeNames = new Set(before.partNames());
  const afterNames = new Set(after.partNames());

  const identical = [];
  const changed = [];
  const removed = [];
  const added = [];

  for (const name of beforeNames) {
    if (!afterNames.has(name)) {
      removed.push(name);
      continue;
    }
    const b = before.read(name);
    const a = after.read(name);
    if (b.length === a.length && digest(b) === digest(a)) {
      identical.push(name);
      continue;
    }
    const isXml = name.endsWith('.xml') || name.endsWith('.rels');
    changed.push({
      name,
      contentType: before.contentTypeOf(name),
      bytesBefore: b.length,
      bytesAfter: a.length,
      xmlEquivalent: isXml ? xmlEquivalent(b.toString('utf8'), a.toString('utf8')) : false,
    });
  }
  for (const name of afterNames) if (!beforeNames.has(name)) added.push(name);

  const total = beforeNames.size;
  return {
    total,
    identical,
    changed,
    added,
    removed,
    // The number that matters: how much of the original survived untouched.
    score: total === 0 ? 1 : identical.length / total,
  };
}

/**
 * Open a file, optionally edit it, write it back, and report the damage.
 *
 * @param {Buffer} buf
 * @param {(pkg: OoxmlPackage) => void|Promise<void>} [edit]
 */
export async function fidelityOf(buf, edit) {
  const pkg = OoxmlPackage.read(buf);
  const declaredKind = pkg.kind();
  if (edit) await edit(pkg);
  const out = pkg.write();
  const diff = comparePackages(buf, out);

  return {
    kind: declaredKind,
    bytesIn: buf.length,
    bytesOut: out.length,
    intendedEdits: pkg.modifiedParts(),
    ...diff,
    // Anything changed that we did not deliberately edit is a bug, not a
    // trade-off. This is the assertion the tests hang on.
    unintendedChanges: diff.changed
      .map((c) => c.name)
      .filter((n) => !pkg.modifiedParts().includes(n)),
    output: out,
  };
}

/** Human-readable one-file report. */
export function formatReport(name, result) {
  const lines = [];
  const pct = (n) => (n * 100).toFixed(1) + '%';
  lines.push(name);
  lines.push('  kind            ' + result.kind);
  lines.push('  parts           ' + result.total);
  lines.push('  byte-identical  ' + result.identical.length + '  (' + pct(result.score) + ')');
  if (result.intendedEdits.length) lines.push('  edited          ' + result.intendedEdits.join(', '));
  if (result.unintendedChanges.length) {
    lines.push('  !! UNINTENDED   ' + result.unintendedChanges.join(', '));
  }
  if (result.removed.length) lines.push('  !! LOST         ' + result.removed.join(', '));
  if (result.added.length) lines.push('  added           ' + result.added.join(', '));
  lines.push('  bytes           ' + result.bytesIn + ' -> ' + result.bytesOut);
  return lines.join('\n');
}

/**
 * Features present in a package that we can preserve but cannot yet render or
 * recalculate. Being explicit about this is the difference between an honest
 * fidelity claim and a misleading one.
 */
const FEATURE_MARKERS = [
  { part: /^xl\/charts?\//, feature: 'charts' },
  { part: /^xl\/pivotCache\//, feature: 'pivot tables' },
  { part: /^xl\/pivotTables?\//, feature: 'pivot tables' },
  { part: /^xl\/drawings\//, feature: 'drawings and shapes' },
  { part: /^xl\/media\//, feature: 'embedded media' },
  { part: /^xl\/threadedComments\//, feature: 'threaded comments' },
  { part: /^xl\/queryTables?\//, feature: 'external data queries' },
  { part: /^xl\/connections\.xml$/, feature: 'external data connections' },
  { part: /^xl\/slicer/, feature: 'slicers' },
  { part: /vbaProject\.bin$/, feature: 'VBA macros' },
  { part: /^customXml\//, feature: 'custom XML' },
  { part: /^_xmlsignatures\//, feature: 'digital signatures' },
  { part: /^word\/footnotes\.xml$/, feature: 'footnotes' },
  { part: /^word\/header/, feature: 'headers' },
  { part: /^word\/footer/, feature: 'footers' },
  { part: /^word\/embeddings\//, feature: 'embedded objects' },
];

/** @returns {string[]} features carried through but not understood */
export function preservedButUnsupported(buf) {
  const pkg = OoxmlPackage.read(buf);
  const found = new Set();
  for (const name of pkg.partNames()) {
    for (const { part, feature } of FEATURE_MARKERS) if (part.test(name)) found.add(feature);
  }
  return [...found].sort();
}
