/**
 * Strict Open XML (ISO/IEC 29500 Strict), read and written through the
 * package's one mapping layer (packages/ooxml/src/strict.js).
 *
 * The Strict files in fixtures/strict are the rich fixtures saved again by
 * Word, Excel and PowerPoint (tools/make-strict-fixtures.ps1). Each must open
 * with everything its Transitional twin has; a save of one must stay Strict
 * unless the usual form is chosen; and a Strict save of ours must pass the
 * same structural check Office's own Strict files pass, which needs no
 * Office. tools/check-strict-office.ps1 has Word, Excel and PowerPoint open
 * our Strict saves as well.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import { join, dirname, basename } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createDocumentService } from '../apps/desktop/main/documents.js';
import { OoxmlPackage } from '../packages/ooxml/src/package.js';
import { readZip } from '../packages/ooxml/src/zip.js';
import { isoToSerial, PERCENT_ATTRS, transitionalUri, strictUri } from '../packages/ooxml/src/strict.js';

const FIX = join(dirname(fileURLToPath(import.meta.url)), 'fixtures');
const PAIRS = [
  ['xlsx', join(FIX, 'rich', 'showcase.xlsx'), join(FIX, 'strict', 'showcase-strict.xlsx')],
  ['docx', join(FIX, 'rich', 'showcase.docx'), join(FIX, 'strict', 'showcase-strict.docx')],
  ['pptx', join(FIX, 'rich', 'showcase.pptx'), join(FIX, 'strict', 'showcase-strict.pptx')],
];
const have = PAIRS.every(([, a, b]) => fs.existsSync(a) && fs.existsSync(b));
const skip = !have && 'fixtures not generated';
const doc = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
const dir = fs.mkdtempSync(join(os.tmpdir(), 'rutba-strict-'));
test.after(() => fs.rmSync(dir, { recursive: true, force: true }));

/** What a window is shown of a file: every cell's text, every block's, every slide's words. */
function seen(id, kind, model) {
  if (kind === 'sheet') {
    const out = {};
    for (const name of model.sheets) {
      doc.apply({ id, ops: [{ op: 'sheet', name }, { op: 'viewport', x: 0, y: 0, width: 2400, height: 1800 }] });
      const m = doc.model({ id });
      out[name] = Object.fromEntries(m.cells.filter((c) => c.text !== '' && c.text != null).map((c) => [c.ref, c.text]));
      out[name + ' drawings'] = m.drawings.map((d) => d.kind).sort();
    }
    return out;
  }
  if (kind === 'deck') {
    return Array.from({ length: model.count }, (_, i) => {
      const s = doc.model({ id, slide: i }).slide;
      return { words: s.shapes.map((sh) => JSON.stringify(sh.text || '')).join('|'), kinds: s.shapes.map((sh) => sh.kind).sort(), notes: s.notes || '' };
    });
  }
  return {
    // Word brings a DATE field's shown result up to the day whenever it saves.
    blocks: model.blocks.map((b) => `${b.style || ''}:${b.text || ''}`.replace(/today is \d{1,2} \w+ \d{4}/, 'today is (the day saved)')),
    notes: [model.footnotes.length, model.endnotes.length, model.comments.length],
    watermark: Boolean(model.bands?.watermark),
  };
}

function opened(file) {
  const s = doc.open({ path: file });
  try {
    return { conformance: s.conformance, seen: seen(s.id, s.kind, s.model) };
  } finally {
    doc.close({ id: s.id });
  }
}

/* ── the structural check ───────────────────────────────────────────── */

const ISO = /^(http:\/\/purl\.oclc\.org\/ooxml\/|http:\/\/schemas\.openxmlformats\.org\/(markup-compatibility|package)\/2006|http:\/\/purl\.org\/dc\/|http:\/\/www\.w3\.org\/)/;
const TRANSITIONAL_URI = /\s(?:xmlns(?::[\w.-]+)?|Type|uri)\s*=\s*"(http:\/\/schemas\.openxmlformats\.org\/(?:officeDocument|wordprocessingml|spreadsheetml|presentationml|drawingml)\/2006[^"]*)"/g;

/** Text with each element whose name `name` (a pattern) matches cut out, nested ones with it. */
function without(xml, name) {
  let out = xml;
  const start = new RegExp(`<(${name})\\b`, 'g');
  for (let found = start.exec(out); found; start.lastIndex = found.index, found = start.exec(out)) {
    const at = found.index;
    const re = new RegExp(`<(/?)${found[1]}\\b[^>]*?(/?)>`, 'g');
    re.lastIndex = at;
    let depth = 0;
    let end = -1;
    for (let m; (m = re.exec(out));) {
      if (m[1]) depth -= 1;
      else if (!m[2]) depth += 1;
      if (depth === 0) { end = m.index + m[0].length; break; }
    }
    if (end < 0) break;
    out = out.slice(0, at) + out.slice(end);
  }
  return out;
}

/**
 * What in a package is not Strict as Office writes Strict: Transitional
 * namespaces or relationship types (a VML drawing's aside), VML outside a
 * choice that requires it, DrawingML percentages in thousandths,
 * WordprocessingML's and SpreadsheetML's left and right where Strict has
 * start and end, an extension namespace a Strict reader is not told it may
 * ignore, and a main part not marked Strict.
 */
function strictProblems(bytes) {
  const problems = [];
  const parts = new Map(readZip(bytes).entries.map((e) => [e.name, e.data.toString('utf8')]));
  const rels = parts.get('_rels/.rels') || '';
  const main = /Type="http:\/\/purl\.oclc\.org\/ooxml\/officeDocument\/relationships\/officeDocument"\s+Target="\/?([^"]+)"|Target="\/?([^"]+)"\s+Type="http:\/\/purl\.oclc\.org\/ooxml\/officeDocument\/relationships\/officeDocument"/.exec(rels);
  if (!main) problems.push('_rels/.rels: no Strict relationship to the main part');
  const mainName = main && (main[1] || main[2]);
  for (const [name, xml] of parts) {
    if (!/\.(xml|rels)$/.test(name)) continue;
    // What a reader that knows no VML and no extension sees: choices and extension lists aside (PowerPoint writes
    // a transition's fallback in Transitional namespaces even in a Strict file).
    const plain = without(without(xml, 'mc:AlternateContent'), '(?:\\w+:)?extLst');
    for (const m of plain.matchAll(TRANSITIONAL_URI)) if (!/relationships\/vmlDrawing$/.test(m[1])) problems.push(`${name}: Transitional ${m[1]}`);
    if (name === mainName && !/<[^?!][^>]*\s(?:w:)?conformance="strict"/.test(xml.slice(0, 4000))) problems.push(`${name}: not marked conformance="strict"`);
    // A sheet's bare legacyDrawing is let be: Excel reads Strict's notes from it, not from a choice.
    if (/<(v|o|w10):\w/.test(plain) || /<(w:pict|w:object)\b/.test(plain)) problems.push(`${name}: VML outside a choice that requires it`);
    for (const m of plain.matchAll(/<(?:(\w+):)?(\w+)\b([^>]*)>/g)) {
      const attrs = PERCENT_ATTRS.get(m[2]);
      if (!attrs || m[1] === 'w') continue;
      for (const a of m[3].matchAll(/\s(\w+)="(-?\d+)"/g)) if (attrs.includes(a[1]) && !/^(c|xdr|wp|w|x)$/.test(m[1] || '')) problems.push(`${name}: ${m[1] || ''}:${m[2]} ${a[1]}="${a[2]}" in thousandths`);
    }
    if (/<w:ind\b[^>]*\sw:(left|right)=/.test(plain)) problems.push(`${name}: w:ind left or right`);
    if (/<w:(jc|lvlJc)\s+w:val="(left|right)"/.test(plain)) problems.push(`${name}: w:jc left or right`);
    if (/<w:(tblBorders|tcBorders|tblCellMar|tcMar)\b[^>]*>(?:(?!<\/w:\1>)[\s\S])*<w:(left|right)\b/.test(plain)) problems.push(`${name}: a table's left or right`);
    if (name === 'xl/styles.xml' && /<border\b[^>]*>(?:(?!<\/border>)[\s\S])*<(left|right)\b/.test(plain)) problems.push(`${name}: a border's left or right`);
    // Every prefix in use is ISO's, or one the part says may be ignored: in a part of ISO's
    // own, that is, not one of Office's (a chart's style, say) a Strict reader does not read.
    const root = /<(?:([\w.-]+):)?[\w.-]+[\s>]/.exec(xml.replace(/<\?[^>]*\?>/g, ''));
    const rootNs = new RegExp(`\\sxmlns${root?.[1] ? ':' + root[1] : ''}\\s*=\\s*"([^"]*)"`).exec(xml)?.[1] ?? '';
    if (!ISO.test(rootNs)) continue;
    const ns = new Map([...plain.matchAll(/\sxmlns:([\w.-]+)\s*=\s*"([^"]*)"/g)].reverse().map((m) => [m[1], m[2]]));
    const ignorable = new Set([...xml.matchAll(/mc:Ignorable="([^"]*)"/g)].flatMap((m) => m[1].split(/\s+/)));
    const used = new Set([...plain.matchAll(/<\/?([\w.-]+):[\w.-]+|\s([\w.-]+):[\w.-]+\s*=/g)].map((m) => m[1] || m[2]).filter((p) => p !== 'xmlns' && p !== 'xml' && p !== 'mc'));
    for (const p of used) {
      if (!ns.has(p)) problems.push(`${name}: prefix ${p} not declared`);
      else if (!ISO.test(ns.get(p)) && !ignorable.has(p)) problems.push(`${name}: ${p} (${ns.get(p)}) neither ISO's nor ignorable`);
    }
  }
  return [...new Set(problems)];
}

/* ── the tests ──────────────────────────────────────────────────────── */

test('a Strict package is known by its relationship to the main part, and read as Transitional with nothing of Strict left', { skip }, () => {
  for (const [, transitional, strict] of PAIRS) {
    assert.equal(OoxmlPackage.read(fs.readFileSync(transitional)).conformance, 'transitional');
    const pkg = OoxmlPackage.read(fs.readFileSync(strict));
    assert.equal(pkg.conformance, 'strict', basename(strict));
    for (const name of pkg.partNames().filter((n) => /\.(xml|rels)$/.test(n))) {
      const xml = pkg.text(name);
      assert.ok(!xml.includes('purl.oclc.org/ooxml'), `${name} still names a Strict namespace`);
      assert.ok(!/\st="d"/.test(xml), `${name} still holds an ISO date`);
      assert.ok(!/(?:\s|w:)conformance="strict"/.test(xml), `${name} still says it is Strict`);
    }
    assert.deepEqual(pkg.modifiedParts(), [], 'reading changes nothing');
  }
});

/** Each part's name and stored bytes, in order: what a save that changed nothing must give back. */
const stored = (bytes) => readZip(bytes).entries.map((e) => [e.name, e.crc, e.compressed.toString('base64')]);

test('a Strict file saved unchanged has every part as it was, and Office\'s Strict passes the structural check', { skip }, async () => {
  for (const [, , strict] of PAIRS) {
    const bytes = fs.readFileSync(strict);
    assert.deepEqual(stored(OoxmlPackage.read(bytes).write()), stored(bytes), basename(strict));
    assert.deepEqual(stored((await OoxmlPackage.readAsync(bytes, { inflate: 'stream' })).write()), stored(bytes), `${basename(strict)}, read as a browser reads it`);
    assert.deepEqual(strictProblems(bytes), [], `the check passes what ${basename(strict)}'s application wrote`);
  }
});

test('Strict made by Word, Excel and PowerPoint opens with what its Transitional twin has: text, tables, styles, pictures, charts, comments and dates', { skip }, () => {
  for (const [ext, transitional, strict] of PAIRS) {
    const a = opened(transitional);
    const b = opened(strict);
    assert.equal(b.conformance, 'strict');
    assert.equal(a.conformance, 'transitional');
    assert.deepEqual(b.seen, a.seen, `${ext}: the Strict file shows what the Transitional one does`);
  }
});

test('Excel\'s ISO 8601 dates are the serial numbers the 1900 date system gives them, and the 1904 system its own', () => {
  assert.equal(isoToSerial('2026-09-10'), 46275);
  assert.equal(isoToSerial('18:00:00.000'), 0.75);
  assert.equal(isoToSerial('2026-09-10T06:00:00'), 46275.25);
  assert.equal(isoToSerial('1900-01-01'), 1);
  assert.equal(isoToSerial('1900-02-28'), 59);
  assert.equal(isoToSerial('1900-03-01'), 61);
  assert.equal(isoToSerial('1900-03-01', { compat: false }), 61);
  assert.equal(isoToSerial('1900-01-01', { compat: false }), 2);
  assert.equal(isoToSerial('1904-01-02', { date1904: true }), 1);
  assert.equal(isoToSerial('not a date'), null);
});

test('the namespace map goes both ways, the VML drawing\'s relationship left Transitional as Excel leaves it', () => {
  for (const uri of [
    'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
    'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail',
  ]) {
    assert.ok(strictUri(uri).startsWith('http://purl.oclc.org/ooxml/'), uri);
    assert.equal(transitionalUri(strictUri(uri)), uri);
  }
  assert.equal(strictUri('http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'), 'http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties');
  assert.equal(strictUri('http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'), 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing');
  assert.equal(strictUri('http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'), 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties');
});

const EDITS = {
  xlsx: (id) => doc.apply({ id, ops: [{ op: 'sheet', name: 'Sales' }, { op: 'setCell', row: 30, col: 0, value: 'MARK' }] }),
  docx: (id) => doc.apply({ id, ops: [{ op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } }, { op: 'insertText', text: 'MARK ' }] }),
  pptx: (id) => doc.apply({ id, ops: [{ op: 'setNotes', slide: 0, text: 'MARK' }] }),
};

test('a Strict file edited and saved stays Strict, the parts nothing edited kept byte for byte, and it reopens with the edit', { skip }, () => {
  for (const [ext, , strict] of PAIRS) {
    const file = join(dir, `edited.${ext}`);
    const s = doc.open({ path: strict });
    EDITS[ext](s.id);
    const saved = doc.save({ id: s.id, path: file });
    doc.close({ id: s.id });
    assert.equal(saved.conformance, 'strict', ext);
    const bytes = fs.readFileSync(file);
    assert.deepEqual(strictProblems(bytes), [], `${ext}: still Strict`);
    const before = new Map(readZip(fs.readFileSync(strict)).entries.map((e) => [e.name, e.compressed]));
    const kept = readZip(bytes).entries.filter((e) => before.get(e.name)?.equals(e.compressed)).length;
    assert.ok(kept >= before.size / 2, `${ext}: most parts kept as they were, ${kept} of ${before.size}`);
    const again = doc.open({ path: file });
    try {
      assert.equal(again.conformance, 'strict');
      assert.ok(JSON.stringify(seen(again.id, again.kind, again.model)).includes('MARK'), `${ext}: the edit is there`);
    } finally {
      doc.close({ id: again.id });
    }
  }
});

test('a Strict file saved as the usual form is Transitional throughout, and shows what it did', { skip }, () => {
  for (const [ext, , strict] of PAIRS) {
    const file = join(dir, `usual.${ext}`);
    const s = doc.open({ path: strict });
    const saved = doc.save({ id: s.id, path: file, conformance: 'transitional' });
    doc.close({ id: s.id });
    assert.equal(saved.conformance, 'transitional');
    const parts = readZip(fs.readFileSync(file)).entries.filter((e) => /\.(xml|rels)$/.test(e.name));
    for (const e of parts) assert.ok(!e.data.toString('utf8').includes('purl.oclc.org/ooxml'), `${ext} ${e.name} names nothing of Strict`);
    assert.deepEqual(opened(file).seen, opened(strict).seen, `${ext}: as it was`);
  }
});

test('a Transitional file saved as Strict passes the structural check, and reopens showing what it did', { skip }, async () => {
  for (const [ext, transitional] of PAIRS) {
    const file = join(dir, `strict.${ext}`);
    const s = doc.open({ path: transitional });
    const saved = await doc.saveAsync({ id: s.id, path: file, conformance: 'strict' });
    doc.close({ id: s.id });
    assert.equal(saved.conformance, 'strict');
    assert.deepEqual(strictProblems(fs.readFileSync(file)), [], `${ext}: Strict as Office writes it`);
    const again = opened(file);
    assert.equal(again.conformance, 'strict');
    assert.deepEqual(again.seen, opened(transitional).seen, `${ext}: as it was`);
    fs.copyFileSync(file, join(dir, `kept-${ext}`));
  }
});
