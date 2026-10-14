// Redact and Inspect for an open document, workbook or deck: what the
// windows' Redact pane and File → Info → Check for Issues ask of the
// document service. Both work on the package as Save would write it.
//
// Redact never touches the open file: Apply writes a copy, then reads the
// copy back and refuses to call it done if any marked words are left in
// any part of it. Inspect removes in place, as Word's Document Inspector
// does, and the window takes the document as it now is.

import fs from 'node:fs';
import { OoxmlPackage } from '@rutba/ooxml';
import { redactPackage, partsHolding } from '@rutba/ooxml/redact';
import { inspectPackage, removeFromPackage } from '@rutba/ooxml/inspect';

const KIND_OF = { doc: 'document', sheet: 'workbook', deck: 'presentation' };

/** The package as Save would write it, unencrypted: a fresh copy, so nothing done to it reaches the open document. */
function packageOf(session) {
  return OoxmlPackage.read(Buffer.from(session.engine.save()));
}

/**
 * Redact: `scan` counts what the marks find, part by part; `apply` writes
 * the redacted copy to `path` and proves it by reading the copy back.
 */
export function redactAction(session, { action = 'scan', marks = [], pictures = {}, path: target = null }, { writeWhole }) {
  if (action === 'scan') {
    const r = redactPackage(packageOf(session), { marks, apply: false });
    return { total: r.total, parts: r.parts, words: [...new Set(r.words)] };
  }
  if (action !== 'apply') throw new Error('Redact does not know "' + action + '".');
  if (!target) throw new Error('The redacted copy needs a name to be saved under.');
  if (session.path && String(target).toLowerCase() === String(session.path).toLowerCase()) throw new Error('The redacted copy is saved as a new file: choose another name, so the original is kept.');
  const scan = redactPackage(packageOf(session), { marks, apply: false });
  const pkg = packageOf(session);
  const pics = {};
  for (const [part, bytes] of Object.entries(pictures || {})) {
    if (!pkg.has(part)) throw new Error('That picture is no longer in the file.');
    pics[part] = Buffer.from(bytes instanceof Uint8Array ? bytes : Object.values(bytes || {}));
  }
  const r = redactPackage(pkg, { marks, pictures: pics });
  const bytes = Buffer.from(pkg.write());
  writeWhole(target, bytes);
  // Read back from the disk: the words the scan found must be in no part of the copy.
  const words = [...new Set(scan.words.map((w) => w.trim()).filter((w) => w.length >= 2))];
  const left = words.length ? partsHolding(fs.readFileSync(target), words) : [];
  if (left.length) throw new Error('The redacted copy still holds marked words in ' + left.join(', ') + '; it was not finished.');
  return { path: target, total: r.total, parts: r.parts, checked: words.length > 0 };
}

/**
 * Inspect: `check` lists what the file holds, each kind with how much;
 * `remove` takes away the kinds named in `ids` from the open document,
 * which `reload` then reopens from the package as it now is.
 */
export function inspectAction(session, { action = 'check', ids = [] }, { reload }) {
  const kind = KIND_OF[session.kind];
  if (!kind) throw new Error('Inspect works on a document, a workbook or a presentation.');
  if (action === 'check') return { kind, items: inspectPackage(packageOf(session), kind) };
  if (action !== 'remove') throw new Error('Inspect does not know "' + action + '".');
  // Hidden rows and columns are deleted by the workbook's own Delete, as Excel's inspector deletes them:
  // the formulas round them follow, as they do when a row is deleted by hand.
  if (session.kind === 'sheet' && ids.includes('hiddenRowsColumns')) deleteHiddenRowsColumns(session.engine);
  // Revisions are accepted by the document's own Accept All first, as Word's inspector accepts them.
  if (session.kind === 'doc' && ids.includes('revisions') && typeof session.engine.acceptChanges === 'function') {
    try { session.engine.acceptChanges({ all: true }); } catch { /* the package pass below accepts what is left */ }
  }
  const pkg = packageOf(session);
  const removed = removeFromPackage(pkg, kind, ids);
  reload(session, Buffer.from(pkg.write()));
  return { kind, removed, items: inspectPackage(packageOf(session), kind) };
}

/** Every hidden row and column of every sheet deleted, from the last up, each sheet in its turn. */
function deleteHiddenRowsColumns(view) {
  const pkg = OoxmlPackage.read(Buffer.from(view.save()));
  const wb = pkg.text('xl/workbook.xml');
  const rels = new Map(pkg.rels('xl/workbook.xml').map((r) => [r.Id, OoxmlPackage.resolveTarget('xl/workbook.xml', r.Target)]));
  const was = view.activeSheet;
  for (const m of wb.matchAll(/<sheet\b([^>]*)\/>/g)) {
    const name = (/\bname="([^"]*)"/.exec(m[1])?.[1] ?? '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'");
    const part = rels.get(/\br:id="([^"]*)"/.exec(m[1])?.[1]);
    if (!part || !pkg.has(part)) continue;
    const xml = pkg.text(part);
    const rows = [...xml.matchAll(/<row\b([^>]*)>/g)].filter((r) => /\bhidden="(?:1|true)"/.test(r[1])).map((r) => Number(/\br="(\d+)"/.exec(r[1])?.[1]) - 1).filter((r) => r >= 0);
    const cols = [];
    for (const c of xml.matchAll(/<col\b([^>]*)\/>/g)) {
      if (!/\bhidden="(?:1|true)"/.test(c[1])) continue;
      const min = Number(/\bmin="(\d+)"/.exec(c[1])?.[1]);
      const max = Number(/\bmax="(\d+)"/.exec(c[1])?.[1] ?? min);
      for (let k = min; k <= max && k - min < 256; k++) cols.push(k - 1);
    }
    if (!rows.length && !cols.length) continue;
    view.selectSheet(name);
    // A merged cell across a hidden row refuses its deletion, as Excel's own Delete does: that row is left.
    for (const r of rows.sort((a, b) => b - a)) { try { view.deleteRows(r, 1); } catch { /* left */ } }
    for (const c of cols.sort((a, b) => b - a)) { try { view.deleteCols(c, 1); } catch { /* left */ } }
  }
  if (was) view.selectSheet(was);
}


/**
 * Review → Compare in Presentations: `compare` merges the reviewed deck at
 * `path` into the open one as changes; `accept` and `reject` decide the
 * changes named; `end` ends the review. The review lives with the open
 * deck; what is saved is the deck as the changes accepted have made it.
 */
export function deckReviewAction(session, { action = 'state', path: file = null, ids = [] }, { Deck, compare }) {
  if (session.kind !== 'deck') throw new Error('Compare works on a presentation.');
  const deck = session.engine;
  if (action === 'compare') {
    if (!file) throw new Error('Choose the reviewed presentation to compare with.');
    const reviewed = Deck.open(fs.readFileSync(file));
    const core = reviewed.pkg.has('docProps/core.xml') ? reviewed.pkg.text('docProps/core.xml') : '';
    const by = /<cp:lastModifiedBy>([^<]*)<\/cp:lastModifiedBy>/.exec(core)?.[1] || /<dc:creator>([^<]*)<\/dc:creator>/.exec(core)?.[1] || null;
    session.review = compare.beginReview(deck, reviewed, { DeckClass: Deck, reviewer: by, name: String(file).split(/[\/]/).pop() });
    return reviewSummary(session, compare);
  }
  if (!session.review) throw new Error('There is no review open: compare with a reviewed copy first.');
  if (action === 'accept' || action === 'reject') {
    compare.decideChanges(deck, session.review, ids, action === 'accept');
    session.dirty = true;
    session.version++;
    return reviewSummary(session, compare);
  }
  if (action === 'end') {
    session.review = null;
    return null;
  }
  return reviewSummary(session, compare);
}

/** What the window shows of a review: who, from which file, and each change as the deck finds it. */
export function reviewSummary(session, compare) {
  const r = session.review;
  if (!r) return null;
  return { reviewer: r.reviewer, name: r.name, changes: compare.reviewState(session.engine, r) };
}
