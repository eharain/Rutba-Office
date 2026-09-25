// Review → Spelling: the pass.
//
// A document is a list of segments — a paragraph, a cell, a run on a slide,
// a footnote — in the order the pass reads them (word.js, sheet.js and
// deck.js make the lists). The pass starts at the caret, reads to the end,
// goes round to the beginning and stops where it started, as Office's does.
// Each step answers the next word that is in neither the dictionary, the
// person's own dictionary nor the words ignored for this document, with the
// sentence around it; the window shows it and the person decides.

import { tokenize, matchCase } from './tokenize.js';

/** A word the person's own dictionary or Ignore All accepts: as written, or lower case for any capitalisation. */
export function acceptedBy(list, word) {
  if (!list || !word) return false;
  const has = (w) => (typeof list.has === 'function' ? list.has(w) : list.includes(w));
  return has(word) || has(word.toLowerCase());
}

/** The index of the segment a position is in; a key that has gone falls back to `fallback`. */
function indexOf(segments, key, fallback) {
  const i = segments.findIndex((s) => s.key === key);
  return i < 0 ? fallback : i;
}

/** The sentence round a word, cut to a readable length for the pane. */
export function contextOf(text, offset, length, room = 60) {
  const s = String(text);
  const startAt = Math.max(0, offset - room);
  const endAt = Math.min(s.length, offset + length + room);
  let before = s.slice(startAt, offset);
  let after = s.slice(offset + length, endAt);
  if (startAt > 0) before = `…${before.replace(/^\S*\s/, '')}`;
  if (endAt < s.length) after = `${after.replace(/\s\S*$/, '')}…`;
  return { before: before.replace(/\s+/g, ' '), word: s.slice(offset, offset + length), after: after.replace(/\s+/g, ' ') };
}

/**
 * The next misspelling from `from`, in pass order.
 *
 *   segments  the document's text, in order
 *   from      `{ key, offset }` — read from here (a word starting at it counts)
 *   stop      `{ key, offset }` — where the pass began; it ends on reaching it again
 *   wrapped   whether the pass has already gone round from the end
 *   misspelt  async (words) => Set of those the dictionary does not know
 *   accepted  (word) => true for the person's own words and those ignored
 *
 * Answers `{ found, next, wrapped }` or `{ done: true, wrapped }`.
 */
export async function nextMisspelling({ segments, from, stop, wrapped = false, misspelt, accepted = () => false, options = {} }) {
  if (!segments.length) return { done: true, wrapped: true };
  const stopIndex = stop ? indexOf(segments, stop.key, segments.length) : segments.length;
  let i = indexOf(segments, from?.key, 0);
  let offset = from?.offset ?? 0;
  let round = wrapped;

  // The dictionary is asked about the words in slices, in the order they
  // come, so a first misspelling near the caret answers quickly.
  const SLICE = 400;
  for (let guard = 0; guard < segments.length * 2 + 2; guard++) {
    if (i >= segments.length) {
      if (round) return { done: true, wrapped: true };
      round = true;
      i = 0;
      offset = 0;
    }
    const upto = Math.min(segments.length, i + SLICE);
    const batch = [];
    for (let k = i; k < upto; k++) {
      const words = tokenize(segments[k].text, options).filter((t) => (k === i ? t.offset >= offset : true));
      batch.push({ k, words });
    }
    const unknown = await misspelt([...new Set(batch.flatMap((b) => b.words.map((t) => t.word)))]);
    for (const { k, words } of batch) {
      for (const t of words) {
        if (round && (k > stopIndex || (k === stopIndex && t.offset >= (stop?.offset ?? 0)))) return { done: true, wrapped: true };
        if (!unknown.has(t.word) || accepted(t.word)) continue;
        const seg = segments[k];
        return {
          found: { key: seg.key, where: seg.where, word: seg.text.slice(t.offset, t.offset + t.length), offset: t.offset, length: t.length, context: contextOf(seg.text, t.offset, t.length) },
          next: { key: seg.key, offset: t.offset + t.length },
          wrapped: round,
        };
      }
      if (round && k >= stopIndex) return { done: true, wrapped: true };
    }
    i = upto;
    offset = 0;
  }
  return { done: true, wrapped: true };
}

/**
 * Change All: every occurrence of `word` in the segments, as edits
 * `{ key, from, to, text }` — the same word in any capitalisation, each given
 * the replacement in its own ("Recieve" becomes "Receive").
 */
export function changeAllEdits(segments, word, replacement, options = {}) {
  const target = String(word).replace(/’/g, "'").toLowerCase();
  const edits = [];
  for (const seg of segments) {
    for (const t of tokenize(seg.text, { ...options, ignoreUppercase: false, ignoreNumbers: false })) {
      if (t.word.toLowerCase() !== target) continue;
      edits.push({ key: seg.key, from: t.offset, to: t.offset + t.length, text: matchCase(seg.text.slice(t.offset, t.offset + t.length), replacement) });
    }
  }
  return edits;
}

/**
 * A personal dictionary as a Hunspell .dic file: the count on the first
 * line, then a word to a line — and back, forgiving a list with no count,
 * blank lines, comments, flags after a slash and a byte-order mark.
 */
export function toDic(words) {
  const list = [...new Set((words || []).map((w) => String(w).trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  return `${list.length}\n${list.join('\n')}\n`;
}

export function fromDic(text) {
  const lines = String(text || '').replace(/^﻿/, '').split(/\r?\n/).map((l) => l.trim());
  if (lines.length && /^\d+$/.test(lines[0])) lines.shift();
  const out = [];
  for (const line of lines) {
    if (!line || line.startsWith('#')) continue;
    const word = line.split('/')[0].trim();
    if (word && !/\s/.test(word)) out.push(word);
  }
  return [...new Set(out)];
}
