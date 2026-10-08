'use strict';

/**
 * Right-to-left text put in the order a page shows it.
 *
 * A PDF draws glyphs left to right in the order it is given them, so a line
 * that mixes Hebrew or Arabic with Latin words and numbers has to be handed
 * over already reordered. This is the Unicode bidirectional algorithm for
 * one line of one paragraph, without the explicit embedding controls a
 * document from Word does not use in its text: characters are typed (strong
 * left, strong right, Arabic letter, European and Arabic numbers, their
 * separators and terminators, marks, neutrals), the weak types resolved
 * (W1–W7), the neutrals given the direction of what surrounds them (N1–N2),
 * levels assigned (I1–I2), trailing spaces put back at the paragraph's level
 * (L1), the runs reversed from the highest level down (L2), and brackets in
 * right-to-left runs mirrored (L4).
 */

function typeOf(cp) {
  if (cp >= 0x30 && cp <= 0x39) return 'EN';
  if (cp >= 0x06F0 && cp <= 0x06F9) return 'EN';
  if ((cp >= 0x0660 && cp <= 0x0669) || cp === 0x066B || cp === 0x066C) return 'AN';
  if (cp === 0x2B || cp === 0x2D) return 'ES';
  if (cp === 0x23 || cp === 0x24 || cp === 0x25 || cp === 0xB0 || cp === 0xA2 || cp === 0xA3 || cp === 0xA5 || cp === 0x20AC || cp === 0x2030 || cp === 0x066A) return 'ET';
  if (cp === 0x2C || cp === 0x2E || cp === 0x3A || cp === 0x2F || cp === 0xA0 || cp === 0x060C) return 'CS';
  if (cp === 0x20 || cp === 0x09 || cp === 0x3000 || (cp >= 0x2000 && cp <= 0x200A)) return 'WS';
  if ((cp >= 0x0591 && cp <= 0x05BD) || cp === 0x05BF || cp === 0x05C1 || cp === 0x05C2 || cp === 0x05C4 || cp === 0x05C5 || cp === 0x05C7) return 'NSM';
  if ((cp >= 0x064B && cp <= 0x065F) || cp === 0x0670 || (cp >= 0x06D6 && cp <= 0x06ED && cp !== 0x06DD && cp !== 0x06DE && cp !== 0x06E5 && cp !== 0x06E6 && cp !== 0x06E9) || (cp >= 0x0610 && cp <= 0x061A)) return 'NSM';
  if ((cp >= 0x0590 && cp <= 0x05FF) || (cp >= 0x07C0 && cp <= 0x085F) || (cp >= 0xFB1D && cp <= 0xFB4F)) return 'R';
  if ((cp >= 0x0600 && cp <= 0x07BF) || (cp >= 0x0860 && cp <= 0x08FF) || (cp >= 0xFB50 && cp <= 0xFDFF) || (cp >= 0xFE70 && cp <= 0xFEFF)) return 'AL';
  if (cp === 0x200F) return 'R';
  if (cp === 0x200E) return 'L';
  const ch = String.fromCodePoint(cp);
  if (/\p{L}|\p{M}/u.test(ch)) return 'L';
  if (/\p{Nd}/u.test(ch)) return 'L';
  return 'ON';
}

const MIRROR = { '(': ')', ')': '(', '[': ']', ']': '[', '{': '}', '}': '{', '<': '>', '>': '<', '«': '»', '»': '«', '‹': '›', '›': '‹' };

/** Whether a string needs reordering at all: anything right to left in it. */
function hasRtl(text) {
  return /[֐-ࣿיִ-﷿ﹰ-﻿‏]/.test(String(text ?? ''));
}

/**
 * The embedding level of each code point of `text`, a line of a paragraph
 * that runs right to left when `rtl`, or that takes its direction from its
 * first strong letter when `rtl` is null.
 */
function levelsOf(cps, rtl = null) {
  const n = cps.length;
  const types = cps.map(typeOf);
  let base = rtl;
  if (base === null || base === undefined) {
    const first = types.find((t) => t === 'L' || t === 'R' || t === 'AL');
    base = first === 'R' || first === 'AL';
  }
  const baseLevel = base ? 1 : 0;
  const sos = base ? 'R' : 'L';
  // W1: a mark takes the type of what it sits on.
  for (let i = 0; i < n; i++) if (types[i] === 'NSM') types[i] = i === 0 ? sos : types[i - 1];
  // W2: a European number after an Arabic letter is an Arabic number.
  let lastStrong = sos;
  for (let i = 0; i < n; i++) {
    const t = types[i];
    if (t === 'L' || t === 'R' || t === 'AL') lastStrong = t;
    else if (t === 'EN' && lastStrong === 'AL') types[i] = 'AN';
  }
  // W3: an Arabic letter is a right-to-left letter.
  for (let i = 0; i < n; i++) if (types[i] === 'AL') types[i] = 'R';
  // W4: one separator between two numbers of a kind joins them.
  for (let i = 1; i < n - 1; i++) {
    if (types[i] === 'ES' && types[i - 1] === 'EN' && types[i + 1] === 'EN') types[i] = 'EN';
    else if (types[i] === 'CS' && types[i - 1] === 'EN' && types[i + 1] === 'EN') types[i] = 'EN';
    else if (types[i] === 'CS' && types[i - 1] === 'AN' && types[i + 1] === 'AN') types[i] = 'AN';
  }
  // W5: terminators next to a European number are part of it.
  for (let i = 0; i < n; i++) {
    if (types[i] !== 'ET') continue;
    let j = i;
    while (j < n && types[j] === 'ET') j++;
    const touches = (i > 0 && types[i - 1] === 'EN') || (j < n && types[j] === 'EN');
    if (touches) for (let k = i; k < j; k++) types[k] = 'EN';
    i = j - 1;
  }
  // W6: any other separator or terminator is neutral.
  for (let i = 0; i < n; i++) if (types[i] === 'ES' || types[i] === 'ET' || types[i] === 'CS') types[i] = 'ON';
  // W7: a European number after a left-to-right letter is left to right.
  lastStrong = sos;
  for (let i = 0; i < n; i++) {
    const t = types[i];
    if (t === 'L' || t === 'R') lastStrong = t;
    else if (t === 'EN' && lastStrong === 'L') types[i] = 'L';
  }
  // N1, N2: neutrals between two of one direction take it; otherwise the paragraph's.
  const dir = (t) => (t === 'L' ? 'L' : t === 'R' || t === 'EN' || t === 'AN' ? 'R' : null);
  for (let i = 0; i < n; i++) {
    if (types[i] !== 'ON' && types[i] !== 'WS') continue;
    let j = i;
    while (j < n && (types[j] === 'ON' || types[j] === 'WS')) j++;
    const before = i > 0 ? dir(types[i - 1]) : sos;
    const after = j < n ? dir(types[j]) : sos;
    const take = before === after ? before : sos;
    for (let k = i; k < j; k++) types[k] = take === 'L' ? 'NL' : 'NR';
    i = j - 1;
  }
  // I1, I2.
  const levels = types.map((t) => {
    if (baseLevel === 0) return t === 'R' || t === 'NR' ? 1 : t === 'EN' || t === 'AN' ? 2 : 0;
    return t === 'L' || t === 'NL' || t === 'EN' || t === 'AN' ? 2 : 1;
  });
  // L1: spaces at the end of the line go back to the paragraph's level.
  for (let i = n - 1; i >= 0 && typeOf(cps[i]) === 'WS'; i--) levels[i] = baseLevel;
  return { levels, baseLevel };
}

/** The visual order of the indices 0..levels.length-1 (L2). */
function visualOrder(levels) {
  const order = levels.map((_, i) => i);
  const top = Math.max(0, ...levels);
  const lowestOdd = levels.reduce((m, l) => (l % 2 ? Math.min(m, l) : m), Infinity);
  for (let level = top; level >= lowestOdd && level > 0; level--) {
    for (let i = 0; i < order.length; i++) {
      if (levels[order[i]] < level) continue;
      let j = i;
      while (j < order.length && levels[order[j]] >= level) j++;
      order.splice(i, j - i, ...order.slice(i, j).reverse());
      i = j - 1;
    }
  }
  return order;
}

/**
 * Pieces of a line in the order they are drawn, left to right. `pieces` are
 * the line's runs in reading order, each `{ text, ... }`; each piece that
 * comes back is a stretch of one run at one level, its text in drawing
 * order with right-to-left brackets mirrored, carrying the run's own fields.
 */
function visualPieces(pieces, { rtl = null } = {}) {
  const cps = [];
  const owner = [];
  pieces.forEach((p, k) => { for (const ch of String(p.text ?? '')) { cps.push(ch.codePointAt(0)); owner.push(k); } });
  if (!cps.length) return [];
  const { levels } = levelsOf(cps, rtl);
  const order = visualOrder(levels);
  const out = [];
  for (const i of order) {
    let ch = String.fromCodePoint(cps[i]);
    if (levels[i] % 2 && MIRROR[ch]) ch = MIRROR[ch];
    const last = out[out.length - 1];
    if (last && last._owner === owner[i] && last._level === levels[i]) last.text += ch;
    else out.push({ ...pieces[owner[i]], text: ch, _owner: owner[i], _level: levels[i] });
  }
  return out.map(({ _owner, _level, ...p }) => ({ ...p, rtl: _level % 2 === 1 }));
}

/** One line of text in drawing order. */
function visualText(text, { rtl = null } = {}) {
  return visualPieces([{ text }], { rtl }).map((p) => p.text).join('');
}

module.exports = { visualPieces, visualText, hasRtl, levelsOf, visualOrder };
