/**
 * Positions in a document.
 *
 * A caret is `{ block, offset }` where offset counts CHARACTERS in the
 * paragraph's plain text — not runs, not XML. That is the coordinate a user
 * thinks in ("third character of the second line") and the one a browser
 * selection reports, so it is the coordinate the model speaks.
 *
 * Runs are an implementation detail of formatting that the position layer
 * translates to and from. A paragraph reading "Dear **Sir** or Madam" is three
 * runs, but offset 7 is offset 7 regardless — and typing there has to find the
 * bold run, splice into it, and leave the other two exactly as they were.
 */

export const runsText = (runs) => runs.map((r) => r.text).join('');

/**
 * Character offset -> which run, and where inside it.
 *
 * `bias` decides the boundary case. At the seam between two runs, offset 5 is
 * both "end of run 0" and "start of run 1". Typing should inherit the formatting
 * of what came BEFORE the caret (type after a bold word and you get bold), so
 * insertion biases left; deleting forward biases right, because it consumes the
 * character that follows.
 */
export function locate(runs, offset, bias = 'left') {
  if (!runs.length) return { runIndex: 0, runOffset: 0, atEnd: true };
  let seen = 0;
  for (let i = 0; i < runs.length; i++) {
    const len = runs[i].text.length;
    const end = seen + len;
    if (offset < end || (offset === end && (bias === 'left' || i === runs.length - 1))) {
      return { runIndex: i, runOffset: offset - seen, atEnd: offset === end };
    }
    seen = end;
  }
  const last = runs.length - 1;
  return { runIndex: last, runOffset: runs[last].text.length, atEnd: true };
}

/** Clamp a position into a document's real extent. */
export function clampPosition(blocks, { block, offset }) {
  // A non-finite index would survive Math.min/Math.max as NaN and then read as
  // paragraph zero further down, which turns a lost caret into an edit in the
  // wrong paragraph. Clamp to the start explicitly instead: a defined floor, not
  // a guess about what the caller meant.
  const requested = Number.isFinite(block) ? block : 0;
  const b = Math.max(0, Math.min(blocks.length - 1, requested));
  const len = blocks[b] ? blocks[b].text.length : 0;
  const at = Number.isFinite(offset) ? offset : 0;
  return { block: b, offset: Math.max(0, Math.min(len, at)) };
}

export const samePosition = (a, b) => a.block === b.block && a.offset === b.offset;

/** Document order: which of two positions comes first. */
export function comparePositions(a, b) {
  if (a.block !== b.block) return a.block < b.block ? -1 : 1;
  if (a.offset !== b.offset) return a.offset < b.offset ? -1 : 1;
  return 0;
}

/** Anchor/focus in document order, so edits never have to think about direction. */
export function orderedRange(anchor, focus) {
  return comparePositions(anchor, focus) <= 0 ? { from: anchor, to: focus } : { from: focus, to: anchor };
}

/**
 * Slice a run list by character range, preserving each run's properties.
 * Used by delete, by copy, and by formatting — all of which need "the runs that
 * cover exactly this span, split at the boundaries".
 */
export function sliceRuns(runs, from, to) {
  const out = [];
  let seen = 0;
  for (const run of runs) {
    const start = seen;
    const end = seen + run.text.length;
    seen = end;
    if (end <= from || start >= to) continue;
    const text = run.text.slice(Math.max(0, from - start), Math.min(run.text.length, to - start));
    if (text !== '') out.push({ ...run, text });
  }
  return out;
}

/** Everything outside [from, to), as one run list. */
export function removeRange(runs, from, to) {
  return [...sliceRuns(runs, 0, from), ...sliceRuns(runs, to, Infinity)];
}

/**
 * Merge adjacent runs that carry identical properties.
 *
 * Without this, every keystroke fragments a paragraph a little further — type
 * ten characters and you have ten runs where one would do. The file still opens
 * correctly, but it grows, diffs become noise, and Word shows a mess in the
 * style inspector.
 */
export function coalesce(runs) {
  const out = [];
  for (const run of runs) {
    if (run.text === '') continue;
    const last = out[out.length - 1];
    // A link is part of a run's identity: merging a linked run into a plain
    // neighbour would stretch or swallow the link.
    if (last && last.rPr === run.rPr && (last.link ?? null) === (run.link ?? null)) last.text += run.text;
    else out.push({ ...run });
  }
  return out;
}
