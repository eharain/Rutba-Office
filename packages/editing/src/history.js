/**
 * Undo, and the rules that make it feel right.
 *
 * Undo is not a feature users ask for; it is one they assume, and its absence
 * is felt in the first minute of use. It is also the single edit operation whose
 * *granularity* is more important than its correctness: an undo that steps back
 * one character at a time is technically correct and unusable.
 *
 * So the interesting part here is COALESCING. Typing a word is one undo. Typing,
 * then clicking elsewhere, then typing again is two. Deleting is its own group,
 * because undoing a deletion and undoing an insertion are different intentions
 * and users do not expect them to merge. Word and Excel both behave this way and
 * users have thirty years of muscle memory for it.
 *
 * The history stores SNAPSHOTS taken BEFORE each edit, not a log of operations.
 * Inverse operations are the more elegant design and the wrong one here: our
 * model is XML held verbatim, so a snapshot is a string we already have, while an
 * inverse-operation log would need every edit to know how to undo itself —
 * exactly the kind of code that goes subtly wrong and corrupts a customer file.
 *
 * Nothing in this file knows what a document or a sheet is. A `state` is opaque:
 * the caller decides what is worth saving, and the caller restores it.
 */

/** How much history to keep. Old entries fall off the bottom. */
export const DEFAULT_LIMIT = 100;

/**
 * How long a run of same-kind edits keeps merging into one undo step.
 *
 * A pause means the user finished a thought, so the next keystroke starts a new
 * undo step even though nothing else changed. Without this, walking away
 * mid-sentence and coming back makes the whole paragraph one undo.
 */
export const DEFAULT_COALESCE_MS = 1200;

export class History {
  /**
   * @param {object}   [opts]
   * @param {number}   [opts.limit]       entries to keep (default 100)
   * @param {number}   [opts.coalesceMs]  merge window for same-group edits
   * @param {Function} [opts.now]         clock, injectable so tests are not timed
   */
  constructor({ limit = DEFAULT_LIMIT, coalesceMs = DEFAULT_COALESCE_MS, now = () => Date.now() } = {}) {
    this.limit = Math.max(1, limit);
    this.coalesceMs = Math.max(0, coalesceMs);
    this.now = now;
    /** @type {Array<{label: string, group: string|null, state: any, meta: any, at: number}>} */
    this.past = [];
    /** @type {Array<{label: string, group: string|null, state: any, meta: any, at: number}>} */
    this.future = [];
    this._lastGroup = null;
  }

  get canUndo() { return this.past.length > 0; }
  get canRedo() { return this.future.length > 0; }
  get depth() { return this.past.length; }

  /** What undo would do next, for a menu label: "Undo typing". */
  get undoLabel() { return this.past.at(-1)?.label ?? null; }
  get redoLabel() { return this.future.at(-1)?.label ?? null; }

  /**
   * Record the state as it was BEFORE an edit.
   *
   * @param {object} entry
   * @param {any}    entry.state   opaque; whatever restores the pre-edit state
   * @param {string} [entry.label] shown in a menu: 'typing', 'delete', 'paste'
   * @param {string} [entry.group] consecutive records sharing a non-null group
   *   merge into one undo step. Pass null — or call `break()` — to force a new
   *   one. A caret move should break the group; that is what makes "type, click
   *   away, type" two undos rather than one.
   * @param {any}    [entry.meta]  restored alongside the state — the selection,
   *   normally, because undo that leaves the caret elsewhere is disorienting.
   * @returns {boolean} true if a new entry was pushed, false if it coalesced
   */
  record({ state, label = 'edit', group = null, meta = null } = {}) {
    // Any new edit invalidates the redo branch. This is the standard linear
    // model; a tree would be more powerful and nobody would understand it.
    this.future = [];

    const at = this.now();
    const previous = this.past.at(-1);
    const merges = Boolean(
      group
      && previous
      && previous.group === group
      && at - previous.at <= this.coalesceMs,
    );

    if (merges) {
      // The earlier snapshot already represents "before the whole run", so it
      // must NOT be overwritten. Only the clock moves, so the window slides
      // with continued typing rather than expiring mid-word.
      previous.at = at;
      this._lastGroup = group;
      return false;
    }

    this.past.push({ label, group, state, meta, at });
    if (this.past.length > this.limit) this.past.shift();
    this._lastGroup = group;
    return true;
  }

  /**
   * End the current coalescing run.
   *
   * Call this when something happens that is not an edit but changes intent — a
   * click, an arrow key, a sheet change. The next edit then starts its own undo
   * step even if it is the same kind as the last.
   */
  break() { this._lastGroup = null; if (this.past.length) this.past.at(-1).group = null; return this; }

  /**
   * Step back. Give the CURRENT state so it can be redone.
   * @returns {{state: any, meta: any, label: string}|null}
   */
  undo(currentState, currentMeta = null) {
    const entry = this.past.pop();
    if (!entry) return null;
    this.future.push({ ...entry, state: currentState, meta: currentMeta });
    this._lastGroup = null;
    return { state: entry.state, meta: entry.meta, label: entry.label };
  }

  /** Step forward again. */
  redo(currentState, currentMeta = null) {
    const entry = this.future.pop();
    if (!entry) return null;
    this.past.push({ ...entry, state: currentState, meta: currentMeta, group: null });
    this._lastGroup = null;
    return { state: entry.state, meta: entry.meta, label: entry.label };
  }

  /** Forget everything — after a save-as, or when a document is reloaded. */
  clear() {
    this.past = [];
    this.future = [];
    this._lastGroup = null;
    return this;
  }

  /** Diagnostics, and what a UI shows on a menu. */
  describe() {
    return {
      canUndo: this.canUndo,
      canRedo: this.canRedo,
      undoLabel: this.undoLabel,
      redoLabel: this.redoLabel,
      depth: this.past.length,
      redoDepth: this.future.length,
    };
  }
}

export default History;
