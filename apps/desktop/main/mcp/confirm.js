// What the AI tools want to do that leaves this computer or tells another
// person, shown in Rutba Office first (owner, 2026-10-26, D-2026-10-26-05:
// "allow sending too, but confirm each message first with choice to create
// drafts only"): a message sent, an invitation or its answer, a contact
// written to a server, a change written to a database engine.
//
// A request waits here with what it would do (`run.act`), what keeping it as
// a draft does instead (`run.draft`, when there is a draft to keep), and
// Cancel, which does nothing. Every window shows the requests waiting (the
// renderer's AiConfirm) and the person's press answers one: only then is
// anything done, and only what was pressed. The tool that asked waits a
// while for the answer and, if the person has not pressed anything yet, says
// the request is waiting with its id (confirmation_answer asks again). A
// request nobody answers is withdrawn after half an hour, nothing done.
//
// Nothing here needs Electron: the service hands in how windows hear of the
// list (`tell`) and are brought to the front (`attend`).

import crypto from 'node:crypto';

export const WAIT_MS = 50_000; // under the minute many AI apps give a call
export const EXPIRE_MS = 30 * 60_000;
const KEEP_ANSWERED = 50;

/**
 * @param {object} o
 * @param {(pending: object[]) => void} [o.tell]  every window hears the list of requests waiting
 * @param {() => void} [o.attend]                  a window brought to the front for a new request
 * @param {() => number} [o.now]
 */
export function createConfirmations({ tell = () => {}, attend = () => {}, now = () => Date.now(), waitMs = WAIT_MS, expireMs = EXPIRE_MS } = {}) {
  const waiting = new Map(); // id -> { request, run, settle, at, timer }
  const answered = new Map(); // id -> { choice, outcome, error }

  /** What the windows draw: the request's words, never how it is done. */
  const shown = () => [...waiting.entries()].map(([id, w]) => ({ id, at: new Date(w.at).toISOString(), ...w.request, draft: Boolean(w.run.draft) }));
  const changed = () => tell(shown());

  const remember = (id, entry) => {
    answered.set(id, entry);
    while (answered.size > KEEP_ANSWERED) answered.delete(answered.keys().next().value);
  };

  /**
   * The person's press: 'act', 'draft' or 'cancel'. Answers what was done,
   * in a sentence; a request already answered, or gone, is refused.
   */
  async function answer(id, choice) {
    const w = waiting.get(id);
    if (!w) throw new Error('That request is no longer waiting.');
    if (!['act', 'draft', 'cancel'].includes(choice) || (choice === 'draft' && !w.run.draft)) throw new Error('That is not one of the choices.');
    waiting.delete(id);
    clearTimeout(w.timer);
    changed();
    let entry;
    try {
      const outcome = choice === 'cancel' ? w.request.cancelled || 'Cancelled in Rutba Office: nothing was done.' : await w.run[choice]();
      entry = { choice, outcome };
    } catch (err) {
      entry = { choice, error: err?.message || String(err) };
    }
    remember(id, entry);
    w.settle(entry);
    return entry;
  }

  return {
    /**
     * A request shown to the person. `request`: { kind, title, thing, action
     * (the button's word), lines: [[label, value]], body?, cancelled? };
     * `run`: { act, draft? }, each answering a sentence of what was done.
     * Answers { id, choice, outcome | error } once pressed, or { id,
     * waiting: true } when the person has not pressed anything within the wait.
     */
    ask(request, run) {
      const id = crypto.randomUUID();
      let settle;
      const done = new Promise((resolve) => { settle = resolve; });
      const timer = setTimeout(() => {
        if (!waiting.has(id)) return;
        waiting.delete(id);
        changed();
        const entry = { choice: 'expired', outcome: 'Nobody answered in Rutba Office within half an hour, so the request was withdrawn and nothing was done.' };
        remember(id, entry);
        settle(entry);
      }, expireMs);
      timer.unref?.();
      waiting.set(id, { request, run, settle, at: now(), timer });
      changed();
      attend();
      const wait = new Promise((resolve) => {
        const t = setTimeout(() => resolve({ waiting: true }), waitMs);
        done.then((entry) => { clearTimeout(t); resolve(entry); });
      });
      return wait.then((entry) => ({ id, ...entry }));
    },
    answer,
    pending: shown,
    /** What became of a request: { choice, outcome | error }, { waiting: true }, or null when it is not known. */
    status(id) {
      if (waiting.has(id)) return { waiting: true };
      return answered.get(id) || null;
    },
    /** Every request withdrawn, nothing done: the app is closing. */
    clear() {
      for (const [id, w] of waiting) {
        clearTimeout(w.timer);
        const entry = { choice: 'cancel', outcome: 'Rutba Office closed before it was answered: nothing was done.' };
        remember(id, entry);
        w.settle(entry);
      }
      waiting.clear();
      changed();
    },
  };
}
