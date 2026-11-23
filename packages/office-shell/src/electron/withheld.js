// Windows kept out of capture (Office 2.27.0, the screen review of 22 Nov
// 2026). A Talk window that holds a call or a meeting says so
// (capture.withhold), and for as long as it does it is no source: it is not
// in the list that Insert → Screenshot, the Recorder and Live Broadcast
// choose from, it cannot be chosen for a recording, and no picture is taken
// of it. With `protect` (a call, or a meeting not set up to be streamed or
// recorded) its contents are also kept out of a capture of the whole
// screen, where the system can do that.
//
// People in a meeting are put on air through Guests in Live Broadcast and
// the Recorder, where everyone is told before they join.
//
// This is what Rutba Office itself does with its own windows. It does not
// stop a camera pointed at the screen, or another program on a system that
// does not heed the protection, and nothing here says it does.
//
// No Electron in this file: the windows are handed in, so the rules are
// tested without one.

export function createWithheld() {
  /** window id -> { win, source, protect } */
  const held = new Map();
  /** Windows whose closing is already watched, so a window that holds many calls is watched once. */
  const watched = new WeakSet();

  const gone = (win) => !win || Boolean(win.isDestroyed?.());
  const sourceOf = (win) => { try { return String(win.getMediaSourceId()); } catch { return null; } };
  const guard = (win, on) => { try { win.setContentProtection?.(Boolean(on)); } catch { /* not on this system */ } };

  const drop = (id) => {
    const h = held.get(id);
    if (!h) return;
    held.delete(id);
    if (h.protect && !gone(h.win)) guard(h.win, false);
  };
  /** Those still open. */
  const all = () => {
    for (const [id, h] of [...held]) if (gone(h.win)) held.delete(id);
    return [...held.values()];
  };

  return {
    /**
     * A window says it holds a call or a meeting (`on`), and whether its
     * contents are to be kept out of a capture of the screen (`protect`);
     * or that it no longer does. A window speaks for itself only.
     * @returns {{ on: boolean, protected: boolean }}
     */
    set(win, { on = false, protect = false } = {}) {
      if (gone(win)) return { on: false, protected: false };
      const id = win.id;
      if (!on) {
        drop(id);
        return { on: false, protected: false };
      }
      const want = Boolean(protect);
      const was = held.get(id);
      if (was) {
        if (was.protect !== want) {
          was.protect = want;
          guard(win, want);
        }
        return { on: true, protected: want };
      }
      held.set(id, { win, source: sourceOf(win), protect: want });
      if (want) guard(win, true);
      if (!watched.has(win)) {
        watched.add(win);
        win.on?.('closed', () => held.delete(id));
      }
      return { on: true, protected: want };
    },

    /** Whether a source (a window's media source id) is one kept out just now. */
    has(sourceId) {
      const id = String(sourceId ?? '');
      return Boolean(id) && all().some((h) => h.source === id);
    },

    /** How many windows are kept out just now. */
    count() { return all().length; },
  };
}

/** One for the run: the shell's capture answers from it (ipc.js). */
export const withheld = createWithheld();
