// The Recorder, driven by the AI tools (apps/desktop/main/mcp/reach-tools.js,
// recorder_*; D-2026-10-30-06). The recording itself is the Recorder
// window's, as when a person presses Record: this only finds that window,
// opens one when there is none, hands it a request as a capture:command and
// waits for its answer (capture.remoteReply). Go Live is never asked for
// here: it publishes, so it stays a person's press.
//
// One per run of Rutba Office: the window manager is handed in when the
// shell's services are built (ipc.js), the AI tools ask it from their pack.

import { CHANNEL_PREFIX } from '../contract.js';

const WAIT = { ready: 20000, sources: 15000, start: 30000, stop: 60000, state: 5000 };

export function createRecorderRemote() {
  let windows = null;
  let next = 1;
  const pending = new Map();
  // Recorder windows that have said they listen, by window id.
  const ready = new Set();
  const readyWaiters = [];

  const recorderWindow = () => {
    const w = windows?.findByApp('recorder');
    return w && !w.isDestroyed() ? w : null;
  };

  /** The Recorder window, opened when there is none, once it listens. */
  async function windowFor({ open }) {
    let w = recorderWindow();
    if (!w) {
      if (!open) return null;
      if (!windows) throw new Error('The Recorder cannot be opened yet: Rutba Office is still starting.');
      windows.open({ app: 'recorder' });
      w = recorderWindow();
      if (!w) throw new Error('The Recorder did not open.');
    }
    if (ready.has(w.id)) return w;
    await new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error('The Recorder opened but did not answer.')), WAIT.ready);
      readyWaiters.push({ id: w.id, done: () => { clearTimeout(timer); resolve(); } });
    });
    return w;
  }

  return {
    attach(manager) { windows = manager; },

    /** The window said it listens (or, gone, no longer does). */
    listening(win, on) {
      if (!win) return;
      if (!on) { ready.delete(win.id); return; }
      ready.add(win.id);
      win.once('closed', () => ready.delete(win.id));
      for (let i = readyWaiters.length - 1; i >= 0; i--) {
        if (readyWaiters[i].id === win.id) { readyWaiters[i].done(); readyWaiters.splice(i, 1); }
      }
    },

    /** The window's answer to request `reply`. */
    answer({ reply, result = null, error = null }) {
      const p = pending.get(reply);
      if (!p) return false;
      pending.delete(reply);
      clearTimeout(p.timer);
      if (error) p.reject(Object.assign(new Error(String(error)), { refusal: true }));
      else p.resolve(result);
      return true;
    },

    /**
     * `op` asked of the Recorder window: 'sources', 'state', 'start' or
     * 'stop'. 'sources' and 'start' open a window; the others, with none
     * open, answer as idle.
     */
    async ask(op, args = {}) {
      const w = await windowFor({ open: op === 'start' || op === 'sources' });
      if (!w) return op === 'state' ? { phase: 'closed' } : Promise.reject(Object.assign(new Error('The Recorder is not open, so nothing is recording.'), { refusal: true }));
      const reply = next++;
      return new Promise((resolve, reject) => {
        const timer = setTimeout(() => { pending.delete(reply); reject(new Error(`The Recorder did not answer ${op} in time.`)); }, WAIT[op] || WAIT.state);
        pending.set(reply, { resolve, reject, timer });
        w.webContents.send(`${CHANNEL_PREFIX}/event/capture:command`, { command: 'remote', op, args, reply });
      });
    },
  };
}

/** The one this run of Rutba Office has. */
export const recorderRemote = createRecorderRemote();
