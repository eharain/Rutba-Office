// One recorder (A7): Video's Record Screen and Record Voice, an introduction's
// scenes and Presentations' Screen Recording all open the Recorder window in a
// mode, and what it records comes back here when it stops. The other half is
// studio/into.js.

import { useEffect, useRef } from 'react';

/**
 * Open the Recorder for this window while `active` is truthy.
 * @param {object} shell
 * @param {*} active  falsy when closed; a new value opens it again
 * @param {object} o
 * @param {'video'|'intro'|'slide'} o.into
 * @param {string} [o.title]  what is being recorded into, for its title bar
 * @param {string} [o.kind]  'voice': the microphone alone
 * @param {() => object} [o.prompter]  an introduction's: { title, script, speed }
 * @param {({ takes }) => void} o.onDone
 * @param {() => void} o.onCancel  closed with nothing recorded
 * @param {(takes) => void} [o.onProgress]  every few seconds while it records
 * @param {(speed: number) => void} [o.onSpeed]
 */
export function useRecorderWindow(shell, active, o) {
  const latest = useRef(o);
  latest.current = o;
  useEffect(() => {
    if (!active) return undefined;
    const channel = `rutba-recorder-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
    const bc = new BroadcastChannel(channel);
    let ended = false;
    const end = () => { ended = true; bc.close(); };
    bc.onmessage = ({ data: m = {} }) => {
      if (ended) return;
      const h = latest.current;
      if (m.type === 'ready') bc.postMessage({ type: 'brief', prompter: h.prompter?.() ?? null });
      else if (m.type === 'progress') h.onProgress?.(m.takes);
      else if (m.type === 'speed') h.onSpeed?.(m.value);
      else if (m.type === 'done') { bc.postMessage({ type: 'took' }); end(); h.onDone({ takes: m.takes }); }
      else if (m.type === 'cancel') { end(); h.onCancel(); }
    };
    // This window gone first: the Recorder keeps what it has and says where.
    const gone = () => { if (!ended) bc.postMessage({ type: 'gone' }); };
    window.addEventListener('pagehide', gone);
    const { into, title = '', kind = '' } = latest.current;
    shell.win.create({ app: 'recorder', query: { into, channel, title, kind } }).catch(() => { if (!ended) { end(); latest.current.onCancel(); } });
    return () => { window.removeEventListener('pagehide', gone); gone(); if (!ended) end(); };
    // Opened once for each `active`; what it calls is read when it is called.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shell, active]);
}
