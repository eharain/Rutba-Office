// Talk, as the notifier hears of it (watch.js, D-2026-11-22-23): a call
// coming in, a call missed, a message.
//
// Nothing here is Talk's own. It listens to what main/talk.js already tells
// every window ('talk:state' and 'talk:event'), and to what a window sends
// back (an answer, a call turned away), and reports to the notifier. Talk
// shows no notification itself.
//
// Where Talk says how many messages are unread and how many calls were
// missed (its state's `unreadTotal` and `missed`), those are the numbers by
// the clock; a Talk that does not say is counted here, from what arrived.
//
// A call rings in the window that heard it. A Talk window opened while a
// call is still ringing (from its notification, or by hand) heard nothing,
// so it is told of the call once, as it starts.

/** How long a call is taken to ring when nothing says it stopped. */
export const RING_MS = 90_000;

/**
 * @param {object} o
 * @param {object} o.notifier
 * @param {(event: string, fn: Function) => Function} o.on   listens to what the main process tells its windows
 * @param {(app: string) => boolean} [o.front]   whether a window of that app is in front now
 * @param {(msg: object) => any} [o.send]   sends to the introducer, as Talk's own service does (a call turned away from its notification)
 * @param {(win: any, msg: object) => void} [o.tellWindow]   one window told of a 'talk:event'
 * @param {() => any[]} [o.talkWindows]   Talk's windows open now: they hear a call as it comes
 */
export function watchTalk({ notifier, on, front = () => false, send = null, tellWindow = () => {}, talkWindows = () => [], setTimer = setTimeout, clearTimer = clearTimeout }) {
  let people = new Map();
  let me = null;
  // Whether Talk's own state says how many wait.
  let talkCounts = false;
  /** The calls ringing here, not yet answered: callId -> { msg, timer, told: Set<window> }. */
  const ringing = new Map();

  const nameOf = (from) => {
    const id = typeof from === 'object' && from ? from.id : from;
    const known = people.get(id);
    const given = typeof from === 'object' && from ? from : {};
    return known?.name || given.name || known?.email || given.email || '';
  };
  const idOf = (from) => (typeof from === 'object' && from ? from.id : from);

  const stopRinging = (callId) => {
    const call = ringing.get(callId);
    if (!call) return null;
    clearTimer(call.timer);
    ringing.delete(callId);
    notifier.withdraw(`call:${callId}`);
    return call;
  };

  const missed = (callId, msg) => {
    notifier.report({
      kind: 'missed',
      id: `missed:${callId}`,
      from: nameOf(msg.from),
      video: Boolean(msg.video),
      counted: !talkCounts,
      app: 'talk',
      open: { app: 'talk', query: { contact: idOf(msg.from), show: 'calls' } },
    });
  };

  const stops = [
    on('talk:state', (state) => {
      if (!state || typeof state !== 'object') return;
      me = state.me?.id ?? null;
      people = new Map((state.contacts || []).map((c) => [c.id, c]));
      if (typeof state.unreadTotal === 'number' && typeof state.missed === 'number') {
        talkCounts = true;
        notifier.setWaiting('chat', state.unreadTotal);
        notifier.setWaiting('missed', state.missed);
      }
      // Signed out: nothing rings, and nothing of Talk's waits.
      if (state.phase === 'signed-out') for (const id of [...ringing.keys()]) stopRinging(id);
    }),

    on('talk:event', (msg) => {
      if (!msg || typeof msg.type !== 'string') return;
      switch (msg.type) {
        case 'messages':
        case 'group.messages': {
          // One notification a person (or a group), with the newest words and how many came.
          const by = new Map();
          for (const m of msg.messages || []) {
            if (!m || m.from === me) continue;
            const key = msg.type === 'group.messages' ? `group:${m.group ?? msg.groupId ?? ''}` : String(m.from);
            const held = by.get(key) || { count: 0, last: null };
            held.count += 1;
            if (!held.last || (m.at || 0) >= (held.last.at || 0)) held.last = m;
            by.set(key, held);
          }
          for (const [key, { count, last }] of by) {
            const group = msg.type === 'group.messages' ? (msg.group?.name || msg.name || last.groupName || '') : '';
            notifier.report({
              kind: 'chat',
              id: `chat:${key}`,
              from: nameOf(last.from),
              group: group || undefined,
              text: last.text,
              count,
              counted: !talkCounts,
              front: front('talk'),
              app: 'talk',
              open: { app: 'talk', query: msg.type === 'group.messages' ? { group: key.slice('group:'.length) } : { contact: last.from } },
            });
          }
          break;
        }
        case 'call.incoming': {
          if (!msg.callId || ringing.has(msg.callId)) break;
          let heard = [];
          try { heard = talkWindows() || []; } catch { heard = []; }
          const call = { msg, told: new Set(heard), timer: null };
          // A call nobody says has stopped is taken to have stopped after a while, and was missed.
          call.timer = setTimer(() => { if (stopRinging(msg.callId)) missed(msg.callId, msg); }, RING_MS);
          call.timer?.unref?.();
          ringing.set(msg.callId, call);
          notifier.report({
            kind: 'call',
            id: `call:${msg.callId}`,
            from: nameOf(msg.from),
            video: Boolean(msg.video),
            app: 'talk',
            open: { app: 'talk', query: { contact: idOf(msg.from), call: msg.callId } },
            on: {
              decline: () => {
                if (!stopRinging(msg.callId)) return;
                try {
                  const sent = send?.({ type: 'call.decline', callId: msg.callId });
                  sent?.catch?.(() => {});
                } catch {
                  /* not connected: the call stops ringing on its own */
                }
              },
            },
          });
          break;
        }
        // Answered on another of the person's computers: not missed.
        case 'call.taken':
          stopRinging(msg.callId);
          break;
        case 'call.ended': {
          // It stopped ringing and nobody here had answered.
          const call = stopRinging(msg.callId);
          if (call) missed(msg.callId, call.msg);
          break;
        }
        default:
          break;
      }
    }),
  ];

  return {
    /** What a window sends to the introducer: an answer here, or the call turned away, ends its ringing. */
    sending(msg) {
      if (!msg || !msg.callId) return;
      if (msg.type === 'call.answer') stopRinging(msg.callId);
      else if (msg.type === 'call.decline') {
        const call = stopRinging(msg.callId);
        // Turned away because another call is on: that one was missed.
        if (call && msg.reason === 'busy') missed(msg.callId, call.msg);
      }
    },

    /** A Talk window has started and listens: it is told once of each call still ringing. */
    windowReady(win) {
      if (!win) return;
      for (const call of ringing.values()) {
        if (call.told.has(win)) continue;
        call.told.add(win);
        try { tellWindow(win, call.msg); } catch { /* a window on its way out */ }
      }
    },

    ringing: () => [...ringing.keys()],

    stop() {
      for (const s of stops) s?.();
      for (const id of [...ringing.keys()]) stopRinging(id);
    },
  };
}
