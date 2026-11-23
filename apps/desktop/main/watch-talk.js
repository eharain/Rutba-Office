// Talk, as the notifier hears of it (watch.js, D-2026-11-22-23): a call
// coming in, a call missed, a message, a meeting someone asks the person
// into.
//
// Nothing here is Talk's own. Talk says for itself what is worth a
// notification: one broadcast, 'talk:notify' (main/talk-notify.js), for a
// call ringing here and no longer ringing, a call missed (also one made
// while this computer was off), a message, a meeting the person is asked
// into and one no longer offered. This file turns each into a report to the
// notifier, or takes its notification down. Talk shows no notification
// itself, and rings in its own window. A call that comes while no Talk
// window is open has no window to ring in, so the notifier is asked to ring
// for it (the report's `ring`), unless the person turned Talk's sounds off.
//
// How many messages are unread and how many calls were missed are Talk's to
// say (its state's `unreadTotal` and `missed`): those are the numbers by the
// clock.

/** What Talk tells that is worth a notification (the same as talk-notify.js's NOTIFY_EVENT). */
export const NOTIFY_EVENT = 'talk:notify';

/**
 * @param {object} o
 * @param {object} o.notifier
 * @param {(event: string, fn: Function) => Function} o.on   listens to what the main process tells its windows
 * @param {(app: string) => boolean} [o.front]   whether a window of that app is in front now
 * @param {(app: string) => boolean} [o.open]   whether a window of that app is open, shown or not; without it Talk's window is taken to ring
 * @param {(msg: object) => any} [o.send]   sends to the introducer, as Talk's own service does (a call turned away from its notification)
 * @param {() => object} [o.state]   Talk's state as its service gives it now; without it, the last one told is used
 */
export function watchTalk({ notifier, on, front = () => false, open = () => true, send = null, state: stateNow = null }) {
  // Whether Talk's own state says how many wait.
  let talkCounts = false;
  // The phase Talk was last told to be in.
  let phase = '';
  // Whether the person wants a sound for calls (Talk's Settings), as Talk last said.
  let sounds = true;
  /** Whether the notifier rings for a call: no Talk window is there to, and the person has Talk's sounds on. */
  const ringsHere = () => {
    let on = sounds;
    try {
      const ring = stateNow?.()?.ring;
      if (typeof ring === 'boolean') on = ring;
    } catch {
      on = sounds;
    }
    if (!on) return false;
    try { return !open('talk'); } catch { return false; }
  };
  /** How many messages have come in each conversation since Talk was last looked at: what its notification says. */
  const fresh = new Map();
  /** The calls and meetings whose notification may still be up. */
  const calls = new Set();
  const meetings = new Set();

  const nameOf = (who) => String(who?.name || '').trim();
  const takeDown = (set, prefix, id) => {
    if (!set.delete(id)) return;
    notifier.withdraw(`${prefix}:${id}`);
  };

  const told = (n) => {
    if (!n || typeof n.kind !== 'string' || n.id == null) return;
    switch (n.kind) {
      case 'call': {
        calls.add(n.id);
        notifier.report({
          kind: 'call',
          id: `call:${n.id}`,
          from: nameOf(n.from),
          video: Boolean(n.video),
          app: 'talk',
          ring: ringsHere(),
          // Talk opens at the person; its window shows the call still ringing, to answer there.
          open: { app: 'talk', query: n.from?.id != null ? { contact: n.from.id } : {} },
          on: {
            decline: () => {
              try {
                const sent = send?.({ type: 'call.decline', callId: n.id });
                sent?.catch?.(() => {});
              } catch {
                /* not connected: the call stops ringing on its own */
              }
            },
          },
        });
        break;
      }
      case 'call-over':
        takeDown(calls, 'call', n.id);
        break;
      case 'missed':
        // One notification a caller: someone who calls again is told again, with how many.
        notifier.report({
          kind: 'missed',
          id: `missed:${n.id}`,
          from: nameOf(n.from),
          video: Boolean(n.video),
          count: Math.max(1, Number(n.count) || 1),
          counted: !talkCounts,
          app: 'talk',
          open: { app: 'talk', query: { ...(n.from?.id != null ? { contact: n.from.id } : {}), show: 'calls' } },
        });
        break;
      case 'message': {
        const inFront = Boolean(front('talk'));
        const key = n.group ? `group:${n.group.id}` : String(n.from?.id ?? '');
        const count = inFront ? 1 : (fresh.get(key) || 0) + 1;
        if (!inFront) fresh.set(key, count);
        // One notification a person (or a group), with the newest words and how many have come.
        notifier.report({
          kind: 'chat',
          id: `chat:${key}`,
          from: nameOf(n.from),
          group: n.group ? nameOf(n.group) || undefined : undefined,
          text: n.text,
          count,
          counted: !talkCounts,
          front: inFront,
          app: 'talk',
          open: { app: 'talk', query: n.group ? { group: n.group.id } : { contact: n.from?.id } },
        });
        break;
      }
      case 'meeting':
        meetings.add(n.id);
        notifier.report({
          kind: 'meeting',
          id: `meet:${n.id}`,
          summary: n.name || nameOf(n.group),
          by: nameOf(n.from),
          live: Boolean(n.live),
          recorded: Boolean(n.recorded),
          join: true,
          front: front('talk'),
          app: 'talk',
          open: { app: 'talk', query: { meet: n.id } },
        });
        break;
      case 'meeting-over':
        takeDown(meetings, 'meet', n.id);
        break;
      default:
        break;
    }
  };

  const stops = [
    on('talk:state', (state) => {
      if (!state || typeof state !== 'object') return;
      if (typeof state.phase === 'string') phase = state.phase;
      if (typeof state.ring === 'boolean') sounds = state.ring;
      if (typeof state.unreadTotal === 'number' && typeof state.missed === 'number') {
        talkCounts = true;
        notifier.setWaiting('chat', state.unreadTotal);
        notifier.setWaiting('missed', state.missed);
        // Everything read: the next message is the first again.
        if (!state.unreadTotal) fresh.clear();
      }
      // Signed out: nothing of Talk's stays up.
      if (state.phase === 'signed-out') {
        for (const id of [...calls]) takeDown(calls, 'call', id);
        for (const id of [...meetings]) takeDown(meetings, 'meet', id);
        fresh.clear();
      }
    }),
    on(NOTIFY_EVENT, (n) => {
      try { told(n); } catch { /* the watch's fault is not Talk's */ }
    }),
  ];

  return {
    /** The calls whose notification may still be up. */
    ringing: () => [...calls],
    /** Whether someone is signed in to Talk here, connected or not at this moment: a call or a message can come. */
    signedIn: () => {
      let is = phase;
      try { is = stateNow?.()?.phase || phase; } catch { is = phase; }
      return Boolean(is) && is !== 'signed-out' && is !== 'signing-in';
    },
    stop() {
      for (const s of stops) s?.();
      for (const id of [...calls]) takeDown(calls, 'call', id);
      for (const id of [...meetings]) takeDown(meetings, 'meet', id);
      fresh.clear();
    },
  };
}
