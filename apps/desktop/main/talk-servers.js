// A meeting held on a Talk Server, as Talk's service asks for it (the Talk
// server contract, "A meeting from start to end").
//
// The introducer says which server carries a meeting and gives each person
// their own pass; the pass is exchanged at the server's door; the meeting's
// signalling then runs through that door. All three are done here, on the
// main side, for two reasons: only here can the door's certificate be checked
// against what the introducer said (talk-door.js), and a window of Office
// connects to nothing outside Office.
//
// What a window is given: the meeting as the introducer describes it, the
// join token for the engine in the server, and a number for the way in that
// is open (`session`). What it is never given: a door's address, the pin, a
// pass, or a host's control token. It cannot name a server, only a meeting.
//
// Nothing here is kept on disk. A pass is used once and forgotten, and a way
// in ends with the window that opened it.

import { AGAIN, DoorError, askDoor, exchange, openSocket } from './talk-door.js';

const MEETING_ID = /^[\w-]{8,64}$/;
const PERSON_ID = /^[\w.:-]{3,128}$/;
const ROLES = ['host', 'member', 'listener'];
const WHO = ['me', 'contacts', 'named', 'org'];
const ANSWER_MS = 15000;
/** The made-up host a window's meeting connects to; only `sessionOf` reads it, and nothing is ever sent there. */
export const DOOR_HOST = 'talk-door.invalid';

/** Where the introducer answers by https: the address of its socket, with the path dropped. */
export function restBase(server) {
  try {
    const u = new URL(String(server || ''));
    if (u.protocol !== 'wss:' && u.protocol !== 'ws:') return null;
    return `${u.protocol === 'wss:' ? 'https:' : 'http:'}//${u.host}`;
  } catch { return null; }
}

/**
 * @param {object} o
 * @param {() => string} o.server the introducer's socket address, as Talk's service keeps it
 * @param {() => string|null} o.token the person's access token
 * @param {(force?: boolean) => Promise<'fresh'|'kept'|'ended'>} o.renew
 * @param {typeof fetch} o.fetchImpl
 * @param {(win: object, event: string, payload: object) => void} o.send to one window
 * @param {object} [o.door] the door's three calls, for a test's stand-in
 */
export function createServers({ server, token, renew, fetchImpl, send = () => {}, door = { exchange, askDoor, openSocket } }) {
  /** Ways in that are open: number -> { owner, id, d, control, sockets }. */
  const sessions = new Map();
  /** The signalling connections: number -> { session, owner, handle, win }. */
  const sockets = new Map();
  let lastSession = 0;
  let lastSocket = 0;
  /** The windows already watched for their closing. */
  const watched = new Set();

  /**
   * One question to the introducer about servers. The answer as it gave it,
   * or { error } with one of: the introducer's own refusals; `offline` when
   * it cannot be reached; `no-service` when it has no servers at all;
   * `sign-in` when the sign-in has ended.
   */
  const call = async (method, path, body = undefined, again = true) => {
    const base = restBase(server());
    const bearer = token();
    if (!base) return { error: 'offline' };
    if (!bearer) return { error: 'sign-in' };
    let res;
    let json = null;
    try {
      res = await fetchImpl(`${base}/servers/${path}`, {
        method,
        headers: { accept: 'application/json', authorization: `Bearer ${bearer}`, ...(body !== undefined ? { 'content-type': 'application/json' } : {}) },
        ...(body !== undefined ? { body: JSON.stringify(body) } : {}),
        redirect: 'error',
        signal: AbortSignal.timeout(ANSWER_MS),
      });
      const text = await res.text();
      try { json = JSON.parse(text); } catch { json = null; }
    } catch { return { error: 'offline' }; }
    if (res.status === 200 && json && typeof json === 'object') return json;
    const error = json && typeof json.error === 'string' && /^[a-z-]{1,40}$/.test(json.error) ? json.error : null;
    if (res.status === 401 && again) {
      // The token ran out between two questions: a fresh one, and the question once more.
      if ((await renew(true)) === 'ended') return { error: 'sign-in' };
      return call(method, path, body, false);
    }
    if (res.status === 401) return { error: 'sign-in' };
    // An introducer with its servers turned off answers as it does for any address it does not have.
    if (res.status === 404 && error !== 'no-meeting' && error !== 'no-server') return { error: 'no-service' };
    if (!error) return { error: res.status >= 500 ? 'failed' : 'refused' };
    return { ...json, error };
  };

  const meetingOf = (m) => (m && typeof m === 'object' ? { id: String(m.id || ''), live: Boolean(m.live), e2ee: m.e2ee !== false, maxPeople: Number(m.maxPeople) || 0, state: String(m.state || '') } : null);
  /** A server of the person's own as a window may know it: no door and no pin. */
  const ownView = (s) => ({
    id: String(s.id || ''),
    name: String(s.name || ''),
    state: String(s.state || ''),
    answering: Boolean(s.answering),
    canCarry: Boolean(s.canCarry),
    why: s.why ? String(s.why) : null,
    who: WHO.includes(s.settings?.who) ? s.settings.who : 'me',
    named: Array.isArray(s.settings?.named) ? s.settings.named.map(String) : [],
  });
  const refusal = (r) => ({
    error: r.error,
    ...(Array.isArray(r.own) ? { own: r.own.map((s) => ({ id: String(s.id || ''), name: String(s.name || ''), why: s.why ? String(s.why) : null })) } : {}),
    ...(r.state ? { state: String(r.state) } : {}),
  });
  const mid = (id) => (MEETING_ID.test(String(id || '')) ? String(id) : null);

  const closeSocket = (n, code = 1000, reason = '') => {
    const s = sockets.get(n);
    if (!s) return;
    sockets.delete(n);
    sessions.get(s.session)?.sockets.delete(n);
    try { s.handle.close(code, reason); } catch { /* gone */ }
  };
  const closeSession = (n) => {
    const s = sessions.get(n);
    if (!s) return;
    for (const k of [...s.sockets]) closeSocket(k, 1001, '');
    sessions.delete(n);
  };
  const ownerOf = (win) => (win && win.id != null ? win.id : null);
  /** A window's ways in end with it. */
  const watch = (win) => {
    const owner = ownerOf(win);
    if (owner == null || watched.has(owner) || typeof win.once !== 'function') return;
    watched.add(owner);
    win.once('closed', () => {
      watched.delete(owner);
      for (const [n, s] of [...sessions]) if (s.owner === owner) closeSession(n);
    });
  };
  /** The way in a window names, only if it is that window's own. */
  const mine = (n, win) => {
    const s = sessions.get(Number(n));
    return s && s.owner === ownerOf(win) ? s : null;
  };

  return {
    /** Every way in closed: the person signed out, or Office is closing. */
    forget() { for (const n of [...sessions.keys()]) closeSession(n); },
    /** How many ways in and connections are open (for a test, and for the proof's last line). */
    open: () => ({ sessions: sessions.size, sockets: sockets.size }),

    api: {
      /**
       * The person's own Talk Servers and what their plan allows:
       * { plan, servers: [{ id, name, state, answering, canCarry, why, who, named }] } | { error }.
       * `no-service` means this introducer has no servers, so nothing about them is offered.
       */
      serverMine: async () => {
        const r = await call('GET', 'mine');
        if (r.error) return refusal(r);
        const p = r.plan || {};
        return { plan: { id: String(p.id || ''), maxPeople: Number(p.maxPeople) || 0, maxServers: p.maxServers == null ? null : Number(p.maxServers), rutbaServers: Boolean(p.rutbaServers), recording: Boolean(p.recording) }, servers: [].concat(r.servers || []).map(ownView) };
      },

      /** Who may hold meetings on a server of the person's own: { server } | { error }. */
      serverShare: async ({ id, who, named } = {}) => {
        if (!PERSON_ID.test(String(id || '')) || (who !== undefined && !WHO.includes(who))) return { error: 'refused' };
        const settings = { ...(who !== undefined ? { who } : {}), ...(named !== undefined ? { named: [].concat(named || []).map(String).filter((x) => PERSON_ID.test(x)).slice(0, 500) } : {}) };
        const r = await call('PATCH', `mine/${encodeURIComponent(id)}`, { settings });
        if (r.error) return refusal(r);
        return { server: ownView(r.server || r) };
      },

      /**
       * A server asked for, for a meeting this person starts (step 1). Asked
       * again for the same meeting it answers the same server, or moves the
       * meeting, whose id then changes: { meeting } | { error, own? }.
       */
      serverCarry: async ({ id, live = false } = {}) => {
        if (!mid(id)) return { error: 'refused' };
        const r = await call('POST', 'meetings', { meetingId: mid(id), live: live === true });
        if (r.error) return refusal(r);
        return { meeting: meetingOf(r.meeting) };
      },

      /** Somebody let in, by the starter or a host (step 2): { person, role } | { error }. */
      serverAdmit: async ({ id, person, role = 'member' } = {}) => {
        if (!mid(id) || !PERSON_ID.test(String(person || '')) || !ROLES.includes(role)) return { error: 'refused' };
        const r = await call('POST', `meetings/${encodeURIComponent(id)}/admit`, { person: { id: String(person), kind: 'office' }, role });
        if (r.error) return refusal(r);
        return { person: String(person), role: String(r.role || role) };
      },

      /** Somebody taken out: they get no more passes, and the server closes their connection. */
      serverRemove: async ({ id, person } = {}) => {
        if (!mid(id) || !PERSON_ID.test(String(person || ''))) return { error: 'refused' };
        const r = await call('POST', `meetings/${encodeURIComponent(id)}/remove`, { person: String(person) });
        return r.error ? refusal(r) : { removed: true };
      },

      /** The meeting ended on its server, for everyone (step 8). */
      serverEnd: async ({ id } = {}) => {
        if (!mid(id)) return { error: 'refused' };
        const r = await call('POST', `meetings/${encodeURIComponent(id)}/end`);
        return r.error ? refusal(r) : { ended: true };
      },

      /** Whether a meeting is still on, for somebody who is in it: { meeting } | { error }. */
      serverInfo: async ({ id } = {}) => {
        if (!mid(id)) return { error: 'refused' };
        const r = await call('GET', `meetings/${encodeURIComponent(id)}`);
        return r.error ? refusal(r) : { meeting: meetingOf(r.meeting) };
      },

      /**
       * Into a meeting on its server (steps 3 and 4): this person's own pass
       * from the introducer, exchanged at the server's door.
       * { session, url, token, room, identity, role, host, e2ee, live, meeting } | { error, state? }.
       *
       * `url` is not the server's address: it names this way in, and the
       * window's connection to it is carried through `doorOpen` below.
       * Errors: the introducer's (`no-pass`, `no-meeting`, `server-gone`,
       * `offline`, `sign-in`, ...) and the door's (`pin`, `certificate`,
       * `plain`, `unreachable`, `ended`, `removed`, `not-serving`, ...).
       */
      serverEnter: async ({ id } = {}, win = null) => {
        if (!mid(id)) return { error: 'refused' };
        for (let attempt = 0; attempt < 2; attempt += 1) {
          const r = await call('POST', `meetings/${encodeURIComponent(id)}/pass`);
          if (r.error) return refusal(r);
          if (typeof r.pass !== 'string' || !r.pass || !r.server) return { error: 'refused' };
          let got;
          try {
            got = await door.exchange({ doors: r.server.doors, pin: r.server.pin ?? null, pass: r.pass });
          } catch (err) {
            const code = err instanceof DoorError ? err.code : 'unreachable';
            // A pass that ran out on the way, or that the server took as old: one new pass, once.
            if (AGAIN.has(code) && attempt === 0) continue;
            return { error: code, ...(err?.extra?.state ? { state: err.extra.state } : {}) };
          }
          const a = got.admitted;
          const n = (lastSession += 1);
          sessions.set(n, { owner: ownerOf(win), id: String(id), d: got.door, control: typeof a.control === 'string' ? a.control : null, sockets: new Set() });
          watch(win);
          return {
            session: n,
            url: `wss://s${n}.${DOOR_HOST}`,
            token: String(a.token),
            room: String(a.room || ''),
            identity: String(a.identity || ''),
            role: ROLES.includes(a.role) ? a.role : String(r.role || 'member'),
            host: Boolean(a.control),
            e2ee: a.e2ee !== false,
            live: Boolean(a.live),
            meeting: meetingOf(r.meeting),
          };
        }
        return { error: 'expired' };
      },

      /** A way in closed, with its connections. */
      serverLeave: ({ session } = {}, win = null) => {
        if (!mine(session, win)) return { left: false };
        closeSession(Number(session));
        return { left: true };
      },

      /**
       * What a host asks the server's door (step 7): who is in the room, one
       * person's microphone off, everybody's off.
       * { people } | { done } | { error }.
       */
      serverControl: async ({ session, do: what, person } = {}, win = null) => {
        const s = mine(session, win);
        if (!s) return { error: 'no-session' };
        if (!s.control) return { error: 'not-host' };
        if (!['roster', 'mute', 'mute-all'].includes(what)) return { error: 'refused' };
        if (what === 'mute' && !PERSON_ID.test(String(person || ''))) return { error: 'refused' };
        let got;
        try {
          got = await door.askDoor(s.d, { method: 'POST', path: '/control', body: { control: s.control, do: what, ...(what === 'mute' ? { person: String(person) } : {}) } });
        } catch (err) { return { error: err instanceof DoorError ? err.code : 'unreachable' }; }
        if (got.status !== 200) return { error: typeof got.body?.error === 'string' && /^[a-z-]{1,40}$/.test(got.body.error) ? got.body.error : 'refused' };
        return what === 'roster' ? { people: [].concat(got.body.people || []).map((p) => ({ id: String(p?.id || ''), name: String(p?.name || ''), muted: Boolean(p?.muted) })) } : { done: true };
      },

      /**
       * The meeting's signalling connection, opened through the door of a
       * way in this window holds. What happens on it is told to the window
       * alone, as `talk:door` { socket, type: 'open' | 'message' | 'close', ... }.
       * { socket } | { error }.
       */
      doorOpen: ({ session, target, protocols = [] } = {}, win = null) => {
        const s = mine(session, win);
        if (!s) return { error: 'no-session' };
        if (s.sockets.size >= 4) return { error: 'too-many' };
        const n = (lastSocket += 1);
        const tell = (payload) => { if (sockets.has(n) || payload.type === 'close') send(win, 'talk:door', { socket: n, ...payload }); };
        const handle = door.openSocket(s.d, String(target || ''), {
          open: (e) => tell({ type: 'open', protocol: e?.protocol || '' }),
          message: (data, binary) => tell(binary ? { type: 'message', binary: true, data: new Uint8Array(data.buffer, data.byteOffset, data.byteLength).slice() } : { type: 'message', binary: false, data: data.toString('utf8') }),
          close: (e) => {
            const had = sockets.delete(n);
            s.sockets.delete(n);
            if (had) send(win, 'talk:door', { socket: n, type: 'close', code: Number(e.code) || 1006, reason: String(e.reason || ''), ...(e.status ? { status: e.status } : {}), ...(e.why ? { why: e.why } : {}) });
          },
        }, { protocols });
        sockets.set(n, { session: Number(session), owner: s.owner, handle });
        s.sockets.add(n);
        return { socket: n };
      },

      /** Something sent on a signalling connection of this window's: { sent }. */
      doorSend: ({ socket, data, binary = true } = {}, win = null) => {
        const s = sockets.get(Number(socket));
        if (!s || s.owner !== ownerOf(win)) return { sent: false };
        if (binary === false || typeof data === 'string') return { sent: s.handle.send(String(data ?? ''), false) };
        let bytes = null;
        if (data instanceof ArrayBuffer) bytes = Buffer.from(data);
        else if (ArrayBuffer.isView(data)) bytes = Buffer.from(data.buffer, data.byteOffset, data.byteLength);
        return { sent: bytes ? s.handle.send(bytes, true) : false };
      },

      /** A signalling connection of this window's closed. The window is told, as for any close. */
      doorClose: ({ socket, code = 1000, reason = '' } = {}, win = null) => {
        const s = sockets.get(Number(socket));
        if (!s || s.owner !== ownerOf(win)) return { closed: false };
        try { s.handle.close(Number(code) || 1000, String(reason || '').slice(0, 100)); } catch { /* gone */ }
        return { closed: true };
      },
    },
  };
}
