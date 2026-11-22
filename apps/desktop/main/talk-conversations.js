// What Talk keeps about its conversations beside their messages (Office
// 2.28, the owner's go-ahead of 2026-11-22 on the Talk, Meet and Live plan):
//
// - how many messages in each conversation have not been read here, for the
//   count on each person and on Talk's own icon;
// - the calls that were missed, whether they rang here unanswered or were
//   made while this computer was off (the introducer keeps a note of those);
// - for each message sent, whether the introducer has it and whether it has
//   arrived at the other side;
// - the group conversations the person is in: three or more people, each put
//   there by one of their own contacts.
//
// talk.js holds the connection and the messages' files and hands this what
// the introducer says; this answers whether the message was one of its own.

/** What this Office says it reads, in its hello: an introducer from before them sends none of it. */
export const CAN = ['groups', 'notices', 'meetings'];
/** How a group's conversation is named among the conversations: never a person's id. */
export const groupKey = (groupId) => `group:${groupId}`;
const CALLS_CAP = 200;
/** The most one line or one entry says were missed. */
const MISSED_MOST = 999;
/** Messages on their way that are remembered, and for how long one is tried again before it is marked as not sent. */
const SENDING_CAP = 500;
export const SENDING_KEEP_MS = 24 * 3600 * 1000;
/** An answer that is also said to other people (a group that changed) is this question's only when it names its ref. */
const byRef = (m, ref) => m.ref === ref;
/** A refusal as a window is told it: why, and what the wording needs. */
const refused = (r) => ({ error: r.error, ...(r.most ? { most: r.most } : {}), ...(Array.isArray(r.who) && r.who.length ? { who: r.who } : {}) });
const REPORT_MESSAGES = 20;
export const REPORT_REASONS = ['spam', 'harassment', 'impersonation', 'other'];

/**
 * @param {object} o
 * @param {() => object|null} o.me the person signed in
 * @param {() => string[]} [o.contacts] the ids of the person's contacts now
 * @param {(key: string) => object[]} o.readHistory a conversation's messages
 * @param {(key: string, list: object[]) => void} o.writeHistory
 * @param {(key: string, entry: object) => void} o.keep one message kept, in place of an earlier one with its id
 * @param {(name: string, fallback: any) => any} o.readFile a small file of the person's own, beside their messages
 * @param {(name: string, value: any) => void} o.writeFile
 * @param {(msg: object) => void} o.send to the introducer; throws when not connected
 * @param {(msg: object, answer: string, match?: Function) => Promise<object>} o.ask
 * @param {(n: object) => void} [o.notify] something worth a notification (talk-notify.js): a call ringing or no longer ringing, a missed call, a message
 * @param {() => number} [o.now]
 */
export function createConversations({ me, contacts = null, readHistory, writeHistory, keep, readFile, writeFile, send, ask, notify = () => {}, now = Date.now }) {
  let groups = [];
  /** Calls ringing here that nobody has answered yet: callId -> { from, video, msg }; msg is what the introducer said, which a window needs to answer it. */
  const ringing = new Map();

  // `sending` is which conversation each message on its way went to (id -> { key, at }), until the
  // introducer says it has it. Kept in the file, so one that was on its way when the connection
  // dropped, or when Office was closed, is sent again and never waits for ever.
  const marks = () => {
    const m = readFile('marks.json', null);
    const own = (o) => (o && typeof o === 'object' && !Array.isArray(o) ? o : {});
    return { unread: own(m?.unread), calls: Array.isArray(m?.calls) ? m.calls : [], sending: own(m?.sending) };
  };
  const save = (m) => { try { writeFile('marks.json', m); } catch { /* the counts are kept next time */ } };
  const trySend = (msg) => { try { send(msg); } catch { /* said again when the connection is back */ } };

  /** A message that came in: one more unread in its conversation. */
  const unread = (key, n = 1) => {
    const m = marks();
    m.unread[key] = (m.unread[key] || 0) + n;
    save(m);
  };

  /** The conversation a message on its way went to, or null once the introducer has answered for it. */
  const pending = (id) => marks().sending[id]?.key ?? null;
  const settle = (id) => {
    const m = marks();
    if (id in m.sending) { delete m.sending[id]; save(m); }
  };

  /**
   * A message that came in, kept once, by who sent it and its id. One that
   * is here already is neither replaced nor counted again (the introducer
   * gives a message again when it did not hear that Office has it), and an
   * id that some other entry holds (a message of mine, or someone else's in
   * the group) leaves that entry as it is: nobody's words can be put in the
   * place of another's. True when the message is new here.
   */
  const keepIncoming = (key, entry) => {
    const list = readHistory(key);
    const own = `${entry.id}~${entry.from}`;
    if (list.some((e) => !e.mine && !e.kind && e.from === entry.from && (e.id === entry.id || e.id === own))) return false;
    list.push(list.some((e) => e.id === entry.id) ? { ...entry, id: own } : entry);
    writeHistory(key, list);
    return true;
  };

  /** Counts that belong to no conversation any more are let go. */
  const prune = (stays) => {
    const m = marks();
    const gone = Object.keys(m.unread).filter((key) => !stays(key));
    if (!gone.length) return;
    for (const key of gone) delete m.unread[key];
    save(m);
  };

  /** The connection is back: what was still on its way is sent again under its own id, or, after a day, marked as not sent. */
  const sendAgain = () => {
    const m = marks();
    let changed = false;
    for (const [id, p] of Object.entries(m.sending)) {
      const key = String(p?.key || '');
      const entry = readHistory(key).find((e) => e.id === id && e.mine);
      if (!entry || entry.state !== 'sending') { delete m.sending[id]; changed = true; continue; }
      if (now() - (p.at || 0) > SENDING_KEEP_MS) {
        mark(key, id, () => ({ state: 'failed' }));
        delete m.sending[id];
        changed = true;
        continue;
      }
      // The introducer takes an id once, and the other side keeps a message once, so this never shows twice.
      trySend(key.startsWith('group:') ? { type: 'group.send', groupId: key.slice(6), id, text: entry.text } : { type: 'msg.send', to: key, id, text: entry.text });
    }
    if (changed) save(m);
  };

  /** A message of mine changed in place: what `change` gives is laid over it. */
  const mark = (key, id, change) => {
    const list = readHistory(key);
    const i = list.findIndex((e) => e.id === id);
    if (i < 0) return false;
    const next = change(list[i]);
    if (!next) return false;
    list[i] = { ...list[i], ...next };
    writeHistory(key, list);
    return true;
  };

  /**
   * A call that reached nobody here: counted once, however it was learned
   * of. Calls missed from the same person fold: into one entry of the list
   * while it has not been looked at, and into one line of the conversation
   * while nothing was said after it, each with a count, so someone calling
   * again and again fills neither. `count` is how many calls the
   * introducer's note stands for, and `note` its id: the same note comes
   * again, with a higher count, when another call was missed before Office
   * had said it has it.
   */
  const missed = ({ callId, from, video = false, why = 'unanswered', at = now(), count = 1, note = null }) => {
    if (!from?.id || !callId) return;
    const stands = Math.max(1, Math.min(Math.floor(Number(count)) || 1, MISSED_MOST));
    const m = marks();
    // It rang here and the introducer noted it too: the one call.
    const had = m.calls.findIndex((c) => c.id === callId || (c.ids || []).includes(callId));
    if (had >= 0) {
      if (note) { m.calls[had] = { ...m.calls[had], note, noted: stands }; save(m); }
      return;
    }
    const told = note ? m.calls.findIndex((c) => c.note === note) : -1;
    // The same note again says only what is new since it was last told.
    const more = told >= 0 ? Math.max(1, stands - (m.calls[told].noted || 0)) : stands;
    let into = told;
    if (into < 0) for (let i = m.calls.length - 1; i >= 0; i--) if (m.calls[i].with === from.id) { into = i; break; }
    if (into >= 0 && m.calls[into].seen) into = -1;
    const base = into >= 0 ? m.calls[into] : null;
    const entry = {
      id: callId,
      ids: [...(base ? base.ids || [base.id] : []), callId].slice(-20),
      with: from.id,
      name: from.name || from.email || base?.name || '',
      video: Boolean(video),
      why,
      at,
      seen: false,
      count: Math.min((base ? base.count || 1 : 0) + more, MISSED_MOST),
      ...(note ? { note, noted: stands } : {}),
    };
    m.calls = [...m.calls.filter((c, i) => i !== into).map((c) => (note && c.note === note ? { ...c, note: null } : c)), entry].slice(-CALLS_CAP);
    save(m);
    const list = readHistory(from.id);
    const last = list[list.length - 1];
    const folds = last?.kind === 'missed-call';
    if (folds) list.pop();
    list.push({ id: `call:${callId}`, kind: 'missed-call', from: from.id, video: Boolean(video), at, count: Math.min((folds ? last.count || 1 : 0) + more, MISSED_MOST) });
    writeHistory(from.id, list);
    // One entry a person while it has not been looked at: told again, under the person's id, with what it now counts.
    notify({ kind: 'missed', id: from.id, callId, from: { id: from.id, name: entry.name }, video: entry.video, why, at, count: entry.count });
  };

  /** A line in a group's conversation saying what changed in it. */
  const noteEvent = (group, event) => {
    if (!event?.kind) return;
    keep(groupKey(group.id), { id: `event:${event.kind}:${event.at || now()}:${(event.who || []).map((w) => w.id).join(',')}`, kind: 'group-event', event, at: event.at || now() });
  };

  const quiet = () => {
    for (const id of [...ringing.keys()]) notify({ kind: 'call-over', id, why: 'ended' });
    ringing.clear();
  };

  const setGroup = (group) => { groups = [...groups.filter((g) => g.id !== group.id), group].sort((a, b) => a.created - b.created); };

  return {
    /** For the state every window is told: the groups, the unread counts, and how many missed calls are not yet looked at. */
    snapshot() {
      if (!me()) return { groups: [], unread: {}, unreadTotal: 0, missed: 0 };
      const m = marks();
      // Only conversations that are there now count: a contact since removed or blocked, or a group the person is no longer in, holds nothing up.
      const known = contacts ? new Set([...contacts(), ...groups.map((g) => groupKey(g.id))]) : null;
      const counts = Object.fromEntries(Object.entries(m.unread).filter(([key, n]) => n > 0 && (!known || known.has(key))));
      return { groups, unread: counts, unreadTotal: Object.values(counts).reduce((a, b) => a + b, 0), missed: m.calls.filter((c) => !c.seen).length };
    },

    /** Nothing rings here any more: signed out, or a new connection (a call that was ringing ended with the old one). */
    forget() { groups = []; quiet(); },

    /** The group conversations the person is in. */
    groups: () => groups,

    /** The calls ringing here now, nobody having answered: each as the introducer said it ('call.incoming'), so a window opened late can show and answer it. */
    ringing: () => [...ringing.values()].map((c) => c.msg),

    /** What the introducer said. True when the message was one of these, so every window's state is told again. */
    handle(msg) {
      switch (msg.type) {
        case 'welcome': {
          groups = [];
          quiet();
          const with_ = new Set((msg.contacts || []).map((c) => c.id));
          prune((key) => key.startsWith('group:') || with_.has(key));
          sendAgain();
          return false;
        }
        case 'contact.removed':
        case 'contact.blocked':
          prune((key) => key !== msg.id);
          return false;
        case 'messages': {
          const ids = [];
          for (const m of msg.messages || []) {
            // Told once: a message given again is neither counted nor told again.
            if (keepIncoming(m.from, { id: m.id, from: m.from, to: me()?.id, text: m.text, at: m.at, mine: false })) { unread(m.from); notify({ kind: 'message', id: m.id, from: m.from, group: null, text: m.text, at: m.at }); }
            ids.push(m.id);
          }
          // Kept here now, so the introducer lets them go.
          if (ids.length) trySend({ type: 'msg.ack', ids });
          return true;
        }
        case 'msg.sent': {
          const key = pending(msg.id);
          if (key) { mark(key, msg.id, (e) => (e.state === 'arrived' ? null : { state: 'sent' })); settle(msg.id); }
          return true;
        }
        case 'groups': {
          groups = (msg.groups || []).slice().sort((a, b) => a.created - b.created);
          const mine = new Set(groups.map((g) => groupKey(g.id)));
          prune((key) => !key.startsWith('group:') || mine.has(key));
          return true;
        }
        case 'group.created':
        case 'group.added':
          if (msg.group) {
            setGroup(msg.group);
            if (msg.type === 'group.added') noteEvent(msg.group, { kind: 'added-you', by: msg.by, at: now() });
          }
          return true;
        case 'group.changed':
          if (msg.group) { setGroup(msg.group); noteEvent(msg.group, msg.event); }
          return true;
        case 'group.left':
          // Off the list; what was said stays on this computer, and shows again if the person is put back.
          groups = groups.filter((g) => g.id !== msg.groupId);
          { const m = marks(); delete m.unread[groupKey(msg.groupId)]; save(m); }
          return true;
        case 'group.messages': {
          const ids = [];
          for (const m of msg.messages || []) {
            if (keepIncoming(groupKey(m.group), { id: m.id, group: m.group, from: m.from, text: m.text, at: m.at, mine: false })) { unread(groupKey(m.group)); notify({ kind: 'message', id: m.id, from: m.from, group: m.group, text: m.text, at: m.at }); }
            ids.push(m.id);
          }
          // Kept here now, so the introducer lets them go, and tells whoever sent them.
          if (ids.length) trySend({ type: 'group.ack', ids });
          return true;
        }
        case 'group.copy':
          if (msg.message) keep(groupKey(msg.message.group), { ...msg.message, mine: true, state: 'sent' });
          return true;
        case 'group.sent':
          mark(groupKey(msg.groupId), msg.id, (e) => ({ state: e.state === 'arrived' ? 'arrived' : 'sent', to: msg.to || [] }));
          settle(msg.id);
          return true;
        case 'error':
          // A message the introducer would not take: it says so on the message, in place of waiting for ever.
          if ((msg.for === 'msg.send' || msg.for === 'group.send') && pending(msg.id)) {
            mark(pending(msg.id), msg.id, () => ({ state: 'failed' }));
            settle(msg.id);
            return true;
          }
          return false;
        case 'notices': {
          const ids = [];
          for (const n of msg.notices || []) {
            ids.push(n.id);
            if (n.kind === 'missed-call') missed({ callId: n.callId, from: n.from, video: n.video, why: n.why, at: n.at, count: n.count, note: n.id });
            else if (n.kind === 'delivered' && n.group) {
              // The introducer says so once, when everyone it was sent to has it: never who, or when.
              if (n.all) for (const id of n.ids || []) { mark(groupKey(n.group), id, () => ({ state: 'arrived' })); settle(id); }
            } else if (n.kind === 'delivered') {
              for (const id of n.ids || []) { mark(n.by, id, () => ({ state: 'arrived' })); settle(id); }
            }
          }
          if (ids.length) trySend({ type: 'notice.ack', ids });
          return true;
        }
        case 'call.incoming':
          if (!msg.callId || ringing.has(msg.callId)) return false;
          ringing.set(msg.callId, { from: msg.from, video: Boolean(msg.video), msg });
          notify({ kind: 'call', id: msg.callId, from: msg.from, video: Boolean(msg.video) });
          return false;
        case 'call.taken':
          // Answered on another of the person's computers.
          if (ringing.delete(msg.callId)) notify({ kind: 'call-over', id: msg.callId, why: 'elsewhere' });
          return false;
        case 'call.ended':
          // It stopped ringing and nobody here had answered.
          if (ringing.has(msg.callId)) {
            const call = ringing.get(msg.callId);
            ringing.delete(msg.callId);
            notify({ kind: 'call-over', id: msg.callId, why: 'missed' });
            missed({ callId: msg.callId, from: call.from, video: call.video });
            return true;
          }
          return false;
        default:
          return false;
      }
    },

    /** What a window sends to the introducer: an answer here means the call was not missed; busy in another call means it was (true then). */
    sending(msg) {
      if (msg.type === 'call.answer') { if (ringing.delete(msg.callId)) notify({ kind: 'call-over', id: msg.callId, why: 'answered' }); }
      else if (msg.type === 'call.decline') {
        const call = ringing.get(msg.callId);
        ringing.delete(msg.callId);
        if (call) notify({ kind: 'call-over', id: msg.callId, why: msg.reason === 'busy' ? 'missed' : 'declined' });
        if (call && msg.reason === 'busy') { missed({ callId: msg.callId, from: call.from, video: call.video, why: 'busy' }); return true; }
      }
      return false;
    },

    /** A message of mine to a person or a group, on its way: marked as sending until the introducer says it has it. */
    sentToPerson(id, key) {
      const m = marks();
      m.sending[id] = { key, at: now() };
      const ids = Object.keys(m.sending);
      for (const old of ids.slice(0, Math.max(0, ids.length - SENDING_CAP))) delete m.sending[old];
      save(m);
    },

    api: {
      /** A conversation looked at: its unread count is cleared. `conversation` is a person's id or a group's key. */
      read: ({ conversation } = {}) => {
        const m = marks();
        if (m.unread[conversation]) { delete m.unread[conversation]; save(m); }
        return { read: true };
      },

      /** The missed calls, newest first; `seen` marks them all looked at. */
      calls: ({ seen = false } = {}) => {
        const m = marks();
        const list = m.calls.slice().reverse();
        if (seen && m.calls.some((c) => !c.seen)) { m.calls = m.calls.map((c) => ({ ...c, seen: true })); save(m); }
        return list;
      },

      // Each of these waits for the answer that names its own ref: what the introducer says to everyone
      // in a group when someone else changes it is never taken for the answer, so a refusal shows as one.
      /** A group conversation made of three or more people: { group } | { error, most?, who? }. */
      groupCreate: async ({ name = '', members = [] } = {}) => {
        const r = await ask({ type: 'group.create', name: String(name || ''), members: [].concat(members || []).map(String) }, 'group.created', byRef);
        return r.error ? refused(r) : { group: r.group };
      },
      groupAdd: async ({ groupId, members = [] } = {}) => {
        const r = await ask({ type: 'group.add', groupId: String(groupId || ''), members: [].concat(members || []).map(String) }, 'group.changed', byRef);
        return r.error ? refused(r) : { group: r.group };
      },
      groupRename: async ({ groupId, name = '' } = {}) => {
        const r = await ask({ type: 'group.rename', groupId: String(groupId || ''), name: String(name || '') }, 'group.changed', byRef);
        return r.error ? refused(r) : { group: r.group };
      },
      groupLeave: async ({ groupId } = {}) => {
        const r = await ask({ type: 'group.leave', groupId: String(groupId || '') }, 'group.left', byRef);
        return r.error ? refused(r) : { left: true };
      },
      /** Someone taken out of a group by the person who started it. */
      groupRemove: async ({ groupId, member } = {}) => {
        const r = await ask({ type: 'group.remove', groupId: String(groupId || ''), member: String(member || '') }, 'group.changed', byRef);
        return r.error ? refused(r) : { group: r.group };
      },

      /**
       * Someone reported to Rutba: why, a note, and, when the person chose
       * so, the last messages that someone sent them. Reporting does not
       * block. { reported } | { error }.
       */
      report: async ({ about, reason, note = '', groupId = null, withMessages = false } = {}) => {
        if (!REPORT_REASONS.includes(reason)) return { error: 'reason' };
        const from = readHistory(groupId ? groupKey(groupId) : String(about || ''));
        const messages = withMessages ? from.filter((m) => !m.kind && m.from === about && typeof m.text === 'string').slice(-REPORT_MESSAGES).map((m) => ({ from: m.from, text: m.text, at: m.at })) : [];
        const r = await ask({ type: 'report', about: String(about || ''), reason, note: String(note || '').slice(0, 1000), ...(groupId ? { groupId: String(groupId) } : {}), messages }, 'report.taken', byRef);
        return r.error ? { error: r.error } : { reported: r.id, messages: messages.length };
      },
    },
  };
}
