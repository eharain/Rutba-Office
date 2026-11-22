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
const REPORT_MESSAGES = 20;
export const REPORT_REASONS = ['spam', 'harassment', 'impersonation', 'other'];

/**
 * @param {object} o
 * @param {() => object|null} o.me the person signed in
 * @param {(key: string) => object[]} o.readHistory a conversation's messages
 * @param {(key: string, list: object[]) => void} o.writeHistory
 * @param {(key: string, entry: object) => void} o.keep one message kept, in place of an earlier one with its id
 * @param {(name: string, fallback: any) => any} o.readFile a small file of the person's own, beside their messages
 * @param {(name: string, value: any) => void} o.writeFile
 * @param {(msg: object) => void} o.send to the introducer; throws when not connected
 * @param {(msg: object, answer: string, match?: Function) => Promise<object>} o.ask
 * @param {() => number} [o.now]
 */
export function createConversations({ me, readHistory, writeHistory, keep, readFile, writeFile, send, ask, now = Date.now }) {
  let groups = [];
  /** Calls ringing here that nobody has answered yet: callId -> { from, video }. */
  const ringing = new Map();
  /** Which conversation a message sent from here went to: id -> key, while its marks are still coming. */
  const sentTo = new Map();

  const marks = () => {
    const m = readFile('marks.json', null);
    return { unread: m?.unread && typeof m.unread === 'object' ? m.unread : {}, calls: Array.isArray(m?.calls) ? m.calls : [] };
  };
  const save = (m) => { try { writeFile('marks.json', m); } catch { /* the counts are kept next time */ } };
  const trySend = (msg) => { try { send(msg); } catch { /* said again when the connection is back */ } };

  /** A message that came in: one more unread in its conversation. */
  const unread = (key, n = 1) => {
    const m = marks();
    m.unread[key] = (m.unread[key] || 0) + n;
    save(m);
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

  /** A call that reached nobody here: listed once, however it was learned of, and a line in that person's conversation. */
  const missed = ({ callId, from, video = false, why = 'unanswered', at = now() }) => {
    if (!from?.id) return;
    const m = marks();
    if (m.calls.some((c) => c.id === callId)) return;
    m.calls = [...m.calls, { id: callId, with: from.id, name: from.name || from.email || '', video: Boolean(video), why, at, seen: false }].slice(-CALLS_CAP);
    save(m);
    keep(from.id, { id: `call:${callId}`, kind: 'missed-call', from: from.id, video: Boolean(video), at });
  };

  /** A line in a group's conversation saying what changed in it. */
  const noteEvent = (group, event) => {
    if (!event?.kind) return;
    keep(groupKey(group.id), { id: `event:${event.kind}:${event.at || now()}:${(event.who || []).map((w) => w.id).join(',')}`, kind: 'group-event', event, at: event.at || now() });
  };

  const setGroup = (group) => { groups = [...groups.filter((g) => g.id !== group.id), group].sort((a, b) => a.created - b.created); };

  return {
    /** For the state every window is told: the groups, the unread counts, and how many missed calls are not yet looked at. */
    snapshot() {
      if (!me()) return { groups: [], unread: {}, unreadTotal: 0, missed: 0 };
      const m = marks();
      const counts = Object.fromEntries(Object.entries(m.unread).filter(([, n]) => n > 0));
      return { groups, unread: counts, unreadTotal: Object.values(counts).reduce((a, b) => a + b, 0), missed: m.calls.filter((c) => !c.seen).length };
    },

    forget() { groups = []; ringing.clear(); sentTo.clear(); },

    /** What the introducer said. True when the message was one of these, so every window's state is told again. */
    handle(msg) {
      switch (msg.type) {
        case 'welcome':
          groups = [];
          return false;
        case 'messages':
          for (const m of msg.messages || []) unread(m.from);
          return false;
        case 'msg.sent':
          if (sentTo.has(msg.id)) mark(sentTo.get(msg.id), msg.id, (e) => (e.state === 'arrived' ? null : { state: 'sent' }));
          return true;
        case 'groups':
          groups = (msg.groups || []).slice().sort((a, b) => a.created - b.created);
          return true;
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
            keep(groupKey(m.group), { id: m.id, group: m.group, from: m.from, text: m.text, at: m.at, mine: false });
            unread(groupKey(m.group));
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
          sentTo.delete(msg.id);
          return true;
        case 'error':
          // A message the introducer would not take: it says so on the message, in place of waiting for ever.
          if ((msg.for === 'msg.send' || msg.for === 'group.send') && sentTo.has(msg.id)) {
            mark(sentTo.get(msg.id), msg.id, () => ({ state: 'failed' }));
            sentTo.delete(msg.id);
            return true;
          }
          return false;
        case 'notices': {
          const ids = [];
          for (const n of msg.notices || []) {
            ids.push(n.id);
            if (n.kind === 'missed-call') missed({ callId: n.callId, from: n.from, video: n.video, why: n.why, at: n.at });
            else if (n.kind === 'delivered' && n.group) {
              // Arrived for one more person; for everyone it was sent to, it has arrived.
              for (const id of n.ids || []) {
                mark(groupKey(n.group), id, (e) => {
                  const arrived = [...new Set([...(e.arrived || []), n.by])];
                  const all = Array.isArray(e.to) && e.to.length > 0 && e.to.every((p) => arrived.includes(p));
                  return { arrived, ...(all ? { state: 'arrived' } : {}) };
                });
              }
            } else if (n.kind === 'delivered') {
              for (const id of n.ids || []) { mark(n.by, id, () => ({ state: 'arrived' })); sentTo.delete(id); }
            }
          }
          if (ids.length) trySend({ type: 'notice.ack', ids });
          return true;
        }
        case 'call.incoming':
          ringing.set(msg.callId, { from: msg.from, video: Boolean(msg.video) });
          return false;
        case 'call.taken':
          ringing.delete(msg.callId);
          return false;
        case 'call.ended':
          // It stopped ringing and nobody here had answered.
          if (ringing.has(msg.callId)) { missed({ callId: msg.callId, ...ringing.get(msg.callId) }); ringing.delete(msg.callId); return true; }
          return false;
        default:
          return false;
      }
    },

    /** What a window sends to the introducer: an answer here means the call was not missed; busy in another call means it was (true then). */
    sending(msg) {
      if (msg.type === 'call.answer') ringing.delete(msg.callId);
      else if (msg.type === 'call.decline') {
        const call = ringing.get(msg.callId);
        ringing.delete(msg.callId);
        if (call && msg.reason === 'busy') { missed({ callId: msg.callId, ...call, why: 'busy' }); return true; }
      }
      return false;
    },

    /** A message of mine to a person or a group, on its way: marked as sending until the introducer says it has it. */
    sentToPerson(id, key) {
      sentTo.set(id, key);
      if (sentTo.size > 500) sentTo.delete(sentTo.keys().next().value);
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

      /** A group conversation made of three or more people: { group } | { error, most? }. */
      groupCreate: async ({ name = '', members = [] } = {}) => {
        const r = await ask({ type: 'group.create', name: String(name || ''), members: [].concat(members || []).map(String) }, 'group.created');
        return r.error ? { error: r.error } : { group: r.group };
      },
      groupAdd: async ({ groupId, members = [] } = {}) => {
        const r = await ask({ type: 'group.add', groupId: String(groupId || ''), members: [].concat(members || []).map(String) }, 'group.changed');
        return r.error ? { error: r.error } : { group: r.group };
      },
      groupRename: async ({ groupId, name = '' } = {}) => {
        const r = await ask({ type: 'group.rename', groupId: String(groupId || ''), name: String(name || '') }, 'group.changed');
        return r.error ? { error: r.error } : { group: r.group };
      },
      groupLeave: async ({ groupId } = {}) => {
        const r = await ask({ type: 'group.leave', groupId: String(groupId || '') }, 'group.left', (m) => m.groupId === String(groupId || ''));
        return r.error ? { error: r.error } : { left: true };
      },
      /** Someone taken out of a group by the person who started it. */
      groupRemove: async ({ groupId, member } = {}) => {
        const r = await ask({ type: 'group.remove', groupId: String(groupId || ''), member: String(member || '') }, 'group.changed');
        return r.error ? { error: r.error } : { group: r.group };
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
        const r = await ask({ type: 'report', about: String(about || ''), reason, note: String(note || '').slice(0, 1000), ...(groupId ? { groupId: String(groupId) } : {}), messages }, 'report.taken');
        return r.error ? { error: r.error } : { reported: r.id, messages: messages.length };
      },
    },
  };
}
