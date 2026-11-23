// Small meetings (Office 2.27, the owner's go-ahead of 2026-11-22 on the
// Talk, Meet and Live plan): a call of up to four, each computer connected
// to each other. The introducer knows that a meeting exists, who may join it
// and who is in it, and passes the set-up between each pair; this is what
// Talk's service asks it and keeps of its answers.
//
// - A meeting has a name, a link and, if it was planned, a time. Its code is
//   in its link. One set up to be streamed live carries `live-` in the code
//   itself (D-2026-11-22-19: chosen at set-up, never while it runs), so the
//   link says so wherever it is kept and nothing here can turn it on or off.
//   One set up to be recorded (an interview in the Recorder) says so the same
//   way to everyone shown it: chosen at set-up, never after.
// - The meetings someone has asked this person into, which are on now, are
//   told to every window, so Talk can offer to join.
// - A meeting started in a group conversation leaves a line there, and one
//   a contact asked this person into leaves a line in the conversation with
//   that contact, so "Not now" does not lose the way in. Once the
//   meeting is known to be over (it was ended while this person was in it,
//   or the introducer no longer knows it), that is kept, so the line stops
//   offering to join.
//
// The meeting itself (sound, picture, its messages) is in the window that
// holds it (renderer/apps/talk/mesh.js) and never comes through here.

import { groupKey } from './talk-conversations.js';

/** A meeting's code: `live-` when it was set up to be streamed live, `meet-` when not. */
export const MEETING_CODE = /^(live|meet)-[A-Za-z0-9_-]{20}$/;
const OVER_FILE = 'meetings.json';
const OVER_CAP = 200;

/**
 * A meeting's code from what someone has: its page's address
 * (https://introducer.rutba.io/m/<code>), the rutba-office:// link, or the
 * code itself. Only the code is ever used: Office asks the introducer it is
 * signed in to, never the address a link names.
 */
export function meetingCode(link) {
  const text = String(link || '').trim();
  const m = /^(?:https?:\/\/[^/\s]+\/m\/|rutba-office:\/\/talk\/meet\/)?((?:live|meet)-[A-Za-z0-9_-]{20})\/?(?:[?#].*)?$/.exec(text);
  return m ? m[1] : null;
}

/**
 * @param {object} o
 * @param {() => object|null} o.me the person signed in
 * @param {(msg: object, answer: string, match?: (m: object, ref: string) => boolean) => Promise<object>} o.ask
 * @param {(key: string, entry: object) => void} o.keep a line kept in a conversation
 * @param {(name: string, fallback: any) => any} [o.readFile] a small file of the person's own, beside their messages
 * @param {(name: string, value: any) => void} [o.writeFile]
 * @param {() => number} [o.now]
 */
export function createMeetings({ me, ask, keep, notify = () => {}, readFile = (name, fallback) => fallback, writeFile = () => {}, now = Date.now }) {
  /** Meetings on now that someone asked this person into: [{ meeting, by, at }]. */
  let asked = [];
  /** The meeting this copy of Office is in, by its code, or null. */
  let inside = null;

  /** The meetings known to be over, by their codes: the last of them, so the file stays small. */
  const overIds = () => {
    const f = readFile(OVER_FILE, null);
    return Array.isArray(f?.over) ? f.over.filter((id) => typeof id === 'string') : [];
  };
  const over = (id) => {
    if (!id || !MEETING_CODE.test(id)) return;
    const list = overIds();
    if (list.includes(id)) return;
    try { writeFile(OVER_FILE, { over: [...list, id].slice(-OVER_CAP) }); } catch { /* its line offers to join once more, and says then that it has ended */ }
  };

  /** No longer offered; `why` is told to whatever shows a notification for it (talk-notify.js). */
  const drop = (id, why = null) => {
    const had = asked.some((a) => a.meeting.id === id);
    asked = asked.filter((a) => a.meeting.id !== id);
    if (had && why) notify({ kind: 'meeting-over', id, why });
  };
  /**
   * A meeting's line, once: in its group's conversation when it started in
   * one; else, for a meeting this person was `asked` into, in their
   * conversation with whoever asked.
   */
  const line = (meeting, by, { asked: wasAsked = false } = {}) => {
    if (!meeting?.id || !by?.id) return;
    const key = meeting.groupId ? groupKey(meeting.groupId) : wasAsked ? by.id : null;
    if (!key) return;
    try {
      keep(key, { id: `meet:${meeting.id}`, kind: 'meeting', meeting: { id: meeting.id, name: meeting.name, live: Boolean(meeting.live), recorded: Boolean(meeting.recorded), link: meeting.link }, by: { id: by.id, name: by.name || '' }, at: now() });
    } catch { /* the meeting goes on without its line */ }
  };
  /** A refusal as a window is given it: why, and what the introducer said with it. */
  const refused = (r) => ({ error: r.error, ...(r.most != null ? { most: r.most } : {}), ...(r.starter ? { starter: r.starter } : {}) });
  // The answer to this question and no other: what a meeting's people are told of it carries no ref.
  const byRef = (m, ref) => m.ref === ref;

  return {
    /** For the state every window is told. */
    snapshot: () => ({ meetings: asked, meeting: inside }),

    forget() { for (const a of [...asked]) drop(a.meeting.id, 'ended'); inside = null; },

    /** The codes of the meetings known to be over. */
    over: () => new Set(overIds()),

    /** What the introducer said. */
    handle(msg) {
      switch (msg.type) {
        case 'welcome':
          // A new connection: the introducer has this copy in no meeting, and says again what is on.
          for (const a of [...asked]) drop(a.meeting.id, 'ended');
          inside = null;
          break;
        case 'meet.invited': {
          if (!msg.meeting?.id) break;
          const by = msg.by || msg.meeting.starter;
          drop(msg.meeting.id);
          asked = [...asked, { meeting: msg.meeting, by, at: now() }];
          line(msg.meeting, by, { asked: true });
          notify({ kind: 'meeting', id: msg.meeting.id, from: by ? { id: by.id, name: by.name || '' } : null, name: msg.meeting.name || '', live: Boolean(msg.meeting.live), recorded: Boolean(msg.meeting.recorded), group: msg.meeting.groupId || null });
          break;
        }
        case 'meet.created':
          line(msg.meeting, me());
          break;
        case 'meet.joined':
          if (msg.meeting?.id) { inside = msg.meeting.id; drop(msg.meeting.id, 'joined'); }
          break;
        case 'meet.left':
        case 'meet.removed':
          if (inside === msg.id) inside = null;
          break;
        case 'meet.ended':
          over(msg.id);
          drop(msg.id, 'ended');
          if (inside === msg.id) inside = null;
          break;
        default:
          break;
      }
    },

    api: {
      /**
       * A meeting set up: its name, its time if it is planned, whether it
       * will be streamed live, whether it will be recorded, the group
       * conversation it is in, the contacts asked. With `now`, it starts now
       * and those people are told at once. With `carried: 'server'` it is
       * to be held on a Talk Server, for up to `most` people; an introducer
       * that holds no meetings on servers sets up a small one all the same,
       * and what it answers says which it is. { meeting } | { error }.
       */
      meetCreate: async ({ name = '', at = null, live = false, recorded = false, groupId = null, invite = [], now: tell = false, carried = null, most = null } = {}) => {
        const held = carried === 'server' ? { carried: 'server', ...(Number.isInteger(most) && most >= 2 ? { most } : {}) } : {};
        const r = await ask({ type: 'meet.create', name: String(name || ''), ...(at != null ? { at: Number(at) } : {}), live: live === true, ...(recorded === true ? { recorded: true } : {}), ...(groupId ? { groupId: String(groupId) } : {}), ...held, invite: [].concat(invite || []).map(String), tell: tell === true }, 'meet.created', byRef);
        return r.error ? refused(r) : { meeting: r.meeting };
      },

      /** More contacts asked into a meeting that is on, by the person who set it up: { told } | { error }. */
      meetInvite: async ({ id, invite = [] } = {}) => {
        const r = await ask({ type: 'meet.invite', id: String(id || ''), invite: [].concat(invite || []).map(String) }, 'meet.asked', byRef);
        return r.error ? refused(r) : { told: Number(r.told) || 0 };
      },

      /** A meeting looked at before joining, by its link or code: { meeting } | { error, starter?, id }. */
      meetPeek: async ({ link } = {}) => {
        const id = meetingCode(link);
        if (!id) return { error: 'no-meeting' };
        const r = await ask({ type: 'meet.peek', id }, 'meet.info', byRef);
        if (r.error === 'no-meeting') over(id);
        return r.error ? { ...refused(r), id } : { meeting: r.meeting };
      },

      /** Into a meeting: { meeting }, its `people` those already there, or { error, most?, starter? }. */
      meetJoin: async ({ id } = {}) => {
        const code = meetingCode(id);
        if (!code) return { error: 'no-meeting' };
        const r = await ask({ type: 'meet.join', id: code }, 'meet.joined', byRef);
        if (r.error === 'no-meeting') over(code);
        return r.error ? refused(r) : { meeting: r.meeting };
      },

      meetLeave: async ({ id } = {}) => {
        const r = await ask({ type: 'meet.leave', id: String(id || '') }, 'meet.left', byRef);
        return r.error ? { error: r.error } : { left: true };
      },

      /** Locked: nobody new comes in. Only for the person who set the meeting up. */
      meetLock: async ({ id, locked = true } = {}) => {
        const r = await ask({ type: 'meet.lock', id: String(id || ''), locked: locked !== false }, 'meet.changed', byRef);
        return r.error ? { error: r.error } : { meeting: r.meeting };
      },

      /** Someone taken out, who cannot come back into it. Only for the person who set the meeting up. */
      meetRemove: async ({ id, who } = {}) => {
        const r = await ask({ type: 'meet.remove', id: String(id || ''), who: String(who || '') }, 'meet.changed', byRef);
        return r.error ? { error: r.error } : { meeting: r.meeting };
      },

      /** The meeting ended for everyone, and its link with it. Only for the person who set it up. */
      meetEnd: async ({ id } = {}) => {
        const r = await ask({ type: 'meet.end', id: String(id || '') }, 'meet.ended', byRef);
        return r.error ? { error: r.error } : { ended: true };
      },

      /** "Not now": a meeting one was asked into is no longer offered. */
      meetDismiss: ({ id } = {}) => { drop(String(id || ''), 'dismissed'); return { dismissed: true }; },
    },
  };
}
