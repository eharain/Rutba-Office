// What Talk says is worth a notification (the owner, 22 Nov 2026,
// D-2026-11-22-23: an icon by the clock to keep an eye on incoming calls,
// meeting reminders and chats, with Windows notifications).
//
// Talk shows no Windows notification, no taskbar flash and no icon by the
// clock itself. The notifier in the main process does all of that, for every
// app; this is what Talk hands it. Each thing is broadcast once, as
// 'talk:notify', beside the 'talk:state' and 'talk:event' every window
// already hears:
//
//   { kind: 'call', id, from: { id, name }, video }
//       A call is ringing here. `id` is the call's.
//   { kind: 'call-over', id, why }
//       That call no longer rings here. why: 'answered' (here), 'declined'
//       (here), 'elsewhere' (answered on another of the person's computers),
//       'missed' (it stopped, or another call was on), 'ended' (signed out,
//       or the connection was made again).
//   { kind: 'missed', id, callId, from: { id, name }, video, why, at, count }
//       A call was missed: it rang here unanswered, came during another
//       call, or was made while this computer was off. `id` is the caller's
//       id: someone who calls again and again is the one entry, told again
//       with a higher `count`, until the missed calls have been looked at.
//   { kind: 'message', id, from: { id, name }, group: { id, name } | null, text, at }
//       A message someone wrote, to the person or in a group. Told once: a
//       message the introducer gives again is not told again.
//   { kind: 'meeting', id, from: { id, name } | null, name, live, recorded, group: { id, name } | null }
//       The person is asked into a meeting that is on now. `id` is the
//       meeting's code (Talk opens on it with ?meet=<id>); `live` says it
//       will be streamed and `recorded` that it will be recorded, which the
//       notification must say too.
//   { kind: 'meeting-over', id, why }
//       No longer offered. why: 'joined', 'dismissed', 'ended'.
//
// How many messages wait and how many calls were missed are in Talk's state
// (`unreadTotal`, `missed`), told with every change.

export const NOTIFY_EVENT = 'talk:notify';
export const NOTIFY_KINDS = ['call', 'call-over', 'missed', 'message', 'meeting', 'meeting-over'];

/** The most of a message's words handed on: a notification shows a line or two. */
export const NOTIFY_TEXT_MOST = 280;

/**
 * @param {object} o
 * @param {(event: string, payload: object) => void} o.broadcast to every window, and to whatever listens in the main process
 * @param {(id: string) => { id: string, name: string }} o.person a person by id, with the name this copy knows them by
 * @param {(id: string) => { id: string, name: string }} o.group a group by id, with its name, or the others in it by name
 * @returns {(n: object) => void} takes what talk-conversations.js and talk-meetings.js say, and tells it with names
 */
export function createNotify({ broadcast, person, group }) {
  const who = (from) => {
    if (from && typeof from === 'object') {
      const known = person(from.id);
      return { id: from.id, name: from.name || known.name || from.email || '' };
    }
    return from ? person(from) : null;
  };
  return (n) => {
    if (!n || !NOTIFY_KINDS.includes(n.kind) || !n.id) return;
    let out = n;
    if (n.kind === 'call' || n.kind === 'missed') out = { ...n, from: who(n.from) };
    else if (n.kind === 'message') out = { ...n, from: who(n.from), group: n.group ? group(n.group) : null, text: String(n.text ?? '').slice(0, NOTIFY_TEXT_MOST) };
    else if (n.kind === 'meeting') out = { ...n, from: who(n.from), group: n.group ? group(n.group) : null };
    try { broadcast?.(NOTIFY_EVENT, out); } catch { /* a notification not shown loses nothing kept */ }
  };
}

const MEETING_LINK = /https?:\/\/[^\s/<>"']+\/m\/((?:live|meet)-[A-Za-z0-9_-]{20})(?![A-Za-z0-9_-])/;

/**
 * The Talk meeting an event carries, for a reminder that a meeting is about
 * to start: Calendar keeps a planned meeting's link in the event's place and
 * notes (calendar-talk.js), and whatever reminds of the event can offer to
 * join it. Given those words, the first meeting link in them as
 * { id, live }, or null. `id` is what Talk opens on (?meet=<id>); a reminder
 * of one with `live` must say it will be streamed.
 */
export function meetingInEvent(...texts) {
  for (const text of texts) {
    const m = MEETING_LINK.exec(String(text ?? ''));
    if (m) return { id: m[1], live: m[1].startsWith('live-') };
  }
  return null;
}
