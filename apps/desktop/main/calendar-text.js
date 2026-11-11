// The words of the messages the calendar sends: an invitation, a change to
// one, a cancellation, an answer, and what a room's mailbox says back.
//
// The window that asks for a message to go hands these over in its own
// language (renderer/apps/calendar-words.js, the same keys through the
// catalogue), with that language's tag for the dates; the last ones handed
// over are kept for a room's mailbox, which answers with no window open.
// English, below, is what is said before any window has spoken.

export const ENGLISH = {
  language: 'en',
  inviteSubject: 'Invitation: {summary}',
  updateSubject: 'Updated invitation: {summary}',
  cancelSubject: 'Cancelled: {summary}',
  acceptedSubject: 'Accepted: {summary}',
  tentativeSubject: 'Tentative: {summary}',
  declinedSubject: 'Declined: {summary}',
  when: 'When: {when}',
  where: 'Where: {where}',
  organizer: 'Organizer: {name}',
  invited: 'Invited: {names}',
  allDay: '{date}, all day',
  attached: 'The invitation is attached: a calendar program shows it with Accept, Tentative and Decline.',
  updated: 'This meeting has changed. The new details are above and attached.',
  cancelled: 'This meeting is cancelled. A calendar program takes it off the calendar with the file attached.',
  acceptedBody: '{name} accepted.',
  tentativeBody: '{name} answered tentative.',
  declinedBody: '{name} declined.',
  roomBooked: '{room} is booked for {when}.',
  roomClash: '{room} is not free for {when}.',
  roomFreeThatDay: 'Free that day: {times}.',
  roomNoneFree: 'It has no free time left that day.',
  roomFreeSubject: 'Free times: {room}',
  roomFreeOn: '{room}, {date}: free {times}.',
  roomBusyAllDay: '{room}, {date}: no free time.',
  roomAsk: 'To ask when {room} is free, write to it with a date, such as "free on 12 March?".',
  roomOffice: 'This answer comes from Rutba Office, which looks after this room\'s mailbox while it is open.',
  freeBusySubject: 'When are you free? {from} to {to}',
  freeBusyBody: 'This asks your calendar when you are busy between {from} and {to}. A calendar program answers with busy times only, never what fills them.',
};

/** A message with its {placeholders} filled. */
export function fill(template, vars = {}) {
  return String(template ?? '').replace(/\{(\w+)\}/g, (m, k) => (vars[k] != null ? String(vars[k]) : m));
}

/** The words to use: those handed over, English for any one missing. */
export function wordsFrom(given) {
  const out = { ...ENGLISH };
  for (const [k, v] of Object.entries(given || {})) if (typeof v === 'string' && v.trim() && k in ENGLISH) out[k] = v;
  return out;
}

/** An event's time, as a person reads it: in the zone it was made in, that zone named when it is not the reader's. */
export function whenText(event, words, { zone = null } = {}) {
  const lang = words.language || 'en';
  if (!event?.start) return '';
  if (event.allDay) {
    const date = new Date(event.start.at).toLocaleDateString(lang, { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC' });
    return fill(words.allDay, { date });
  }
  const timeZone = zone || event.start.tzid || undefined;
  const opts = (o) => {
    try {
      return new Intl.DateTimeFormat(lang, { ...o, timeZone });
    } catch {
      return new Intl.DateTimeFormat(lang, o);
    }
  };
  const day = opts({ weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
  const time = opts({ hour: '2-digit', minute: '2-digit' });
  const start = new Date(event.start.at);
  const end = new Date(event.end?.at ?? event.start.at);
  const sameDay = day.format(start) === day.format(end);
  const span = sameDay ? `${day.format(start)}, ${time.format(start)} – ${time.format(end)}` : `${day.format(start)} ${time.format(start)} – ${day.format(end)} ${time.format(end)}`;
  return timeZone ? `${span} (${timeZone})` : span;
}

const person = (p) => (p ? (p.name ? `${p.name} <${p.email}>` : p.email) : '');

/** The plain text of an invitation, a change or a cancellation. */
export function inviteText(event, words, { method = 'REQUEST', update = false } = {}) {
  const lines = [event.summary || ''];
  lines.push('');
  lines.push(fill(words.when, { when: whenText(event, words) }));
  if (event.location) lines.push(fill(words.where, { where: event.location }));
  if (event.organizer?.email) lines.push(fill(words.organizer, { name: person(event.organizer) }));
  const invited = (event.attendees || []).filter((a) => a.email).map((a) => a.name || a.email);
  if (invited.length) lines.push(fill(words.invited, { names: invited.join(', ') }));
  if (event.description) lines.push('', event.description);
  lines.push('', method === 'CANCEL' ? words.cancelled : update ? words.updated : words.attached);
  return lines.join('\n');
}

/** The subject of each kind of message. */
export function subjectOf(kind, event, words) {
  const key = { REQUEST: 'inviteSubject', UPDATE: 'updateSubject', CANCEL: 'cancelSubject', ACCEPTED: 'acceptedSubject', TENTATIVE: 'tentativeSubject', DECLINED: 'declinedSubject' }[kind] || 'inviteSubject';
  return fill(words[key], { summary: event?.summary || '' });
}

/** The words of an answer. */
export function answerText(partstat, name, words) {
  const key = { ACCEPTED: 'acceptedBody', TENTATIVE: 'tentativeBody', DECLINED: 'declinedBody' }[partstat] || 'acceptedBody';
  return fill(words[key], { name });
}
