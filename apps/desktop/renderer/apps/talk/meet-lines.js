// What Talk, Calendar and Mail say about a small meeting (Office 2.28):
// its link read out of an event's words, the notice that it will be streamed
// live, what goes into an invitation, why joining was refused, and the marks
// beside a person in it. Words only, no drawing, so they are tested without
// a window.
//
// Kept plain (owner, 2026-10-21: "for chats and conference call etc build
// clean simple user interface do not complicate them like Ms teams").

import { t } from '@rutba/office-ui/messages';

const CODE = '(?:live|meet)-[A-Za-z0-9_-]{20}';
const LINK = new RegExp(`https?://[^\\s/<>"']+/m/(${CODE})(?![A-Za-z0-9_-])`);
const ASKED = new RegExp(`^(?:https?://[^/\\s]+/m/|rutba-office://talk/meet/)?(${CODE})/?(?:[?#].*)?$`);

/** A meeting's code from its page's address, its rutba-office:// link, or the code itself; null for anything else. */
export function meetingCode(link) {
  const m = ASKED.exec(String(link || '').trim());
  return m ? m[1] : null;
}

/** Whether a meeting was set up to be streamed live: its code says so, wherever the link is kept. */
export const isLive = (code) => String(code || '').startsWith('live-');

/** The first meeting link in some words (an event's place and notes): { code, link, live } or null. */
export function meetingIn(...texts) {
  for (const text of texts) {
    const m = LINK.exec(String(text || ''));
    if (m) return { code: m[1], link: m[0], live: isLive(m[1]) };
  }
  return null;
}

/** The notice a meeting set up to be streamed live carries, in an invitation and before anyone joins. */
export const liveNotice = () => t('talk.meet-lines.meeting-streamed-live');

/** What an invitation says of its meeting: how to join, and the notice when it will be streamed live. */
export function meetingNote(meeting) {
  const lines = [t('talk.meet-lines.join-meeting-rutba', { link: meeting.link })];
  if (meeting.live) lines.push(liveNotice());
  return lines.join('\n');
}

/** An event's place and notes with its meeting put in: the link where the place was empty, the note under what was written. */
export function withMeeting({ location = '', description = '' } = {}, meeting) {
  const note = meetingNote(meeting);
  return {
    location: String(location).trim() ? location : meeting.link,
    description: String(description).includes(meeting.link) ? description : [String(description).trimEnd(), note].filter(Boolean).join('\n\n'),
  };
}

/** A name for a meeting started now: the group's, or the person's own. */
export function meetingName({ group = null, me = null } = {}) {
  if (group) return group;
  const first = String(me?.name || '').split(/\s+/)[0];
  return first ? t('talk.meet-lines.meeting-with-name', { name: first }) : t('common.meeting');
}

/** Why a meeting could not be set up, looked at or joined, in words. */
export function meetRefusal(code, { starter = null, most = null } = {}) {
  const who = starter?.name || '';
  switch (code) {
    case 'no-meeting': return t('talk.meet-lines.meeting-ended-link');
    case 'not-allowed': return who ? t('talk.meet-lines.only-contacts-name', { name: who }) : t('talk.meet-lines.only-contacts-person');
    case 'removed': return t('talk.meet-lines.taken-out-meeting-cannot');
    case 'locked': return t('talk.meet-lines.meeting-locked-nobody');
    case 'full': return t('talk.meet-lines.meeting-full-holds', { count: most || 4 });
    case 'not-yours': return t('talk.meet-lines.only-person-set');
    case 'nobody': return t('talk.meet-lines.person-not-meeting');
    case 'name': return t('talk.meet-lines.give-meeting-name');
    case 'time': return t('talk.meet-lines.time-cannot-used');
    case 'not-a-contact': return t('talk.meet-lines.only-own-contacts');
    case 'no-group': return t('talk.lines.no-longer-group');
    case 'too-many': return t('talk.lines.too-many-now');
    case 'unknown': return t('talk.lines.talk-server-does');
    case 'offline':
    case 'no-answer': return t('talk.invite.talk-not-connected');
    default: return t('talk.invite.talk-could-not');
  }
}

/** Why this person is out of a meeting they were in, in words. */
export function goneText(why, by = null) {
  if (why === 'removed') return by?.name ? t('talk.meet-lines.name-took-out', { name: by.name }) : t('talk.meet-lines.taken-out-meeting');
  if (why === 'elsewhere') return t('talk.meet-lines.joined-meeting-another');
  if (why === 'lost') return t('talk.meet-lines.connection-lost-meeting');
  if (why === 'ended') return by?.name ? t('talk.meet-lines.name-ended-meeting', { name: by.name }) : t('talk.meet-lines.meeting-ended');
  return '';
}

/** The marks beside a person in a meeting, as words: muted, hand up, showing their screen. */
export function personMarks(person, { starter = false, me = false } = {}) {
  const marks = [];
  if (me) marks.push(t('common.you'));
  if (starter) marks.push(t('talk.meet-lines.set-meeting'));
  if (person?.hand) marks.push(t('talk.meet-lines.hand-raised'));
  if (person?.sharing) marks.push(t('talk.meet-lines.showing-their-screen'));
  if (person?.muted) marks.push(t('talk.meet-lines.muted'));
  return marks.join(' · ');
}

/** The line a group's conversation keeps when a meeting starts in it. */
export function meetingLine(entry, meId) {
  const name = entry?.meeting?.name || t('common.meeting');
  return entry?.by?.id === meId ? t('talk.meet-lines.started-meeting-title', { title: name }) : t('talk.meet-lines.name-started-meeting', { name: entry?.by?.name || t('common.someone'), title: name });
}

/** How many tiles sit side by side for this many people. */
export const tilesAcross = (n) => (n <= 1 ? 1 : 2);

/** When a planned meeting is, as its card says it; nothing for one with no time. */
export function whenText(at, lang = undefined) {
  if (!at) return '';
  return new Date(at).toLocaleString(lang, { weekday: 'short', day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' });
}
