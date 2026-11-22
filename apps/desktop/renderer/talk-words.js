// What the rest of Office says of Talk, kept with the core so that Home,
// Calendar and Mail take nothing from Talk's own files: the count of what
// waits in Talk, a meeting's link read out of an event's words, what an
// invitation says of its meeting, why a meeting was refused and why a person
// is out of one (the studio's guests are told too). Talk's own
// files (apps/talk/lines.js, apps/talk/meet-lines.js) hand these on under the
// same names. Words only, no drawing, so they are tested without a window.

import { t } from '@rutba/office-ui/messages';

/** A count as a small mark shows it: nothing for none, never wider than three characters. */
export const countText = (n) => (n > 99 ? '99+' : n > 0 ? String(n) : '');

/** What waits in Talk: the messages not read and the missed calls not looked at. */
export const waitingIn = (state) => (state?.unreadTotal || 0) + (state?.missed || 0);

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

/** The notice a meeting set up to be recorded carries, the same way. */
export const recordedNotice = () => t('talk.meet-lines.meeting-recorded-people');

/** Who can join a meeting one sets up oneself, as Talk has it today: one's own contacts, and no more than it holds. */
export const ownMeetingRule = (most = 4) => t('talk.meet-lines.only-talk-contacts-join', { count: most });

/**
 * What an invitation says of its meeting: first the notice when it will be
 * streamed live or recorded, so nobody reads past it, then how to join and
 * who can.
 */
export function meetingNote(meeting) {
  const lines = [];
  if (meeting.live) lines.push(liveNotice());
  if (meeting.recorded) lines.push(recordedNotice());
  lines.push(t('talk.meet-lines.join-meeting-rutba', { link: meeting.link }), t('talk.meet-lines.only-talk-contacts'));
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
