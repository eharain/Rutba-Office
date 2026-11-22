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

/** The notice a meeting set up to be recorded carries, the same way. */
export const recordedNotice = () => t('talk.meet-lines.meeting-recorded-people');

/**
 * What someone in a meeting is told while it is on air or being recorded,
 * and whether they are in the picture; nothing while it is neither.
 * `air` is what their computer was told: { on, you, how }.
 */
export function airText(air) {
  if (!air?.on) return '';
  if (air.how === 'record') return air.you ? t('talk.meet-lines.being-recorded') : t('talk.meet-lines.recording-running-not');
  return air.you ? t('talk.meet-lines.air-people-watching') : t('talk.meet-lines.stream-air-not');
}

/** The app the person who set a meeting up takes it to: Live Broadcast for one set up to be streamed, the Recorder for one set up to be recorded, none for any other. */
export const studioFor = (meeting) => (meeting?.live ? 'broadcast' : meeting?.recorded ? 'recorder' : null);

/**
 * Why someone looking at a meeting cannot come into it now, from what they
 * were shown of it: 'full', 'locked', or null when they can. Someone already
 * in it on another computer can always bring it here.
 */
export function cannotJoin(meeting, meId) {
  const people = meeting?.people || [];
  if (people.some((p) => p.id === meId)) return null;
  if (meeting?.locked && meeting?.starter?.id !== meId) return 'locked';
  return people.length >= (meeting?.most || 4) ? 'full' : null;
}

/** Who can join a meeting one sets up oneself, as Talk has it today: one's own contacts, and no more than it holds. */
export const ownMeetingRule = (most = 4) => t('talk.meet-lines.only-talk-contacts-join', { count: most });

/**
 * Who can join a meeting with its link, in words, for whoever holds the
 * link: the contacts of the person who set it up and, for one started in a
 * group conversation, the people of that group.
 */
export function whoCanJoin(meeting, meId) {
  const count = meeting?.most || 4;
  const group = Boolean(meeting?.groupId);
  if (meeting?.starter?.id === meId) return group ? t('talk.meet-lines.only-people-group-talk', { count }) : ownMeetingRule(count);
  const name = meeting?.starter?.name || t('talk.meet-lines.person-set');
  return group ? t('talk.meet-lines.only-people-group', { name, count }) : t('talk.meet-lines.only-talk-contacts-name', { name, count });
}

/** A microphone or a camera that could not be used, in words: what it means for the others, and what to check. */
export function deviceWords(kind, fault) {
  if (!fault) return '';
  if (kind === 'camera') return fault === 'none' ? t('talk.meet-lines.no-camera-found') : t('talk.meet-lines.talk-could-not');
  return fault === 'none' ? t('talk.meet-lines.no-microphone-found') : t('talk.meet-lines.talk-could-not-use');
}

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
