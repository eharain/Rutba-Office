// What Talk, Calendar and Mail say about a small meeting (Office 2.28):
// its link read out of an event's words, the notice that it will be streamed
// live, what goes into an invitation, why joining was refused, and the marks
// beside a person in it. Words only, no drawing, so they are tested without
// a window.
//
// Kept plain (owner, 2026-10-21: "for chats and conference call etc build
// clean simple user interface do not complicate them like Ms teams").

import { t } from '@rutba/office-ui/messages';
import { ownMeetingRule } from '../../talk-words.js';

// What Home, Calendar and Mail use of these is the core's (talk-words.js), handed on here under the same names.
export { meetingCode, isLive, meetingIn, noticeBy, liveNotice, recordedNotice, meetingNote, ownMeetingRule, withMeeting, meetRefusal, goneText } from '../../talk-words.js';

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

/** A name for a meeting started now: the group's, or the person's own. */
export function meetingName({ group = null, me = null } = {}) {
  if (group) return group;
  const first = String(me?.name || '').split(/\s+/)[0];
  return first ? t('talk.meet-lines.meeting-with-name', { name: first }) : t('common.meeting');
}

/** The marks beside a person in a meeting, as words: muted, hand up, showing their screen. */
export function personMarks(person, { starter = false, me = false } = {}) {
  const marks = [];
  if (me) marks.push(t('common.you'));
  if (starter) marks.push(t('talk.meet-lines.organiser'));
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

/**
 * The mark of a meeting that will be streamed live or recorded, in a few
 * words, for where the whole notice has no room (a conversation's line,
 * the card that asks someone in): [{ kind: 'live'|'recorded', words }].
 */
export function airMarks(meeting) {
  const marks = [];
  if (meeting?.live) marks.push({ kind: 'live', words: t('talk.meeting.streamed-live') });
  if (meeting?.recorded) marks.push({ kind: 'recorded', words: t('talk.meeting.will-be-recorded') });
  return marks;
}

/** What is said of a meeting set up neither to be streamed nor to be recorded, before joining: it cannot become either. */
export const plainNotice = () => t('talk.meet-lines.meeting-not-set');

/** What the "streamed live" switch means where a meeting is set up, said for the way it stands. */
export const liveSwitchText = (on) => (on
  ? t('talk.meet-lines.everyone-asked-told')
  : t('talk.meet-lines.choose-now-meeting'));

/** What the word under a meeting's round button says, for the way it stands. */
export function controlWord(which, on) {
  switch (which) {
    case 'microphone': return on ? t('common.microphone') : t('talk.meet-lines.muted');
    case 'camera': return on ? t('common.camera') : t('talk.meet-lines.camera-off');
    case 'screen': return on ? t('talk.meet-lines.sharing') : t('common.screen');
    case 'hand': return on ? t('talk.meet-lines.hand-raised') : t('talk.meet-lines.hand');
    default: return '';
  }
}

/** A person who could not be connected to, in words: who, and the likely reason. */
export const notConnectedText = (name) => t('talk.meet-lines.could-not-connect', { name: name || t('common.someone') });

/** How many tiles sit side by side for this many people. */
export const tilesAcross = (n) => (n <= 1 ? 1 : 2);

/** When a planned meeting is, as its card says it; nothing for one with no time. */
export function whenText(at, lang = undefined) {
  if (!at) return '';
  return new Date(at).toLocaleString(lang, { weekday: 'short', day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' });
}
