// What Talk, Calendar and Mail say about a small meeting (Office 2.28):
// its link read out of an event's words, the notice that it will be streamed
// live, what goes into an invitation, why joining was refused, and the marks
// beside a person in it. Words only, no drawing, so they are tested without
// a window.
//
// Kept plain (owner, 2026-10-21: "for chats and conference call etc build
// clean simple user interface do not complicate them like Ms teams").

import { t } from '@rutba/office-ui/messages';

// What Home, Calendar and Mail use of these is the core's (talk-words.js), handed on here under the same names.
export { meetingCode, isLive, meetingIn, liveNotice, recordedNotice, meetingNote, withMeeting, meetRefusal, goneText } from '../../talk-words.js';

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
