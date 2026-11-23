// What Talk, Calendar and Mail say about a small meeting (Office 2.27):
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

/**
 * What this window asks of the shell's capture while it holds a call or a
 * meeting (capture.withhold): { on, protect }. `on`: it is no source for a
 * screenshot, a recording or a broadcast. `protect`: its contents are kept
 * out of a capture of the whole screen as well, for a call and for a meeting
 * not set up to be streamed or recorded; a meeting not yet known (still
 * coming in) is taken as such.
 */
export function withholdOf({ call = null, room = false, meeting = null } = {}) {
  const inCall = Boolean(call) && call !== 'ended';
  if (!inCall && !room) return { on: false, protect: false };
  return { on: true, protect: inCall || !(meeting?.live || meeting?.recorded) };
}

/** The app the person who set a meeting up takes it to: Live Broadcast for one set up to be streamed, the Recorder for one set up to be recorded, none for any other. */
export const studioFor = (meeting) => (meeting?.live ? 'broadcast' : meeting?.recorded ? 'recorder' : null);

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
  if (person?.unreadable) marks.push(t('talk.meet-lines.cannot-heard'));
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

/** How many tiles sit side by side for this many people: a small meeting as it always was, a larger one in three and then four. */
export const tilesAcross = (n) => (n <= 1 ? 1 : n <= 4 ? 2 : n <= 9 ? 3 : 4);

/* ── a meeting held on a Talk Server ──────────────────────────────── */

// Whether a meeting is held on a Talk Server (isCarried), and why someone
// cannot come into one (cannotJoin), are the core's, beside the room that
// reads them (renderer/meeting/meet-rules.js), and handed on here.
export { cannotJoin, isCarried } from '../../meeting/meet-rules.js';

/** The most contacts one set-up asks into a meeting (the introducer takes no more at once). */
const ASKED_MOST = 50;

/**
 * What the person setting a meeting up can do with Talk Servers, from what
 * Talk's service said (serverMine): { can, most, asked, why }. `can`: the
 * Talk service knows a meeting held on a server (`knows`) and says a server
 * would hold this person's meeting now (`carry`), their own or one that is
 * not theirs (somebody chose them, or it is their organisation's, or the
 * Talk service's): which, is never said. `most`: how many people the
 * meeting would hold there; `asked`: how many contacts can be asked at
 * set-up. `why`: when they have servers of their own and no server can
 * hold a meeting now, the first one's reason in words. Everybody else,
 * and everybody at a Talk service that knows no such meeting or did not
 * answer, gets { can: false } and nothing to say: nothing is offered that
 * leads nowhere.
 */
export function serverChoice(mine) {
  const none = { can: false, most: 0, asked: 0, why: '' };
  if (!mine || mine.error || mine.knows !== true) return none;
  const servers = Array.isArray(mine.servers) ? mine.servers : [];
  const carry = mine.carry;
  if (carry?.can !== true) return { ...none, why: servers.length ? serverWhyText(servers[0]) : '' };
  // A server that holds no more than a small meeting does: nothing to choose.
  const most = Number.isInteger(carry.maxPeople) && carry.maxPeople > 4 ? carry.maxPeople : 0;
  if (!most) return none;
  return { can: true, most, asked: Math.min(ASKED_MOST, most - 1), why: '' };
}

/** Why a server of the person's own cannot hold a meeting now, in a sentence; `s` is { name, why }. */
export function serverWhyText(s) {
  const name = s?.name || t('talk.meet-lines.your-talk-server');
  switch (s?.why) {
    case 'not-answering': return t('talk.meet-lines.name-not-answering', { name });
    case 'not-reachable': return t('talk.meet-lines.name-cannot-reached', { name });
    case 'starting': return t('talk.meet-lines.name-still-starting', { name });
    case 'meetings-off': return t('talk.meet-lines.name-set-hold', { name });
    case 'full': return t('talk.meet-lines.name-holding-many', { name });
    case 'outside-hours': return t('talk.meet-lines.name-outside-hours', { name });
    case 'paused': return t('talk.meet-lines.name-is-paused', { name });
    case 'outdated': return t('talk.meet-lines.name-must-updated', { name });
    case 'pending': return t('talk.meet-lines.name-waiting-code', { name });
    default: return t('talk.meet-lines.name-cannot-hold', { name });
  }
}

/** What the "held on a Talk Server" switch means where a meeting is set up, said for the way it stands; `count` is how many people it holds. */
export const serverSwitchText = (on, count) => (on
  ? t('talk.meet-lines.count-people-sound', { count })
  : t('talk.meet-lines.more-four-people'));

/** What a meeting held on a server says of itself where it is set up, in place of "a meeting holds four people". */
export const serverLead = (count) => t('talk.meet-lines.meeting-holds-count', { count });

/** Said when a meeting asked for on a server was set up as a small one: the Talk service holds none on servers. */
export const notCarriedText = () => t('talk.meet-lines.talk-service-cannot');

/**
 * What someone looking at a meeting held on a server is told before joining:
 * how it is held and, while the person who set it up is not in it, that they
 * can come in once that person is.
 */
export function serverSeenText(meeting, meId) {
  const held = t('talk.meet-lines.held-talk-server');
  const starter = meeting?.starter;
  if (!starter?.id || starter.id === meId || (meeting.people || []).some((p) => p.id === starter.id)) return held;
  return [held, starter.name ? t('talk.meet-lines.name-lets-people', { name: starter.name }) : t('talk.meet-lines.person-set-lets')].join(' ');
}

/** How far this person is into a meeting held on a server, in words; nothing once they are in. `starter` is who set it up. */
export function serverStepText(step, starter = null) {
  const name = starter?.name || '';
  switch (step) {
    case 'starting': return t('talk.meet-lines.finding-talk-server');
    case 'waiting': return name ? t('talk.meet-lines.waiting-name-lets', { name }) : t('talk.meet-lines.waiting-person-set');
    case 'asking': return name ? t('talk.meet-lines.asking-name-let', { name }) : t('talk.meet-lines.asking-let');
    case 'entering': return t('talk.meet-lines.connecting-meetings-talk');
    default: return '';
  }
}

/** Said all the while someone is in a meeting held on a server, while everything they receive is encrypted by the others' apps. */
export const serverNotice = () => t('talk.meet-lines.held-talk-server-sound');

/**
 * Sound or picture that does not open on this computer, in words: `all`
 * when nothing anybody sends opens (this copy holds another key than the
 * others), otherwise somebody else's does not.
 */
export const unreadableText = (all) => (all
  ? t('talk.meet-lines.cannot-hear-see')
  : t('talk.meet-lines.somebody-meeting-cannot'));

/**
 * Why a meeting held on a server could not be started, come into or stayed
 * in, in a sentence. `fault` is what the meeting kept of it (renderer/meeting/meet-room.js):
 * { why, mine, own? }; `starter` is who set it up.
 */
export function serverFaultText(fault, starter = null) {
  const name = starter?.name || '';
  switch (fault?.why) {
    case 'offline':
    case 'no-answer': return t('talk.invite.talk-not-connected');
    case 'sign-in': return t('talk.meet-lines.sign-talk-again');
    case 'no-service': return t('talk.meet-lines.talk-service-no');
    case 'no-server': {
      if (!fault.mine) return t('talk.meet-lines.no-talk-server-take');
      const own = (fault.own || []).map(serverWhyText);
      return [t('talk.meet-lines.no-talk-server'), ...own].join(' ');
    }
    case 'server-gone':
    case 'not-serving': return t('talk.meet-lines.meetings-talk-server');
    case 'unreachable': return t('talk.meet-lines.talk-could-not-reach');
    case 'expired':
    case 'used':
    case 'stale': return t('talk.meet-lines.pass-meeting-ran');
    case 'no-pass': return fault.mine ? t('talk.meet-lines.talk-server-did-not-take-meeting') : t('talk.meet-lines.not-let-meeting');
    case 'no-meeting':
    case 'meeting-taken': return t('talk.meet-lines.meeting-over-talk');
    case 'pin': return t('talk.meet-lines.talk-could-not-confirm');
    case 'certificate': return t('talk.meet-lines.talk-servers-certificate');
    case 'plain': return t('talk.meet-lines.talk-server-offered');
    case 'full': return fault.max ? t('talk.meet-lines.meeting-full-holds', { count: fault.max }) : t('talk.meet-lines.meeting-full-talk');
    case 'turned-away': return t('talk.meet-lines.talk-server-did-not-take');
    case 'not-admitted': return fault.mine ? t('talk.meet-lines.talk-server-did-not-take-meeting') : name ? t('talk.meet-lines.talk-server-did-not', { name }) : t('talk.meet-lines.talk-server-did');
    case 'refused': return fault.mine ? t('talk.meet-lines.talk-service-did') : name ? t('talk.meet-lines.name-s-talk', { name }) : t('talk.meet-lines.could-not-let');
    default: return t('talk.meet-lines.talk-could-not-connect');
  }
}

/** When a planned meeting is, as its card says it; nothing for one with no time. */
export function whenText(at, lang = undefined) {
  if (!at) return '';
  return new Date(at).toLocaleString(lang, { weekday: 'short', day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' });
}
