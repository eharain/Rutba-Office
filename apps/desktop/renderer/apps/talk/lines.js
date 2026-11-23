// What Talk says about a conversation beside its messages (Office 2.27):
// the count of what has not been read, a missed call, whether a message has
// arrived, what changed in a group conversation, and why something was
// refused. Words only, no drawing, so they are tested without a window.
//
// Kept plain (owner, 2026-10-21: "for chats and conference call etc build
// clean simple user interface do not complicate them like Ms teams"): one
// short line each, no states a person has to learn.

import { t, tn } from '@rutba/office-ui/messages';
import { countText } from '../../talk-words.js';

/** How a group's conversation is named among the conversations (as the service names it). */
export const groupKey = (groupId) => `group:${groupId}`;
/** The group a conversation's key names, or null for a person's. */
export const groupIdOf = (key) => (typeof key === 'string' && key.startsWith('group:') ? key.slice(6) : null);

// The count and what waits are the core's (talk-words.js): Home shows them too.
export { countText, waitingIn } from '../../talk-words.js';

/** Talk's name with what waits in it, for the window's title and so the taskbar. */
export const titleWithCount = (title, n) => (n > 0 ? `${title} (${countText(n)})` : title);

const nameOf = (p) => p?.name || p?.email || t('common.someone');

/** A group's name: the one it was given, or the others in it by their first names. */
export function groupTitle(group, meId) {
  if (group?.name) return group.name;
  const others = (group?.members || []).filter((m) => m.id !== meId).map((m) => nameOf(m).split(/\s+/)[0]);
  return others.length ? others.join(', ') : t('common.group');
}

/** The people of a group in one line, the person themself last. */
export function groupPeopleLine(group, meId) {
  const others = (group?.members || []).filter((m) => m.id !== meId).map(nameOf);
  return [...others, t('talk.lines.you')].join(', ');
}

/** Whose message this is, in a group: the member's name, or that they have left. */
export function senderName(group, id) {
  const m = (group?.members || []).find((x) => x.id === id);
  return m ? nameOf(m) : t('talk.lines.someone-who-left');
}

/** The contacts who can be put in a group: not in it already. */
export function canJoin(contacts, group) {
  const inIt = new Set((group?.members || []).map((m) => m.id));
  return (contacts || []).filter((c) => !inIt.has(c.id));
}

const listOf = (people) => (people || []).map(nameOf).join(', ');

/** A line in a group's conversation saying what changed in it, as the person reading sees it. */
export function eventText(event, meId) {
  const mine = event?.by?.id === meId;
  const by = nameOf(event?.by);
  const who = listOf(event?.who);
  switch (event?.kind) {
    case 'added-you': return t('talk.lines.name-added-group', { name: by });
    case 'added': return mine ? t('talk.lines.you-added-who', { who }) : t('talk.lines.name-added-who', { name: by, who });
    case 'renamed':
      if (!event.name) return mine ? t('talk.lines.took-name-off') : t('talk.lines.name-took-name', { name: by });
      return mine ? t('talk.lines.named-group-title', { title: event.name }) : t('talk.lines.name-named-group', { name: by, title: event.name });
    case 'removed': return mine ? t('talk.lines.took-out-group', { who }) : t('talk.lines.name-took-out', { name: by, who });
    case 'left': return t('talk.lines.left-group', { who: who || by });
    default: return '';
  }
}

/** A missed call in one line: what kind it was, or how many when the same person called more than once. */
export const missedText = (call) => (call?.count > 1 ? tn(call.count, 'talk.lines.count-missed-call') : call?.video ? t('talk.lines.missed-video-call') : t('talk.lines.missed-call'));

/** Why a call was missed, for the list of them; nothing to add when it simply rang unanswered. */
export function missedWhy(why) {
  if (why === 'busy') return t('talk.lines.another-call');
  if (why === 'offline') return t('talk.lines.talk-not-open');
  return '';
}

/** Why a message of mine was not sent, in words; only "Not sent" when the reason is not one a person can act on. */
export function notSentText(why) {
  switch (why) {
    case 'late': return t('talk.lines.not-sent-talk');
    case 'not-a-contact': return t('talk.lines.not-sent-no');
    case 'no-group': return t('talk.lines.not-sent-no-longer');
    case 'too-many': return t('talk.lines.not-sent-too');
    case 'text': return t('talk.lines.not-sent-message');
    default: return t('talk.lines.not-sent');
  }
}

/**
 * Whether a message of mine has gone and arrived: { icon, double, words } or
 * null for one from before the marks. One on its way while Talk is not
 * connected says that it waits (it goes by itself once Talk is back); one
 * that was not sent says why, and `failed` has the screen offer to try
 * again or to remove it.
 */
export function markOf(message, { online = true } = {}) {
  if (!message?.mine) return null;
  switch (message.state) {
    case 'sending': return online ? { icon: null, words: t('mail.sending') } : { icon: null, waiting: true, words: t('talk.lines.waiting-talk-not') };
    case 'sent': return { icon: 'check', double: false, words: t('mail.sent') };
    case 'arrived': return { icon: 'check', double: true, words: t('talk.lines.arrived') };
    case 'failed': return { icon: null, bad: true, failed: true, words: notSentText(message.why) };
    default: return null;
  }
}

/** A refusal about a group or a report, in words. */
export function refusalText(code, { most, who } = {}) {
  switch (code) {
    case 'not-a-contact': return t('talk.lines.only-own-contacts');
    case 'too-few': return t('talk.lines.group-three-people');
    case 'too-many': return most ? t('talk.lines.group-holds-count', { count: most }) : t('talk.lines.too-many-now');
    case 'too-many-groups': return Array.isArray(who) && who.length ? t('talk.lines.cannot-put-any', { who: listOf(who) }) : t('talk.lines.many-groups-talk');
    case 'no-group': return t('talk.lines.no-longer-group');
    case 'not-yours': return t('talk.lines.only-person-started');
    case 'nobody': return t('talk.lines.person-cannot-chosen');
    case 'offline':
    case 'no-answer': return t('talk.invite.talk-not-connected');
    case 'unknown': return t('talk.lines.talk-cannot-do');
    default: return t('talk.invite.talk-could-not');
  }
}

/** How many of the last messages a report carries at most, when the person ticks so (the service sends no more: main/talk-conversations.js). */
export const REPORT_MESSAGES = 20;

/**
 * What a report's tick for messages says: how many at most, whose, and
 * written where. In a group it is what they wrote in the group, not to
 * this person alone.
 */
export const reportMessagesText = (name, { group = false } = {}) => (group
  ? t('talk.lines.send-count-last', { count: REPORT_MESSAGES, name })
  : t('talk.lines.send-count-last-messages', { count: REPORT_MESSAGES, name }));

/**
 * What leaving a group means, said before it is done. The person who
 * started it is told what only they could do: once they have left, nobody
 * can take anyone out.
 */
export const leaveGroupText = (group, meId) => [
  t('talk.groups.no-longer-get'),
  group?.starter === meId ? t('talk.lines.started-group-once') : '',
].filter(Boolean).join(' ');

/** Why someone is reported: the choices, as the dialog lists them. */
export const reportReasons = () => [
  ['spam', t('talk.lines.spam-or-advertising')],
  ['harassment', t('talk.lines.harassment-or')],
  ['impersonation', t('talk.lines.pretending-someone-else')],
  ['other', t('talk.lines.something-else')],
];

/** What a member of a group is marked with in its list of people, or ''. */
export function memberNote(member, group, meId) {
  if (member.id === meId) return t('common.you');
  const notes = [];
  if (member.id === group?.starter) notes.push(t('talk.lines.started-the-group'));
  if (!member.contact) notes.push(t('talk.lines.not-contacts'));
  if (member.needsUpdate) notes.push(t('talk.lines.needs-newer-rutba'));
  return notes.join(' · ');
}
