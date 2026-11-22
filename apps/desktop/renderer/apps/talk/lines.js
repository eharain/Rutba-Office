// What Talk says about a conversation beside its messages (Office 2.28):
// the count of what has not been read, a missed call, whether a message has
// arrived, what changed in a group conversation, and why something was
// refused. Words only, no drawing, so they are tested without a window.
//
// Kept plain (owner, 2026-10-21: "for chats and conference call etc build
// clean simple user interface do not complicate them like Ms teams"): one
// short line each, no states a person has to learn.

import { t, tn } from '@rutba/office-ui/messages';

/** How a group's conversation is named among the conversations (as the service names it). */
export const groupKey = (groupId) => `group:${groupId}`;
/** The group a conversation's key names, or null for a person's. */
export const groupIdOf = (key) => (typeof key === 'string' && key.startsWith('group:') ? key.slice(6) : null);

/** A count as a small mark shows it: nothing for none, never wider than three characters. */
export const countText = (n) => (n > 99 ? '99+' : n > 0 ? String(n) : '');

/** Talk's name with what waits in it, for the window's title and so the taskbar. */
export const titleWithCount = (title, n) => (n > 0 ? `${title} (${countText(n)})` : title);

/** What waits in Talk: the messages not read and the missed calls not looked at. */
export const waitingIn = (state) => (state?.unreadTotal || 0) + (state?.missed || 0);

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

/** Whether a message of mine has gone and arrived: { icon, double, words } or null for one from before the marks. */
export function markOf(message) {
  if (!message?.mine) return null;
  switch (message.state) {
    case 'sending': return { icon: null, words: t('mail.sending') };
    case 'sent': return { icon: 'check', double: false, words: t('mail.sent') };
    case 'arrived': return { icon: 'check', double: true, words: t('talk.lines.arrived') };
    case 'failed': return { icon: null, bad: true, words: t('talk.lines.not-sent') };
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
    case 'unknown': return t('talk.lines.talk-server-does');
    default: return t('talk.invite.talk-could-not');
  }
}

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
