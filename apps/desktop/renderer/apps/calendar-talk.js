// Talk, behind this one file, so only it changes if Talk's calls move.
// Talk's contract (2.25): `talk.connections()` answers the people Talk is
// connected to, `[{ id, name, email, presence }]`, and a Talk window opened
// with `{ inviteEmail, inviteName }` asks someone in. A Talk not installed
// has no `talk` on the shell: then nobody comes from it and no Talk button
// is offered. A Talk without `connections` yet gives nobody.

import { meetingIn, meetingCode, withMeeting } from './talk/meet-lines.js';

const lower = (s) => String(s || '').trim().toLowerCase();

/** Whether Talk is there to ask someone into. */
export function canInviteToTalk(shell) {
  return Boolean(shell?.talk) && typeof shell?.win?.create === 'function';
}

/**
 * The people Talk is connected to whose name or address holds `query`, as
 * Contacts' suggestions are: `[{ name, email, source: 'talk', presence }]`.
 * Nobody without an address; nobody when Talk cannot say.
 */
export async function talkPeople(shell, query = '', limit = 6) {
  const connections = shell?.talk?.connections;
  if (typeof connections !== 'function') return [];
  let rows;
  try {
    rows = await connections();
  } catch {
    return [];
  }
  const q = lower(query);
  const out = [];
  for (const r of Array.isArray(rows) ? rows : []) {
    const email = String(r?.email || '').trim();
    if (!/@/.test(email)) continue;
    const name = String(r?.name || '').trim() || null;
    if (q && !lower(email).includes(q) && !lower(name).includes(q)) continue;
    out.push({ name, email, source: 'talk', presence: r?.presence || null });
    if (out.length >= limit) break;
  }
  return out;
}

/** Two lists of suggestions as one: the first's people first, each address once. */
export function mergePeople(first = [], second = [], limit = 6) {
  const seen = new Set();
  const out = [];
  for (const p of [...(first || []), ...(second || [])]) {
    const key = lower(p?.email);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    out.push(p);
    if (out.length >= limit) break;
  }
  return out;
}

/** Contacts' suggestions for `query`, with Talk's people after them. */
export async function invitePeople(shell, query, limit = 6) {
  const [cards, talk] = await Promise.all([
    Promise.resolve(shell?.contacts?.suggest?.({ query, limit })).catch(() => []),
    talkPeople(shell, query, limit),
  ]);
  return mergePeople(cards || [], talk, limit);
}

/** Whether a contact card can be asked into Talk: a person (not a room or a resource) with an address. */
export function talkableCard(card) {
  return Boolean(card?.emails?.some?.((e) => /@/.test(String(e?.value || '')))) && card.kind !== 'room' && card.kind !== 'resource';
}

/** A contact card asked into Talk, by its first address and the name it shows. */
export function inviteCardToTalk(shell, card) {
  if (!talkableCard(card)) return false;
  const email = card.emails.find((e) => /@/.test(String(e?.value || ''))).value;
  return inviteToTalk(shell, { email, name: card.display || null });
}

/** Someone asked into Talk: a Talk window opened on its invitation, their address and name filled. */
export function inviteToTalk(shell, { email, name } = {}) {
  const address = String(email || '').trim();
  if (!address || !canInviteToTalk(shell)) return false;
  shell.win.create({ app: 'talk', query: { inviteEmail: address, ...(name ? { inviteName: String(name) } : {}) } });
  return true;
}

/* ── a meeting on an event (Office 2.28) ──────────────────────────── */

/** Whether this Talk can hold a meeting for an event. */
export function canMeet(shell) {
  return typeof shell?.talk?.meetCreate === 'function';
}

/** The meeting an event carries, read from its place and its notes: { code, link, live } or null. */
export function meetingOf(event) {
  return meetingIn(event?.location, event?.description);
}

/**
 * Talk made ready for a meeting to be set up: 'online', or why not
 * ('signed-out': nobody is signed in to Talk; 'offline': it could not
 * connect). A Talk that is signed in but not connected yet (Office was
 * started without it, or the connection dropped) is connected first and
 * waited for a few seconds, so Calendar never needs Talk's window open.
 */
export async function talkReady(shell, { wait = 8000, step = 250, sleep = (ms) => new Promise((r) => { setTimeout(r, ms); }) } = {}) {
  const talk = shell?.talk;
  // An earlier Talk does not say how it stands: asked as before.
  if (typeof talk?.state !== 'function' || typeof talk?.resume !== 'function') return 'online';
  let s = await Promise.resolve().then(() => talk.resume()).catch(() => null);
  for (let waited = 0; ; waited += step) {
    if (!s) return 'offline';
    if (s.phase === 'online') return 'online';
    if (s.phase === 'signed-out' || s.phase === 'signing-in') return 'signed-out';
    if (waited >= wait) return 'offline';
    await sleep(step);
    s = await Promise.resolve().then(() => talk.state()).catch(() => null);
  }
}

/** Talk's window brought forward: to sign in, or to see why it is not connected. */
export function openTalk(shell) {
  if (!canJoinMeeting(shell)) return false;
  shell.win.create({ app: 'talk', reuse: true });
  return true;
}

/**
 * What an event's invitees mean for its Talk meeting, as Talk has it today:
 * a meeting holds `most` people with the organiser, and only the
 * organiser's Talk contacts can come in. { count, tooMany, strangers }:
 * `strangers` are the invited addresses Talk knows no contact by, or null
 * when Talk cannot say (it is not connected).
 */
export async function meetingGuests(shell, emails = [], most = 4) {
  const list = [...new Set((emails || []).map(lower).filter((e) => e.includes('@')))];
  const out = { count: list.length, tooMany: list.length > most - 1, strangers: null };
  let s = null;
  try { s = await shell.talk.state(); } catch { s = null; }
  if (s?.phase !== 'online') return out;
  const known = new Set((s.contacts || []).map((c) => lower(c.email)).filter(Boolean));
  const mine = lower(s.me?.email);
  return { ...out, strangers: list.filter((e) => e !== mine && !known.has(e)) };
}

/**
 * A meeting set up in Talk for an event, and the event's place and notes
 * with it put in: the link, how to join, and the notice when it will be
 * streamed live, so the invitation carries all three. Whether it will be
 * streamed live is chosen here and never after (D-2026-11-22-19). Talk is
 * connected first if it is not ('signed-out' or 'offline' when it cannot be).
 * { meeting, location, description } | { error }.
 */
export async function addMeeting(shell, { name, at, live = false, location = '', description = '' } = {}, ready = {}) {
  if (!canMeet(shell)) return { error: 'no-talk' };
  const talk = await talkReady(shell, ready);
  if (talk !== 'online') return { error: talk };
  let r;
  try {
    r = await shell.talk.meetCreate({ name: String(name || '').trim(), at: at ?? null, live: live === true });
  } catch {
    return { error: 'offline' };
  }
  if (!r || r.error || !r.meeting?.link) return { error: r?.error || 'offline' };
  return { meeting: r.meeting, ...withMeeting({ location, description }, r.meeting) };
}

/** Whether a meeting can be joined from here: Talk is installed. */
export function canJoinMeeting(shell) {
  return Boolean(shell?.talk) && typeof shell?.win?.create === 'function';
}

/** Join: Talk brought forward on the meeting (its window already open, if there is one), where it is looked at before joining. */
export function joinMeeting(shell, link) {
  const code = meetingCode(link);
  if (!code || !canJoinMeeting(shell)) return false;
  shell.win.create({ app: 'talk', query: { meet: code }, reuse: true });
  return true;
}
