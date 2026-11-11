// Talk, behind this one file, so only it changes if Talk's calls move.
// Talk's contract (2.25): `talk.connections()` answers the people Talk is
// connected to, `[{ id, name, email, presence }]`, and a Talk window opened
// with `{ inviteEmail, inviteName }` asks someone in. A Talk not installed
// has no `talk` on the shell: then nobody comes from it and no Talk button
// is offered. A Talk without `connections` yet gives nobody.

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

/** Someone asked into Talk: a Talk window opened on its invitation, their address and name filled. */
export function inviteToTalk(shell, { email, name } = {}) {
  const address = String(email || '').trim();
  if (!address || !canInviteToTalk(shell)) return false;
  shell.win.create({ app: 'talk', query: { inviteEmail: address, ...(name ? { inviteName: String(name) } : {}) } });
  return true;
}
