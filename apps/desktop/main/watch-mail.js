// Mail, as the notifier hears of it (watch.js, D-2026-11-22-23): new email.
//
// Nothing here fetches mail. Mail already checks for it while Rutba Office
// runs (a few seconds after it starts, every few minutes after, and at once
// where the server can say so), and tells the other apps in this process of
// each arrival (its `onArrived`): what came, after its junk filter, with who
// each message is from and what it is about. This file turns an arrival
// into one report to the notifier.
//
// Told of: mail that is unread and newly written, arriving in the Inbox or
// in a folder of the person's own. Not told of: what the junk filter took,
// what arrives in Sent, Drafts, Trash, Junk or Archive, mail already read
// somewhere else, and old mail a first fetch of an account brings down.

/** Folders whose arrivals nobody is told of. */
export const QUIET_ROLES = new Set(['sent', 'drafts', 'trash', 'junk', 'archive', 'outbox', 'templates']);
/** How old a message's own date may be and still be new mail: an account's first fetch brings years of it. */
export const FRESH_MS = 3 * 24 * 3600_000;
/** How many senders are named when several messages come at once. */
export const NAMES_MOST = 3;

const nameOf = (from) => String(from?.name || from?.address || '').trim();

/**
 * The messages of an arrival that are news to the person: unread, and
 * written lately. An arrival from a Mail that does not say what came (older
 * than this file) is null: then only how many came is known.
 */
export function newsOf(arrival, now = Date.now()) {
  if (!Array.isArray(arrival?.messages)) return null;
  return arrival.messages.filter((m) => {
    if (!m || m.unread === false) return false;
    const at = m.date != null ? new Date(m.date).getTime() : NaN;
    return Number.isNaN(at) || now - at <= FRESH_MS;
  });
}

/**
 * @param {object} o
 * @param {object} o.notifier
 * @param {object} o.mail   Mail's service: onArrived(listener), accounts()
 * @param {(app: string) => boolean} [o.front]  whether a window of that app is in front now
 * @param {() => number} [o.now]
 * @returns {{ has: () => boolean, stop: () => void } | null}  null when this Mail cannot tell of arrivals
 */
export function watchMail({ notifier, mail, front = () => false, now = () => Date.now() }) {
  if (typeof mail?.onArrived !== 'function') return null;

  const addressOf = (accountId) => {
    try {
      const a = (mail.accounts?.() || []).find((x) => x.id === accountId);
      return a?.address || a?.email || a?.name || '';
    } catch {
      return '';
    }
  };

  const heard = (arrival) => {
    if (!arrival || QUIET_ROLES.has(arrival.role)) return;
    const news = newsOf(arrival, now());
    const count = news ? news.length : Math.max(0, Number(arrival.count) || 0);
    if (!count) return;
    const one = news && count === 1 ? news[0] : null;
    // A rule may have filed it: then it is no longer where it arrived, and Mail is opened on the account.
    const there = !arrival.filedByRules;
    const names = news ? [...new Set(news.map((m) => nameOf(m.from)).filter(Boolean))] : [];
    notifier.report({
      kind: 'mail',
      id: `mail:${arrival.accountId}:${one?.id ?? arrival.ids?.[0] ?? arrival.at ?? ''}`,
      from: one ? nameOf(one.from) : names.slice(0, NAMES_MOST).join(', ') + (names.length > NAMES_MOST ? '…' : ''),
      subject: one ? one.subject : '',
      account: addressOf(arrival.accountId),
      count,
      front: front('mail'),
      app: 'mail',
      open: one && there
        // One message: in a window of its own, as a double-click in the list opens it.
        ? { app: 'mail', fresh: true, query: { message: JSON.stringify({ accountId: arrival.accountId, folder: arrival.folder, id: one.id }) } }
        // Several: Mail on that account, in the folder they came to (the Inbox unless another is named).
        : { app: 'mail', query: { account: arrival.accountId, ...(there && arrival.folder && arrival.role !== 'inbox' ? { folder: arrival.folder } : {}) } },
    });
  };

  let stop = null;
  try {
    stop = mail.onArrived((arrival) => {
      try { heard(arrival); } catch { /* the watch's fault is not Mail's */ }
    });
  } catch {
    return null;
  }
  return {
    /** Whether mail can arrive by itself: an account, and Mail set to fetch it or to be told by the server. */
    has: () => {
      try {
        if (!(mail.accounts?.() || []).length) return false;
        const fetch = mail.fetchSettings?.();
        return !fetch || Number(fetch.every) > 0 || Boolean(fetch.push);
      } catch {
        return false;
      }
    },
    stop: () => { try { stop?.(); } catch { /* gone already */ } },
  };
}
