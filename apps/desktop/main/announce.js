// The one thing this application says to the outside world.
//
// Rutba Office does not phone home. No account, no identifier, no telemetry, no
// crash reports, no document ever leaving the machine — that is the promise the
// mail client's tracker report is written against, and it would be worthless if
// the suite itself did quietly what it accuses newsletters of doing.
//
// This is the single, stated exception: once a day, at most, the launcher asks
// office.rutba.io whether there is an announcement — a release worth knowing
// about, a security note, a change to the licence. Two things follow from that
// request, and both are said plainly in the interface:
//
//   1. It is how we know the suite is being used. The server counts requests.
//      A request carries no identifier of any kind, so what is counted is
//      "a copy of Rutba Office opened somewhere today", never "this copy" or
//      "this person". There is no cookie, no id in the URL, no id in a file.
//   2. It sends the version and the platform, because an announcement about a
//      Windows 1.0 problem should not be shown to somebody on Linux running
//      1.4, and because knowing which platforms are actually in use is how the
//      next build gets prioritised.
//
// It can be turned off, in one press, in the launcher. Turning it off turns off
// the announcement too — there is no version of this that reports without
// telling you anything back, because that would be telemetry with extra steps.
//
// Everything here is best-effort. No network, a redirect, a 500, a timeout, a
// body that is not JSON, a body that is JSON but nonsense: all of them mean
// "no announcement today", and none of them may ever reach the person as an
// error. A launcher that shows a red toast because a marketing endpoint was
// down is worse than one that shows nothing.

import { app } from 'electron';
import { SITE } from '@rutba/office-formats/registry';

/** Once a day is plenty for a notice board, and is the cheapest useful signal. */
const EVERY = 20 * 60 * 60 * 1000;

/** A slow answer is the same as no answer, and must not delay the window. */
const TIMEOUT = 4000;

/** Nothing here is trusted: the server is ours, the parsing is not credulous. */
function clean(payload) {
  if (!payload || typeof payload !== 'object') return null;

  const text = (value, max) => (typeof value === 'string' ? value.trim().slice(0, max) : '');
  const title = text(payload.title, 140);
  if (!title) return null;

  const url = text(payload.url, 400);
  let link = null;
  try {
    // Only ever our own site, and only ever https. An announcement endpoint
    // that has been tampered with must not become a way to send people
    // anywhere, and must never be able to hand back a file: or javascript: URL.
    const parsed = new URL(url, SITE.home);
    if (parsed.protocol === 'https:' && /(^|\.)rutba\.io$/.test(parsed.hostname)) link = parsed.href;
  } catch {
    /* no link, which is fine — most announcements are a sentence */
  }

  return {
    // The identity of the announcement, so one that has been read stays read.
    // If the server does not give one, the title is one: an announcement whose
    // words have not changed is not a new announcement.
    id: text(payload.id, 80) || title,
    title,
    body: text(payload.body, 400),
    link,
    linkLabel: text(payload.linkLabel, 40) || 'Read more',
    // A hint at how loudly to say it. Anything unfamiliar is a plain notice.
    kind: ['news', 'release', 'security'].includes(payload.kind) ? payload.kind : 'news',
    at: text(payload.at, 40) || new Date().toISOString(),
  };
}

export function createAnnouncementService({ stores, broadcast }) {
  const enabled = () => stores.settings.get('announcements.enabled', true) !== false;

  /**
   * Ask, or say why not. Never throws, never reports, never blocks.
   * @param {{ force?: boolean }} [opts] `force` ignores the once-a-day rule.
   */
  async function check({ force = false } = {}) {
    const seen = stores.settings.get('announcements.seen', null);
    const cached = stores.settings.get('announcements.last', null);
    const checkedAt = stores.settings.get('announcements.checkedAt', 0);

    // What the window should draw right now, from what we already have. This is
    // computed first so a launcher with no network still shows the notice it
    // was shown yesterday, and shows nothing at all once it has been read.
    const held = cached && cached.id !== seen ? cached : null;

    if (!enabled()) return { announcement: null, enabled: false, checkedAt };
    if (!force && Date.now() - checkedAt < EVERY) return { announcement: held, enabled: true, checkedAt };

    const url = new URL('/announcement', SITE.home);
    url.searchParams.set('v', app.getVersion());
    url.searchParams.set('os', process.platform);

    try {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), TIMEOUT);
      const response = await fetch(url, {
        signal: controller.signal,
        redirect: 'follow',
        headers: { Accept: 'application/json', 'User-Agent': `RutbaOffice/${app.getVersion()}` },
        // No credentials, no cookies, nothing that could become an identifier.
        credentials: 'omit',
        cache: 'no-store',
      }).finally(() => clearTimeout(timer));

      // The request has now been made, which is the count. Whether it answered
      // usefully is a separate question, and the time is recorded either way so
      // a broken endpoint is not asked again every time a window opens.
      stores.settings.set('announcements.checkedAt', Date.now());

      if (!response.ok) return { announcement: held, enabled: true, checkedAt: Date.now() };

      const announcement = clean(await response.json());
      if (!announcement) {
        // An empty answer is a real answer: there is nothing on the board.
        stores.settings.delete('announcements.last');
        return { announcement: null, enabled: true, checkedAt: Date.now() };
      }

      stores.settings.set('announcements.last', announcement);
      const fresh = announcement.id !== seen ? announcement : null;
      if (fresh) broadcast?.('announce:new', fresh);
      return { announcement: fresh, enabled: true, checkedAt: Date.now() };
    } catch {
      // Offline, blocked by a firewall, DNS not answering, the endpoint not
      // built yet, a proxy returning HTML. All the same thing from here.
      stores.settings.set('announcements.checkedAt', Date.now());
      return { announcement: held, enabled: true, checkedAt: Date.now() };
    }
  }

  return {
    check,

    /** Read. It does not come back unless the board changes. */
    dismiss: ({ id }) => {
      stores.settings.set('announcements.seen', id ?? stores.settings.get('announcements.last', {})?.id ?? null);
      return { seen: stores.settings.get('announcements.seen', null) };
    },

    /** What the setting is, and what turning it off actually means. */
    status: () => ({
      enabled: enabled(),
      checkedAt: stores.settings.get('announcements.checkedAt', 0),
      endpoint: new URL('/announcement', SITE.home).href,
      sends: ['the version of Rutba Office', 'the operating system name'],
      sendsNot: ['any identifier', 'any document', 'anything about what you opened'],
    }),

    setEnabled: ({ on }) => {
      stores.settings.set('announcements.enabled', Boolean(on));
      return { enabled: Boolean(on) };
    },
  };
}
