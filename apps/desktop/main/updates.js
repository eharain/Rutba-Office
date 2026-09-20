// Staying current.
//
// The suite checks GitHub for a newer release, downloads it in the background,
// and installs it the next time you quit. That is the whole behaviour, and it
// is worth being precise about what it costs in privacy, because the rest of
// this application makes a point of contacting nothing:
//
//   - One HTTPS request to api.github.com for the release feed, carrying no
//     identifier beyond what any HTTP request carries.
//   - Nothing about your documents, your mail or your machine is sent.
//   - It can be turned off, and when it is off no request is made at all.
//
// Updates are never installed while you are working. The download is silent,
// the install happens on quit, and a release you have not been offered cannot
// be applied — `autoInstallOnAppQuit` is the only path.
//
// And you are TOLD. A release the check finds is announced in every window
// the moment it is found (owner, 2026-09-14: a copy that updates in silence
// is a copy whose owner never learns what changed): the prompt shows the
// download's progress, offers "Restart and update" once it is downloaded,
// and "Not now" puts that version away for a day — after which it asks
// again, because a copy that is never quit never installs on quit. A newer
// release is never covered by an older "not now".
//
// And you are told AFTERWARDS (owner, 2026-09-20: "the auto update is very
// silent"). The install happens on quit, so the next launch looked exactly
// like the last one. The service remembers which version ran last; the
// first launch of a new one carries the arrival in its state, every window
// shows a card naming both versions with the release's notes a click away,
// and `seen` puts the card away everywhere. While a download runs, every
// window's status bar counts it and the taskbar button shows it.

import { app } from 'electron';
import { isSnoozed, snoozeFor } from './update-snooze.js';

const FEED = { provider: 'github', owner: 'eharain', repo: 'Rutba-Office' };
const SIX_HOURS = 6 * 60 * 60 * 1000;

/**
 * @param {object} o
 * @param {object} o.stores      settings, for the on/off switch
 * @param {Function} o.broadcast tell every window what the state is
 * @param {Function} [o.taskbar]  show a download's progress on the windows
 *                                (a fraction, or -1 for none)
 */
export function createUpdateService({ stores, broadcast, taskbar = null }) {
  /** @type {'idle'|'checking'|'available'|'downloading'|'ready'|'current'|'error'|'off'|'unpackaged'} */
  let state = 'idle';
  let info = null;
  let progress = null;
  let error = null;
  let updater = null;
  let timer = null;

  /**
   * On unless somebody turns it off.
   *
   * A suite that ships a release a day and never tells anyone leaves every
   * installed copy on the version it was installed with, defects and all. So
   * the check is on from the first launch: one request to GitHub's release
   * feed, carrying nothing about the person or the machine (the README's
   * privacy table lists it, and the announcement, and nothing else). The
   * switch in About turns it off, and off means no request at all.
   */
  const enabled = () => stores.settings.get('updates.automatic', true) !== false;


  /**
   * The version that ran last time, so the first launch of a new one can
   * say so. A copy installed fresh has nothing to say; a copy that was
   * updated keeps the arrival until somebody has seen it.
   */
  const current = app.getVersion();
  const lastRun = stores.settings.get('updates.lastRun', null);
  if (lastRun !== current) {
    stores.settings.set('updates.lastRun', current);
    if (lastRun) stores.settings.set('updates.arrived', { from: lastRun, to: current, at: Date.now(), seen: false });
  }
  const arrived = () => {
    const a = stores.settings.get('updates.arrived', null);
    return a && a.to === current ? a : null;
  };

  const showProgress = (fraction) => {
    try {
      taskbar?.(fraction);
    } catch {
      // The taskbar is a nicety; a window on its way out must not fail a download.
    }
  };

  /** The version "not now" covers right now, or null. */
  const snoozedVersion = () => {
    const entry = stores.settings.get('updates.snoozed', null);
    return entry && isSnoozed(entry, entry.version) ? String(entry.version) : null;
  };

  const publish = () => {
    const payload = {
      state,
      version: app.getVersion(),
      available: info?.version || null,
      // Whether the windows should keep quiet about `available`: the person
      // said "not now" to this very version less than a day ago.
      snoozed: isSnoozed(stores.settings.get('updates.snoozed', null), info?.version),
      snoozedVersion: snoozedVersion(),
      notes: info?.releaseNotes ? String(info.releaseNotes).slice(0, 4000) : null,
      releasedAt: info?.releaseDate || null,
      percent: progress?.percent ?? null,
      transferred: progress?.transferred ?? null,
      total: progress?.total ?? null,
      error,
      automatic: enabled(),
      channel: FEED.owner + '/' + FEED.repo,
      // The version this copy replaced, until the card has been put away.
      arrived: arrived(),
    };
    broadcast?.('update:state', payload);
    return payload;
  };

  /** electron-updater is loaded only when it is actually going to be used. */
  async function load() {
    if (updater) return updater;
    const mod = await import('electron-updater');
    updater = mod.autoUpdater ?? mod.default?.autoUpdater;

    updater.autoDownload = true;
    updater.autoInstallOnAppQuit = true;
    // Every release is published as a pre-release while the suite is in beta
    // (owner's decision, 2026-09-08), so an installed copy has to accept them
    // or it would never see an update at all.
    updater.allowPrerelease = true;

    updater.logger = null;
    updater.setFeedURL(FEED);

    updater.on('checking-for-update', () => {
      state = 'checking';
      error = null;
      publish();
    });
    updater.on('update-available', (next) => {
      state = 'available';
      info = next;
      publish();
    });
    updater.on('update-not-available', () => {
      state = 'current';
      info = null;
      publish();
    });
    updater.on('download-progress', (p) => {
      state = 'downloading';
      progress = p;
      showProgress((p?.percent ?? 0) / 100);
      publish();
    });
    updater.on('update-downloaded', (next) => {
      state = 'ready';
      info = next;
      progress = null;
      showProgress(-1);
      publish();
    });
    updater.on('error', (err) => {
      // A failed check is not an event worth interrupting anybody for; it is
      // reported where somebody has asked to look, and nowhere else.
      state = 'error';
      error = err?.message || String(err);
      showProgress(-1);
      publish();
    });

    return updater;
  }

  async function check({ manual = false } = {}) {
    if (!app.isPackaged) {
      state = 'unpackaged';
      error = null;
      return publish();
    }
    if (!enabled() && !manual) {
      state = 'off';
      return publish();
    }
    // Somebody who asks for a check wants to be told what it finds.
    if (manual) stores.settings.set('updates.snoozed', null);
    try {
      const u = await load();
      await u.checkForUpdates();
    } catch (err) {
      state = 'error';
      error = err?.message || String(err);
    }
    return publish();
  }

  /** Start the background rhythm: once shortly after launch, then periodically. */
  function start() {
    if (!app.isPackaged) return;
    if (!enabled()) {
      state = 'off';
      return;
    }
    // Not at the moment of launch: the first seconds belong to the window.
    setTimeout(() => check().catch(() => {}), 25_000);
    timer = setInterval(() => check().catch(() => {}), SIX_HOURS);
    timer.unref?.();
  }

  return {
    state: () => publish(),
    check: ({ manual = true } = {}) => check({ manual }),
    install: async () => {
      if (state !== 'ready') return { installed: false, reason: 'No update has been downloaded yet.' };
      const u = await load();
      // Quit, install, and come back — the only way an update is ever applied.
      setImmediate(() => u.quitAndInstall(false, true));
      return { installed: true };
    },
    // "Not now": this version, for a day. The download and the install on
    // quit go on regardless; only the asking stops.
    snooze: ({ version } = {}) => {
      stores.settings.set('updates.snoozed', snoozeFor(version ?? info?.version ?? null));
      return publish();
    },
    // The arrival card has been read (or put away): every window drops it.
    seen: () => {
      const a = arrived();
      if (a && !a.seen) stores.settings.set('updates.arrived', { ...a, seen: true });
      return publish();
    },
    setAutomatic: ({ on }) => {
      stores.settings.set('updates.automatic', Boolean(on));
      if (on && app.isPackaged) {
        state = 'idle';
        check().catch(() => {});
      } else {
        state = 'off';
        clearInterval(timer);
      }
      return publish();
    },
    start,
  };
}
