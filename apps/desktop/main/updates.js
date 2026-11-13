// Staying current.
//
// The suite checks GitHub for a newer release, downloads it in the background,
// and installs it the next time you quit. That is the whole behaviour, and it
// is worth being precise about what it costs in privacy, because the rest of
// this application makes a point of contacting nothing:
//
//   - HTTPS requests to GitHub for the release feed and for a newer template
//     set, carrying no identifier beyond what any HTTP request carries.
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
//
// And it waits for the person's voices and languages (private design:
// office-pack-compatibility-design.md 5, C11). The download is never
// automatic: between finding a version and fetching it, the add-ons service
// is asked whether every add-on on this computer is ready for it. With none
// installed the answer is yes at once, with no request. Otherwise that
// version's packages feed is read; a version some add-on is not ready for
// waits (`waiting`, with `waitingFor` and `retired` in the state), is asked
// again at each check, and is offered the moment they are ready. "Update
// anyway" is a button, for that version only. The packs that are ready are
// staged before the download (C12), for the new version to move in at its
// first start.

import { app, net } from 'electron';
import { isSnoozed, snoozeFor } from './update-snooze.js';
import { updateErrorKind, whatStopsRestart } from './update-error.js';
import { guardDownloads, signatureAddress, keepAnonymous } from './update-signature.js';
import { gateStep } from './update-gate.js';

export const FEED = { provider: 'github', owner: 'eharain', repo: 'Rutba-Office' };
const SIX_HOURS = 6 * 60 * 60 * 1000;

/** Whether this is the portable build, which electron-builder starts with these set. */
const portable = () => Boolean(process.env.PORTABLE_EXECUTABLE_DIR || process.env.PORTABLE_EXECUTABLE_FILE);

/**
 * @param {object} o
 * @param {object} o.stores      settings, for the on/off switch
 * @param {Function} o.broadcast tell every window what the state is
 * @param {Function} [o.taskbar]  show a download's progress on the windows
 *                                (a fraction, or -1 for none)
 */
export function createUpdateService({ stores, broadcast, taskbar = null, templates = null, openWindows, addons = null }) {
  /** @type {'idle'|'checking'|'available'|'waiting'|'downloading'|'ready'|'current'|'error'|'off'|'unpackaged'} */
  let state = 'idle';
  // What the add-ons said of the version found (C11): who it waits for, which add-ons it no longer has.
  let gateAnswer = null;
  let gating = null;
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
      error: error?.detail ?? null,
      errorKind: error?.kind ?? null,
      automatic: enabled(),
      channel: FEED.owner + '/' + FEED.repo,
      // The version this copy replaced, until the card has been put away.
      arrived: arrived(),
      // The add-ons the version found waits for, and those it no longer has (C11).
      waitingFor: gateAnswer?.waitingFor || [],
      retired: gateAnswer?.retired || [],
      addonsChecked: gateAnswer ? !gateAnswer.feedError : null,
      // "Update anyway" was chosen for this version.
      anyway: Boolean(info?.version) && stores.settings.get('updates.anyway', null) === info?.version,
      // A newer template set: ready for the next start, or needing a newer Office.
      templates: templates ? (({ phase, set, needs }) => ({ phase, set, needs }))(templates.state()) : null,
    };
    broadcast?.('update:state', payload);
    return payload;
  };

  /** electron-updater is loaded only when it is actually going to be used. */
  async function load() {
    if (updater) return updater;
    const mod = await import('electron-updater');
    updater = mod.autoUpdater ?? mod.default?.autoUpdater;

    // Never by itself: the add-ons are asked first (gate below). A portable copy downloads nothing at all.
    updater.autoDownload = false;
    // Automatic updates turned off means nothing is installed unasked, a download already made included.
    updater.autoInstallOnAppQuit = enabled() && !portable();
    keepAnonymous(updater);
    // Only Windows goes through this door; a Windows copy installs nothing the makers did not sign.
    if (process.platform === 'win32') {
      guardDownloads(updater, {
        current: () => updater.updateInfoAndProvider?.info ?? info,
        fetchSignature: async ({ version, file }) => {
          const response = await net.fetch(signatureAddress({ owner: FEED.owner, repo: FEED.repo, version, file }), { signal: AbortSignal.timeout(60_000) });
          if (!response.ok) throw new Error(`${response.status}`);
          return response.text();
        },
      });
    }
    // Full releases only from 2.25.0 on. A copy of 2.24.0 or before still looks at the newest entry and finds the full release.
    updater.allowPrerelease = false;

    updater.logger = null;
    updater.setFeedURL(FEED);

    updater.on('checking-for-update', () => {
      state = 'checking';
      error = null;
      publish();
    });
    updater.on('update-available', (next) => {
      state = 'available';
      // A newer version replaces one waiting, and what it waits for is asked again.
      if (info?.version !== next?.version) gateAnswer = null;
      info = next;
      publish();
      gate(next).catch((err) => {
        state = 'error';
        error = updateErrorKind(err?.message || String(err));
        publish();
      });
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
      error = updateErrorKind(err?.message || String(err));
      showProgress(-1);
      publish();
    });

    return updater;
  }

  /**
   * Between finding a version and downloading it: whether the add-ons are
   * ready for it. No add-ons service, or none installed, is ready at once.
   */
  async function gate(next) {
    const version = next?.version;
    if (!version || gating === version) return;
    gating = version;
    try {
      const answer = addons ? await addons.readiness(version) : { offer: true, ready: [], waiting: [], retired: [], waitingFor: [] };
      if (info?.version !== version) return; // a newer one was found meanwhile
      gateAnswer = answer;
      const step = gateStep({ answer, version, anyway: stores.settings.get('updates.anyway', null), portable: portable() });
      if (step === 'wait') state = 'waiting';
      // A portable copy is told, and installs nothing.
      if (step !== 'download') {
        publish();
        return;
      }
      await download(version);
    } finally {
      gating = null;
    }
  }

  async function download(version) {
    // The ready packs staged first, for the new version's first start; one not staged is fetched by that version itself.
    if (addons && gateAnswer?.ready?.length) await addons.stageFor(version).catch(() => {});
    const u = await load();
    if (info?.version !== version) return;
    state = 'downloading';
    publish();
    await u.downloadUpdate();
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
      error = updateErrorKind(err?.message || String(err));
    }
    // The templates are looked for whatever Office's own check found: a set fails on its own.
    if (templates) await templates.check().catch(() => {});
    // And the add-ons' own revisions and withdrawals for this version (C4), with the same switch: off means no request.
    if (addons) await addons.check({ manual }).catch(() => {});
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
      const stopped = whatStopsRestart({ ready: state === 'ready', windows: openWindows?.() || [] });
      if (stopped) return stopped;
      const u = await load();
      // Quit, install, and come back — the only way an update is ever applied.
      // Silently: the person pressed Restart and update, and has no pages to answer.
      setImmediate(() => u.quitAndInstall(true, true));
      return { installed: true };
    },
    // "Update anyway" (C11): this version, though some add-ons are not ready
    // for it. Asked once, in the window; kept like a snooze, for that version only.
    anyway: async ({ version } = {}) => {
      const v = version ?? info?.version ?? null;
      if (!v) return publish();
      stores.settings.set('updates.anyway', v);
      if (state === 'waiting' && info?.version === v && !portable()) {
        download(v).catch((err) => {
          state = 'error';
          error = updateErrorKind(err?.message || String(err));
          publish();
        });
      }
      return publish();
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
      if (updater) updater.autoInstallOnAppQuit = Boolean(on);
      if (on && app.isPackaged) {
        state = 'idle';
        check().catch(() => {});
        // Off at launch means no timer was started, and on again means another.
        clearInterval(timer);
        timer = setInterval(() => check().catch(() => {}), SIX_HOURS);
        timer.unref?.();
      } else {
        state = 'off';
        clearInterval(timer);
      }
      return publish();
    },
    start,
  };
}
