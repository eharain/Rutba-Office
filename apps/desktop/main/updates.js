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

import { app } from 'electron';

const FEED = { provider: 'github', owner: 'eharain', repo: 'Rutba-Office' };
const SIX_HOURS = 6 * 60 * 60 * 1000;

/**
 * @param {object} o
 * @param {object} o.stores      settings, for the on/off switch
 * @param {Function} o.broadcast tell every window what the state is
 */
export function createUpdateService({ stores, broadcast }) {
  /** @type {'idle'|'checking'|'available'|'downloading'|'ready'|'current'|'error'|'off'|'unpackaged'} */
  let state = 'idle';
  let info = null;
  let progress = null;
  let error = null;
  let updater = null;
  let timer = null;

  const enabled = () => stores.settings.get('updates.automatic', true) !== false;

  const publish = () => {
    const payload = {
      state,
      version: app.getVersion(),
      available: info?.version || null,
      notes: info?.releaseNotes ? String(info.releaseNotes).slice(0, 4000) : null,
      releasedAt: info?.releaseDate || null,
      percent: progress?.percent ?? null,
      transferred: progress?.transferred ?? null,
      total: progress?.total ?? null,
      error,
      automatic: enabled(),
      channel: FEED.owner + '/' + FEED.repo,
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
      publish();
    });
    updater.on('update-downloaded', (next) => {
      state = 'ready';
      info = next;
      progress = null;
      publish();
    });
    updater.on('error', (err) => {
      // A failed check is not an event worth interrupting anybody for; it is
      // reported where somebody has asked to look, and nowhere else.
      state = 'error';
      error = err?.message || String(err);
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
