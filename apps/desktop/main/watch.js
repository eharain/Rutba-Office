// Keeping watch (the owner, 22 November 2026, D-2026-11-22-23): Rutba
// Office's icon by the clock, the Windows notifications, and Office kept
// running there after its last window closes, so a call, a message, an
// email or a reminder is not missed because no window was open.
//
// This file joins the parts: the notifier and the tray (office-shell's
// notifier.js and tray.js), the apps that report to them (watch-talk.js,
// watch-mail.js, watch-calendar.js), the two settings that are the system's (kept
// running by the clock, started with Windows), and the `notify` calls a
// window makes. Electron's own parts are handed in (office-shell's
// notify-electron.js makes them in the application), so a test of this file
// shows nothing on the screen.
//
// What is watched is watched only while Rutba Office runs. Closed from its
// icon's Quit, or never started, it watches nothing.

import { createNotifier } from '@rutba/office-shell/electron/notifier';
import { createTray, stillRunningWords } from '@rutba/office-shell/electron/tray';
import { watchTalk } from './watch-talk.js';
import { watchMail } from './watch-mail.js';
import { watchCalendar } from './watch-calendar.js';

/** Whether the person has been told once that Office stays by the clock. */
export const TOLD_KEY = 'notifications.toldStillRunning';
/** The flag a copy started with Windows carries (the same as notify-electron.js's). */
export const TRAY_FLAG = '--tray';
/** The apps the icon's menu opens, in its order. */
export const WATCHED_APPS = ['talk', 'mail', 'calendar'];
/** A window that holds one thing by itself (Mail's message alone, a message being written): not the one an app is opened at. */
const ALONE = /[?&](?:message|compose)=/;

/**
 * @param {object} o
 * @param {object} o.stores        the profile's stores (settings)
 * @param {object} o.parts         { show, flash, makeTray, buildMenu, icons(), startup: { can, get, set }, supported() }
 * @param {object} o.windows       the shell's window manager: findByApp, open, all, infoFor
 * @param {(win, event, payload) => void} o.sendEvent
 * @param {(event, payload) => void} o.broadcast
 * @param {string} [o.name]        the application's name, or the profile's
 * @param {() => string[]} [o.apps]  the apps installed and on in this profile
 * @param {boolean} [o.check]      a check run: no icon, nothing kept running, nothing shown
 * @param {() => void} o.quit
 * @param {() => ({ version, run } | null)} [o.update]  an update that has come down and waits for a restart
 * @param {() => void} [o.speak]   called before words are written
 */
export function createWatch({ stores, parts, windows, sendEvent, broadcast, name = 'Rutba Office', apps = () => WATCHED_APPS, check = false, quit, update = () => null, speak = () => {}, now = () => Date.now() }) {
  const listeners = new Map();
  const on = (event, fn) => {
    if (!listeners.has(event)) listeners.set(event, new Set());
    listeners.get(event).add(fn);
    return () => listeners.get(event)?.delete(fn);
  };
  /** What a service tells the windows, heard here too. */
  const heard = (event, payload) => {
    for (const fn of [...(listeners.get(event) || [])]) {
      try { fn(payload); } catch { /* one listener's fault is its own */ }
    }
  };

  const front = (app) => {
    try {
      return windows.all().some((w) => !w.isDestroyed() && w.isFocused() && windows.infoFor(w).app === app);
    } catch {
      return false;
    }
  };
  const windowsOf = (app) => {
    try {
      return windows.all().filter((w) => !w.isDestroyed() && windows.infoFor(w).app === app);
    } catch {
      return [];
    }
  };

  /** The window an app is opened at: its most recently used, unless that one holds one thing by itself. */
  const mainWindowOf = (app) => {
    const isMain = (w) => {
      try { return !ALONE.test(String(w.webContents?.getURL?.() || '')); } catch { return true; }
    };
    const used = windows.findByApp(app);
    if (used && isMain(used)) return used;
    return used ? windowsOf(app).find(isMain) || null : null;
  };

  /**
   * An app opened at a thing: its window already open is brought forward
   * and told where to go ('app:link', as a link opened for it is); else one
   * opens on it. `fresh` always opens a window of its own (one email).
   */
  const goTo = ({ app, query = null, fresh = false }) => {
    const open = fresh ? null : mainWindowOf(app);
    if (open) {
      if (open.isMinimized()) open.restore();
      if (!open.isVisible()) open.show();
      open.focus();
      if (query) sendEvent(open, 'app:link', query);
      return open;
    }
    return windows.open(query ? { app, query } : { app });
  };

  const openSettings = () => goTo({ app: 'home', query: { notifications: 1 } });

  const notifier = createNotifier({ settings: stores.settings, show: parts.show, flash: parts.flash, open: goTo, speak, now });

  const chosen = () => {
    let on = WATCHED_APPS;
    try { on = apps() || []; } catch { on = WATCHED_APPS; }
    return WATCHED_APPS.filter((a) => on.includes(a));
  };

  /** The notifier as one app reports to it: an app the person has turned off tells them nothing. */
  const reportsOf = (app) => ({
    ...notifier,
    report: (note) => (chosen().includes(app) ? notifier.report(note) : { shown: false, why: 'off' }),
    setWaiting: (kind, count) => notifier.setWaiting(kind, chosen().includes(app) ? count : 0),
  });

  let tray = null;
  let talk = null;
  let mail = null;
  let calendar = null;
  const makeTray = () => createTray({
    make: parts.makeTray,
    buildMenu: parts.buildMenu,
    icons: parts.icons(),
    notifier,
    name,
    apps: chosen,
    speak,
    does: {
      open: goTo,
      pause: () => notifier.pause(),
      resume: () => notifier.resume(),
      settings: openSettings,
      quit,
      update,
    },
  });

  /** The kinds this copy watches for: those of the apps installed and on. */
  const watched = () => {
    const on = chosen();
    const talks = Boolean(talk) && on.includes('talk');
    const plans = Boolean(calendar) && on.includes('calendar');
    return [
      ...(talks ? ['call', 'missed', 'chat'] : []),
      ...(mail && on.includes('mail') ? ['mail'] : []),
      ...(plans ? ['reminder'] : []),
      ...(plans || talks ? ['meeting'] : []),
    ];
  };

  /** Everything a window is shown: the notifier's state, and what of it is the system's. */
  const state = () => ({
    ...notifier.state(),
    watched: watched(),
    supported: Boolean(parts.supported?.()),
    tray: Boolean(tray?.shown()),
    startup: { can: Boolean(parts.startup?.can), on: Boolean(parts.startup?.can) && notifier.settings().startup },
  });

  // Windows is told when the setting and what Windows holds differ, never on every start.
  const applyStartup = () => {
    if (!parts.startup?.can) return;
    const want = notifier.settings().startup;
    try {
      if (parts.startup.get() !== want) parts.startup.set(want);
    } catch {
      /* Windows refused: the setting stands, and is tried again at the next start */
    }
  };

  let started = false;
  let stopTelling = null;

  const watch = {
    notifier,
    on,
    goTo,

    /** A service's `broadcast`, heard here as well as by the windows. */
    hear: (event, payload) => {
      broadcast(event, payload);
      heard(event, payload);
    },

    /**
     * Talk's service, once it is made. What Talk says is worth a notification
     * ('talk:notify', main/talk-notify.js) is already heard; the service is
     * used only to turn a call away from its notification, and is left as it
     * is: Talk's own window shows a call still ringing and opens at what a
     * notification points to.
     */
    watchTalk(service) {
      if (!service || talk) return;
      talk = watchTalk({
        notifier: reportsOf('talk'),
        on,
        front,
        send: (msg) => service.send?.(msg),
      });
    },

    /** Mail's service, once it is made: each arrival of new mail is reported. Mail itself is left as it is. */
    watchMail(service) {
      if (!service || mail) return;
      mail = watchMail({ notifier: reportsOf('mail'), mail: service, front, now });
    },

    /** Calendar's service, once it is made: its events' reminders, a meeting about to start, and the day's events for the menu. */
    watchCalendar(service) {
      if (!service || calendar) return;
      calendar = watchCalendar({
        notifier: reportsOf('calendar'),
        calendar: service,
        settings: stores.settings,
        on,
        open: goTo,
        talk: () => Boolean(talk) && chosen().includes('talk'),
        now,
      });
    },

    /** The first window is up: the icon goes by the clock, and Windows is told whether Office starts with it. */
    start() {
      if (started) return;
      started = true;
      if (!check) {
        tray = makeTray();
        tray.show();
      }
      applyStartup();
      stopTelling = notifier.onChange(() => broadcast('notify:state', state()));
    },

    /** A window of this app came to the front: what it shows has been seen. */
    focused(app) {
      notifier.seenApp(app);
    },

    /** The apps chosen, or the language, changed: the menu is drawn again. */
    refresh() {
      tray?.refresh();
    },

    /**
     * The last window closed. True when Office stays by the clock (the
     * setting, and an icon there to open it and quit it from); the first
     * time, the person is told so.
     */
    lastWindowClosed() {
      if (check || !tray?.shown() || !notifier.settings().keepRunning) return false;
      if (!stores.settings.get(TOLD_KEY, false)) {
        stores.settings.set(TOLD_KEY, true);
        speak();
        notifier.tell({ id: 'office:still-running', ...stillRunningWords(name), open: { app: 'home' } });
      }
      return true;
    },

    /** Whether this start goes to the clock with no window: started with Windows, and kept running there. */
    startsQuiet(argv = []) {
      return !check && argv.includes(TRAY_FLAG) && Boolean(tray?.shown()) && notifier.settings().keepRunning;
    },

    /** The `notify` calls a window makes. */
    service: {
      state,
      set: ({ settings } = {}) => {
        const change = settings && typeof settings === 'object' ? settings : {};
        // Only the installed copy starts with Windows; elsewhere the switch stays off.
        notifier.setSettings({ ...change, ...(change.startup && !parts.startup?.can ? { startup: false } : {}) });
        applyStartup();
        return state();
      },
      pause: ({ minutes = 60 } = {}) => {
        notifier.pause(Math.round(Number(minutes) || 60) * 60_000);
        return state();
      },
      resume: () => {
        notifier.resume();
        return state();
      },
      test: () => notifier.test(),
    },

    /** Office is quitting: the icon leaves the clock and the notifications come down. */
    stop() {
      stopTelling?.();
      talk?.stop();
      mail?.stop();
      calendar?.stop();
      tray?.hide();
      notifier.close();
    },
  };
  return watch;
}
