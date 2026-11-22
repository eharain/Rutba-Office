// Rutba Office's icon by the clock (D-2026-11-22-23).
//
// It says what waits (missed calls, unread messages, new mail, the events
// still to come today), opens Talk, Mail or Calendar at them, pauses the
// notifications for an hour, opens their settings, and quits. It wears a
// mark while something waits. What waits is the notifier's to say
// (notifier.js); this file draws it.
//
// The icon and its menu are made by functions handed in (Electron's Tray and
// Menu in the application, stand-ins in a test), so a test puts no icon by
// the clock on this computer.

import { t, tn } from '@rutba/office-ui/messages';

/** The most of the day's events the menu lists. */
export const EVENTS_MOST = 4;

const defaultTime = (ms) => new Date(ms).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

/**
 * The mark an icon wears while something waits: a red dot with a white rim
 * in its lower right corner, painted into a copy of the picture's bytes
 * (four a pixel, blue first, as Electron's bitmaps are). The picture handed
 * in is left as it was.
 */
export function paintMark(bitmap, width, height) {
  const out = Buffer.from(bitmap);
  const size = Math.min(width, height);
  const r = size * 0.27;
  const rim = Math.max(1, size / 16);
  const cx = width - r - rim * 0.5;
  const cy = height - r - rim * 0.5;
  // Each pixel sampled nine times, so the dot's edge is smooth at sixteen pixels.
  const steps = [1 / 6, 0.5, 5 / 6];
  for (let y = Math.max(0, Math.floor(cy - r - rim)); y < height; y++) {
    for (let x = Math.max(0, Math.floor(cx - r - rim)); x < width; x++) {
      let dot = 0;
      let ring = 0;
      for (const sy of steps) {
        for (const sx of steps) {
          const d = Math.hypot(x + sx - cx, y + sy - cy);
          if (d <= r) dot += 1;
          else if (d <= r + rim) ring += 1;
        }
      }
      if (!dot && !ring) continue;
      const i = (y * width + x) * 4;
      const a = (dot + ring) / 9;
      // The dot's red (229, 72, 77) and the rim's white, mixed by how much of each the pixel holds.
      const b = (77 * dot + 255 * ring) / (dot + ring);
      const g = (72 * dot + 255 * ring) / (dot + ring);
      const red = (229 * dot + 255 * ring) / (dot + ring);
      out[i] = Math.round(b * a + out[i] * (1 - a));
      out[i + 1] = Math.round(g * a + out[i + 1] * (1 - a));
      out[i + 2] = Math.round(red * a + out[i + 2] * (1 - a));
      out[i + 3] = Math.round(255 * a + out[i + 3] * (1 - a));
    }
  }
  return out;
}

/** The apps the menu opens, as it names them. */
const appWords = (app) => ({ talk: t('common.talk'), mail: t('common.mail'), calendar: t('common.calendar') })[app] || app;

/**
 * What Office says the first time it stays by the clock after its last
 * window closed, so nobody wonders why it is still running.
 */
export function stillRunningWords(name) {
  return {
    title: t('shell.tray.name-still-running', { name }),
    body: t('shell.tray.keeps-watch-calls'),
  };
}

/** What waits, in a few words, for the icon's tip: nothing when nothing does. */
export function waitingWords(waiting) {
  const parts = [];
  if (waiting.missed) parts.push(tn(waiting.missed, 'shell.count-missed-call'));
  if (waiting.chat) parts.push(tn(waiting.chat, 'shell.tray.count-unread-message'));
  if (waiting.mail) parts.push(tn(waiting.mail, 'shell.count-new-email'));
  return parts;
}

/**
 * The menu, as a list of rows Electron's Menu takes: what waits, the day's
 * events, the apps, the pause, the settings and Quit.
 *
 * @param {object} o
 * @param {object} o.state     the notifier's state()
 * @param {string} o.name      "Rutba Office", or the profile's name
 * @param {string[]} o.apps    the apps the menu opens ('talk', 'mail', 'calendar'), those installed and on
 * @param {object} o.does      { open({ app, query }), pause(), resume(), settings(), quit(), update?: { version, run } }
 */
export function menuRows({ state, name, apps = [], does, timeOf = defaultTime }) {
  const rows = [];
  const has = (app) => apps.includes(app);
  const w = state.waiting;
  if (w.missed && has('talk')) rows.push({ id: 'missed', label: tn(w.missed, 'shell.count-missed-call'), click: () => does.open({ app: 'talk', query: { show: 'calls' } }) });
  if (w.chat && has('talk')) rows.push({ id: 'chat', label: tn(w.chat, 'shell.tray.count-unread-message'), click: () => does.open({ app: 'talk', query: { show: 'unread' } }) });
  if (w.mail && has('mail')) rows.push({ id: 'mail', label: tn(w.mail, 'shell.count-new-email'), click: () => does.open({ app: 'mail', query: { show: 'new' } }) });
  const events = has('calendar') ? state.events.slice(0, EVENTS_MOST) : [];
  if (events.length) {
    if (rows.length) rows.push({ type: 'separator' });
    rows.push({ id: 'today', label: t('common.today'), enabled: false });
    for (const e of events) {
      const when = e.allDay ? t('calendar.all-day') : timeOf(e.start);
      rows.push({ id: `event:${e.id}`, label: `${when}  ${e.summary || t('ai-tools.event')}`, click: () => does.open({ app: 'calendar', query: { event: e.id, at: e.original ?? e.start } }) });
    }
    if (state.events.length > events.length) rows.push({ id: 'more', label: tn(state.events.length - events.length, 'shell.tray.count-more'), click: () => does.open({ app: 'calendar', query: { show: 'today' } }) });
  }
  if (!rows.length) rows.push({ id: 'nothing', label: t('shell.tray.nothing-is-waiting'), enabled: false });
  rows.push({ type: 'separator' });
  rows.push({ id: 'home', label: t('shell.tray.show-name', { name }), click: () => does.open({ app: 'home' }) });
  for (const app of apps) rows.push({ id: `app:${app}`, label: appWords(app), click: () => does.open({ app }) });
  rows.push({ type: 'separator' });
  if (state.paused) {
    rows.push({ id: 'paused', label: t('common.notifications-paused', { time: timeOf(state.pausedUntil) }), enabled: false });
    rows.push({ id: 'resume', label: t('common.turn-notifications-back'), click: () => does.resume() });
  } else {
    rows.push({ id: 'pause', label: t('shell.tray.pause-notifications-hour'), click: () => does.pause() });
  }
  rows.push({ id: 'settings', label: t('shell.tray.notification'), click: () => does.settings() });
  // An update that has come down waits for Office to quit, which a copy kept by the clock seldom does.
  if (does.update) rows.push({ id: 'update', label: does.update.version ? t('shell.tray.restart-update-version', { version: does.update.version }) : t('shell.restart-and'), click: () => does.update.run() });
  rows.push({ type: 'separator' });
  rows.push({ id: 'quit', label: t('shell.tray.quit-name', { name }), click: () => does.quit() });
  return rows;
}

/**
 * The icon by the clock.
 *
 * @param {object} o
 * @param {(icon: any) => object} o.make        makes the icon: { setImage, setToolTip, setContextMenu, on, destroy }
 * @param {(rows: object[]) => any} o.buildMenu  a menu from rows
 * @param {{ plain: any, waiting: any }} o.icons the icon, and the icon with its mark
 * @param {object} o.notifier
 * @param {() => string[]} [o.apps]
 * @param {object} o.does      as menuRows takes; `update()` may answer { version, run } or null each time the menu is made
 * @param {() => void} [o.speak]
 */
export function createTray({ make, buildMenu, icons, notifier, name = 'Rutba Office', apps = () => [], does, speak = () => {}, timeOf = defaultTime }) {
  let tray = null;
  let marked = null;
  let stop = null;

  const draw = () => {
    if (!tray) return;
    speak();
    const state = notifier.state();
    const waits = state.waiting.total > 0;
    if (waits !== marked) {
      marked = waits;
      try { tray.setImage(waits ? icons.waiting : icons.plain); } catch { /* the plain icon stays */ }
    }
    const words = waitingWords(state.waiting);
    tray.setToolTip(words.length ? `${name}: ${words.join(', ')}` : name);
    let update = null;
    try { update = does.update?.() || null; } catch { update = null; }
    tray.setContextMenu(buildMenu(menuRows({ state, name, apps: apps(), does: { ...does, update }, timeOf })));
  };

  return {
    /** Puts the icon by the clock. Answers whether it is there. */
    show() {
      if (tray) return true;
      try {
        tray = make(icons.plain);
      } catch {
        tray = null;
        return false;
      }
      marked = false;
      // A press on the icon opens Office; its menu is on the right button, as Windows has it.
      tray.on?.('click', () => does.open({ app: 'home' }));
      stop = notifier.onChange(draw);
      draw();
      return true;
    },
    /** Draws the icon and its menu again: the apps chosen, or the language, changed. */
    refresh: draw,
    shown: () => Boolean(tray),
    hide() {
      stop?.();
      stop = null;
      try { tray?.destroy?.(); } catch { /* gone already */ }
      tray = null;
      marked = null;
    },
  };
}
