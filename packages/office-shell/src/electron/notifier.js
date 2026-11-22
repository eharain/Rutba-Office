// The one place Rutba Office tells the person something outside its own
// windows (the owner, 22 November 2026, D-2026-11-22-23: "the system should
// have a tray icon to keep an eye on incoming calls, calendar events meeting
// reminders, chats as well as emails with windows notifications options").
//
// Every app reports here; none shows a notification by itself. What comes in
// is a fact (a call from this person, a message with these words, mail with
// this subject); this file counts what waits, decides whether the person is
// told now (the kind turned off, notifications paused, the quiet hours), and
// writes the words. The Windows notification itself, the taskbar's flash and
// the clock are handed in, so nothing here touches Electron and a test
// raises nothing on the screen.

import { t, tn } from '@rutba/office-ui/messages';

/** What the person can be told of, each with a switch of its own. */
export const KINDS = ['call', 'missed', 'chat', 'mail', 'reminder', 'meeting'];
/** The kinds that wait until looked at, counted on the icon and in its menu. */
export const COUNTED = ['missed', 'chat', 'mail'];
/** Which app's window, looked at, has seen which of them. */
export const APP_KINDS = { talk: ['missed', 'chat'], mail: ['mail'] };
export const SETTINGS_KEY = 'notifications';
/** "Pause notifications for an hour". */
export const PAUSE_MS = 60 * 60_000;
/** The most of a message's words a notification carries. */
export const TEXT_MOST = 240;

const CLOCK = /^([01]\d|2[0-3]):[0-5]\d$/;
const clock = (v, fallback) => (CLOCK.test(String(v)) ? String(v) : fallback);
const minutesOf = (hhmm) => Number(hhmm.slice(0, 2)) * 60 + Number(hhmm.slice(3));

/**
 * The settings as kept, whole, whatever was stored: each kind on unless
 * turned off, sound on, the words of a message and the subject of an email
 * shown, no quiet hours, Office kept running by the clock after its last
 * window closes, and not started with Windows.
 */
export function readSettings(raw) {
  const s = raw && typeof raw === 'object' ? raw : {};
  const kinds = {};
  for (const k of KINDS) kinds[k] = s.kinds?.[k] !== false;
  return {
    kinds,
    sound: s.sound !== false,
    text: s.text !== false,
    quiet: { on: Boolean(s.quiet?.on), from: clock(s.quiet?.from, '22:00'), to: clock(s.quiet?.to, '07:00') },
    pausedUntil: Number.isFinite(s.pausedUntil) && s.pausedUntil > 0 ? s.pausedUntil : 0,
    keepRunning: s.keepRunning !== false,
    startup: Boolean(s.startup),
  };
}

/**
 * Whether `at` (a time in milliseconds) is inside the quiet hours, by the
 * computer's own clock. Hours that end before they begin run past midnight;
 * the same time twice is no quiet hours at all.
 */
export function inQuietHours(quiet, at) {
  if (!quiet?.on) return false;
  const from = minutesOf(clock(quiet.from, '22:00'));
  const to = minutesOf(clock(quiet.to, '07:00'));
  if (from === to) return false;
  const d = new Date(at);
  const now = d.getHours() * 60 + d.getMinutes();
  return from < to ? now >= from && now < to : now >= from || now < to;
}

const short = (text, most = TEXT_MOST) => {
  const s = String(text || '').replace(/\s+/g, ' ').trim();
  return s.length > most ? `${s.slice(0, most - 1)}…` : s;
};

/** How long until `start`, in words: now, in so many minutes, or at a time. */
function whenWords(start, now, timeOf) {
  const minutes = Math.round((start - now) / 60_000);
  if (minutes <= 0) return t('shell.notifier.starting-now');
  if (minutes < 60) return tn(minutes, 'shell.notifier.in-count-minute');
  return t('shell.notifier.at-time', { time: timeOf(start) });
}

/**
 * The words of a notification: { title, body, actions: [{ id, label }] }.
 * `showText` off keeps a message's words and an email's subject out of it;
 * who it is from is still said.
 */
export function wordsOf(note, { showText = true, now = Date.now(), timeOf = defaultTime } = {}) {
  const from = short(note.from, 80);
  const count = Math.max(1, Number(note.count) || 1);
  switch (note.kind) {
    case 'call':
      return {
        title: note.video ? t('shell.notifier.incoming-video-call') : t('talk.incoming-call'),
        body: from || t('shell.notifier.someone-in-talk'),
        actions: [{ id: 'answer', label: t('common.answer') }, { id: 'decline', label: t('common.decline') }],
      };
    case 'missed':
      return { title: count > 1 ? tn(count, 'shell.count-missed-call') : t('shell.notifier.missed-call'), body: from || t('shell.notifier.someone-in-talk'), actions: [] };
    case 'chat': {
      const who = from || t('common.talk');
      const title = note.group ? t('shell.notifier.name-in-group', { name: who, group: short(note.group, 60) }) : who;
      const words = showText ? short(note.text) : '';
      const many = tn(count, 'mail.count-new-message');
      return { title, body: words ? (count > 1 ? `${words} (${many})` : words) : many, actions: [] };
    }
    case 'mail': {
      if (count > 1) return { title: tn(count, 'shell.count-new-email'), body: short(note.account || from, 120), actions: [] };
      const subject = showText ? short(note.subject) || t('mail.no-subject') : '';
      return { title: from || t('common.new-email'), body: subject || (from ? t('common.new-email') : short(note.account, 120)), actions: [] };
    }
    case 'reminder':
    case 'meeting': {
      const when = note.start != null ? whenWords(note.start, now, timeOf) : '';
      const where = short(note.location, 80);
      return {
        title: short(note.summary, 120) || (note.kind === 'meeting' ? t('shell.notifier.meeting') : t('ai-tools.event')),
        body: [when, where].filter(Boolean).join(' · '),
        actions: note.kind === 'meeting' && note.join ? [{ id: 'join', label: t('shell.notifier.join') }] : [],
      };
    }
    default:
      return { title: short(note.title, 120), body: short(note.body), actions: [] };
  }
}

function defaultTime(ms) {
  return new Date(ms).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
}

/**
 * The notifier.
 *
 * @param {object} o
 * @param {{ get: Function, set: Function }} o.settings  the profile's settings store
 * @param {(note: object) => ({ close?: () => void } | null)} [o.show]  shows a Windows notification: { id, kind, title, body, silent, sticky, actions: [{ id, label }], onClick, onAction(id) }
 * @param {(app: string) => void} [o.flash]   the taskbar button of that app's window, when it is not in front
 * @param {(target: { app: string, query?: object }) => void} [o.open]  opens an app at a thing
 * @param {() => number} [o.now]
 * @param {() => void} [o.speak]  called before any words are written, so they are in the person's language
 * @param {(ms: number) => string} [o.timeOf]
 * @param {Function} [o.setTimer]  setTimeout, for a test to hold
 */
export function createNotifier({ settings, show = () => null, flash = () => {}, open = () => {}, now = () => Date.now(), speak = () => {}, timeOf = defaultTime, setTimer = setTimeout, clearTimer = clearTimeout }) {
  const waiting = Object.fromEntries(COUNTED.map((k) => [k, 0]));
  // A kind whose own app says how many wait (Talk's unread count) is not counted here.
  const owned = new Set();
  const shown = new Map();
  const listeners = new Set();
  let events = [];
  let serial = 0;
  let pauseTimer = null;

  const read = () => readSettings(settings.get(SETTINGS_KEY, null));
  const changed = () => {
    for (const fn of [...listeners]) {
      try { fn(); } catch { /* one listener's fault is its own */ }
    }
  };
  const total = () => COUNTED.reduce((n, k) => n + waiting[k], 0);

  /** Why a notification of this kind is not shown now, or null when it is. */
  const held = (kind, s = read()) => {
    if (KINDS.includes(kind) && !s.kinds[kind]) return 'off';
    if (s.pausedUntil > now()) return 'paused';
    if (inQuietHours(s.quiet, now())) return 'quiet';
    return null;
  };

  const armPause = () => {
    clearTimer(pauseTimer);
    pauseTimer = null;
    const left = read().pausedUntil - now();
    if (left <= 0) return;
    // The icon's menu says "paused" until then, and stops saying it when the hour is up.
    pauseTimer = setTimer(() => { pauseTimer = null; changed(); }, left + 500);
    pauseTimer?.unref?.();
  };
  armPause();

  const put = (note, words, s) => {
    const id = String(note.id || `${note.kind || 'note'}-${++serial}`);
    shown.get(id)?.close?.();
    const run = (what) => {
      const own = note.on?.[what];
      if (typeof own === 'function') {
        try { own(); } catch { /* the app's own fault */ }
        return;
      }
      // Decline has no window to open; everything else goes to the thing.
      if (what !== 'decline' && note.open) open(note.open);
    };
    let handle = null;
    try {
      handle = show({
        id,
        kind: note.kind || 'note',
        title: words.title,
        body: words.body,
        silent: !s.sound || Boolean(note.silent),
        sticky: note.kind === 'call',
        actions: words.actions,
        onClick: () => run('open'),
        onAction: (actionId) => run(actionId),
        onClose: () => { if (shown.get(id) === handle) shown.delete(id); },
      }) || null;
    } catch {
      handle = null;
    }
    if (handle) shown.set(id, handle);
    return id;
  };

  const notifier = {
    /**
     * Something happened that the person may want to know of. It is counted
     * when its kind waits (a missed call, a message, an email), and shown as
     * a Windows notification unless its kind is off, notifications are paused
     * or the quiet hours are on. Answers { shown, id } or { shown: false, why }.
     *
     * `note`: { kind, id?, from?, text?, subject?, account?, group?, video?,
     * summary?, start?, location?, join?, count?, counted?, app?, open?:
     * { app, query }, on?: { open, answer, decline, join }, front? }. `app`
     * is whose taskbar button flashes; `front` says that app's window is in
     * front now, so the person is looking at it: nothing is shown or counted.
     */
    report(note) {
      if (!note || !KINDS.includes(note.kind)) throw new TypeError('A report names its kind.');
      if (note.front) return { shown: false, why: 'front' };
      if (COUNTED.includes(note.kind) && note.counted !== false && !owned.has(note.kind)) waiting[note.kind] += Math.max(1, Number(note.count) || 1);
      const s = read();
      const why = held(note.kind, s);
      if (!why) {
        speak();
        const id = put(note, wordsOf(note, { showText: s.text, now: now(), timeOf }), s);
        if (note.app) {
          try { flash(note.app); } catch { /* a window on its way out */ }
        }
        changed();
        return { shown: true, id };
      }
      changed();
      return { shown: false, why };
    },

    /** Office's own word to the person (it went to the icon by the clock): never a kind, shown once by whoever asks. */
    tell({ id, title, body, open: target = null, silent = true }) {
      const s = read();
      if (s.pausedUntil > now() || inQuietHours(s.quiet, now())) return { shown: false };
      return { shown: true, id: put({ id, kind: 'office', open: target, silent }, { title, body, actions: [] }, s) };
    },

    /** A notification shown on purpose, from the settings: the person sees that Windows lets Office show them, and hears whether they sound. */
    test() {
      speak();
      const words = { title: t('shell.notifier.notifications-are-on'), body: t('shell.notifier.how-rutba-office'), actions: [] };
      return { shown: true, id: put({ id: 'office:test', kind: 'office' }, words, read()) };
    },

    /** A notification taken back: the call was answered elsewhere, or stopped ringing. */
    withdraw(id) {
      const handle = shown.get(String(id));
      if (!handle) return false;
      shown.delete(String(id));
      try { handle.close?.(); } catch { /* gone already */ }
      return true;
    },

    /** How many of a kind wait, said by the app that knows: from then on that kind is its app's to count. */
    setWaiting(kind, count) {
      if (!COUNTED.includes(kind)) throw new TypeError('Only missed calls, messages and emails wait.');
      owned.add(kind);
      const n = Math.max(0, Math.floor(Number(count) || 0));
      if (waiting[kind] === n) return;
      waiting[kind] = n;
      changed();
    },

    /** A kind looked at: nothing of it waits now. */
    seen(kind) {
      if (!COUNTED.includes(kind) || owned.has(kind) || !waiting[kind]) return;
      waiting[kind] = 0;
      changed();
    },

    /** An app's window came to the front: what it shows has been seen (unless the app counts for itself). */
    seenApp(app) {
      for (const kind of APP_KINDS[app] || []) notifier.seen(kind);
    },

    /** The day's events still to come, for the icon's menu: [{ id, summary, start, end, allDay }]. */
    setEvents(list) {
      const next = (Array.isArray(list) ? list : []).map((e) => ({ id: e.id ?? null, summary: String(e.summary || ''), start: Number(e.start), end: Number(e.end), allDay: Boolean(e.allDay), original: e.original ?? null }));
      if (JSON.stringify(next) === JSON.stringify(events)) return;
      events = next;
      changed();
    },

    settings: read,

    /** A change to the settings, each part checked; answers them whole. */
    setSettings(change = {}) {
      const before = read();
      const next = readSettings({
        ...before,
        ...change,
        kinds: { ...before.kinds, ...(change.kinds || {}) },
        quiet: { ...before.quiet, ...(change.quiet || {}) },
        pausedUntil: before.pausedUntil,
      });
      // Starting with Windows is starting by the clock: it needs Office kept running there.
      if (!next.keepRunning) next.startup = false;
      settings.set(SETTINGS_KEY, next);
      changed();
      return next;
    },

    /** No notifications for a while (an hour, from the icon's menu); what comes in is still counted. */
    pause(ms = PAUSE_MS) {
      const until = now() + Math.max(60_000, Math.min(Number(ms) || PAUSE_MS, 24 * 60 * 60_000));
      settings.set(SETTINGS_KEY, { ...read(), pausedUntil: until });
      armPause();
      changed();
      return until;
    },

    resume() {
      settings.set(SETTINGS_KEY, { ...read(), pausedUntil: 0 });
      armPause();
      changed();
    },

    /** Everything the icon, its menu and the settings window show. */
    state() {
      const s = read();
      const at = now();
      return {
        settings: s,
        waiting: { ...waiting, total: total() },
        events: events.filter((e) => e.end > at),
        paused: s.pausedUntil > at,
        pausedUntil: s.pausedUntil > at ? s.pausedUntil : 0,
        quiet: inQuietHours(s.quiet, at),
      };
    },

    held,

    /** Told each time what waits, the events or the settings change. Answers the function that stops it. */
    onChange(fn) {
      listeners.add(fn);
      return () => listeners.delete(fn);
    },

    /** Every notification still up taken down (Office is quitting). */
    close() {
      clearTimer(pauseTimer);
      for (const id of [...shown.keys()]) notifier.withdraw(id);
      listeners.clear();
    },
  };
  return notifier;
}
