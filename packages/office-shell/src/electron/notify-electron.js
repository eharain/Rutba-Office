// Electron's side of the notifier and the icon by the clock (notifier.js,
// tray.js): the Windows notification, the taskbar's flash, the icon and its
// menu, and starting with Windows. Everything those two files are handed in
// the application is made here, and only here, so they can be tested with
// stand-ins and nothing on the screen.
//
// Under a check run nothing is shown and nothing sounds: no icon by the
// clock, no notification and no ring. What would have been shown is kept in
// globalThis.__rutbaNotifications for a check to read, and each ring that
// would have sounded in globalThis.__rutbaRings.

import crypto from 'node:crypto';
import { app, BrowserWindow, Notification, Tray, Menu, nativeImage } from 'electron';
import { paintMark } from './tray.js';
import { RING_LIMIT_MS, RING_PAGE, ringScript } from './call-ring.js';

/** The flag a copy started with Windows carries: it goes to the icon by the clock, with no window. */
export const TRAY_FLAG = '--tray';
/** How many notifications are kept alive for a press that comes late, from Windows' notification centre. */
const KEPT = 40;
/** The icon's sizes: sixteen points, at each scale Windows draws the clock's corner in. */
const SCALES = [1, 1.25, 1.5, 2, 3];

/** A notification's tag, as Windows takes one: short, and the same for the same thing. */
const tagOf = (id) => crypto.createHash('sha1').update(String(id)).digest('hex').slice(0, 16);

/**
 * The icon, and the icon with its mark, from the application's own picture.
 * A picture that cannot be read gives the file's path for both, which the
 * tray draws as it is.
 */
export function trayIcons(iconPath) {
  try {
    const base = nativeImage.createFromPath(iconPath);
    if (base.isEmpty()) return { plain: iconPath, waiting: iconPath };
    const plain = nativeImage.createEmpty();
    const waiting = nativeImage.createEmpty();
    for (const scaleFactor of SCALES) {
      const px = Math.round(16 * scaleFactor);
      const small = base.resize({ width: px, height: px, quality: 'best' });
      const bitmap = small.toBitmap();
      const size = small.getSize();
      if (bitmap.length !== size.width * size.height * 4) continue;
      plain.addRepresentation({ scaleFactor, width: size.width, height: size.height, buffer: small.toPNG() });
      const marked = nativeImage.createFromBitmap(paintMark(bitmap, size.width, size.height), { width: size.width, height: size.height });
      waiting.addRepresentation({ scaleFactor, width: size.width, height: size.height, buffer: marked.toPNG() });
    }
    if (plain.isEmpty() || waiting.isEmpty()) return { plain: iconPath, waiting: iconPath };
    return { plain, waiting };
  } catch {
    return { plain: iconPath, waiting: iconPath };
  }
}

/**
 * What the notifier and the tray are handed in the application.
 *
 * @param {object} o
 * @param {string} o.iconPath
 * @param {object} o.windows      the shell's window manager
 * @param {boolean} [o.check]     a check run: nothing is shown
 * @param {boolean} [o.installed] the person's installed copy: only that one can start with Windows
 * @param {string[]} [o.startArgs]  what a copy started with Windows is started with, beside --tray (its profile)
 * @param {string} [o.startName]  the name Windows keeps that start under, when it is not the application's own
 */
export function notifyParts({ iconPath, windows, check = false, installed = app.isPackaged, startArgs = [], startName = null }) {
  const kept = [];
  const recorded = [];
  if (check) globalThis.__rutbaNotifications = recorded;

  const show = (note) => {
    if (check) {
      const entry = { id: note.id, kind: note.kind, title: note.title, body: note.body, silent: note.silent, sticky: Boolean(note.sticky), actions: note.actions.map((a) => a.id), at: Date.now(), closed: false };
      // What a person can do to it, for a check to do: press it, or press one of its buttons.
      Object.defineProperties(entry, { click: { value: () => note.onClick?.() }, act: { value: (id) => note.onAction?.(id) } });
      recorded.push(entry);
      if (recorded.length > 200) recorded.shift();
      return { close: () => { entry.closed = true; } };
    }
    if (!Notification.isSupported()) return null;
    const n = new Notification({
      id: tagOf(note.id),
      title: note.title,
      body: note.body,
      silent: Boolean(note.silent),
      icon: iconPath,
      // A call stays up until it is answered, turned away or stops ringing.
      timeoutType: note.sticky ? 'never' : 'default',
      urgency: note.sticky ? 'critical' : 'normal',
      actions: (note.actions || []).map((a) => ({ type: 'button', text: a.label })),
    });
    n.on('click', () => note.onClick?.());
    n.on('action', (details, index) => {
      const at = details?.actionIndex ?? index;
      const action = (note.actions || [])[at];
      if (action) note.onAction?.(action.id);
    });
    n.on('close', () => note.onClose?.());
    // Windows cannot show it (notifications off for Office, or no Start menu entry): nothing more to do.
    n.on('failed', () => note.onClose?.());
    n.show();
    kept.push(n);
    if (kept.length > KEPT) kept.shift();
    return { close: () => { try { n.close(); } catch { /* gone already */ } } };
  };

  /**
   * The ring for a call no window rings for (call-ring.js), in a window that
   * is never shown: it holds no page of Office's, no preload and nothing of
   * the profile's. Answers what stops it; it also stops by itself after
   * RING_LIMIT_MS. A computer with no sound, or a window that cannot be
   * made, is no fault to show anyone: the notification is still there.
   */
  const rung = [];
  if (check) globalThis.__rutbaRings = rung;
  const ring = () => {
    if (check) {
      const entry = { at: Date.now(), stopped: false };
      rung.push(entry);
      if (rung.length > 200) rung.shift();
      return () => { entry.stopped = true; };
    }
    let win = null;
    let limit = null;
    let over = false;
    const stop = () => {
      if (over) return;
      over = true;
      clearTimeout(limit);
      try { if (win && !win.isDestroyed()) win.destroy(); } catch { /* gone already */ }
      win = null;
    };
    try {
      win = new BrowserWindow({
        show: false,
        width: 240,
        height: 120,
        skipTaskbar: true,
        focusable: false,
        webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false, backgroundThrottling: false, autoplayPolicy: 'no-user-gesture-required', partition: 'call-ring' },
      });
      limit = setTimeout(stop, RING_LIMIT_MS + 1000);
      limit.unref?.();
      const made = win;
      made.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(RING_PAGE)}`)
        .then(() => (over || made.isDestroyed() ? null : made.webContents.executeJavaScript(ringScript(), true)))
        .catch(() => stop());
    } catch {
      stop();
    }
    return stop;
  };

  /** The taskbar button of that app's windows, until one of them is looked at. */
  const flash = (appKey) => {
    if (check) return;
    for (const w of windows.all()) {
      if (w.isDestroyed() || windows.infoFor(w).app !== appKey || w.isFocused() || !w.isVisible()) continue;
      w.flashFrame(true);
      w.once('focus', () => { if (!w.isDestroyed()) w.flashFrame(false); });
    }
  };

  // Only the person's installed copy starts with Windows: a copy run from the source would start Electron itself.
  const canStart = Boolean(installed) && !check && (process.platform === 'win32' || process.platform === 'darwin');
  const startOptions = () => ({ args: [TRAY_FLAG, ...startArgs], ...(startName && process.platform === 'win32' ? { name: startName } : {}) });
  const startup = {
    can: canStart,
    get: () => {
      if (!canStart) return false;
      try { return Boolean(app.getLoginItemSettings(startOptions()).openAtLogin); } catch { return false; }
    },
    set: (on) => {
      if (!canStart) return false;
      try {
        app.setLoginItemSettings({ openAtLogin: Boolean(on), ...startOptions() });
        return Boolean(on);
      } catch {
        return false;
      }
    },
  };

  return {
    show,
    flash,
    ring,
    supported: () => check || Notification.isSupported(),
    makeTray: (icon) => new Tray(icon),
    buildMenu: (rows) => Menu.buildFromTemplate(rows),
    icons: () => trayIcons(iconPath),
    startup,
  };
}
