// Windows.
//
// One BrowserWindow per open document, which is what people expect from an
// office suite: two spreadsheets side by side, a message open next to the
// message list. The window is frameless everywhere and draws its own caption,
// because a ribbon that starts below a native title bar wastes the one strip of
// vertical space a document editor cannot spare. On macOS the traffic lights
// stay native and inset, since replacing those is a fight nobody wins.
//
// Window position and size are remembered per app, not per document: reopening
// a spreadsheet should land where your spreadsheets live.

import { BrowserWindow, screen, nativeTheme } from 'electron';
import path from 'node:path';
import { SCHEME, encodePath } from './protocol.js';

const isMac = process.platform === 'darwin';

/** Default geometry per app — a mail client and a picture viewer want different rooms. */
const GEOMETRY = {
  home: { width: 1180, height: 760, minWidth: 900, minHeight: 560 },
  mail: { width: 1440, height: 900, minWidth: 1000, minHeight: 600 },
  word: { width: 1280, height: 900, minWidth: 820, minHeight: 560 },
  sheets: { width: 1440, height: 900, minWidth: 900, minHeight: 560 },
  slides: { width: 1400, height: 880, minWidth: 900, minHeight: 600 },
  pictures: { width: 1280, height: 860, minWidth: 720, minHeight: 520 },
  image: { width: 1360, height: 900, minWidth: 900, minHeight: 620 },
  video: { width: 1400, height: 880, minWidth: 900, minHeight: 600 },
};

export function createWindowManager({ stores, preloadPath, iconPath, onWindowEvent, confirmClose }) {
  /** @type {Map<number, { app: string, file: string|null, dirty: boolean, name: string, closing: boolean }>} */
  const meta = new Map();

  function savedBounds(appKey) {
    const saved = stores.settings.get(`window.${appKey}`, null);
    if (!saved) return null;
    // A monitor that has gone away must not strand a window off-screen.
    const areas = screen.getAllDisplays().map((d) => d.workArea);
    const visible = areas.some(
      (a) =>
        saved.x + saved.width > a.x + 40 &&
        saved.x < a.x + a.width - 40 &&
        saved.y + saved.height > a.y &&
        saved.y < a.y + a.height - 40
    );
    return visible ? saved : null;
  }

  function rememberBounds(win, appKey) {
    if (!win || win.isDestroyed() || win.isMinimized()) return;
    const b = win.getBounds();
    stores.settings.set(`window.${appKey}`, { ...b, maximized: win.isMaximized() });
  }

  function urlFor({ app: appKey, file, query }) {
    const params = new URLSearchParams({ app: appKey || 'home' });
    if (file) params.set('file', file);
    for (const [k, v] of Object.entries(query || {})) if (v != null) params.set(k, String(v));
    return `${SCHEME}://app/index.html?${params.toString()}`;
  }

  function create({ app: appKey = 'home', file = null, query = null, parentId = null } = {}) {
    const geo = GEOMETRY[appKey] || GEOMETRY.home;
    const saved = savedBounds(appKey);
    const dark = nativeTheme.shouldUseDarkColors;
    // A capture harness can ask for a room of its own size — a page-tall
    // window for a rendering comparison — without touching the saved bounds.
    const forced = /^(\d+)x(\d+)$/.exec(process.env.RUTBA_WINDOW_SIZE || '');

    const win = new BrowserWindow({
      width: forced ? Number(forced[1]) : saved?.width ?? geo.width,
      height: forced ? Number(forced[2]) : saved?.height ?? geo.height,
      x: saved?.x,
      y: saved?.y,
      minWidth: geo.minWidth,
      minHeight: geo.minHeight,
      show: false,
      backgroundColor: dark ? '#161719' : '#f3f3f5',
      icon: iconPath,
      title: 'Rutba Office',
      autoHideMenuBar: true,
      frame: false,
      titleBarStyle: isMac ? 'hiddenInset' : 'hidden',
      trafficLightPosition: isMac ? { x: 14, y: 16 } : undefined,
      parent: parentId ? BrowserWindow.fromId(parentId) ?? undefined : undefined,
      webPreferences: {
        preload: preloadPath,
        contextIsolation: true,
        sandbox: true,
        nodeIntegration: false,
        webviewTag: false,
        spellcheck: true,
        plugins: true, // Chromium's PDF viewer, for previewing what we export
        backgroundThrottling: false,
      },
    });

    meta.set(win.id, { app: appKey, file, dirty: false, name: '', closing: false });
    if (saved?.maximized) win.maximize();

    win.once('ready-to-show', () => win.show());

    const pushState = () => {
      if (win.isDestroyed()) return;
      onWindowEvent?.(win, 'win:state', {
        maximized: win.isMaximized(),
        fullscreen: win.isFullScreen(),
        focused: win.isFocused(),
      });
    };
    for (const ev of ['maximize', 'unmaximize', 'enter-full-screen', 'leave-full-screen', 'focus', 'blur']) {
      win.on(ev, pushState);
    }
    let saveTimer = null;
    for (const ev of ['resize', 'move']) {
      win.on(ev, () => {
        clearTimeout(saveTimer);
        saveTimer = setTimeout(() => rememberBounds(win, appKey), 400);
      });
    }
    /**
     * A window with unsaved work does not just close.
     *
     * This is the one defect a document application must not have. The close is
     * cancelled, the question is asked with the platform's own dialog, and the
     * window closes only once the answer is in — after saving, if that is what
     * was asked for. `close({ force: true })` is how the renderer says the
     * question has been answered.
     */
    win.on('close', (event) => {
      const info = meta.get(win.id);
      if (info?.dirty && !info.closing && typeof confirmClose === 'function') {
        event.preventDefault();
        info.closing = true;
        Promise.resolve(confirmClose(win, info))
          .then((decision) => {
            // 'cancel' leaves the window exactly as it was, including the flag,
            // so a second attempt asks again rather than closing silently.
            if (decision === 'cancel') info.closing = false;
          })
          .catch(() => {
            info.closing = false;
          });
        return;
      }
      clearTimeout(saveTimer);
      rememberBounds(win, appKey);
    });
    win.on('closed', () => meta.delete(win.id));

    // Nothing in this app navigates. A link goes to the browser, a popup is
    // refused, and a dropped file is handled by the renderer, not by Chromium
    // replacing the document with it.
    win.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
    win.webContents.on('will-navigate', (e, target) => {
      if (!target.startsWith(`${SCHEME}://app/`)) e.preventDefault();
    });

    win.loadURL(urlFor({ app: appKey, file, query }));
    return win;
  }

  /** The window showing this file already, if there is one — focus beats duplicate. */
  function findByFile(file) {
    for (const [id, m] of meta) {
      if (m.file && path.resolve(m.file) === path.resolve(file)) {
        const w = BrowserWindow.fromId(id);
        if (w && !w.isDestroyed()) return w;
      }
    }
    return null;
  }

  function open({ app: appKey, file, query }) {
    if (file) {
      const existing = findByFile(file);
      if (existing) {
        if (existing.isMinimized()) existing.restore();
        existing.focus();
        return existing;
      }
    }
    return create({ app: appKey, file, query });
  }

  function setFile(win, file) {
    const m = meta.get(win.id);
    if (m) m.file = file;
  }

  /** The renderer tells us whether this window is holding unsaved work. */
  function setDirty(win, { dirty, name }) {
    const m = meta.get(win.id);
    if (!m) return;
    m.dirty = Boolean(dirty);
    if (name) m.name = name;
  }

  /** Close past the guard, once the question has been answered. */
  function forceClose(win) {
    const m = meta.get(win.id);
    if (m) {
      m.dirty = false;
      m.closing = true;
    }
    win.close();
  }

  function infoFor(win) {
    return meta.get(win.id) || { app: 'home', file: null };
  }

  function all() {
    return BrowserWindow.getAllWindows();
  }

  return { create, open, findByFile, setFile, setDirty, forceClose, infoFor, all, encodePath, GEOMETRY };
}
