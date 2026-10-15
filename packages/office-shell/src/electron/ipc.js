// The backend half of the contract.
//
// One handler per channel, registered from the same METHODS list the preload
// builds its bridge from. A namespace the host supplies (`mail`) is merged in
// here so the desktop app can own its own protocols without the shell knowing
// anything about IMAP.
//
// Every handler receives the calling window, because almost everything an
// office app asks for is about *this* document window, not about the app.

import { app, BrowserWindow, dialog, shell, nativeTheme, clipboard, desktopCapturer, screen, nativeImage, session, globalShortcut, Tray, Menu } from 'electron';
import fsp from 'node:fs/promises';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { ipcMain } from 'electron';
import { METHODS, CHANNEL_PREFIX } from '../contract.js';
import { blobOwner } from './blobs.js';
import { runsWhenOpened } from '../runs.js';
import { writeWhole } from '../write-whole.js';
import { grantPlace } from './grants.js';
import { speakAs, t } from './words.js';
import { readHTML, readImage, writeImage, writeParts } from './clipboard.js';

const isMac = process.platform === 'darwin';

function statOf(p) {
  try {
    const s = fs.statSync(p);
    return {
      path: p,
      name: path.basename(p),
      dir: s.isDirectory(),
      size: s.size,
      mtime: s.mtimeMs,
      ctime: s.ctimeMs,
      ext: path.extname(p).toLowerCase(),
      readonly: !canWrite(p),
    };
  } catch {
    return null;
  }
}

/**
 * The same, without holding the main process: a folder of thousands of
 * pictures, or one on a slow drive, was stat'd one file at a time while
 * every window waited.
 */
async function statOfAsync(p) {
  try {
    const s = await fsp.stat(p);
    const writable = await fsp.access(p, fs.constants.W_OK).then(() => true, () => false);
    return { path: p, name: path.basename(p), dir: s.isDirectory(), size: s.size, mtime: s.mtimeMs, ctime: s.ctimeMs, ext: path.extname(p).toLowerCase(), readonly: !writable };
  } catch {
    return null;
  }
}

function canWrite(p) {
  try {
    fs.accessSync(p, fs.constants.W_OK);
    return true;
  } catch {
    return false;
  }
}

export function buildImplementations({ stores, windows, quitting, thumbnailer = null, iconPath = null }) {
  const impl = {};

  impl.app = {
    version: () => ({
      name: app.getName(),
      version: app.getVersion(),
      electron: process.versions.electron,
      chrome: process.versions.chrome,
      node: process.versions.node,
      v8: process.versions.v8,
      platform: process.platform,
      arch: process.arch,
      locale: app.getLocale(),
      dark: nativeTheme.shouldUseDarkColors,
    }),
    paths: () => {
      const get = (k, fallback) => {
        try {
          return app.getPath(k);
        } catch {
          return fallback;
        }
      };
      return {
        home: get('home', os.homedir()),
        documents: get('documents', os.homedir()),
        pictures: get('pictures', os.homedir()),
        videos: get('videos', os.homedir()),
        music: get('music', os.homedir()),
        downloads: get('downloads', os.homedir()),
        desktop: get('desktop', os.homedir()),
        userData: app.getPath('userData'),
        temp: app.getPath('temp'),
        separator: path.sep,
      };
    },
    recent: () => stores.recent.list(),
    addRecent: (p) => stores.recent.add(p),
    removeRecent: (p) => stores.recent.remove(p),
    clearRecent: () => stores.recent.clear(),
    quit: () => {
      quitting.value = true;
      app.quit();
    },
    relaunch: () => {
      app.relaunch();
      quitting.value = true;
      app.quit();
    },
  };

  // Edit in Image: the window a picture was opened from, by the window that edits it.
  const returnTo = new Map();

  impl.win = {
    create: (p, win) => {
      const created = windows.open({ app: p.app, file: p.file, query: p.query, parentId: p.modal ? win?.id : null });
      if (p.query?.editFor && win) returnTo.set(created.id, win.id);
      return { id: created.id };
    },
    handBack: (p, win) => {
      const to = win ? returnTo.get(win.id) : null;
      const target = to ? BrowserWindow.fromId(to) : null;
      if (!target || target.isDestroyed()) return { delivered: false };
      const bytes = p?.bytes instanceof Uint8Array ? p.bytes : new Uint8Array(p?.bytes || []);
      sendEvent(target, 'picture:returned', { token: String(p?.token || ''), bytes, contentType: String(p?.contentType || 'image/png'), width: Number(p?.width) || 0, height: Number(p?.height) || 0 });
      return { delivered: true };
    },
    close: (p, win) => {
      if (!win) return;
      // force means the save prompt has already been answered.
      if (p?.force) windows.forceClose(win);
      else win.close();
    },
    minimize: (_p, win) => void win?.minimize(),
    toggleMaximize: (_p, win) => {
      if (!win) return { maximized: false };
      if (win.isMaximized()) win.unmaximize();
      else win.maximize();
      return { maximized: win.isMaximized() };
    },
    isMaximized: (_p, win) => Boolean(win?.isMaximized()),
    setTitle: (p, win) => void win?.setTitle(p.title || 'Rutba Office'),
    setDirty: (p, win) => void (win && windows.setDirty(win, p)),
    setDocumentEdited: (p, win) => {
      if (!win) return;
      if (isMac) win.setDocumentEdited(Boolean(p.edited));
      // Windows and Linux have no such concept; the renderer shows the dot itself.
    },
    fullscreen: (p, win) => (win ? windows.fullscreen(win, p) : { fullscreen: false }),
    displays: (_p, win) => windows.displaysFor(win),
    arrange: (p, win) => (win ? windows.arrange(win, p.mode) : { count: 0 }),
    list: (_p, win) => (win ? windows.listFor(win) : []),
    focus: (p) => windows.focusById(p.id),
    hide: (_p, win) => (win ? windows.hideWindow(win) : { hidden: false }),
    unhide: (p) => windows.unhide(p.id),
    zoom: (p, win) => {
      if (!win) return { factor: 1 };
      const wc = win.webContents;
      if (p.reset) wc.setZoomFactor(1);
      else {
        const next = Math.min(3, Math.max(0.5, wc.getZoomFactor() + (p.delta || 0)));
        wc.setZoomFactor(next);
      }
      return { factor: wc.getZoomFactor() };
    },
    state: (_p, win) => ({
      maximized: Boolean(win?.isMaximized()),
      fullscreen: Boolean(win?.isFullScreen()),
      focused: Boolean(win?.isFocused()),
      platform: process.platform,
      dark: nativeTheme.shouldUseDarkColors,
    }),
    devtools: (_p, win) => void win?.webContents.toggleDevTools(),
  };

  // The scratch files `temp` made in this run: the only ones `append` and `dropTemp` touch.
  const temps = new Set();
  // The recordings capture.take made: appended to the same way, but their folder is a person's, kept.
  const takes = new Set();

  impl.fs = {
    read: async ({ path: p }) => {
      const bytes = await fsp.readFile(p);
      return { bytes: new Uint8Array(bytes), stat: statOf(p) };
    },
    // The first N bytes: what a header needs. The photo viewer read a whole
    // fifty-megabyte picture across the bridge to learn its size and its
    // EXIF, which live in the first quarter megabyte.
    // `offset` because a header is not always at the front: an MP4 written by
    // a camera keeps the atom that holds its duration at the END of the file.
    readHead: async ({ path: p, bytes = 262144, offset = 0 }) => {
      const handle = await fsp.open(p, 'r');
      try {
        const buf = Buffer.alloc(Math.max(0, Number(bytes) || 0));
        const { bytesRead } = await handle.read(buf, 0, buf.length, Math.max(0, Number(offset) || 0));
        return { bytes: new Uint8Array(buf.subarray(0, bytesRead)), stat: statOf(p) };
      } finally {
        await handle.close();
      }
    },

    readText: async ({ path: p, encoding = 'utf8' }) => ({
      text: await fsp.readFile(p, encoding),
      stat: statOf(p),
    }),
    write: async ({ path: p, bytes }) => {
      await fsp.mkdir(path.dirname(p), { recursive: true });
      await writeWhole(p, Buffer.from(bytes));
      return { stat: statOf(p) };
    },
    writeText: async ({ path: p, text }) => {
      await fsp.mkdir(path.dirname(p), { recursive: true });
      await writeWhole(p, text, 'utf8');
      return { stat: statOf(p) };
    },
    stat: ({ path: p }) => statOf(p),
    exists: ({ path: p }) => fs.existsSync(p),
    list: async ({ path: p, filter }) => {
      const names = await fsp.readdir(p, { withFileTypes: true });
      // A folder browsed: its pictures may be drawn, a share's among them (grants.js).
      grantPlace(p);
      const exts = filter ? new Set(filter.map((e) => (e.startsWith('.') ? e : `.${e}`).toLowerCase())) : null;
      const wanted = [];
      for (const e of names) {
        if (e.name.startsWith('.')) continue;
        const isDir = e.isDirectory();
        if (!isDir && exts && !exts.has(path.extname(e.name).toLowerCase())) continue;
        wanted.push(path.join(p, e.name));
      }
      // Looked at sixty-four at a time, each without waiting on the main process.
      const out = [];
      for (let i = 0; i < wanted.length; i += 64) {
        for (const s of await Promise.all(wanted.slice(i, i + 64).map(statOfAsync))) if (s) out.push(s);
      }
      out.sort((a, b) => (a.dir === b.dir ? a.name.localeCompare(b.name, undefined, { numeric: true }) : a.dir ? -1 : 1));
      return out;
    },
    mkdirp: async ({ path: p }) => void (await fsp.mkdir(p, { recursive: true })),
    // Deleting is always to the OS trash. An office suite must never be the
    // reason a file is unrecoverable.
    remove: async ({ path: p }) => void (await shell.trashItem(p)),
    temp: async ({ ext = '', bytes = null }) => {
      const dir = await fsp.mkdtemp(path.join(app.getPath('temp'), 'rutba-office-'));
      const p = path.join(dir, `scratch${ext.startsWith('.') || !ext ? ext : `.${ext}`}`);
      if (bytes) await fsp.writeFile(p, Buffer.from(bytes));
      temps.add(p);
      return { path: p };
    },
    // A scratch file written a piece at a time — a screen recording as it is
    // made — so no one message carries the whole; only one `temp` made.
    append: async ({ path: p, bytes }) => {
      if (!temps.has(p)) throw new Error('Only a scratch file made in this run is added to.');
      await fsp.appendFile(p, Buffer.from(bytes));
      return { size: (await fsp.stat(p)).size };
    },
    // An export streamed to a scratch file writes one thing twice: a size
    // near the front, known only once the rest is in.
    patch: async ({ path: p, offset, bytes }) => {
      if (!temps.has(p) || takes.has(p)) throw new Error('Only a scratch file made in this run is written over.');
      const handle = await fsp.open(p, 'r+');
      try {
        const buf = Buffer.from(bytes);
        await handle.write(buf, 0, buf.length, Math.max(0, Number(offset) || 0));
        return { size: (await handle.stat()).size };
      } finally {
        await handle.close();
      }
    },
    // The finished file to its place: renamed where it can be, copied across drives where it cannot.
    keepTemp: async ({ path: p, to }) => {
      if (!temps.has(p) || takes.has(p)) throw new Error('Only a scratch file made in this run is kept.');
      if (!to || !path.isAbsolute(String(to))) throw new Error('A file needs a place to go.');
      await fsp.mkdir(path.dirname(to), { recursive: true });
      try {
        await fsp.rename(p, to);
      } catch {
        await fsp.copyFile(p, to);
        await fsp.rm(p, { force: true });
      }
      temps.delete(p);
      await fsp.rmdir(path.dirname(p)).catch(() => {});
      return { stat: statOf(to) };
    },
    dropTemp: async ({ path: p }) => {
      if (!temps.delete(p)) return;
      await fsp.rm(p, { force: true });
      if (takes.delete(p)) return;
      await fsp.rmdir(path.dirname(p)).catch(() => {});
    },
  };

  impl.thumbs = {
    put: async ({ path: p, bytes }) => (thumbnailer ? thumbnailer.put(p, bytes) : { stored: false }),
  };

  impl.dialog = {
    open: async (p, win) => {
      const planned = plannedAnswer('open');
      if (planned !== undefined) { for (const f of [].concat(planned || [])) if (typeof f === 'string') grantPlace(f); return planned; }
      const props = ['openFile'];
      if (p.multiple) props.push('multiSelections');
      if (p.directory) {
        props.length = 0;
        props.push('openDirectory');
      }
      const r = await dialog.showOpenDialog(win ?? undefined, {
        title: p.title,
        defaultPath: p.defaultPath,
        filters: p.filters,
        properties: props,
      });
      if (r.canceled) return [];
      // Picked by the person: drawn from, on a share too (grants.js).
      for (const f of r.filePaths) grantPlace(f);
      return r.filePaths;
    },
    save: async (p, win) => {
      const planned = plannedAnswer('save');
      if (planned !== undefined) return planned;
      const r = await dialog.showSaveDialog(win ?? undefined, {
        title: p.title,
        defaultPath: p.defaultPath,
        filters: p.filters,
        properties: ['createDirectory', 'showOverwriteConfirmation'],
      });
      return r.canceled ? null : r.filePath;
    },
    message: async (p, win) => {
      const planned = plannedAnswer('message');
      if (planned !== undefined) return planned;
      const r = await dialog.showMessageBox(win ?? undefined, {
        type: p.type || 'none',
        message: p.message || '',
        detail: p.detail,
        buttons: p.buttons || ['OK'],
        defaultId: p.defaultId ?? 0,
        cancelId: p.cancelId ?? (p.buttons ? p.buttons.length - 1 : 0),
        checkboxLabel: p.checkboxLabel,
        noLink: true,
      });
      return { response: r.response, checked: r.checkboxChecked };
    },
    error: (p) => void dialog.showErrorBox(p.title || 'Rutba Office', p.content || ''),
  };

  impl.shell = {
    // Only protocols a document can legitimately point at.
    openExternal: ({ url }) => {
      const ok = /^(https?|mailto|tel):/i.test(url || '');
      if (!ok) throw new Error(`refused to open: ${url}`);
      return shell.openExternal(url);
    },
    showInFolder: ({ path: p }) => void shell.showItemInFolder(p),
    // A program is opened only after a warning shown from here, not from
    // the window: a window that asked to open one cannot also have answered
    // for the person. It is still their file and their choice, so it is
    // asked, not refused.
    openPath: async ({ path: p }, win) => {
      if (runsWhenOpened(p)) {
        const { response } = await impl.dialog.message({
          type: 'warning',
          message: `${path.basename(String(p))} is a program. Opening it runs it on this computer.`,
          detail: 'Rutba Office cannot tell whether a program is safe. Open it only if you were expecting it from someone you know; a message that presses you to open an attachment is the commonest way a computer is taken over.',
          buttons: ['Open', 'Cancel'],
          defaultId: 1,
          cancelId: 1,
        }, win);
        if (response !== 0) return false;
      }
      const failed = await shell.openPath(p);
      if (failed) throw new Error(failed);
      return true;
    },
    beep: () => void shell.beep(),
  };

  impl.store = {
    get: ({ key, fallback }) => stores.settings.get(key, fallback ?? null),
    set: ({ key, value }) => void stores.settings.set(key, value),
    delete: ({ key }) => void stores.settings.delete(key),
    all: () => stores.settings.all(),
  };

  impl.print = {
    toPDF: async (p, win) => {
      const bytes = await win.webContents.printToPDF({
        landscape: Boolean(p.landscape),
        printBackground: p.printBackground ?? true,
        pageSize: p.pageSize || 'A4',
        margins: p.margins || { marginType: 'default' },
        preferCSSPageSize: p.preferCSSPageSize ?? true,
      });
      return { bytes: new Uint8Array(bytes) };
    },
    print: (p, win) =>
      new Promise((resolve) => {
        win.webContents.print({ silent: Boolean(p.silent), printBackground: true }, (ok, reason) =>
          resolve({ ok, reason })
        );
      }),
  };

  impl.clipboard = {
    // Every call answers once the clipboard has taken or given its part, so a
    // window that awaits a write can read it back at once.
    writeText: async ({ text }) => {
      await clipboard.writeText(text ?? '');
    },
    readText: () => clipboard.readText(),
    write: async ({ text, html }) => {
      await writeParts({ text: text ?? '', html });
    },
    read: async () => ({ text: await clipboard.readText(), html: await readHTML() }),
    // Image: Copy puts the picture on the clipboard, Paste takes one off it.
    writeImage: async ({ bytes }) => {
      const picture = nativeImage.createFromBuffer(Buffer.from(bytes || []));
      if (picture.isEmpty()) throw new Error('That is not a picture the clipboard can hold.');
      await writeImage(picture);
      return picture.getSize();
    },
    readImage: async () => {
      const picture = await readImage();
      if (picture.isEmpty()) return null;
      const { width, height } = picture.getSize();
      return { bytes: new Uint8Array(picture.toPNG()), width, height };
    },
  };

  // Insert → Screenshot. Only on a press of the button, and only into the
  // document of the window that asked: the list leaves that window out, as
  // Office's does, and a picture is taken of the one chosen and nothing else.
  // What each window's next getDisplayMedia is given (capture.choose): a
  // source and whether the computer's own sound comes with it.
  const chosen = new Map();
  let displayHandler = false;
  const ownerOf = (frame) => {
    const top = frame?.top || frame;
    if (!top) return null;
    return BrowserWindow.getAllWindows().find((w) => {
      if (w.isDestroyed()) return false;
      const m = w.webContents.mainFrame;
      return m === top || (m.processId === top.processId && m.routingId === top.routingId);
    }) || null;
  };
  const installDisplayHandler = () => {
    if (displayHandler) return;
    displayHandler = true;
    // A page's getDisplayMedia is answered here, never by a picker of
    // Chromium's: with the source this window chose in its own list, and on
    // Windows the computer's own sound ('loopback'). A request no window
    // chose for is refused.
    session.defaultSession.setDisplayMediaRequestHandler(async (request, callback) => {
      try {
        const win = ownerOf(request.frame);
        const choice = win ? chosen.get(win.webContents.id) : null;
        if (!choice) return callback({});
        chosen.delete(win.webContents.id);
        const kind = String(choice.id).startsWith('screen:') ? 'screen' : 'window';
        const found = await desktopCapturer.getSources({ types: [kind], thumbnailSize: { width: 0, height: 0 }, fetchWindowIcons: false });
        const source = found.find((x) => x.id === choice.id);
        if (!source) return callback({});
        const answer = { video: source };
        if (choice.audio && request.audioRequested && process.platform === 'win32') answer.audio = 'loopback';
        return callback(answer);
      } catch {
        return callback({});
      }
    });
  };

  // While a window records: the global keys and the tray icon that stop and
  // pause it, released when it stops or the window goes. A check run never
  // takes a key from the person's keyboard or puts an icon on their desktop:
  // it is handed the press instead (globalThis.__rutbaCaptureControls).
  const KEYS = { stop: 'CommandOrControl+Shift+F10', pause: 'CommandOrControl+Shift+F9' };
  const SHOWN_KEYS = { stop: isMac ? 'Cmd+Shift+F10' : 'Ctrl+Shift+F10', pause: isMac ? 'Cmd+Shift+F9' : 'Ctrl+Shift+F9' };
  let held = null; // { win, tray, minimised, release }
  const CHECK = Boolean(process.env.RUTBA_OFFICE_VERIFY_APPS || process.env.RUTBA_OFFICE_VERIFY_EDIT || process.env.RUTBA_SCREENS_FEATURES);
  const releaseControls = () => {
    const h = held;
    held = null;
    if (!h) return;
    if (!CHECK) {
      for (const k of Object.values(KEYS)) { try { globalShortcut.unregister(k); } catch { /* not ours */ } }
      try { h.tray?.destroy(); } catch { /* gone */ }
      if (h.minimised && !h.win.isDestroyed()) { h.win.restore(); h.win.focus(); }
    }
    if (!h.win.isDestroyed()) h.win.removeListener('closed', h.onClosed);
    if (CHECK) globalThis.__rutbaCaptureControls = { on: false, keys: SHOWN_KEYS, press: () => false };
  };
  const trayMenu = (win, paused) => {
    speakAs(stores);
    const send = (command) => () => { if (!win.isDestroyed()) win.webContents.send(`${CHANNEL_PREFIX}/event/capture:command`, { command }); };
    return Menu.buildFromTemplate([
      { label: t('Stop recording'), accelerator: KEYS.stop, click: send('stop') },
      { label: paused ? t('Resume recording') : t('Pause recording'), accelerator: KEYS.pause, click: send('pause') },
    ]);
  };

  // The pointer while a window records (R7): where it is over the screen or
  // window being recorded, polled at the frame rate, and where it was
  // pressed, as fractions of that source, on the recording's clock (the
  // time paused left out). A screen is mapped by its display's bounds and
  // one of the suite's own windows by its own; another program's window
  // cannot be, so the pointer is not recorded for it. A press is seen in the
  // suite's own windows only: nothing outside them is listened to. A check
  // run reads no mouse: it hands the positions and the presses in
  // (globalThis.__rutbaPointerFeed).
  let tracking = null; // { timer, samples, clicks, started, heldAt, held, toSource, listeners }
  const stopTracking = () => {
    const tr = tracking;
    tracking = null;
    if (!tr) return null;
    clearInterval(tr.timer);
    for (const [wc, fn] of tr.listeners) { try { if (!wc.isDestroyed()) wc.removeListener('input-event', fn); } catch { /* gone */ } }
    if (CHECK) globalThis.__rutbaPointerFeed = null;
    return { samples: tr.samples, clicks: tr.clicks };
  };
  const pointerClock = (tr) => {
    const now = Date.now();
    const end = tr.heldAt != null ? tr.heldAt : now;
    return Math.max(0, (end - tr.started - tr.held) / 1000);
  };

  impl.capture = {
    pointer: async ({ on = false, id = '', fps = 30, paused = undefined } = {}) => {
      if (!on) return stopTracking() || { samples: [], clicks: [] };
      if (tracking && paused !== undefined && !id) {
        // Pause and resume: the clock holds while paused, and nothing is written down.
        if (paused && tracking.heldAt == null) tracking.heldAt = Date.now();
        if (!paused && tracking.heldAt != null) { tracking.held += Date.now() - tracking.heldAt; tracking.heldAt = null; }
        return { tracking: true };
      }
      stopTracking();
      const sid = String(id || '');
      let toSource = null;
      if (sid.startsWith('screen:')) {
        const found = await desktopCapturer.getSources({ types: ['screen'], thumbnailSize: { width: 0, height: 0 }, fetchWindowIcons: false }).catch(() => []);
        const src = found.find((s) => s.id === sid);
        const displays = screen.getAllDisplays();
        const d = displays.find((x) => String(x.id) === String(src?.display_id)) || (displays.length === 1 ? displays[0] : null);
        if (d) toSource = (p) => ({ x: (p.x - d.bounds.x) / d.bounds.width, y: (p.y - d.bounds.y) / d.bounds.height });
      } else {
        const w = BrowserWindow.getAllWindows().find((x) => !x.isDestroyed() && x.getMediaSourceId() === sid);
        if (w) toSource = (p) => { if (w.isDestroyed()) return null; const b = w.getBounds(); return { x: (p.x - b.x) / b.width, y: (p.y - b.y) / b.height }; };
      }
      if (!toSource && !CHECK) return { tracking: false };
      const tr = { timer: null, samples: [], clicks: [], started: Date.now(), heldAt: null, held: 0, listeners: [] };
      tracking = tr;
      const r3 = (n) => Math.round(n * 1000) / 1000;
      const feed = CHECK ? { at: null, move(x, y) { feed.at = { x, y }; return true; }, click(x, y) { if (tr.heldAt == null) tr.clicks.push([r3(pointerClock(tr)), r3(x), r3(y)]); return true; } } : null;
      if (CHECK) globalThis.__rutbaPointerFeed = feed;
      tr.timer = setInterval(() => {
        if (tr.heldAt != null) return;
        let at = null;
        if (feed) at = feed.at;
        else { const p = toSource(screen.getCursorScreenPoint()); at = p; }
        if (!at) return;
        tr.samples.push([r3(pointerClock(tr)), r3(at.x), r3(at.y)]);
      }, Math.max(15, Math.round(1000 / Math.max(5, Math.min(60, Number(fps) || 30)))));
      if (!CHECK && toSource) {
        // A press in one of the suite's windows, where the screen recorded can see it.
        for (const w of BrowserWindow.getAllWindows()) {
          if (w.isDestroyed()) continue;
          const wc = w.webContents;
          const fn = (_e, input) => {
            if (input?.type !== 'mouseDown' || tr.heldAt != null) return;
            const b = w.getContentBounds();
            const p = toSource({ x: b.x + (input.x || 0), y: b.y + (input.y || 0) });
            if (p && p.x >= 0 && p.x <= 1 && p.y >= 0 && p.y <= 1) tr.clicks.push([r3(pointerClock(tr)), r3(p.x), r3(p.y)]);
          };
          wc.on('input-event', fn);
          tr.listeners.push([wc, fn]);
        }
      }
      return { tracking: true };
    },
    sources: async (_p, win) => {
      const own = win && !win.isDestroyed() ? win.getMediaSourceId() : null;
      const found = await desktopCapturer.getSources({ types: ['window', 'screen'], thumbnailSize: { width: 320, height: 200 }, fetchWindowIcons: false });
      const displays = screen.getAllDisplays();
      const ours = CHECK ? new Set(BrowserWindow.getAllWindows().filter((w) => !w.isDestroyed()).map((w) => w.getMediaSourceId())) : null;
      const sizeOf = (s) => {
        const d = displays.find((x) => String(x.id) === String(s.display_id)) || (displays.length === 1 ? displays[0] : null);
        return d ? { width: Math.round(d.size.width * d.scaleFactor), height: Math.round(d.size.height * d.scaleFactor) } : null;
      };
      return found
        // A check run lists only this application's own windows: nothing of the
        // person's desktop is looked at, or drawn in a picture the run keeps.
        .filter((s) => s.id !== own && !s.thumbnail.isEmpty() && (!CHECK || ours.has(s.id)))
        .map((s) => {
          const kind = s.id.startsWith('screen:') ? 'screen' : 'window';
          return { id: s.id, name: s.name, kind, thumbnail: s.thumbnail.toDataURL(), ...(kind === 'screen' ? { display: sizeOf(s) } : {}) };
        })
        .sort((a, b) => (a.kind === b.kind ? 0 : a.kind === 'window' ? -1 : 1));
    },
    choose: ({ id, audio = false }, win) => {
      if (!win || !id) throw new Error('Nothing was chosen to record.');
      installDisplayHandler();
      chosen.set(win.webContents.id, { id: String(id), audio: Boolean(audio) });
      return { loopback: process.platform === 'win32' };
    },
    controls: ({ on = false, paused = false, minimise = false }, win) => {
      if (!on || !win) {
        if (!held || !win || held.win === win) releaseControls();
        return { stop: SHOWN_KEYS.stop, pause: SHOWN_KEYS.pause, tray: false };
      }
      if (held && held.win === win) {
        // Only the tray's words change between pause and resume.
        if (held.tray) held.tray.setContextMenu(trayMenu(win, paused));
        return { stop: SHOWN_KEYS.stop, pause: SHOWN_KEYS.pause, tray: Boolean(held.tray) };
      }
      releaseControls();
      const send = (command) => { if (!win.isDestroyed()) win.webContents.send(`${CHANNEL_PREFIX}/event/capture:command`, { command }); };
      const onClosed = () => releaseControls();
      win.once('closed', onClosed);
      held = { win, tray: null, minimised: false, onClosed };
      if (CHECK) {
        globalThis.__rutbaCaptureControls = { on: true, keys: SHOWN_KEYS, press: (command) => { send(command); return true; } };
        return { stop: SHOWN_KEYS.stop, pause: SHOWN_KEYS.pause, tray: false };
      }
      try { globalShortcut.register(KEYS.stop, () => send('stop')); } catch { /* taken by another program */ }
      try { globalShortcut.register(KEYS.pause, () => send('pause')); } catch { /* taken by another program */ }
      try {
        if (iconPath) {
          speakAs(stores);
          const tray = new Tray(iconPath);
          tray.setToolTip(t('Rutba Office is recording'));
          tray.setContextMenu(trayMenu(win, paused));
          tray.on('click', () => send('stop'));
          held.tray = tray;
        }
      } catch {
        held.tray = null;
      }
      if (minimise && !win.isMinimized()) { win.minimize(); held.minimised = true; }
      return { stop: SHOWN_KEYS.stop, pause: SHOWN_KEYS.pause, tray: Boolean(held.tray) };
    },
    space: async ({ path: p }) => {
      // The nearest folder of the path that is there: a recordings folder not made yet is on its parent's disk.
      let at = p ? path.resolve(String(p)) : app.getPath('temp');
      for (let i = 0; i < 40 && !fs.existsSync(at); i++) {
        const up = path.dirname(at);
        if (up === at) break;
        at = up;
      }
      const st = await fsp.statfs(at);
      return { free: Number(st.bavail) * Number(st.bsize), total: Number(st.blocks) * Number(st.bsize) };
    },
    take: async ({ dir, name = 'Recording.webm' }) => {
      if (!dir || !path.isAbsolute(String(dir))) throw new Error('A recording needs a folder to go in.');
      await fsp.mkdir(dir, { recursive: true });
      const safe = String(name).replace(/[\\/:*?"<>|]+/g, ' ').trim() || 'Recording.webm';
      const ext = path.extname(safe);
      const stem = safe.slice(0, safe.length - ext.length);
      let p = path.join(dir, safe);
      for (let n = 2; fs.existsSync(p); n++) p = path.join(dir, `${stem} (${n})${ext}`);
      await fsp.writeFile(p, Buffer.alloc(0), { flag: 'wx' });
      temps.add(p);
      takes.add(p);
      return { path: p };
    },
    grab: async ({ id }) => {
      // As large as the largest screen, so a window or a screen comes at its own size.
      const big = screen.getAllDisplays().reduce((m, d) => ({ width: Math.max(m.width, Math.round(d.size.width * d.scaleFactor)), height: Math.max(m.height, Math.round(d.size.height * d.scaleFactor)) }), { width: 0, height: 0 });
      // Only the kind the id names: a window's grab drew every screen at full
      // size as well, and a screen's every window, to keep one of them.
      const kind = String(id).startsWith('screen:') ? 'screen' : 'window';
      const found = await desktopCapturer.getSources({ types: [kind], thumbnailSize: big, fetchWindowIcons: false });
      const source = found.find((s) => s.id === id);
      if (!source) throw new Error('That window has closed since the list was made.');
      if (source.thumbnail.isEmpty()) throw new Error(`${source.name} could not be captured — a minimised window shows nothing to take.`);
      const { width, height } = source.thumbnail.getSize();
      return { bytes: new Uint8Array(source.thumbnail.toPNG()), width, height, name: source.name };
    },
  };

  return impl;
}

/**
 * Answers a check run hands the system's dialogs, taken in order by kind.
 *
 * A native Open or Save dialog is a window no check can click, so a button
 * that opens one could not be checked at all, and the checks steered round
 * every such button. In a check run — and only then — a check puts the
 * answer a person would give on this queue first: `{ kind: 'save', answer:
 * path }`, `'open'` with a list of paths, `'message'` with `{ response }`.
 * A dialog with no answer waiting opens as it always does.
 */
const CHECK_RUN = Boolean(process.env.RUTBA_OFFICE_VERIFY_APPS || process.env.RUTBA_OFFICE_VERIFY_EDIT || process.env.RUTBA_SCREENS_FEATURES);
if (CHECK_RUN) globalThis.__rutbaCheckDialogAnswers = globalThis.__rutbaCheckDialogAnswers || [];
function plannedAnswer(kind) {
  const queue = CHECK_RUN ? globalThis.__rutbaCheckDialogAnswers : null;
  const at = queue ? queue.findIndex((a) => a.kind === kind) : -1;
  return at < 0 ? undefined : queue.splice(at, 1)[0].answer;
}

/**
 * Register one ipcMain handler per contract channel.
 * @param {object} impls namespace -> { method(payload, win) }
 */
export function installIpc(impls) {
  // RUTBA_IPC_TRACE=1: count every request by channel and window, and print
  // the busiest every five seconds. A main process that is never idle is
  // answering somebody; this says whom, and for what.
  const trace = process.env.RUTBA_IPC_TRACE ? new Map() : null;
  if (trace) {
    setInterval(() => {
      if (!trace.size) return;
      const rows = [...trace.entries()].sort((a, b) => b[1] - a[1]).slice(0, 6).map(([k, n]) => `${k} ${n}`);
      console.log(`     [ipc] last 5 s: ${rows.join(' · ')}`);
      trace.clear();
    }, 5000).unref();
  }
  for (const [ns, names] of Object.entries(METHODS)) {
    for (const name of names) {
      const channel = `${CHANNEL_PREFIX}/${ns}:${name}`;
      const fn = impls[ns]?.[name];
      ipcMain.handle(channel, async (event, payload) => {
        if (typeof fn !== 'function') {
          throw new Error(`Rutba Office: no backend for ${ns}:${name}`);
        }
        const win = BrowserWindow.fromWebContents(event.sender);
        if (trace) {
          const key = `${ns}:${name}#${win?.id ?? '?'}`;
          trace.set(key, (trace.get(key) || 0) + 1);
        }
        try {
          // Anything held for the window during the request is owned by it.
          return await blobOwner.run(win?.id ?? null, () => fn(payload ?? {}, win));
        } catch (err) {
          // A window that has gone cannot be answered, and its last request
          // failing is not a fault: the documents it held were freed the
          // moment it closed, so a viewport request already in flight lands
          // on a session that is no longer there. Every closed window logged
          // a stack for it. Nothing is waiting for the answer either.
          if (!win || win.isDestroyed() || event.sender.isDestroyed()) return null;
          // Errors cross IPC as a plain shape; an Error instance loses its
          // message on the way and the renderer shows "an object".
          const wrapped = new Error(err?.message || String(err));
          wrapped.name = err?.name || 'ShellError';
          wrapped.code = err?.code;
          throw wrapped;
        }
      });
    }
  }
}

export function sendEvent(win, event, payload) {
  if (!win || win.isDestroyed()) return;
  win.webContents.send(`${CHANNEL_PREFIX}/event/${event}`, payload);
}

export function broadcast(event, payload) {
  for (const w of BrowserWindow.getAllWindows()) sendEvent(w, event, payload);
}
