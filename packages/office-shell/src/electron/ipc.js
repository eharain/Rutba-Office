// The backend half of the contract.
//
// One handler per channel, registered from the same METHODS list the preload
// builds its bridge from. A namespace the host supplies (`mail`) is merged in
// here so the desktop app can own its own protocols without the shell knowing
// anything about IMAP.
//
// Every handler receives the calling window, because almost everything an
// office app asks for is about *this* document window, not about the app.

import { app, BrowserWindow, dialog, shell, nativeTheme, clipboard } from 'electron';
import fsp from 'node:fs/promises';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { ipcMain } from 'electron';
import { METHODS, CHANNEL_PREFIX } from '../contract.js';

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

function canWrite(p) {
  try {
    fs.accessSync(p, fs.constants.W_OK);
    return true;
  } catch {
    return false;
  }
}

export function buildImplementations({ stores, windows, quitting }) {
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

  impl.win = {
    create: (p, win) => {
      const created = windows.open({ app: p.app, file: p.file, query: p.query, parentId: p.modal ? win?.id : null });
      return { id: created.id };
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
    fullscreen: (p, win) => {
      if (!win) return { fullscreen: false };
      win.setFullScreen(p.on ?? !win.isFullScreen());
      return { fullscreen: win.isFullScreen() };
    },
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
      await fsp.writeFile(p, Buffer.from(bytes));
      return { stat: statOf(p) };
    },
    writeText: async ({ path: p, text }) => {
      await fsp.mkdir(path.dirname(p), { recursive: true });
      await fsp.writeFile(p, text, 'utf8');
      return { stat: statOf(p) };
    },
    stat: ({ path: p }) => statOf(p),
    exists: ({ path: p }) => fs.existsSync(p),
    list: async ({ path: p, filter }) => {
      const names = await fsp.readdir(p, { withFileTypes: true });
      const exts = filter ? new Set(filter.map((e) => (e.startsWith('.') ? e : `.${e}`).toLowerCase())) : null;
      const out = [];
      for (const e of names) {
        if (e.name.startsWith('.')) continue;
        const full = path.join(p, e.name);
        const isDir = e.isDirectory();
        if (!isDir && exts && !exts.has(path.extname(e.name).toLowerCase())) continue;
        const s = statOf(full);
        if (s) out.push(s);
      }
      out.sort((a, b) => (a.dir === b.dir ? a.name.localeCompare(b.name, undefined, { numeric: true }) : a.dir ? -1 : 1));
      return out;
    },
    mkdirp: async ({ path: p }) => void (await fsp.mkdir(p, { recursive: true })),
    // Deleting is always to the OS trash. An office suite must never be the
    // reason a file is unrecoverable.
    remove: async ({ path: p }) => void (await shell.trashItem(p)),
    rename: async ({ from, to }) => void (await fsp.rename(from, to)),
    copy: async ({ from, to }) => void (await fsp.copyFile(from, to)),
    temp: async ({ ext = '', bytes = null }) => {
      const dir = await fsp.mkdtemp(path.join(app.getPath('temp'), 'rutba-office-'));
      const p = path.join(dir, `scratch${ext.startsWith('.') || !ext ? ext : `.${ext}`}`);
      if (bytes) await fsp.writeFile(p, Buffer.from(bytes));
      return { path: p };
    },
  };

  impl.dialog = {
    open: async (p, win) => {
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
      return r.canceled ? [] : r.filePaths;
    },
    save: async (p, win) => {
      const r = await dialog.showSaveDialog(win ?? undefined, {
        title: p.title,
        defaultPath: p.defaultPath,
        filters: p.filters,
        properties: ['createDirectory', 'showOverwriteConfirmation'],
      });
      return r.canceled ? null : r.filePath;
    },
    message: async (p, win) => {
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
    openPath: ({ path: p }) => shell.openPath(p),
    beep: () => void shell.beep(),
  };

  impl.secrets = {
    available: () => stores.secrets.available(),
    get: ({ key }) => stores.secrets.get(key),
    set: ({ key, value }) => void stores.secrets.set(key, value),
    delete: ({ key }) => void stores.secrets.delete(key),
    keys: () => stores.secrets.keys(),
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
    writeText: ({ text }) => void clipboard.writeText(text ?? ''),
    readText: () => clipboard.readText(),
  };

  return impl;
}

/**
 * Register one ipcMain handler per contract channel.
 * @param {object} impls namespace -> { method(payload, win) }
 */
export function installIpc(impls) {
  for (const [ns, names] of Object.entries(METHODS)) {
    for (const name of names) {
      const channel = `${CHANNEL_PREFIX}/${ns}:${name}`;
      const fn = impls[ns]?.[name];
      ipcMain.handle(channel, async (event, payload) => {
        if (typeof fn !== 'function') {
          throw new Error(`Rutba Office: no backend for ${ns}:${name}`);
        }
        const win = BrowserWindow.fromWebContents(event.sender);
        try {
          return await fn(payload ?? {}, win);
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
