// The Electron backend.
//
// `createShell()` is everything the desktop app needs from the operating
// system: one instance, the rutba:// scheme, windows, the menu and its
// accelerators, and the file-association plumbing that makes double-clicking a
// .docx open Word rather than a second copy of the launcher.
//
// The app supplies its own namespaces (mail) and its command handler. Nothing
// about documents, formats or IMAP is known here — that is the point of the
// seam.

import { app, BrowserWindow, Menu, dialog, nativeTheme, shell as electronShell } from 'electron';
import path from 'node:path';
import fs from 'node:fs';
import { createStores } from './store.js';
import { registerSchemePrivileges, installProtocol, fileUrl, holdBlob, releaseBlob, SCHEME } from './protocol.js';
import { createWindowManager } from './windows.js';
import { buildImplementations, installIpc, sendEvent, broadcast } from './ipc.js';

const isMac = process.platform === 'darwin';

// Registered before app ready, or the scheme is not privileged when the first
// window loads. Callers get this by importing the module, so it must be safe to
// run at import time.
registerSchemePrivileges();

function buildMenu({ send, appName }) {
  const cmd = (command, args) => () => send(command, args);
  const template = [
    ...(isMac
      ? [
          {
            label: appName,
            submenu: [
              { role: 'about' },
              { type: 'separator' },
              { label: 'Settings…', accelerator: 'Cmd+,', click: cmd('app.settings') },
              { type: 'separator' },
              { role: 'services' },
              { type: 'separator' },
              { role: 'hide' },
              { role: 'hideOthers' },
              { role: 'unhide' },
              { type: 'separator' },
              { role: 'quit' },
            ],
          },
        ]
      : []),
    {
      label: '&File',
      submenu: [
        { label: 'New', accelerator: 'CmdOrCtrl+N', click: cmd('file.new') },
        { label: 'Open…', accelerator: 'CmdOrCtrl+O', click: cmd('file.open') },
        { label: 'Open Recent', role: 'recentDocuments', submenu: [{ label: 'Clear', role: 'clearRecentDocuments' }] },
        { type: 'separator' },
        { label: 'Save', accelerator: 'CmdOrCtrl+S', click: cmd('file.save') },
        { label: 'Save As…', accelerator: 'CmdOrCtrl+Shift+S', click: cmd('file.saveAs') },
        { label: 'Export as PDF…', accelerator: 'CmdOrCtrl+Shift+E', click: cmd('file.exportPdf') },
        { type: 'separator' },
        { label: 'Print…', accelerator: 'CmdOrCtrl+P', click: cmd('file.print') },
        { type: 'separator' },
        { label: 'Close Window', accelerator: 'CmdOrCtrl+W', role: 'close' },
        ...(isMac ? [] : [{ label: 'Exit', role: 'quit' }]),
      ],
    },
    {
      label: '&Edit',
      submenu: [
        { label: 'Undo', accelerator: 'CmdOrCtrl+Z', click: cmd('edit.undo') },
        { label: 'Redo', accelerator: isMac ? 'Cmd+Shift+Z' : 'Ctrl+Y', click: cmd('edit.redo') },
        { type: 'separator' },
        { role: 'cut' },
        { role: 'copy' },
        { role: 'paste' },
        { label: 'Paste as Plain Text', accelerator: 'CmdOrCtrl+Shift+V', role: 'pasteAndMatchStyle' },
        { role: 'selectAll' },
        { type: 'separator' },
        { label: 'Find…', accelerator: 'CmdOrCtrl+F', click: cmd('edit.find') },
        { label: 'Replace…', accelerator: 'CmdOrCtrl+H', click: cmd('edit.replace') },
      ],
    },
    {
      label: '&View',
      submenu: [
        { label: 'Zoom In', accelerator: 'CmdOrCtrl+Plus', click: cmd('view.zoomIn') },
        { label: 'Zoom Out', accelerator: 'CmdOrCtrl+-', click: cmd('view.zoomOut') },
        { label: 'Actual Size', accelerator: 'CmdOrCtrl+0', click: cmd('view.zoomReset') },
        { type: 'separator' },
        { label: 'Toggle Theme', accelerator: 'CmdOrCtrl+Shift+D', click: cmd('view.toggleTheme') },
        { label: 'Full Screen', accelerator: isMac ? 'Ctrl+Cmd+F' : 'F11', click: cmd('view.fullscreen') },
        { type: 'separator' },
        { label: 'Reload', accelerator: 'CmdOrCtrl+R', role: 'reload' },
        { label: 'Developer Tools', accelerator: isMac ? 'Alt+Cmd+I' : 'Ctrl+Shift+I', role: 'toggleDevTools' },
      ],
    },
    {
      label: '&Apps',
      submenu: [
        { label: 'Home', accelerator: 'CmdOrCtrl+Shift+H', click: cmd('open.app', { app: 'home' }) },
        { type: 'separator' },
        { label: 'Mail', click: cmd('open.app', { app: 'mail' }) },
        { label: 'Word', click: cmd('open.app', { app: 'word' }) },
        { label: 'Worksheets', click: cmd('open.app', { app: 'sheets' }) },
        { label: 'Presentation', click: cmd('open.app', { app: 'slides' }) },
        { label: 'Pictures', click: cmd('open.app', { app: 'pictures' }) },
        { label: 'Image', click: cmd('open.app', { app: 'image' }) },
        { label: 'Video', click: cmd('open.app', { app: 'video' }) },
      ],
    },
    {
      label: '&Help',
      submenu: [
        { label: 'Rutba Office Help', accelerator: 'F1', click: cmd('help.show') },
        { label: 'Keyboard Shortcuts', click: cmd('help.shortcuts') },
        { type: 'separator' },
        { label: 'Rutba Office Website', click: () => electronShell.openExternal('https://office.rutba.io') },
        { label: 'Contact Us', click: () => electronShell.openExternal('https://office.rutba.io/contact') },
        { label: 'Source Code (AGPL)', click: () => electronShell.openExternal('https://github.com/eharain/Rutba-Office') },
        { type: 'separator' },
        { label: 'Check for Updates…', click: cmd('help.updates') },
        { type: 'separator' },
        { label: 'About Rutba Office', click: cmd('help.about') },
      ],
    },
  ];
  return Menu.buildFromTemplate(template);
}

/**
 * @param {object} o
 * @param {string} o.rendererDir  directory holding index.html + bundle
 * @param {string} o.preloadPath  bundled CommonJS preload
 * @param {string} [o.iconPath]
 * @param {Record<string,string>} [o.appIcons]  per-app .ico paths on disk, for the taskbar
 * @param {Record<string,string>} [o.appNames]  per-app display names, for a pinned taskbar entry
 * @param {string} [o.appName]
 * @param {(ctx) => object} [o.namespaces] extra IPC namespaces, given the shell context
 * @param {(path: string) => string} [o.appForFile] which app opens this file
 */
export function createShell({
  rendererDir,
  preloadPath,
  iconPath,
  appIcons = {},
  appNames = {},
  appName = 'Rutba Office',
  namespaces,
  appForFile = () => 'home',
  onReady,
}) {
  app.setName(appName);
  const appUserModelId = 'co.techstyle.rutba.office';
  if (process.platform === 'win32') app.setAppUserModelId(appUserModelId);

  const quitting = { value: false };

  // An uncaught error in the main process must never become a dialog. With
  // no listener Electron shows "A JavaScript error occurred in the main
  // process" — a stack trace in a box, in front of a person's document, for
  // an error they can do nothing about (a fetch whose stream Node closed
  // twice, say). It is written to errors.log in the profile instead, with
  // the time, and the application carries on. The log is the bug report.
  const errorLog = () => {
    try {
      return path.join(app.getPath('userData'), 'errors.log');
    } catch {
      return null;
    }
  };
  const record = (kind, err) => {
    const line = `${new Date().toISOString()} ${kind}: ${err?.stack || err?.message || String(err)}\n`;
    console.error(line.trim());
    const file = errorLog();
    if (!file) return;
    try {
      // Capped: a fault that repeats every second must not fill the disk.
      if (fs.existsSync(file) && fs.statSync(file).size > 512 * 1024) fs.truncateSync(file, 0);
      fs.appendFileSync(file, line);
    } catch {
      /* the log is best effort */
    }
  };
  process.on('uncaughtException', (err) => record('uncaught exception', err));
  process.on('unhandledRejection', (reason) => record('unhandled rejection', reason));

  /** Files handed to us before the app was ready (double-click at cold start). */
  const pending = [];

  const gotLock = app.requestSingleInstanceLock();
  if (!gotLock) {
    app.quit();
    return { app };
  }

  let stores = null;
  let windows = null;

  function send(command, args) {
    const win = BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0];
    if (win) sendEvent(win, 'app:command', { command, args: args || null });
  }

  function openPath(p) {
    if (!p || !fs.existsSync(p)) return;
    const which = appForFile(p);
    windows.open({ app: which, file: p });
    stores.recent.add({ path: p, app: which });
  }

  /** `--app=word` — how a pinned per-app taskbar entry relaunches into its app. */
  function argvApp(argv) {
    const flag = argv.find((a) => a.startsWith('--app='));
    return flag ? flag.slice('--app='.length) : null;
  }

  function argvFiles(argv) {
    return argv
      .slice(1)
      .filter((a) => !a.startsWith('-') && !a.startsWith(`${SCHEME}:`))
      .filter((a) => {
        try {
          return fs.statSync(a).isFile();
        } catch {
          return false;
        }
      });
  }

  // macOS delivers documents through this event, before and after ready.
  app.on('open-file', (event, p) => {
    event.preventDefault();
    if (windows) openPath(p);
    else pending.push(p);
  });

  app.on('second-instance', (_e, argv) => {
    const files = argvFiles(argv);
    if (files.length) files.forEach(openPath);
    else if (argvApp(argv)) windows.open({ app: argvApp(argv) });
    else {
      const win = BrowserWindow.getAllWindows()[0];
      if (win) {
        if (win.isMinimized()) win.restore();
        win.focus();
      } else windows.open({ app: 'home' });
    }
  });

  app.on('window-all-closed', () => {
    if (!isMac) app.quit();
  });

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) windows.open({ app: 'home' });
  });

  app.on('before-quit', () => {
    quitting.value = true;
  });

  // A check run keeps its windows off the desktop (windows.js). Chromium's
  // native occlusion tracker would count a window on no screen as hidden and
  // stop painting it; told not to, the page renders and takes input as if
  // it were in front.
  if (process.env.RUTBA_WINDOW_DISPLAY === 'offscreen') app.commandLine.appendSwitch('disable-features', 'CalculateNativeWinOcclusion');

  app.whenReady().then(async () => {
    stores = createStores();

    installProtocol({ rendererDir, allowFile: () => true });

    windows = createWindowManager({
      stores,
      preloadPath,
      iconPath,
      appIcons,
      appNames,
      appUserModelId,
      onWindowEvent: (win, event, payload) => sendEvent(win, event, payload),

      /**
       * Asked before a window holding unsaved work is allowed to close.
       *
       * The dialog is the platform's own, because this is the one moment where
       * looking like every other application matters more than looking like
       * ours: people answer this question by muscle memory. "Save" hands back
       * to the window, which knows how to save itself and closes when it has.
       */
      confirmClose: async (win, info) => {
        const { response } = await dialog.showMessageBox(win, {
          type: 'warning',
          message: `Save changes to ${info.name || 'this document'}?`,
          detail: 'If you don’t save, your changes will be lost.',
          buttons: ['Save', "Don't save", 'Cancel'],
          defaultId: 0,
          cancelId: 2,
          noLink: true,
        });
        if (response === 0) {
          sendEvent(win, 'app:command', { command: 'file.saveAndClose', args: null });
          return 'save';
        }
        if (response === 1) {
          windows.forceClose(win);
          return 'discard';
        }
        return 'cancel';
      },
    });

    const context = {
      stores,
      windows,
      openPath,
      fileUrl,
      holdBlob,
      releaseBlob,
      broadcast,
      sendEvent,
      send,
    };

    const base = buildImplementations({ stores, windows, quitting });
    const extra = namespaces ? await namespaces(context) : {};
    installIpc({ ...base, ...extra });

    Menu.setApplicationMenu(buildMenu({ send, appName }));

    nativeTheme.on('updated', () => broadcast('theme:changed', { dark: nativeTheme.shouldUseDarkColors }));

    await onReady?.(context);

    const startFiles = [...pending, ...argvFiles(process.argv)];
    if (startFiles.length) startFiles.forEach(openPath);
    else windows.open({ app: argvApp(process.argv) || 'home' });
  });

  return { app, get windows() { return windows; }, get stores() { return stores; }, openPath, fileUrl, holdBlob };
}

export { fileUrl, holdBlob, releaseBlob, sendEvent, broadcast, path };
