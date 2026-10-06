// The bridge, generated from the contract.
//
// Bundled to CommonJS by the desktop app's build so it can run in a sandboxed
// preload. Nothing here decides policy — it only forwards the calls the
// contract names, which is why adding a capability means editing one list.

import { contextBridge, ipcRenderer, webUtils } from 'electron';
import { METHODS, EVENTS, CHANNEL_PREFIX } from '../contract.js';

const api = { __contract: { methods: METHODS, events: EVENTS } };

for (const [ns, names] of Object.entries(METHODS)) {
  api[ns] = {};
  for (const name of names) {
    const channel = `${CHANNEL_PREFIX}/${ns}:${name}`;
    api[ns][name] = (payload) => ipcRenderer.invoke(channel, payload ?? {});
  }
}

// Events are one-way, backend to renderer. Handlers are stored per channel so a
// window that reloads does not leak the previous document's listeners.
api.on = (event, handler) => {
  if (!EVENTS.includes(event)) throw new Error(`unknown event: ${event}`);
  const channel = `${CHANNEL_PREFIX}/event/${event}`;
  const wrapped = (_e, payload) => handler(payload);
  ipcRenderer.on(channel, wrapped);
  return () => ipcRenderer.removeListener(channel, wrapped);
};

// Where a dropped file lives on disk. Electron 32 took `File.path` away and
// left this in its place, and it can only be asked from the preload: without
// it every drop arrived with no path, and dragging a file onto a window did
// nothing at all. A file that is not on disk answers null.
api.pathOf = (file) => {
  try {
    return webUtils.getPathForFile(file) || null;
  } catch {
    return null;
  }
};

// The window's own identity, handed over at creation time so the renderer knows
// which app it is before it has asked anything.
api.boot = () => {
  const q = new URLSearchParams(globalThis.location?.search || '');
  return {
    app: q.get('app') || 'home',
    file: q.get('file') || null,
    platform: process.platform,
    arch: process.arch,
    versions: { electron: process.versions.electron, chrome: process.versions.chrome, node: process.versions.node },
  };
};

contextBridge.exposeInMainWorld('rutbaOffice', api);
