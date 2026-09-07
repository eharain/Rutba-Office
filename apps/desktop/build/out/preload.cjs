// ../../packages/office-shell/src/electron/preload.js
var import_electron = require("electron");

// ../../packages/office-shell/src/contract.js
var METHODS = {
  app: [
    "version",
    // () -> { name, version, electron, chrome, node, platform, arch }
    "paths",
    // () -> { home, documents, pictures, videos, downloads, userData, temp }
    "recent",
    // () -> [{ path, name, app, at }]
    "addRecent",
    // ({ path, app }) -> [recent]
    "clearRecent",
    // () -> []
    "quit",
    // () -> void
    "relaunch"
    // () -> void
  ],
  win: [
    "create",
    // ({ app, file, query }) -> { id }
    "close",
    // () -> void
    "minimize",
    // () -> void
    "toggleMaximize",
    // () -> { maximized }
    "isMaximized",
    // () -> boolean
    "setTitle",
    // ({ title }) -> void
    "setDocumentEdited",
    // ({ edited }) -> void
    "fullscreen",
    // ({ on }) -> { fullscreen }
    "zoom",
    // ({ delta, reset }) -> { factor }
    "state",
    // () -> { maximized, fullscreen, focused, platform }
    "devtools"
    // () -> void
  ],
  fs: [
    "read",
    // ({ path }) -> { bytes: Uint8Array, stat }
    "readText",
    // ({ path, encoding }) -> { text, stat }
    "write",
    // ({ path, bytes }) -> { stat }
    "writeText",
    // ({ path, text }) -> { stat }
    "stat",
    // ({ path }) -> stat | null
    "list",
    // ({ path, filter }) -> [{ name, path, dir, size, mtime, ext }]
    "mkdirp",
    // ({ path }) -> void
    "remove",
    // ({ path }) -> void        (to the OS trash, never unlink)
    "rename",
    // ({ from, to }) -> void
    "copy",
    // ({ from, to }) -> void
    "temp",
    // ({ ext, bytes }) -> { path }
    "exists"
    // ({ path }) -> boolean
  ],
  dialog: [
    "open",
    // ({ title, filters, multiple, directory, defaultPath }) -> [paths]
    "save",
    // ({ title, filters, defaultPath }) -> path | null
    "message",
    // ({ type, message, detail, buttons, defaultId, cancelId }) -> { response, checked }
    "error"
    // ({ title, content }) -> void
  ],
  shell: [
    "openExternal",
    // ({ url }) -> void   (http/https/mailto only)
    "showInFolder",
    // ({ path }) -> void
    "openPath",
    // ({ path }) -> void
    "beep"
    // () -> void
  ],
  secrets: [
    "available",
    // () -> boolean
    "get",
    // ({ key }) -> string | null
    "set",
    // ({ key, value }) -> void
    "delete",
    // ({ key }) -> void
    "keys"
    // () -> [key]
  ],
  store: [
    "get",
    // ({ key, fallback }) -> value
    "set",
    // ({ key, value }) -> void
    "delete",
    // ({ key }) -> void
    "all"
    // () -> object
  ],
  // Documents live in the backend, not in the renderer.
  //
  // The OOXML engine is a Node engine — it inflates with zlib and works in
  // Buffers — so a session opens, edits and saves in the main process and the
  // renderer draws whatever view model comes back. That also keeps a 200 MB
  // workbook out of the window's heap: the grid asks for the viewport it is
  // about to paint, and nothing else crosses.
  doc: [
    "new",
    // ({ kind, template }) -> { id, kind, model, meta }
    "open",
    // ({ path, kind }) -> { id, kind, model, meta }
    "close",
    // ({ id }) -> void
    "meta",
    // ({ id }) -> meta
    "model",
    // ({ id, part }) -> model
    "apply",
    // ({ id, ops }) -> { version, model?, patch? }
    "viewport",
    // ({ id, sheet, top, left, rows, cols }) -> { cells, geometry }
    "undo",
    // ({ id }) -> { version, model }
    "redo",
    // ({ id }) -> { version, model }
    "save",
    // ({ id, path }) -> { path, stat }
    "export",
    // ({ id, format, path, options }) -> { path }
    "search",
    // ({ id, query, options }) -> [hit]
    "asset",
    // ({ id, ref }) -> { url, type, name }
    "sessions"
    // () -> [{ id, kind, path, dirty }]
  ],
  mail: [
    "accounts",
    // () -> [account]
    "addAccount",
    // ({ account, password }) -> account
    "updateAccount",
    // ({ id, patch }) -> account
    "removeAccount",
    // ({ id }) -> void
    "testAccount",
    // ({ account, password }) -> { imap, smtp, error }
    "autodiscover",
    // ({ email }) -> { imap, smtp, source }
    "folders",
    // ({ accountId }) -> [folder]
    "sync",
    // ({ accountId, folder, limit }) -> { added, total }
    "messages",
    // ({ accountId, folder, offset, limit, query }) -> { rows, total }
    "message",
    // ({ accountId, uid, folder }) -> message
    "flag",
    // ({ accountId, folder, uids, add, remove }) -> void
    "move",
    // ({ accountId, folder, uids, to }) -> void
    "delete",
    // ({ accountId, folder, uids }) -> void
    "send",
    // ({ accountId, draft }) -> { messageId }
    "saveDraft",
    // ({ accountId, draft }) -> { id }
    "import",
    // ({ path, accountId }) -> { folders, messages }
    "importScan",
    // ({ path }) -> { kind, folders, messages }
    "export",
    // ({ accountId, folder, path, format }) -> { messages }
    "attachment",
    // ({ accountId, folder, id, index }) -> { url, name, type, size }
    "search"
    // ({ accountId, query, limit }) -> [header]
  ],
  print: [
    "toPDF",
    // ({ landscape, margins, pageSize }) -> { bytes }
    "print"
    // ({ silent }) -> void
  ],
  clipboard: [
    "writeText",
    // ({ text }) -> void
    "readText"
    // () -> string
  ]
};
var EVENTS = [
  "win:state",
  // { maximized, fullscreen, focused }
  "app:open-file",
  // { path }            OS asked us to open a document
  "app:command",
  // { command, args }   menu / accelerator
  "mail:progress",
  // { accountId, folder, done, total, phase }
  "mail:new",
  // { accountId, folder, count }
  "theme:changed"
  // { dark }
];
var CHANNEL_PREFIX = "rutba-office";

// ../../packages/office-shell/src/electron/preload.js
var api = { __contract: { methods: METHODS, events: EVENTS } };
for (const [ns, names] of Object.entries(METHODS)) {
  api[ns] = {};
  for (const name of names) {
    const channel = `${CHANNEL_PREFIX}/${ns}:${name}`;
    api[ns][name] = (payload) => import_electron.ipcRenderer.invoke(channel, payload ?? {});
  }
}
api.on = (event, handler) => {
  if (!EVENTS.includes(event)) throw new Error(`unknown event: ${event}`);
  const channel = `${CHANNEL_PREFIX}/event/${event}`;
  const wrapped = (_e, payload) => handler(payload);
  import_electron.ipcRenderer.on(channel, wrapped);
  return () => import_electron.ipcRenderer.removeListener(channel, wrapped);
};
api.boot = () => {
  const q = new URLSearchParams(globalThis.location?.search || "");
  return {
    app: q.get("app") || "home",
    file: q.get("file") || null,
    platform: process.platform,
    arch: process.arch,
    versions: { electron: process.versions.electron, chrome: process.versions.chrome, node: process.versions.node }
  };
};
import_electron.contextBridge.exposeInMainWorld("rutbaOffice", api);
