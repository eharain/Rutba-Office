// The platform contract.
//
// Every operating-system capability Rutba Office needs is named here once, as
// data. A backend (Electron today, Tauri tomorrow) implements this list; the
// preload bridge and the renderer client are both generated from it, so a
// capability cannot exist on one side and be missing on the other.
//
// Rules that keep the seam honest:
//   - Arguments and results cross as structured-clone-safe values only.
//   - No method takes a callback; events travel on the separate event list.
//   - A method that touches the file system takes an absolute path, never a
//     path relative to some ambient working directory the renderer cannot see.

/** Namespaced methods: renderer calls, backend answers. */
export const METHODS = {
  app: [
    'version',      // () -> { name, version, electron, chrome, node, platform, arch }
    'paths',        // () -> { home, documents, pictures, videos, downloads, userData, temp }
    'recent',       // () -> [{ path, name, app, at }]
    'addRecent',    // ({ path, app }) -> [recent]
    'clearRecent',  // () -> []
    'quit',         // () -> void
    'relaunch',     // () -> void
  ],
  win: [
    'create',       // ({ app, file, query }) -> { id }
    'close',        // ({ force }) -> void   force answers the save prompt
    'minimize',     // () -> void
    'toggleMaximize',// () -> { maximized }
    'isMaximized',  // () -> boolean
    'setTitle',     // ({ title }) -> void
    'setDocumentEdited', // ({ edited }) -> void
    'setDirty',     // ({ dirty, name }) -> void   guards the close
    'fullscreen',   // ({ on }) -> { fullscreen }
    'zoom',         // ({ delta, reset }) -> { factor }
    'state',        // () -> { maximized, fullscreen, focused, platform }
    'devtools',     // () -> void
  ],
  fs: [
    'read',         // ({ path }) -> { bytes: Uint8Array, stat }
    'readHead',     // ({ path, bytes }) -> { bytes: Uint8Array, stat } — the first N bytes only, for a header

    'readText',     // ({ path, encoding }) -> { text, stat }
    'write',        // ({ path, bytes }) -> { stat }
    'writeText',    // ({ path, text }) -> { stat }
    'stat',         // ({ path }) -> stat | null
    'list',         // ({ path, filter }) -> [{ name, path, dir, size, mtime, ext }]
    'mkdirp',       // ({ path }) -> void
    'remove',       // ({ path }) -> void        (to the OS trash, never unlink)
    'rename',       // ({ from, to }) -> void
    'copy',         // ({ from, to }) -> void
    'temp',         // ({ ext, bytes }) -> { path }
    'exists',       // ({ path }) -> boolean
  ],
  dialog: [
    'open',         // ({ title, filters, multiple, directory, defaultPath }) -> [paths]
    'save',         // ({ title, filters, defaultPath }) -> path | null
    'message',      // ({ type, message, detail, buttons, defaultId, cancelId }) -> { response, checked }
    'error',        // ({ title, content }) -> void
  ],
  shell: [
    'openExternal', // ({ url }) -> void   (http/https/mailto only)
    'showInFolder', // ({ path }) -> void
    'openPath',     // ({ path }) -> void
    'beep',         // () -> void
  ],
  secrets: [
    'available',    // () -> boolean
    'get',          // ({ key }) -> string | null
    'set',          // ({ key, value }) -> void
    'delete',       // ({ key }) -> void
    'keys',         // () -> [key]
  ],
  store: [
    'get',          // ({ key, fallback }) -> value
    'set',          // ({ key, value }) -> void
    'delete',       // ({ key }) -> void
    'all',          // () -> object
  ],
  // Documents live in the backend, not in the renderer.
  //
  // The OOXML engine is a Node engine — it inflates with zlib and works in
  // Buffers — so a session opens, edits and saves in the main process and the
  // renderer draws whatever view model comes back. That also keeps a 200 MB
  // workbook out of the window's heap: the grid asks for the viewport it is
  // about to paint, and nothing else crosses.
  doc: [
    'new',          // ({ kind, template }) -> { id, kind, model, meta }
    'open',         // ({ path, kind }) -> { id, kind, model, meta }
    'close',        // ({ id }) -> void
    'meta',         // ({ id }) -> meta
    'thumbnails',   // ({ id, indexes }) -> { [index]: svg | null } — a deck's slide thumbnails the open model left out

    'model',        // ({ id, part }) -> model
    'apply',        // ({ id, ops }) -> { version, model?, patch? }
    'viewport',     // ({ id, sheet, top, left, rows, cols }) -> { cells, geometry }
    'undo',         // ({ id }) -> { version, model }
    'redo',         // ({ id }) -> { version, model }
    'save',         // ({ id, path }) -> { path, stat }
    'export',       // ({ id, format, path, options }) -> { path }
    'search',       // ({ id, query, options }) -> [hit]
    'asset',        // ({ id, ref }) -> { url, type, name }
    'pageSetup',    // ({ id, sheet }) -> the page setup the workbook carries
    'sessions',     // () -> [{ id, kind, path, dirty }]
    'recoverable',  // () -> [{ file, kind, name, path, at, size }] — what a crash left behind
    'recover',      // ({ file }) -> { id, kind, model, meta, recoveredFrom }
    'discardRecovery', // ({ file }) -> { discarded }
  ],
  mail: [
    'accounts',     // () -> [account]
    'addAccount',   // ({ account, password }) -> account
    'updateAccount',// ({ id, patch }) -> account
    'removeAccount',// ({ id }) -> void
    'testAccount',  // ({ account, password }) -> { imap, smtp, error }
    'autodiscover', // ({ email }) -> { imap, smtp, source }
    'importAccounts', // ({ path, only, test }) -> { added, skipped, failed, total } — a file of accounts, set up at once
    'attachmentText', // ({ accountId, folder, id, index }) -> { name, type, text }
    'folders',      // ({ accountId }) -> [folder]
    'sync',         // ({ accountId, folder, limit }) -> { added, total }
    'messages',     // ({ accountId, folder, offset, limit, query }) -> { rows, total }
    'message',      // ({ accountId, uid, folder }) -> message
    'flag',         // ({ accountId, folder, uids, add, remove }) -> void
    'move',         // ({ accountId, folder, uids, to }) -> void
    'delete',       // ({ accountId, folder, uids }) -> void
    'send',         // ({ accountId, draft }) -> { messageId }
    'saveDraft',    // ({ accountId, draft }) -> { id }
    'import',       // ({ path, accountId }) -> { folders, messages }
    'importScan',   // ({ path }) -> { kind, folders, messages }
    'export',       // ({ accountId, folder, path, format }) -> { messages }
    'attachment',   // ({ accountId, folder, id, index }) -> { url, name, type, size }
    'search',       // ({ accountId, query, limit }) -> [header]
    'unified',      // ({ role, query }) -> one list across every account
    'insight',      // ({ accountId, folder, id }) -> { trackers, unsubscribe, auth, bulk }
    'files',        // ({ accountId, query }) -> every attachment in the mailbox
    'people',       // ({ accountId }) -> the addresses seen, ranked
    'markAllRead',  // ({ accountId, folder }) -> { changed }
    'emptyFolder',  // ({ accountId, folder }) -> { removed }
    'queue',        // ({ accountId, draft, at, holdSeconds }) -> outbox item
    'outbox',       // () -> items still waiting to go
    'unsend',       // ({ id }) -> the item, taken back
    'rules',        // () -> [rule]
    'saveRule',     // ({ rule }) -> rule
    'deleteRule',   // ({ id }) -> { removed }
    'testRules',    // ({ accountId, folder }) -> { matched, of, sample }  changes nothing
    'runRules',     // ({ accountId, folder }) -> { matched, moved, starred, read, deleted }
  ],
  calendar: [
    'calendars',      // () -> [{ id, name, colour, visible, count }]
    'saveCalendar',   // ({ calendar }) -> [calendar]
    'removeCalendar', // ({ id }) -> { removed }
    'events',         // ({ from, to, all }) -> [occurrence]
    'get',            // ({ id }) -> event
    'save',           // ({ calendarId, event, scope, original }) -> event
    'remove',         // ({ id, scope, original }) -> { removed }
    'openFile',       // ({ path, from, to }) -> { name, method, count, events, invitation }
    'importFile',     // ({ path, calendarId }) -> { added, updated }
    'exportFile',     // ({ path, calendarId }) -> { path, count }
    'respond',        // ({ path | id, partstat }) -> { kept, organizer, subject, replyPath }
    'invitationFile', // ({ id }) -> { path, to, subject }
  ],
  contacts: [
    'list',       // ({ query }) -> [contact]
    'get',        // ({ id }) -> contact
    'save',       // ({ contact }) -> contact
    'remove',     // ({ id | ids }) -> { removed }
    'suggest',    // ({ query, limit }) -> [{ name, email, source }]
    'peek',       // ({ path }) -> { name, contacts }
    'importFile', // ({ path }) -> { added, updated, same, total }
    'importText', // ({ text }) -> { added, updated, same, total }
    'exportFile', // ({ path, ids }) -> { path, count }
    'fromMail',   // ({ name, email }) -> contact
    'count',      // () -> number
  ],
  update: [
    'state',        // () -> { state, version, available, percent, automatic }
    'check',        // ({ manual }) -> state
    'install',      // () -> { installed }   quits, installs, returns
    'setAutomatic', // ({ on }) -> state
  ],
  // Making this the application a file opens with. What is possible
  // differs by platform, and the service says which.
  defaults: [
    'status',       // () -> { platform, canSet, ours, total, formats, instructions }
    'set',          // ({ exts }) -> { changed, opened, message }
  ],
  // What other mail clients have left on this computer.
  // Where a slide show is up to, shared between the audience window and the
  // presenter window because two renderers cannot tell each other anything.
  present: [
    'state',        // () -> { id, index, running, startedAt, blank }
    'set',          // ({ index, running, blank, restart }) -> state
  ],
  // Signing in to Gmail and Outlook.com, which no longer accept a password.
  oauth: [
    'provider',     // ({ email }) -> { id, label, configured, imap, smtp } | null
    'signIn',       // ({ email, provider }) -> { email, provider, imap, smtp }
    'signOut',      // ({ email, provider }) -> { removed, note }
    'setClientId',  // ({ provider, clientId }) -> { configured }
    'clientIds',    // () -> { google, microsoft }
  ],
  // The one request this application makes of the outside world.
  announce: [
    'check',        // ({ force }) -> { announcement, enabled, checkedAt }
    'dismiss',      // ({ id }) -> { seen }
    'status',       // () -> { enabled, endpoint, sends, sendsNot }
    'setEnabled',   // ({ on }) -> { enabled }
  ],
  discover: [
    'scan',         // () -> { accounts, files, scanned }
  ],
  print: [
    'summary',      // ({ id, options }) -> { kind, pages, setup, sheets? } — before anything is drawn
    'printers',     // () -> [{ name, description, status, default }]
    'pdf',          // ({ id, path, options }) -> { path } — a workbook or deck laid out on pages
    'document',     // ({ id, options, printer, copies, silent }) -> { ok, reason }
    'toPDF',        // ({ landscape, margins, pageSize }) -> { bytes }
    'print',        // ({ silent }) -> void
  ],
  clipboard: [
    'writeText',    // ({ text }) -> void
    'readText',     // () -> string
  ],
};

/** Backend-to-renderer events. */
export const EVENTS = [
  'win:state',        // { maximized, fullscreen, focused }
  'app:open-file',    // { path }            OS asked us to open a document
  'app:command',      // { command, args }   menu / accelerator
  'mail:progress',    // { accountId, folder, done, total, phase }
  'mail:new',         // { accountId, folder, count }
  'present:state',    // { id, index, running, startedAt, blank }
  'mail:oauth',       // { phase, provider, email }
  'announce:new',     // { id, title, body, link, kind }
  'contacts:changed', // { count }           the address book was written
  'calendar:changed', // {}                  a calendar or an event was written
  'mail:sent',        // { id, to }
  'mail:sendFailed',  // { id, message, attempts, gaveUp }         // { accountId, folder, count }
  'theme:changed',    // { dark }
  'update:state',     // { state, version, available, percent, automatic }
];

/** Every method as a flat channel list, e.g. 'fs:read'. */
export function channels() {
  const out = [];
  for (const [ns, names] of Object.entries(METHODS)) for (const n of names) out.push(`${ns}:${n}`);
  return out;
}

export const CHANNEL_PREFIX = 'rutba-office';
