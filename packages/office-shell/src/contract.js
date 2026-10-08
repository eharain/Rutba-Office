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
//   - A window is given only what a window uses. Stored passwords and sign-in
//     tokens stay in the main process: no window ever asked for them, and a
//     page that read a hostile document would otherwise be one call away from
//     every account's credentials. The same went for renaming and copying
//     files, which nothing in a window did.

/** Namespaced methods: renderer calls, backend answers. */
export const METHODS = {
  app: [
    'version',      // () -> { name, version, electron, chrome, node, platform, arch }
    'paths',        // () -> { home, documents, pictures, videos, downloads, userData, temp }
    'recent',       // () -> [{ path, name, app, at }]
    'addRecent',    // ({ path, app }) -> [recent]
    'removeRecent', // ({ path }) -> [recent]
    'renameRecent', // ({ path, name }) -> [recent]   renames the file on disk too
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
    'fullscreen',   // ({ on, display, presenter }) -> { fullscreen, display }   display: 'automatic' | 'primary' | a display's id
    'displays',     // () -> [{ id, name, primary, current, width, height }]   for Slide Show → Monitor
    'zoom',         // ({ delta, reset }) -> { factor }
    'state',        // () -> { maximized, fullscreen, focused, platform }
    'devtools',     // () -> void
    'arrange',      // ({ mode }) -> { count }   this app's windows: stack, columns, tile, cascade, sideBySide
    'list',         // () -> [{ id, name, current, hidden }]   this app's windows, for Switch Windows
    'focus',        // ({ id }) -> boolean
    'hide',         // () -> { hidden, reason }   refused for the last window on screen
    'unhide',       // ({ id }) -> boolean
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
    'temp',         // ({ ext, bytes }) -> { path }
    'exists',       // ({ path }) -> boolean
  ],
  // Thumbnails are served over rutba://thumb; a window that drew one itself
  // (a frame of a clip the system had no thumbnail for) hands it in here.
  thumbs: [
    'put',          // ({ path, bytes }) -> { stored }
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
    'new',          // ({ kind, template, design }) -> { id, kind, model, meta } — design: Set as Default's theme and styles, for a blank document
    'design',       // ({ id }) -> { theme, styles, name } — a document's theme and styles, for Design → Set as Default
    'open',         // ({ path, kind, password }) -> { id, kind, model, meta } | { locked, name, wrong } — a protected file asks for its password
    'setPassword',  // ({ id, password }) -> meta — File → Info → Encrypt with Password; '' goes back to a plain save
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
    'shapeClip',    // ({ id, slide, shape }) -> { xml, tag, rels } — a slide's shape as something to paste
    'deckFind',     // ({ id, query, options }) -> [hit] — every occurrence across a deck, for the find pane
    'deckEquation', // ({ id, slide, shape, linear, display }) -> { model, opResult } — Insert → Equation on a slide, or one edited again
    'deckDesign',   // ({ id, slide, kind }) -> { info, items } — Design's gallery: themes, variants, colours, fonts or effects, with live previews
    'designIdeas',  // ({ id, slide, width }) -> [{ id, name, svg }] — Design Ideas: the layouts the slide suits, each drawn
    'trace',        // ({ id, kind, row, col }) -> { kind, at, arrows, elsewhere } — a cell's precedents or dependents
    'evaluateFormula', // ({ id, row, col, actions }) -> { levels, canEvaluate, canStepIn, canStepOut, done, message } — Evaluate Formula after the presses
    'sessions',     // () -> [{ id, kind, path, dirty }]
    'recoverable',  // () -> [{ file, kind, name, path, at, size }] — what a crash left behind
    'recover',      // ({ file, password }) -> { id, kind, model, meta, recoveredFrom } | { locked, name, wrong }
    'discardRecovery', // ({ file }) -> { discarded }
    'adopt',        // ({ id }) -> { id, kind, model, meta } — a window takes over a session made for it (a merge's Letters1)
    'photoAlbum',   // ({ files, perSlide, captions, title, subtitle }) -> { id, name, slides } — Insert → Photo Album, a new deck for a window to adopt
    'compare',      // ({ original, revised, author }) -> { id, name, changes } — Review → Compare, a new document with the changes marked, for a window to adopt
    'objectFile',   // ({ id, part }) -> { path } — Insert → Object, double-clicked: the embedded document written to a file for its app to open
    'mailMerge',    // ({ id, action, ... }) -> Mailings: records, attach, sheets, attachContacts, createList, errors, finish, messages
    'proof',        // ({ id, action, ... }) -> Review: accessibility, spellStart, spellNext, spellCheckWord, ignoreAll, addWord, dictionary, options
  ],
  mail: [
    'accounts',     // () -> [account]
    'addAccount',   // ({ account, password }) -> account
    'updateAccount',// ({ id, patch }) -> account
    'removeAccount',// ({ id }) -> void
    'testAccount',  // ({ account, password }) -> { imap, smtp, error }
    'autodiscover', // ({ email }) -> { imap, smtp, source }
    'providers',    // () -> the providers worth a tile in Add account, as plain data
    'provider',     // ({ id }) -> one provider's full table entry, or null
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
    'deliverTest',  // ({ accountId, folder, raw }) -> { added, autoReplied, junked } — a fake message arriving; refused outside a check run
    'junk',         // ({ accountId? }) -> { level, safe, blocked, trustContacts, safeRecipients, blockedTlds, blockedEncodings, encodings, learned, ready, minimum, accountId, own }
    'setJunk',      // ({ accountId?, patch }) -> the same, changed — the account's own from then on
    'listSender',   // ({ accountId?, address, list: 'safe'|'blocked' }) -> the same
    'listRecipient', // ({ accountId, folder, ids } | { accountId, address }) -> the same, with { added }  Never block this group or mailing list
    'notJunk',      // ({ accountId, folder, ids }) -> { moved }  back to the Inbox, learned as good
    'learnJunk',    // ({ accountId }) -> the same, with what was learned now
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
  // Calendar and contacts accounts: CalDAV and CardDAV servers kept in step.
  dav: [
    'accounts',       // () -> [{ id, url, user, name, lastSync, lastError, calendars, books }]
    'add',            // ({ url, user, password, name }) -> account, its calendars and cards brought in
    'remove',         // ({ id }) -> { removed }   its calendars and cards with it
    'sync',           // ({ id }) -> [{ id, sent, received, removed, conflicts } | { id, error }]
    'books',          // () -> { books, defaultBook }
    'setDefaultBook', // ({ id }) -> id   where a new card goes; null for this computer only
  ],
  update: [
    'state',        // () -> { state, version, available, percent, automatic }
    'check',        // ({ manual }) -> state
    'install',      // () -> { installed }   quits, installs, returns
    'setAutomatic', // ({ on }) -> state
    'snooze',       // ({ version }) -> state   "not now": that version, for a day
    'seen',         // () -> state   the arrival of this version has been seen
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
    'write',        // ({ text, html }) -> void   the words, and a table for a spreadsheet to read by cell
    'read',         // () -> { text, html }
  ],
  // Insert → Screenshot: the windows and screens there are, and one of them as a picture.
  capture: [
    'sources',      // () -> [{ id, name, kind, thumbnail }]   every window but the asking one, then each screen
    'grab',         // ({ id }) -> { bytes, width, height, name }   that window or screen as a PNG
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
  'dav:changed',      // {}                  an account synced, was added or was removed
  'mail:sent',        // { id, to }
  'mail:sendFailed',  // { id, message, attempts, gaveUp }         // { accountId, folder, count }
  'theme:changed',    // { dark }
  'update:state',     // { state, version, available, percent, automatic, snoozed, snoozedVersion, arrived }
];

/** Every method as a flat channel list, e.g. 'fs:read'. */
export function channels() {
  const out = [];
  for (const [ns, names] of Object.entries(METHODS)) for (const n of names) out.push(`${ns}:${n}`);
  return out;
}

export const CHANNEL_PREFIX = 'rutba-office';
