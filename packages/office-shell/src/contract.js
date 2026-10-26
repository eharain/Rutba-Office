// The platform contract.
//
// Every operating-system capability Rutba Office needs is named here once, as
// data. A backend (Electron today, Tauri tomorrow) implements this list; the
// preload bridge and the renderer client are both generated from it, so a
// capability cannot exist on one side and be missing on the other.
//
// The list is divided by owner: the core's namespaces, and each app pack's
// (NAMESPACE_PACK below). One run offers the core's and those of the packs it
// loaded, no others: the main process registers only those channels, and the
// preload asks for that list once, as a window starts, rather than being
// built with the whole of it.
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
    'frameAction',  // ({ action: 'copy' }) -> { done }   in a frame with no scripts of its own (Mail's message): the words chosen there, to the clipboard
    'state',        // () -> { maximized, fullscreen, focused, platform }
    'devtools',     // () -> void
    'arrange',      // ({ mode }) -> { count }   this app's windows: stack, columns, tile, cascade, sideBySide
    'list',         // () -> [{ id, name, current, hidden }]   this app's windows, for Switch Windows
    'focus',        // ({ id }) -> boolean
    'hide',         // () -> { hidden, reason }   refused for the last window on screen
    'unhide',       // ({ id }) -> boolean
    'handBack',     // ({ token, bytes, contentType, width, height }) -> { delivered }   Edit in Image: the edited picture to the window that opened this one
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
    'append',       // ({ path, bytes }) -> { size }: only to a file `temp` or capture.take made, in this run
    'patch',        // ({ path, offset, bytes }) -> { size }: bytes written over what is at offset, in a file `temp` made in this run (an MP4's mdat size, once its samples are in)
    'keepTemp',     // ({ path, to }) -> { stat }: a finished file `temp` made, moved to where it was asked to go (replacing what is there), its scratch folder gone
    'dropTemp',     // ({ path }) -> void — a file `temp` made, and its folder, gone
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
    'gallery',      // ({ app }) -> { categories, templates: [{ id, app, name, description, categories, thumbnail }] }: File > New's templates; doc.new({ template: id }) makes an untitled copy of one
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
    'save',         // ({ id, path, conformance? }) -> { path, stat, conformance }: 'strict' or 'transitional' chooses the form, the file's own when left out
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
    'keepCopy',     // ({ key, kind, name, path, bytes }) -> entry: a window that holds its own document (Video's project) leaves a recovery copy
    'dropCopy',     // ({ key }) -> { dropped }: that copy, saved or let go of
    'readCopy',     // ({ file }) -> { bytes, name, path, at, kind }: a recovered copy, for the window that opens it
    'adopt',        // ({ id }) -> { id, kind, model, meta } — a window takes over a session made for it (a merge's Letters1)
    'photoAlbum',   // ({ files, perSlide, captions, title, subtitle }) -> { id, name, slides } — Insert → Photo Album, a new deck for a window to adopt
    'compare',      // ({ original, revised, author }) -> { id, name, changes } — Review → Compare, a new document with the changes marked, for a window to adopt
    'objectFile',   // ({ id, part }) -> { path } — Insert → Object, double-clicked: the embedded document written to a file for its app to open
    'redact',       // ({ id, action: 'scan' | 'apply', marks, pictures, path }) -> { total, parts, words } | { path, total, parts, checked }, Redact: what the marks find, or the redacted copy saved and read back
    'inspect',      // ({ id, action: 'check' | 'remove', ids }) -> { kind, items, removed? }, File → Info → Check for Issues → Inspect
    'deckReview',   // ({ id, action: 'compare' | 'accept' | 'reject' | 'end', path, ids }) -> { reviewer, name, changes } | null, Presentations: Review → Compare and its changes
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
    'setupAccount', // ({ email, password, name, confirmed }) -> { outcome, account, provider, server, foreign, tried } the servers found and tried until one signs in; 'confirm' names servers outside the domain to agree to first; nothing is kept
    'addAccounts',  // ({ text, job }) -> { job, results: [{ line, email, outcome, provider, server, id, foreign }] } rows of addresses and passwords set up at once; progress as mail:accountsProgress
    'confirmAccounts', // ({ job, index, hosts }) -> result: a row whose domain named servers elsewhere, set up once the person agreed to those hosts
    'forgetAccountsJob', // ({ job }) -> { forgotten } rows still waiting for agreement forgotten, passwords and all
    'discoveryForTest', // ({ domain, found }) -> { staged } a domain whose records name servers elsewhere, for a check run; refused otherwise
    'attachmentText', // ({ accountId, folder, id, index }) -> { name, type, text }
    'source',       // ({ accountId, folder, id }) -> { text, bytes, from: 'server' | 'rebuilt', reason? }   the message as it came, fetched from its server when shown; rebuilt from what is kept when there is no server
    'fetchPicture', // ({ url }) -> { bytes, type, name }   a picture a message shows from the web, fetched when it is saved (http and https only)
    'attachmentReading', // ({ accountId, folder, id, index }) -> reading | null: what was read from an attachment (an invoice's details) when its message was opened
    'keepReading', // ({ accountId, folder, id, index, reading }) -> { kept: 'message' | 'run' }: kept with the message, so it is not read again
    'folders',      // ({ accountId }) -> [folder]
    'sync',         // ({ accountId, folder, limit }) -> { added, total }   the changes made here sent, then what is new fetched by UID, with marks and deletions made at the server elsewhere
    'messages',     // ({ accountId, folder, offset, limit, query }) -> { rows, total }
    'message',      // ({ accountId, uid, folder }) -> message
    'flag',         // ({ accountId, folder, ids, patch }) -> counts   read, unread, flagged here and at the server; a pin here only
    'move',         // ({ accountId, folder, ids, to }) -> { moved }   here and at the server
    'delete',       // ({ accountId, folder, ids, toTrash }) -> { deleted, trash? }   for good (EXPUNGE at the server), or to the Trash with toTrash
    'flushChanges', // ({ accountId }) -> { left, error? }   the changes made here sent to the server now
    'syncState',    // ({ accountId }) -> { waiting, offline, error, at }   where the account stands with its server
    'fetchSettings',    // () -> { every, push }   how Mail fetches by itself: every so many minutes (0 only with Get mail), and as the server says mail came
    'setFetchSettings', // ({ every, push }) -> the same, changed
    'fetchNow',     // () -> { accounts }   every account fetched now, quietly, as the Mail window opens
    'serverAutoReply', // ({ accountId, action: 'state' | 'on' | 'off', settings }) -> { supported, on, settings?, reason? }   the mail server's own out of office, over ManageSieve
    'send',         // ({ accountId, draft }) -> { messageId }
    'saveDraft',    // ({ accountId, draft }) -> { id }
    'import',       // ({ path, accountId }) -> { folders, messages }
    'importScan',   // ({ path }) -> { kind, folders, messages }
    'export',       // ({ accountId, folder, path, format }) -> { messages }
    'attachment',   // ({ accountId, folder, id, index }) -> { url, name, type, size }
    'search',       // ({ accountId, query, limit }) -> [header]
    'unified',      // ({ role, query }) -> one list across every account
    'insight',      // ({ accountId, folder, id }) -> { trackers, unsubscribe, auth, bulk }
    'unsubscribe',  // ({ accountId, folder, id }) -> { done, status?, reason? }  the one RFC 8058 one-click POST
    'files',        // ({ accountId, query }) -> every attachment in the mailbox
    'people',       // ({ accountId }) -> the addresses seen, ranked
    'markAllRead',  // ({ accountId, folder }) -> { changed }
    'emptyFolder',  // ({ accountId, folder }) -> { removed }
    'createFolder', // ({ accountId, parent, name }) -> { path }  made at the server too
    'renameFolder', // ({ accountId, folder, name }) -> { path }  with the folders inside it
    'deleteFolder', // ({ accountId, folder }) -> { folders, removed }  with what is inside it
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
    'archives',     // () -> [{ id, path, name, kind, label, bytes, folders, messages, missing, index: { state, done, total } }] archives opened in place
    'openArchive',  // ({ path }) -> that archive: a .pst, .ost, .olm or mbox read where it lies, never written to; its id begins archive:
    'closeArchive', // ({ id }) -> { closed }  forgotten and its part of the index dropped; the file is untouched
    'stopIndexing', // ({ id }) -> { state, done, total }  an archive's indexing paused; what is indexed stays searchable
    'resumeIndexing', // ({ id }) -> { state, done, total }  carried on where it was paused
    'searchAll',    // ({ query, limit, offset, sources }) -> { rows, total, exact, indexing: [{ id, name, state, done, total }] } one search over every account and archive, from what is indexed so far
    'exportHold',   // ({ job, query, sources, folders, items, dir, mbox, pdf, attachments, labels, selection }) -> { dir, exported, requested, cancelled, sources, attachments, pdfs, files }
    'cancelExport', // ({ job }) -> { cancelled }
    'smimeIdentities',  // () -> [{ name, email, emails, issuer, validFrom, validTo, thumbprint, signs, encrypts }] one's own certificates (their keys kept in the system's secret store)
    'importIdentity',   // ({ path, password }) -> that identity: a .p12/.pfx opened with its password
    'removeIdentity',   // ({ thumbprint }) -> { removed }
    'importCertificate', // ({ path }) -> [certificate] someone else's, from a .cer/.crt/.pem
    'smimeCertificates', // () -> { [email]: [certificate] } other people's, kept from their signed mail or imported
    'forgetCertificates', // ({ emails }) -> { forgotten }
    'pendingCertificates', // () -> { [email]: certificate } a different certificate arrived for an address that has one, held aside
    'acceptCertificate', // ({ email }) -> { accepted } use the one held aside instead
    'dismissCertificate', // ({ email }) -> { dismissed } keep the one kept, forget the one held aside
    'canEncryptFor',    // ({ addresses }) -> { ok, missing } is there a certificate for each
    'trustForTest',     // ({ pem }) -> { anchors } a test authority trusted for a check run; refused otherwise
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
    'peekCsv',    // ({ path, columns }) -> { name, count, columns, skipped, unread, sample }  a CSV checked before it comes in
    'importFile', // ({ path, columns }) -> { added, updated, same, total, skipped, unread }
    'importText', // ({ text }) -> { added, updated, same, total }
    'exportFile', // ({ path, ids }) -> { path, count }
    'fromMail',   // ({ name, email }) -> contact
    'count',      // () -> number
  ],
  // Calendar and contacts accounts: CalDAV and CardDAV servers kept in step.
  dav: [
    'accounts',       // () -> [{ id, url, user, name, lastSync, lastError, calendars, books, needsPassword }]
    'add',            // ({ url, user, password, name }) -> account, its calendars and cards brought in; named after its server when no name is given
    'rename',         // ({ id, name }) -> account
    'remove',         // ({ id }) -> { removed }   its calendars and cards with it
    'sync',           // ({ id }) -> [{ id, sent, received, removed, conflicts } | { id, error }]
    'books',          // () -> { books, defaultBook }
    'setDefaultBook', // ({ id }) -> id   where a new card goes; null for this computer only (until chosen, the first book of an account added from 2.19 on; a store from before keeps this computer)
    'setPassword',    // ({ id, password }) -> account   its password typed again (not saved on this computer any more), and a sync
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
    'follow',       // ({ apps }) -> { skipped, withdrawn, restored }   Windows told only about the chosen apps' types (an installed copy)
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
    'browsers',     // () -> [{ id, name, isDefault }] the browsers on this computer a sign-in can open in
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
  // The PDF app: the files a window has open, read once, and new PDFs made
  // from their pages without drawing anything (apps/desktop/main/pdf.js).
  pdf: [
    'open',         // ({ path } | { bytes, name }) -> { id, name, path, size, version, encrypted, locked, access, permissions, problem, repaired, title, pages: [{ width, height, rotate }] }
    'unlock',       // ({ id, password }) -> open's answer, or with { wrong: true }: a protected file read with its password (the user's or the owner's), held in memory with it
    'close',        // ({ id }) -> { closed }
    'assemble',     // ({ pages: [{ source, page, rotate }], path, title }) -> { path, size } | { bytes, size }: written whole, or handed back
    'text',         // ({ id, page }) -> string: the words a page's content draws
    'addText',      // ({ id, pages: [{ page, lines: [{ words: [{ text, x, y, width, size, angle }] }] }] }) -> { id, ...open's answer, written, bytes }: an invisible text layer as an incremental update, a new source
    'inspect',      // ({ path, password }) -> { pages: [{ width, height, rotate, text }], encrypted, locked, access, title }: a file read and let go, for the checks
    'annotations',  // ({ id }) -> [{ id, page, type, rect, contents, author, modified, color, quads, ink, line, arrow, replyTo, status, ... }]: the comments on every page
    'annotate',     // ({ id, changes: { add, reply, resolve, edit, remove }, author }) -> { id, ...open's answer, added, annotations, fields, bytes }: comments changed as an incremental update, a new source
    'fields',       // ({ id }) -> [{ id, name, type, value, options, multiline, readOnly, widgets: [{ id, page, rect, onState, state }], ... }]: the form's fields
    'fill',         // ({ id, values: { [fieldId]: value } }) -> { id, ...open's answer, changed, annotations, fields, bytes }: form values with appearances, as an incremental update, a new source
    'addCheckboxes', // ({ id, boxes: [{ page, rect, name, checked }] }) -> { id, ...open's answer, added, annotations, fields, bytes }: check box fields where a scanned form's tick boxes are, as an incremental update, a new source
    'navigation',   // ({ id }) -> { outline: [{ title, page, url, bold, italic, open, kids }], labels: [string] | null, links: [{ page, rect, target, url }] }
    'compressPlan', // ({ pages, dpi }) -> { token, size, plan: [{ ref, width, height, toWidth, toHeight, filter, comps, data }] }: pictures shown at more than dpi
    'compress',     // ({ token, dpi, images: { [ref]: { data, width, height, filter, comps } }, path }) -> { path, before, after, replaced }: written whole as a new file
    'redactPlan',   // ({ sources }) -> [{ source, page, key, width, height, filter, rects, data }]: pictures under redaction marks only the window can decode
    'redact',       // ({ pages: [{ source, page, rotate }], path, images, title }) -> { path, size, removed: { glyphs, pictures, images, annotations, pages }, elsewhere: { [kind]: count }, words, unpainted: [{ page, count }] }: the marks applied, the marked words taken out of everything else, written whole as a new file
    'redactPreview', // ({ pages, images, title }) -> { removed, elsewhere, words, unpainted, size }: what applying would do, nothing written
  ],
  // Text recognition (apps/desktop/main/ocr.js, @rutba/ocr): read on a pool of
  // threads, never this process's or the window's; nothing leaves the computer.
  ocr: [
    'recognize',    // ({ image: { width, height, data } | pngBytes, languages, dpi, mode, table, document, job }) -> { text, lines, words: [{ text, box, confidence, line }], width, height, languages, angle, ms, table?: { rows, columns, cells, confidence, ruled }, document?: { kind, fields, items, checks, currency, checkboxes, words } }
    'cancel',       // ({ job }) -> { cancelled }   a job this window named, waiting or reading
    'languages',    // () -> [{ code, name, script, direction, beta? }]
    'stats',        // () -> { threads, busy, queued, limit, done, failed, cancelled }   for the checks
  ],
  // Recorder → Go Live: the destinations a person streams to (their keys kept in the
  // keystore and never handed back), and the program's pieces sent to every one.
  // Databases: a database is one SQLite file, held here while a window has it open (apps/desktop/main/base.js).
  base: [
    'create',       // ({ file }) -> summary   a new database at file
    'createFromTemplate', // ({ file, template }) -> summary   a new database from a ready-made one
    'readyMade',    // () -> { set, office, templates: [{ id, needs }], samples: [{ id, needs }] }   which need a newer Office
    'createFromSample', // ({ file, sample, withData }) -> summary   a new database from a sample business's, with its invented records or empty
    'open',         // ({ file }) -> summary   a .rbase, or any SQLite database
    'close',        // ({ id }) -> { closed }
    'summary',      // ({ id }) -> { id, file, name, tables: [{ name, caption, count }], relationships, steps }
    'fields',       // ({ id, table }) -> [field]
    'rows',         // ({ id, table, offset?, limit?, order?, where?, search?, filters? }) -> { rows, count }   a page of records, searched and filtered (database/src/filter.js)
    'designUpdate', // ({ id }) -> { available, tables, fields, added, replaced, beside }   a newer design from the template set in use
    'applyDesignUpdate', // ({ id }) -> { done, ..., summary }   that design brought in, records kept, the person's changes kept
    'pin',          // ({ id, kind: 'table' | 'query' | 'form' | 'report' | 'document', name, pinned }) -> summary   pinned first in the side panel, kept in the file
    'fold',         // ({ id, section, folded }) -> summary   a side panel section folded away, kept in the file
    'signIn',       // ({ id, name, password }) -> summary   a protected database opened as that person's role allows (D-2026-10-22-11)
    'signOut',      // ({ id }) -> summary   locked again until someone signs in
    'protect',      // ({ id, name, password }) -> summary   protected from now on, that person its first Admin
    'unprotect',    // ({ id, password }) -> summary   the protection taken away by an Admin; the people go, the records stay
    'seeAs',        // ({ id, role | null }) -> summary   an Admin sees the database as a role would, reading only
    'roles',        // ({ id }) -> [{ name, rights, builtIn }]
    'saveRole',     // ({ id, name, rights }) -> { name }
    'removeRole',   // ({ id, name }) -> { removed }
    'people',       // ({ id }) -> [{ id, name, role, created }]
    'addPerson',    // ({ id, name, role, password }) -> { id }
    'changePerson', // ({ id, person, name?, role? }) -> { changed }
    'removePerson', // ({ id, person }) -> { removed }
    'setPassword',  // ({ id, person, password, current? }) -> { changed }   your own with the one before it; anyone's, for an Admin
    'trailOf',      // ({ id, table, rowid }) -> [{ at, person, action, fields }]   who added and changed a record
    'allFields',    // ({ id }) -> { [table]: [field name] }   every field, for the People and roles grid
    'engines',      // () -> [{ id, name, port, ready }]   the database engines a table can be linked to (D-2026-10-22-13)
    'engineTables', // ({ engine, settings, password }) -> { tables: [{ schema, name, kind }] }   an engine tried; nothing kept
    'saveLink',     // ({ id, link?, name, engine, settings, password? }) -> { link, summary }   its password in this computer's protected store
    'removeLink',   // ({ id, link }) -> summary
    'linkTables',   // ({ id, link, remotes: [{ schema, name, as? }] }) -> { tables, summary }   linked live, read at once
    'unlinkTable',  // ({ id, table }) -> summary   the engine keeps its records
    'refreshLinked',// ({ id, table?, now? }) -> summary   read again from the engine, a copy synced, a source read when due (or now)
    'setLinkMode',  // ({ id, table, mode: 'live' | 'copy' }) -> summary   read live or kept as a copy in the file, its records kept
    'setLinkStamp', // ({ id, table, field | null }) -> summary   a copy's updated-at field, to read only what changed
    'clashes',      // ({ id, table }) -> [{ id, key, at, action, mine, theirs, why }]   versions from here the engine's replaced
    'settleClash',  // ({ id, table, clash, keep: 'engine' | 'mine' }) -> summary
    'sourceItems',  // ({ where, person?, password? }) -> { protected, items: [{ kind, name }] }   another database's tables and queries
    'trySource',    // ({ kind, settings, password?, headers?, item? }) -> { fields, rows, count, sheets?, items? }   nothing kept
    'linkSource',   // ({ id, kind, name, settings, password?, headers?, item? }) -> { table, summary }   a file, another database or JSON, kept as a copy
    'setSourceEvery', // ({ id, table, every: 5 | 15 | 60 | 0 }) -> summary   how often a source is read again
    'sourceSignIn', // ({ id, table, person, password }) -> summary   who you are in a protected database a source reads
    'refreshDue',   // ({ id }) -> summary | null   sources due on their schedules read again
    'sourceKeys',   // ({ id, table, headers }) -> summary   a web address's keys or tokens typed again, kept for that address only
    'allowSources', // ({ id, addresses }) -> summary   the sources and engines asked about allowed for this database on this computer, then read
    'queryRows',    // ({ id, name? | design?, offset?, limit?, order?, search?, filters? }) -> { columns: [{ name, type }], rows, count }   a page of a query's results, as rows reads a table's
    'insert',       // ({ id, table, values }) -> { rowid }
    'update',       // ({ id, table, rowid, values }) -> { changed }
    'remove',       // ({ id, table, rowids }) -> { removed }
    'createTable',  // ({ id, name, fields }) -> summary
    'addField',     // ({ id, table, field }) -> summary
    'changeField',  // ({ id, table, name, change }) -> summary
    'removeField',  // ({ id, table, name }) -> summary
    'renameTable',  // ({ id, table, to }) -> summary
    'dropTable',    // ({ id, table }) -> summary
    'addRelationship',    // ({ id, from, to, cascadeDelete? }) -> summary
    'removeRelationship', // ({ id, relation }) -> summary
    'steps',        // ({ id }) -> [{ n, at, label }]   the design's history
    'query',        // ({ id, sql, limit? }) -> { columns, rows, count }   SELECT only, read-only
    'importFile',   // ({ id, file }) -> { tables, notes, summary }   CSV, TSV and the old files, a table each
    'queries',      // ({ id }) -> [{ name, design }]   the queries kept in the file
    'saveQuery',    // ({ id, name, design }) -> summary   a designed query kept (made or replaced)
    'removeQuery',  // ({ id, name }) -> summary
    'runQuery',     // ({ id, name?, design?, limit? }) -> { columns, rows, count, sql }
    'sqlOf',        // ({ id, design }) -> string   the SQL a design stands for, its values in place
    'forms',        // ({ id }) -> [{ name, layout }]
    'saveForm',     // ({ id, layout }) -> summary
    'removeForm',   // ({ id, name }) -> summary
    'makeForm',     // ({ id, table, name? }) -> { layout, summary }   the form wizard's form, kept
    'reports',      // ({ id }) -> [{ name, layout }]
    'saveReport',   // ({ id, layout }) -> summary
    'removeReport', // ({ id, name }) -> summary
    'makeReport',   // ({ id, table? | query?, name? }) -> { layout, summary }   the report wizard's report, kept
    'reportData',   // ({ id, name? | layout? }) -> { layout, groups, totals, count }   grouped, with totals
    'documents',    // ({ id }) -> [{ name, layout }]   the print templates
    'saveDocument', // ({ id, layout }) -> summary
    'removeDocument', // ({ id, name }) -> summary
    'makeDocument', // ({ id, table, name? }) -> { layout, summary }   the wizard's print template, kept
    'documentData', // ({ id, name? | layout?, keys?, limit? }) -> { layout, letterhead, documents }   one a record
    'lookup',       // ({ id, table, field, show }) -> [{ value, text }]   a lookup list's choices
    'lookupPage',   // ({ id, table, field, search?, values?, limit? }) -> { items: [{ value, text }], more }   a long list a page at a time, searched, or given records' words
    'saveAs',       // ({ id, file }) -> summary   a copy at file, then the one open; another program's database made a Data Files copy (summary.copied)
    'allowChanges', // ({ id }) -> summary   another program's database opened so its records can be changed in place (A2.20)
  ],
  // Talk: the one connection to the introducer, the messages kept here, the relays (apps/desktop/main/talk.js).
  talk: [
    'state',        // () -> { phase, me, contacts, requests, blocked, relays, server, relay, presence, error }
    'signIn',       // ({ token, server? }) -> state   a token given outright, for a test or self-run introducer
    'signInBrowser',// () -> state   Rutba sign-in in the system browser, then Rutba's introducer
    'signOut',      // () -> state
    'resume',       // () -> state   connects again with the sign-in kept
    'send',         // (msg) -> { sent }   a contact request, presence, a call's set-up, to the introducer
    'message',      // ({ to, text }) -> { id, from, to, text, at, mine }   kept here and sent
    'history',      // ({ contactId, limit? }) -> [message]
    'relays',       // () -> { choice, relays, own, ice }
    'setRelay',     // ({ choice: 'auto' | id | 'own', own?: { urls, username, credential } }) -> { choice, ice }
    'callSetup',    // () -> { iceServers, offerRelay }   what a call is made with
    'invite',       // ({ email, note, how: 'own' | 'rutba' | 'none', accountId?, letter? }) -> { request, told, allowance } | { contact } | { error }   asked, and told by email
    'allowance',    // () -> { day: { used, limit }, week: { used, limit } } | { error }   what Rutba will still email for this person
    'inviteLink',   // ({ note }) -> { invite: { code, url, note, created, expires, uses, maxUses } } | { error }
    'revokeInvite', // ({ code }) -> { revoked } | { error }
    'peekInvite',   // ({ link }) -> { code, from, note, kind } | { error }   who an invitation is from, by its link or code
    'acceptInvite', // ({ link }) -> { contact } | { error }
    'removeContact', // ({ contactId }) -> { removed } | { error }
    'block',        // ({ contactId }) -> { blocked } | { error }
    'withdraw',     // ({ requestId }) -> { withdrawn } | { error }   an invitation asked, taken back
    'remind',       // ({ requestId, letter? }) -> { request, told } | { error, next }   sent again, the same way, once the introducer allows it
  ],
  live: [
    'destinations',     // () -> [{ id, platform, name, url, on, hasKey, keyHint }]
    'platforms',        // () -> { youtube: { name, url, keyHelp }, ... }
    'saveDestination',  // ({ id?, platform, name?, url?, on?, key? }) -> destinations   a key given replaces the one kept
    'removeDestination',// ({ id }) -> destinations
    'start',            // ({ video: { width, height, fps, kbps }, audio: { kbps, sampleRate, channels } }) -> { started: [id], skipped: [{ id, name, reason }] }
    'send',             // ({ kind: 'video' | 'audio', timestamp, data, key, config }) -> { live }   one encoded piece, to every live destination
    'status',           // () -> { live, destinations: [{ id, name, state, detail, bytes, dropped, since, reconnects }] }
    'stop',             // () -> status
  ],
  // Profiles (D-2026-10-26-02): which profile a window belongs to, the others, and managing them.
  // A password is set, changed and taken away from the profile's own windows only.
  profiles: [
    'current',      // () -> { id, name, colour, count, protected, isDefault }
    'list',         // () -> { current, openFirst, colours, profiles: [{ id, name, colour, protected, openFirst, current, running, isDefault }] }
    'open',         // ({ id, file }) -> { here } | { started, running }   that profile's Rutba Office, started or brought to the front
    'make',         // ({ name, colour }) -> list + { made }
    'rename',       // ({ id, name }) -> list
    'setColour',    // ({ id, colour }) -> list
    'setOpenFirst', // ({ id }) -> list
    'removal',      // ({ id }) -> { refused, size }   asked before Remove: why not, or how much goes
    'remove',       // ({ id }) -> list   its folder to the Recycle Bin
    'setPassword',  // ({ password }) -> current
    'changePassword', // ({ old, password }) -> current
    'removePassword', // ({ old }) -> current
    'keystoreNotice', // () -> { at, why } | null   saved passwords that could not be opened on this computer, given to one window once
  ],
  ai: [
    'state',        // () -> { mode: 'off' | 'session' | 'always', folders, connect: { command, args, env }, log, connections: [{ client, clientVersion, where, since, last, lastTool, pid }], reach: 'local' | 'remote', remote: { port, key, addresses, listening, error } | null }
    'setMode',      // ({ mode }) -> state
    'addFolder',    // ({ path }) -> state
    'removeFolder', // ({ path }) -> state
    'setReach',     // ({ reach: 'local' | 'remote', port? }) -> state   this computer only, or other machines too, with a key
    'renewKey',     // () -> state   a new key; the old one stops working at once
    // What each mail account, archive, calendar, address book, database and folder allows (D-2026-10-26-06):
    // one setting, kept with the thing, shown in its own app's settings and in AI Tools.
    'things',       // () -> { addon, things: [{ kind, id, name, step, steps, server?, local?, role?, protected? }] }
    'setStep',      // ({ kind, id, step, role? }) -> { kind, id, step, steps, role }   'off' | 'read' | 'prepare' | 'act'
    'stepOf',       // ({ kind, id }) -> { addon, step, steps, role }
    'pending',      // () -> [{ id, at, kind, thing, fields, body, draft, change }]   what the tools asked to do, waiting for the person
    'answer',       // ({ id, choice: 'act' | 'draft' | 'cancel' }) -> { choice, outcome | error }
  ],
  // Add or remove apps (docs/APP-PACKS.md 5.4): the packs on this computer, and changes to them for the next start.
  packs: [
    'list',         // () -> { apps: [{ id, name, state, size, reason }], restart, perMachine, source }   state: installed, adding, removing, absent, refused
    'add',          // ({ id }) -> list   its pack fetched into a staging folder, checked there, and put in place for the next start
    'remove',       // ({ id }) -> list   gone at the next start; the person's files and settings stay
    'keep',         // ({ id }) -> list   an addition or a removal not yet started, taken back
    'restart',      // () -> { restarting }   started again to finish; a per-machine copy asks for the administrator's permission first
  ],
  clipboard: [
    'writeText',    // ({ text }) -> void
    'readText',     // () -> string
    'write',        // ({ text, html }) -> void   the words, and a table for a spreadsheet to read by cell
    'read',         // () -> { text, html }
    'writeImage',   // ({ bytes }) -> { width, height }   a picture (PNG or JPEG bytes) on the clipboard, as Image's Copy puts it
    'readImage',    // () -> { bytes, width, height } | null   the clipboard's picture as PNG, for Image's Paste
  ],
  // Insert → Screenshot: the windows and screens there are, and one of them as a picture.
  // The recorder (Video → Record, Presentations → Screen Recording) records one of them.
  capture: [
    'sources',      // () -> [{ id, name, kind, thumbnail, display?, capture? }]   every window but the asking one, then each screen (display: its size in pixels; capture: what it is recorded from, a check run's stand-in screen)
    'grab',         // ({ id }) -> { bytes, width, height, name }   that window or screen as a PNG
    'choose',       // ({ id, audio }) -> { loopback }   what this window's next getDisplayMedia is given: that source, and the computer's own sound where the platform hears it
    'controls',     // ({ on, paused, minimise }) -> { stop, pause, tray }   while recording: a global Stop and Pause key and a tray icon, sent back as capture:command; off releases them
    'pictureKey',   // ({ on }) -> { key }   while Recorder is open: a global key for Take Picture, sent back as capture:command 'picture'
    'space',        // ({ path }) -> { free, total }   bytes on the disk that holds path
    'take',         // ({ dir, name }) -> { path }   a new file in dir for a recording, written to by fs.append and let go of by fs.dropTemp
    'pointer',      // ({ on, id, fps, paused }) -> on: { tracking } | off: { samples, clicks }   the pointer over the source being recorded, a sample a frame, as fractions of it
  ],
};

/** Backend-to-renderer events. */
export const EVENTS = [
  'win:state',        // { maximized, fullscreen, focused }
  'win:link',         // { url }             a link followed inside a frame (a message's body), for the window to open or warn about
  'talk:state',       // state                 Talk's connection, contacts and requests changed
  'talk:event',       // msg                   what the introducer said: messages, a call's set-up
  'win:frameMenu',    // { x, y, link, linkText, src, media, selection }   a right-click inside such a frame: what was under the pointer, for the window's own menu
  'live:status',      // { live, destinations }  Go Live's destinations as they connect, stream, drop and come back
  'app:open-file',    // { path }            OS asked us to open a document
  'app:command',      // { command, args }   menu / accelerator
  'app:link',         // { invite }          a rutba-office:// link Office was opened with, for the window of its app already open
  'mail:progress',    // { accountId, folder, done, total, phase }
  'mail:archives',    // { id, closed } an archive opened in place or closed
  'mail:indexing',    // { id, state, done, total } how far an archive's indexing has got
  'mail:exportProgress', // { job, done, total, dir } an export for legal hold under way
  'mail:accountsProgress', // { job, index, line, email, state, outcome, provider, server } one row of addAccounts
  'mail:new',         // { accountId, folder, count }
  'present:state',    // { id, index, running, startedAt, blank }
  'mail:oauth',       // { phase, provider, email }
  'announce:new',     // { id, title, body, link, kind }
  'contacts:changed', // { count }           the address book was written
  'calendar:changed', // {}                  a calendar or an event was written
  'dav:changed',      // {}                  an account synced, was added or was removed
  'mail:sent',        // { id, to }
  'mail:sendFailed',  // { id, message, attempts, gaveUp }         // { accountId, folder, count }
  'mail:syncProblem', // { accountId, kind, folder, to, count }   a change the mail server refused, put back as the server has it
  'mail:syncState',   // { accountId, offline, error, at }   an account's fetches and changes reaching its server, or not
  'mail:sealing',     // { done, total }   the mail store being sealed again with the profile's password (set, changed or taken away)
  'theme:changed',    // { dark }
  'profiles:changed', // { id, name, colour, count, protected }   a profile made, renamed, coloured or removed, here or in another running profile
  'ai:changed',       // { kind, id, step }    a thing's AI tools setting changed, in its own app or in AI Tools
  'ai:pending',       // [{ id, kind, thing, fields, body, draft }]   what the AI tools asked to do, waiting for the person's press
  'update:state',     // { state, version, available, percent, automatic, snoozed, snoozedVersion, arrived }
  'capture:command',  // { command: 'stop' | 'pause' }   a recording's global key or tray icon was used
  'picture:returned', // { token, bytes, contentType, width, height }   Edit in Image saved the picture back
  'packs:changed',    // { apps, restart }   Add or remove apps changed what the next start loads
];

/**
 * The app packs (docs/APP-PACKS.md): one per app, by the key the registry and
 * the windows already use, and text recognition, which four apps' windows
 * offer. An app not installed is absent: its namespaces below are not
 * registered in the main process and not offered to a window.
 */
export const PACKS = [
  { id: 'word', name: 'Documents' },
  { id: 'sheets', name: 'Worksheets' },
  { id: 'slides', name: 'Presentations' },
  { id: 'pdf', name: 'PDF' },
  { id: 'pictures', name: 'Pictures' },
  { id: 'image', name: 'Image' },
  { id: 'video', name: 'Video' },
  { id: 'recorder', name: 'Recorder' },
  { id: 'mail', name: 'Mail' },
  { id: 'calendar', name: 'Calendar' },
  { id: 'contacts', name: 'Contacts' },
  { id: 'base', name: 'Data Files' },
  { id: 'talk', name: 'Talk' },
  { id: 'ai', name: 'AI Tools' },
  { id: 'ocr', name: 'Text recognition' },
];

/**
 * Which pack each namespace is. A namespace not named here is the core's:
 * there whatever is installed (the document service the three document apps
 * share, printing, the CalDAV and CardDAV accounts both Calendar and
 * Contacts sync through, recording the screen for three apps).
 */
export const NAMESPACE_PACK = {
  mail: 'mail',
  discover: 'mail',
  calendar: 'calendar',
  contacts: 'contacts',
  pdf: 'pdf',
  ocr: 'ocr',
  base: 'base',
  talk: 'talk',
  live: 'recorder',
  ai: 'ai',
  present: 'slides',
};

/**
 * The events a pack's own service sends. The rest are the core's: Mail's
 * sign-in progress (`mail:oauth`) among them, since signing in is the core's
 * oauth service.
 */
export const EVENT_PACK = {
  'talk:state': 'talk',
  'talk:event': 'talk',
  'live:status': 'recorder',
  'mail:progress': 'mail',
  'mail:archives': 'mail',
  'mail:indexing': 'mail',
  'mail:exportProgress': 'mail',
  'mail:accountsProgress': 'mail',
  'mail:new': 'mail',
  'mail:sent': 'mail',
  'mail:sendFailed': 'mail',
  'mail:syncProblem': 'mail',
  'mail:syncState': 'mail',
  'mail:sealing': 'mail',
  'ai:changed': 'ai',
  'ai:pending': 'ai',
  'present:state': 'slides',
  'contacts:changed': 'contacts',
  'calendar:changed': 'calendar',
};

/** A pack's own namespaces and their methods, as its manifest carries them (`ipc`). */
export function packIpc(id) {
  const out = {};
  for (const [ns, owner] of Object.entries(NAMESPACE_PACK)) if (owner === id && METHODS[ns]) out[ns] = [...METHODS[ns]];
  return out;
}

/** A pack's own events, as its manifest carries them. */
export function packEvents(id) {
  return EVENTS.filter((e) => EVENT_PACK[e] === id);
}

/** The core's methods: every namespace no pack owns. */
export function coreMethods() {
  const out = {};
  for (const [ns, names] of Object.entries(METHODS)) if (!NAMESPACE_PACK[ns]) out[ns] = [...names];
  return out;
}

/** The core's events. */
export function coreEvents() {
  return EVENTS.filter((e) => !EVENT_PACK[e]);
}

/**
 * The methods one run offers: the core's and those of the packs named, as
 * the contract lists them. A run that loads packs from their files builds
 * the same thing from each pack's manifest (main.js); this is the list for
 * packs built in.
 */
export function methodsFor(packIds = []) {
  const out = coreMethods();
  for (const id of packIds) Object.assign(out, packIpc(id));
  return out;
}

/** The events one run offers: the core's and those of the packs named. */
export function eventsFor(packIds = []) {
  const want = new Set(packIds);
  return EVENTS.filter((e) => !EVENT_PACK[e] || want.has(EVENT_PACK[e]));
}

/** Every method as a flat channel list, e.g. 'fs:read': of the whole contract, or of the methods given. */
export function channels(methods = METHODS) {
  const out = [];
  for (const [ns, names] of Object.entries(methods)) for (const n of names) out.push(`${ns}:${n}`);
  return out;
}

export const CHANNEL_PREFIX = 'rutba-office';

/** The one synchronous request a window's preload makes as it starts: the methods, events and packs of this run. */
export const CONTRACT_CHANNEL = `${CHANNEL_PREFIX}/contract`;
