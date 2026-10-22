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
    'sync',         // ({ accountId, folder, limit }) -> { added, total }
    'messages',     // ({ accountId, folder, offset, limit, query }) -> { rows, total }
    'message',      // ({ accountId, uid, folder }) -> message
    'flag',         // ({ accountId, folder, uids, add, remove }) -> void
    'move',         // ({ accountId, folder, uids, to }) -> void
    'delete',       // ({ accountId, folder, uids }) -> void
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
    'open',         // ({ path } | { bytes, name }) -> { id, name, path, size, version, encrypted, repaired, title, pages: [{ width, height, rotate }] }
    'close',        // ({ id }) -> { closed }
    'assemble',     // ({ pages: [{ source, page, rotate }], path, title }) -> { path, size } | { bytes, size }: written whole, or handed back
    'text',         // ({ id, page }) -> string: the words a page's content draws
    'addText',      // ({ id, pages: [{ page, lines: [{ words: [{ text, x, y, width, size, angle }] }] }] }) -> { id, ...open's answer, written, bytes }: an invisible text layer as an incremental update, a new source
    'inspect',      // ({ path }) -> { pages: [{ width, height, rotate, text }], encrypted, title }: a file read and let go, for the checks
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
    'saveAs',       // ({ id, file }) -> summary   a copy at file, then the one open
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
  ai: [
    'state',        // () -> { mode: 'off' | 'session' | 'always', folders, connect: { command, args, env }, log, connections: [{ client, clientVersion, where, since, last, lastTool, pid }], reach: 'local' | 'remote', remote: { port, key, addresses, listening, error } | null }
    'setMode',      // ({ mode }) -> state
    'addFolder',    // ({ path }) -> state
    'removeFolder', // ({ path }) -> state
    'setReach',     // ({ reach: 'local' | 'remote', port? }) -> state   this computer only, or other machines too, with a key
    'renewKey',     // () -> state   a new key; the old one stops working at once
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
    'sources',      // () -> [{ id, name, kind, thumbnail, display? }]   every window but the asking one, then each screen (display: its size in pixels)
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
  'theme:changed',    // { dark }
  'update:state',     // { state, version, available, percent, automatic, snoozed, snoozedVersion, arrived }
  'capture:command',  // { command: 'stop' | 'pause' }   a recording's global key or tray icon was used
  'picture:returned', // { token, bytes, contentType, width, height }   Edit in Image saved the picture back
];

/** Every method as a flat channel list, e.g. 'fs:read'. */
export function channels() {
  const out = [];
  for (const [ns, names] of Object.entries(METHODS)) for (const n of names) out.push(`${ns}:${n}`);
  return out;
}

export const CHANNEL_PREFIX = 'rutba-office';
