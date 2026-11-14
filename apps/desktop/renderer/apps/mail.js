// Rutba Mail.
//
// Three panes and one rule: a message body is never trusted. It is rendered in
// a sandboxed frame with no scripts, no same-origin access and a policy that
// blocks every remote fetch, so a tracking pixel cannot report that you opened
// the mail and a script cannot reach anything at all. Remote images load only
// when you ask, per message — and the frame's neighbour says who wanted them.
//
// The other half of this app is the part no free client does well: opening what
// you already have. A .pst, a .ost, an Outlook for Mac archive, an mbox, a
// folder of .eml files — found on your disk without being asked, scanned so you
// can see what is in there, then imported into the same local store live
// accounts synchronise into.
//
// Everything below is drawn from that one local store. That is what lets this
// offer a unified inbox across accounts, conversations that hold together
// across an imported archive, an attachment view, and a list of the people who
// write to you — none of which a client that only proxies IMAP can do.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  Ribbon, Group, Button, Separator, Icon, Spacer, Chip, Empty, Spinner, Panel, Content, List, Item,
  Search, Dialog, Field, Input, Select, useToast, useMenu, useCommands, useWheelZoom, menuItems,
  formatWhen, formatBytes, t, tn, tFilled, prettyKey,
} from '@rutba/office-ui';
import { AppFrame, useAppMenu, useFileDrop } from '../shell.js';
import { pathOf } from '@rutba/office-shell/client';
import { appFor, kindFromExtension } from '@rutba/office-formats/sniff';
import { DefaultsDialog } from '../defaults.js';
import { avatarFor, displayName, buildThreads, arrange, FILTERS, SORTS, installStyles, stripTags, safeDraftHtml, freeNameIn, pictureInline, nextAfterRemoved, rowDeleteTargets, syncProblemWords, openMessage, rowKey, menuRows, afterOpenStep, rowAsListed, createKeepOnce } from './mail/parts.js';
import { useAttachmentNotes, AttachmentNotes } from './mail/docnote.js';
import { linkKind, linkRisks, mailtoFields, fromFrameUrl } from './mail/links.js';
import Reader from './mail/reader.js';
import { InviteCard } from './calendar-invite-card.js';
import { frameMenuItems, pictureBytes, savePicture, SourceDialog, SAFETY_CSS } from './mail/safety.js';
import Compose from './mail/compose.js';
import { AccountDialog, AccountsDialog, ImportDialog, ImportPreview, ImportingDialog, FilesView, PeopleView, SignatureDialog, OutOfOfficeDialog, JunkDialog } from './mail/dialogs.js';
import { RulesDialog } from './mail/rules.js';
import { PromptDialog } from '@rutba/office-ui/proofing';
import { attachmentExtension } from '@rutba/mailbox/mime';
import { withSignature, withSignatureHtml, signatureTextToHtml } from '@rutba/mailbox/signature';
import { textToHtml } from './mail/richtext.js';
import { isOpenedId, openedFileOf, carriesFiles, emlSources, emlNameFor, OpenedBar, installEmlStyles } from './mail/eml.js';
import { isArchiveId, ArchiveBanner, ExportHoldDialog, SearchHelp, installArchiveStyles, indexedCount } from './mail/archive.js';
import { CertificatesDialog, installSmimeStyles } from './mail/smime.js';
import { BulkAccountsDialog, installBulkStyles } from './mail/accounts-bulk.js';
import { PasswordAgain } from './mail/password-again.js';
import { createReader } from '../read-aloud/reader.js';
import { ReadAloudBar } from '../read-aloud/bar.js';
import { VoiceOffer } from '../read-aloud/voice-offer.js';
import { createMailReading, queueOf, placeOfList, openingDoes, readButtons, readerStalled, marksReadAfterReading, createQuietOpen, readingStopsFor, refocusAfterReading } from './mail/read-aloud.js';

installStyles();
installArchiveStyles();
installEmlStyles();
installSmimeStyles();
installBulkStyles();

/** Every account at once. Not an id any account can have. */
const EVERYTHING = '*';

const keyOf = (row) => `${row.accountId}|${row.folder}|${row.id}`;

/** What takes a message out of the folder shown: it leaves the list at once. */
const LEAVES = ['delete', 'archive', 'junk', 'move', 'notJunk', 'block'];

function MailWindow({ app, shell, opened = null }) {
  const toast = useToast();
  const menu = useMenu();

  // A message opened in a window of its own, by a double-click in the list
  // (owner, 2026-10-19): the window shows that message alone, with the
  // ribbon to answer, file or delete it, and closes once it is moved away.
  const alone = useMemo(() => {
    // A .eml opened from Explorer: that message alone, read only (Mail 2.25).
    if (opened) return { accountId: opened.accountId, folder: opened.folder, id: opened.id };
    try {
      const m = JSON.parse(new URLSearchParams(location.search).get('message') || 'null');
      return m && m.id != null ? { accountId: m.accountId, folder: m.folder, id: m.id } : null;
    } catch {
      return null;
    }
  }, []);

  const [accounts, setAccounts] = useState([]);
  const [accountId, setAccountId] = useState(alone?.accountId ?? null);
  const [folders, setFolders] = useState([]);
  // Bumped when mail arrives or a message moves: a folder the move made, and
  // every folder's count, are read again.
  const [folderTick, setFolderTick] = useState(0);
  const [folder, setFolder] = useState(alone?.folder ?? null);
  const [role, setRole] = useState('inbox');
  const [list, setList] = useState({ rows: [], total: 0 });
  const [busy, setBusy] = useState(false);

  const [checked, setChecked] = useState(() => new Set());
  const [selected, setSelected] = useState(alone); // the row object, so unified lists know where it lives
  const [message, setMessage] = useState(null);
  // Where the message the reading pane shows lives, or the one it could not open: what the pane's own buttons act on,
  // and, once it is the one selected, what the ribbon and its keys act on (openMessage).
  const [shownAt, setShownAt] = useState(null);
  const shownFor = shownAt ? rowKey(shownAt) : null;
  // Why the message selected could not be opened, said in the pane where it would have been.
  const [openError, setOpenError] = useState(null);
  // One's own certificates for signed and encrypted mail, and a count that
  // reopens the message read when one is imported (it may open now).
  const [identities, setIdentities] = useState([]);
  const [smimeTick, setSmimeTick] = useState(0);
  const [insight, setInsight] = useState(null);
  const [expanded, setExpanded] = useState(() => new Set());
  const lastClicked = useRef(null);

  const [query, setQuery] = useState('');
  const [tab, setTab] = useState('home');
  const [view, setView] = useState('mail');
  const [remote, setRemote] = useState(false);
  const [plain, setPlain] = useState(false);
  const [progress, setProgress] = useState(null);
  const [dialog, setDialog] = useState(() => (new URLSearchParams(location.search).has('import') ? { kind: 'import' } : null));
  const [compose, updateCompose] = useState(null);
  // Each message opened to write gets a form of its own (composeKey): the body
  // is seeded once, so a form reused for another draft kept the last one's words
  // under the new one's recipients. updateCompose is for the message already there.
  const composeKey = useRef(0);
  // Read Aloud as of the last render (set where it is defined, below): whether it reads, whether a docked form is
  // out of sight for it, and how to stop it.
  const readingNow = useRef(false);
  const formHidden = useRef(false);
  const stopReadingRef = useRef(null);
  const setCompose = useCallback((next) => {
    // A form opened while Read Aloud reads stops the reading: the form is seen, not opened out of sight.
    if (readingStopsFor({ reading: readingNow.current, opening: Boolean(next) })) stopReadingRef.current?.();
    // The form held out of sight is gone or shown now: an open asked for before the next render keeps nothing again.
    formHidden.current = false;
    composeKey.current += 1;
    updateCompose(next);
  }, []);
  const composeOpen = useRef(null);
  const editAsked = useRef(0);
  const selectedWas = useRef(null);
  // Every new form (Reply, Reply all, Forward, New message, a message taken back, a mailto link) opens through here.
  // While Read Aloud keeps a docked form out of sight, what was typed in it is kept in Drafts first, as opening a
  // draft does (keepWriting); then the reading stops (setCompose) and the new form shows. Not kept: the reading
  // stops and that form shows again, as it was, so nothing in it goes unseen.
  // A form is kept once however many ask at a time (createKeepOnce): Reply pressed twice saves one draft, not two.
  const keepWritingNow = useRef(null);
  const [keepFormOnce] = useState(() => createKeepOnce((writing) => keepWritingNow.current(writing)));
  const openForm = useCallback(async (next) => {
    const hidden = formHidden.current ? composeOpen.current?.() : null;
    // The last asked wins: a form being left, or a draft being opened, meanwhile does not close this one.
    const asked = (editAsked.current += 1);
    if (hidden) {
      if (!(await keepFormOnce(composeKey.current, hidden))) {
        stopReadingRef.current?.();
        return false;
      }
      // Another form was asked for while that one was being kept.
      if (asked !== editAsked.current) return false;
    }
    setCompose(next);
    return true;
  }, [setCompose, keepFormOnce]);
  // A message is written docked in the reading pane or in the whole window,
  // as last chosen (a choice of the whole window before is kept as one); or,
  // popped out, in a window of its own.
  const [composePlace, setComposePlace] = useState('docked');
  useEffect(() => {
    Promise.all([shell.store.get({ key: 'mail.composePlace', fallback: null }), shell.store.get({ key: 'mail.composeFull', fallback: false })])
      .then(([placeKept, fullKept]) => setComposePlace(placeKept === 'full' || placeKept === 'docked' ? placeKept : fullKept ? 'full' : 'docked'))
      .catch(() => {});
  }, [shell]);
  // A window of its own for one message being written: its draft handed over through the store, then forgotten there.
  const composeAlone = useMemo(() => new URLSearchParams(location.search).get('compose') || null, []);
  const hadCompose = useRef(false);
  useEffect(() => {
    if (!composeAlone) return;
    shell.store.get({ key: composeAlone, fallback: null })
      .then((draft) => {
        if (draft) setCompose(draft);
        return shell.store.set({ key: composeAlone, value: null });
      })
      .catch(() => {});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  // Sent, saved or discarded, the window of its own goes with it.
  useEffect(() => {
    if (compose) hadCompose.current = true;
    else if (composeAlone && hadCompose.current) shell.win.close({}).catch(() => {});
  }, [compose, composeAlone, shell]);
  const [scan, setScan] = useState(null);
  const [outbox, setOutbox] = useState([]);
  // Archives opened where they lie, and the search box's reach: the folder
  // open, or all mail (every account and every archive). `found` is what the
  // last search over all mail said: how many, whether counted exactly, and
  // which archives were still being indexed.
  const [archives, setArchives] = useState([]);
  const [scope, setScope] = useState('folder');
  const [found, setFound] = useState(null);
  const quietSearch = useRef(false);
  const searchSeq = useRef(0);

  // How the list looks. Remembered, because nobody wants to choose twice.
  const [prefs, setPrefs] = useState({
    filter: 'all',
    sort: 'date',
    ascending: false,
    threaded: true,
    layout: 'right',
    density: 'cosy',
    // View → Zoom: the message alone, not the window.
    zoom: 1,
  });
  const setPref = useCallback(
    (patch) =>
      setPrefs((p) => {
        const next = { ...p, ...patch };
        shell.store.set({ key: 'mail.view', value: next }).catch(() => {});
        return next;
      }),
    [shell]
  );

  const dark = document.documentElement.dataset.theme === 'dark';
  const unified = accountId === EVERYTHING;
  // New mail goes from the account being read when it can send, else from
  // "Send new mail from" in Accounts (D-2026-11-11-02): in the combined
  // Inbox, in a mailbox brought in from a file, and from a mailto: link.
  const here = accounts.find((a) => a.id === accountId);
  // With no account that can send, it stays the one being read, as before.
  const newMailFrom = here?.canSend ? here.id : (accounts.find((a) => a.isDefault) || accounts.find((a) => a.canSend) || here || accounts[0])?.id || null;
  // An archive opened in place, or a message opened from a file: read only alike.
  const archiveId = isArchiveId(accountId) || isOpenedId(accountId) ? accountId : null;
  // An answer to a message from an archive or a file goes from an account that can send.
  const sendingFrom = (id) => (accounts.some((a) => a.id === id) ? id : newMailFrom);
  const openArchive = archives.find((a) => a.id === archiveId) || null;
  const allMail = scope === 'all' && query.trim().length > 0;

  // Office Mail as the mail app Windows opens mailto: links and .eml files with (Mail 2.25):
  // Settings opened at Rutba Office, where the person chooses; Windows lets nobody else.
  const chooseMailAppRef = useRef(null);
  // File, Open brings .eml files into the folder open; Save As saves the messages chosen as .eml (Mail 2.25).
  const fileMenu = useRef({});
  const appMenu = useAppMenu({
    shell,
    appKey: 'mail',
    onOpen: () => fileMenu.current.open?.(),
    saveAs: () => fileMenu.current.saveAs?.(),
    own: [
      { label: t('mail.import-mail'), icon: 'import', run: () => setDialog({ kind: 'import' }) },
      { label: t('mail.add-account'), icon: 'plus', run: () => setDialog({ kind: 'account' }) },
      { label: t('mail.open-files-rutba'), icon: 'settings', run: () => setDialog({ kind: 'defaults' }) },
      { label: t('mail.make-office-mail'), icon: 'mail', run: () => chooseMailAppRef.current?.() },
      { label: t('mail.certificates-signed'), icon: 'shield', run: () => setDialog({ kind: 'certificates' }) },
    ],
  });

  /* ── loading ──────────────────────────────────────────────────────────── */

  useEffect(() => {
    shell.store
      .get({ key: 'mail.view', fallback: null })
      .then((saved) => saved && setPrefs((p) => ({ ...p, ...saved })))
      .catch(() => {});
    // A window an earlier build zoomed stays zoomed across restarts — the
    // level is kept per origin — and the message carries the zoom now, so
    // the window itself goes back to 100%.
    shell.win.zoom({ reset: true }).catch(() => {});
  }, [shell]);

  const loadAccounts = useCallback(async () => {
    const found = await shell.mail.accounts();
    setAccounts(found);
    setAccountId((current) => current || found[0]?.id || null);
    return found;
  }, [shell]);

  useEffect(() => {
    loadAccounts().catch(() => {});
  }, [loadAccounts]);

  const loadArchives = useCallback(async () => {
    const list = await shell.mail.archives();
    setArchives(list);
    return list;
  }, [shell]);

  useEffect(() => {
    loadArchives().catch(() => {});
    const offList = shell.on('mail:archives', () => loadArchives().catch(() => {}));
    // How far an archive's indexing has got, for its strip and the status bar.
    const offIndex = shell.on('mail:indexing', (p) =>
      setArchives((list) => list.map((a) => (a.id === p.id ? { ...a, index: { ...a.index, state: p.state, done: p.done, total: p.total } } : a)))
    );
    return () => {
      offList?.();
      offIndex?.();
    };
  }, [shell, loadArchives]);

  useEffect(() => {
    if (!accountId || unified) return;
    shell.mail
      .folders({ accountId })
      .then((f) => {
        setFolders(f);
        // The inbox first. An imported archive has none: then the ordinary
        // folder holding the most mail, which is the archive itself — not
        // Sent (once a message had been sent from it, the window opened on an
        // empty Sent) and not a folder a rule has only just made.
        const busiest = f
          .filter((x) => x.role === 'folder')
          .reduce((best, x) => ((x.total ?? 0) > (best?.total ?? -1) ? x : best), null);
        setFolder((current) =>
          f.some((x) => x.path === current)
            ? current
            : f.find((x) => x.role === 'inbox')?.path || busiest?.path || f[0]?.path || null
        );
      })
      .catch(() => setFolders([]));
  }, [accountId, unified, shell, folderTick]);

  const refreshList = useCallback(async () => {
    if (view !== 'mail') return;
    // Asked again as an archive's indexing goes on: no spinner over the list.
    const quiet = quietSearch.current;
    quietSearch.current = false;
    if (allMail) {
      // One search over every account and every archive; an answer that
      // arrives after a newer one was asked for is dropped.
      const seq = ++searchSeq.current;
      if (!quiet) setBusy(true);
      try {
        const r = await shell.mail.searchAll({ query, limit: 500 });
        if (seq !== searchSeq.current) return;
        setList({ rows: r.rows, total: r.total });
        setFound({ ...r, query });
      } catch (err) {
        toast(err.message, { tone: 'bad' });
      } finally {
        if (seq === searchSeq.current) setBusy(false);
      }
      return;
    }
    setFound(null);
    if (unified) {
      if (!quiet) setBusy(true);
      try {
        setList(await shell.mail.unified({ role, query, limit: 500 }));
      } catch (err) {
        toast(err.message, { tone: 'bad' });
      } finally {
        setBusy(false);
      }
      return;
    }
    if (!accountId || !folder) {
      setList({ rows: [], total: 0 });
      return;
    }
    if (!quiet) setBusy(true);
    try {
      const next = await shell.mail.messages({ accountId, folder, query, limit: 500 });
      // Rows from a single folder do not carry where they came from; every
      // action downstream needs it, so it is stamped on here once.
      setList({ ...next, rows: next.rows.map((r) => ({ ...r, accountId, folder })) });
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    } finally {
      setBusy(false);
    }
  }, [allMail, unified, role, accountId, folder, query, view, shell, toast]);

  // While an archive searched is still being indexed, the search is asked
  // again every second and a half, so its results and its count grow as the
  // archive is indexed; and once more when an archive's indexing has changed
  // since (resumed, stopped or finished), so the answer says where it is now.
  const searchStale = useMemo(
    () =>
      Boolean(
        allMail &&
          found?.indexing?.some((a) => {
            const now = archives.find((x) => x.id === a.id)?.index?.state;
            return a.state === 'indexing' || now === 'indexing' || (now && now !== a.state);
          })
      ),
    [allMail, found, archives]
  );
  // A folder listed from what has been indexed of it so far (a search in an
  // archive, a large .olm) is asked for again as indexing goes on.
  const listGrows = !allMail && list?.indexed?.state === 'indexing';
  useEffect(() => {
    if (!listGrows) return undefined;
    const timer = setTimeout(() => {
      quietSearch.current = true;
      refreshList();
    }, 2000);
    return () => clearTimeout(timer);
  }, [listGrows, list, refreshList]);

  useEffect(() => {
    if (!searchStale) return undefined;
    const timer = setTimeout(() => {
      quietSearch.current = true;
      refreshList();
    }, 1500);
    return () => clearTimeout(timer);
  }, [searchStale, found, refreshList]);

  useEffect(() => {
    refreshList();
    setChecked(new Set());
    // A message's own window keeps its message whatever the list does.
    if (alone) return;
    setSelected(null);
    setMessage(null);
    setInsight(null);
  }, [refreshList, alone]);

  // A server's folder opened is asked for what the server has in it, quietly: one never
  // opened here (mail a server's filter put in Junk, a folder made in webmail) is otherwise empty.
  useEffect(() => {
    if (alone || unified || !accountId || !folder) return undefined;
    if (accounts.find((a) => a.id === accountId)?.local) return undefined;
    let live = true;
    shell.mail.sync({ accountId, folder })
      .then((r) => { if (live && (r?.added || 0) > 0) refreshList(); })
      .catch(() => { /* offline: Send/Receive says so when asked */ });
    return () => { live = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [accountId, folder, unified, alone]);

  // A message Read Aloud shows as its reading starts (D-2026-11-13-34): its row's key, until the selection takes it.
  const [quietOpen] = useState(createQuietOpen);
  // The message Read Aloud loaded last, by its key, for the pane to open without loading it again.
  const readLoaded = useRef(new Map());
  // Messages marked read or unread by hand since the reading began: what a person sets wins (review 2).
  const handMarked = useRef(new Set());

  // Opening a message: the body and the report on it come together, because
  // showing the body first and the warning a moment later is exactly backwards.
  useEffect(() => {
    const selectedNow = selected ? JSON.stringify([selected.accountId, selected.folder, selected.id]) : null;
    const moved = selectedNow !== selectedWas.current;
    selectedWas.current = selectedNow;
    const formWhenAsked = composeKey.current;
    if (!selected) {
      setMessage(null);
      setShownAt(null);
      setOpenError(null);
      quietOpen.take(null, smimeTick);
      setInsight(null);
      // A draft's form goes with the draft when the folder is left. A new message being written stays.
      if (moved && composeOpen.current?.()?.draft.replaces) leaveFormRef.current?.();
      return;
    }
    const { accountId: a, folder: f, id } = selected;
    // Shown by Read Aloud: no draft opened, no form left, not marked read until read to its end.
    const quiet = quietOpen.take(keyOf(selected), smimeTick);
    // The message Read Aloud has just loaded to read, opened as it is rather than loaded again.
    // Only as the selection moves to it: a re-run for the certificates loads it again, to read it with them.
    const loaded = quiet && moved ? readLoaded.current.get(keyOf(selected)) : undefined;
    readLoaded.current.clear();
    setRemote(false);
    let current = true;
    Promise.all([loaded ? Promise.resolve(loaded) : shell.mail.message({ accountId: a, folder: f, id }), shell.mail.insight({ accountId: a, folder: f, id })])
      .then(([m, i]) => {
        // Another message was clicked while this one was on its way.
        if (!current) return;
        setMessage(m);
        setShownAt({ accountId: a, folder: f, id });
        setOpenError(null);
        setInsight(i);
        const does = openingDoes({ quiet, readOnly: Boolean(m?.readOnly), unread: Boolean(m?.unread), inDrafts: Boolean(m) && isDraftsFolder(a, f), moved, formUntouched: formWhenAsked === composeKey.current });
        // A draft is opened to be finished, not read (owner, 2026-10-20).
        if (does.editDraft) editDraftRef.current?.({ accountId: a, folder: f, id }, m);
        // Anything else clicked (Sent, the Inbox) is there to be read, in the place the form was holding.
        else if (does.leaveForm) leaveFormRef.current?.();
        if (does.markRead) {
          shell.mail.flag({ accountId: a, folder: f, ids: [id], patch: { unread: false } }).then(() => {
            setList((current) => ({
              ...current,
              rows: current.rows.map((r) => (r.id === id && r.folder === f ? { ...r, unread: false } : r)),
            }));
            loadAccounts();
          });
        }
      }, (err) => {
        // The message could not be opened (a throw in the lines above is not this, and does not blank one that did).
        toast(err.message, { tone: 'bad' });
        // Not left showing the message before it: the one clicked is the one Delete and the rest then take.
        if (!current) return;
        setMessage(null);
        setInsight(null);
        setShownAt({ accountId: a, folder: f, id });
        setOpenError(err?.message || String(err));
      })
      .catch((err) => toast(err.message, { tone: 'bad' }));
    return () => { current = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selected, smimeTick]);

  // A PDF or picture attached to the message open, read as an invoice or a receipt now it is opened.
  const attachmentNotes = useAttachmentNotes({ shell, where: shownAt, message });

  useEffect(() => {
    shell.mail.smimeIdentities().then(setIdentities).catch(() => {});
  }, [shell, smimeTick]);

  useEffect(() => {
    const offProgress = shell.on('mail:progress', (p) => setProgress(p));
    const offNew = shell.on('mail:new', () => {
      setProgress(null);
      loadAccounts();
      refreshList();
      setFolderTick((n) => n + 1);
    });
    const offSent = shell.on('mail:sent', ({ to, refused, notKept }) => {
      if (refused?.length) toast(t('mail.sent-server-refused', { who: refused.join(', ') }), { tone: 'bad', ms: 9000 });
      else if (notKept) toast(t('mail.sent-copy-could'), { tone: 'bad', ms: 9000 });
      else toast(t('mail.sent-to-to', { to }), { tone: 'good' });
      shell.mail.outbox().then(setOutbox).catch(() => {});
      refreshList();
      // The first message sent makes Sent: the folders are read again so it
      // shows at once, not after the window is opened again.
      loadAccounts();
      setFolderTick((n) => n + 1);
    });
    const offFailed = shell.on('mail:sendFailed', ({ message: why, gaveUp }) => {
      toast(gaveUp ? t('mail.could-not-send', { why }) : t('mail.send-failed-trying', { why }), {
        tone: 'bad',
        ms: 9000,
      });
      shell.mail.outbox().then(setOutbox).catch(() => {});
    });
    // A change the mail server would not make: said, and the list shown as the server has it.
    const offProblem = shell.on('mail:syncProblem', (p) => {
      toast(syncProblemWords(p), { tone: 'bad', ms: 9000 });
      loadAccounts();
      refreshList();
      setFolderTick((n) => n + 1);
    });
    return () => {
      offProgress?.();
      offNew?.();
      offSent?.();
      offFailed?.();
      offProblem?.();
    };
  }, [shell, loadAccounts, refreshList, toast]);

  useEffect(() => {
    shell.mail.outbox().then(setOutbox).catch(() => {});
  }, [shell]);

  // Mail fetches by itself as it opens (A6.2), and keeps doing so in the
  // background; how often is a choice on Send / Receive. A message's own
  // window leaves that to the main one.
  const [fetching, setFetching] = useState({ every: 5, push: true });
  useEffect(() => {
    shell.mail.fetchSettings?.().then(setFetching).catch(() => {});
    if (!alone) shell.mail.fetchNow?.().catch(() => {});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shell]);
  const changeFetching = useCallback(
    (patch) => shell.mail.setFetchSettings({ ...fetching, ...patch }).then(setFetching).catch((err) => toast(err.message, { tone: 'bad' })),
    [shell, fetching, toast]
  );
  const fetchWhen = (every) => (every ? tn(every, 'mail.every-minute') : t('mail.only-get-mail'));

  /* ── the list, as it is actually drawn ────────────────────────────────── */

  const rows = useMemo(() => arrange(list.rows, prefs), [list.rows, prefs]);
  const threads = useMemo(() => (prefs.threaded ? buildThreads(rows) : rows.map((r) => ({ ...r, messages: [r], count: 1 }))), [rows, prefs.threaded]);

  const allRowsIn = useCallback((thread) => (prefs.threaded ? thread.messages : [thread]), [prefs.threaded]);

  /* ── Read Aloud (D-2026-11-13-34) ─────────────────────────────────────── */

  // The message open, every message in the list, or its unread ones, read on
  // the reader the apps share. The list as drawn decides: its filter and its
  // search. Each message is shown as its reading starts.
  const speakerRef = useRef(null);
  const playerRef = useRef(null);
  const [readState, setReadState] = useState(null);
  const [reading, setReading] = useState(null); // { mode, key, n, count } while reading
  const [addedVoices, setAddedVoices] = useState(null); // the added voices, for the offer of one
  const speaker = () => (speakerRef.current ||= createReader({
    shell,
    onState: (s) => {
      setReadState(s);
      // A voice that failed, or no speech at all: the shared reader stops and never says it ended.
      if (readerStalled(s)) playerRef.current?.stalled();
    },
  }));
  const readingModeRef = useRef(null); // the mode of the run under way
  /** A message read to its end, marked read, here and in the list (the PM, on review 1). */
  const markReadAfterReading = useCallback((row, m) => {
    const local = isArchiveId(row.accountId) || isOpenedId(row.accountId);
    if (!marksReadAfterReading({ mode: readingModeRef.current, unread: m?.unread, readOnly: m?.readOnly, local, touched: handMarked.current.has(keyOf(row)) })) return;
    shell.mail.flag({ accountId: row.accountId, folder: row.folder, ids: [row.id], patch: { unread: false } }).then(() => {
      setList((current) => ({ ...current, rows: current.rows.map((r) => (keyOf(r) === keyOf(row) ? { ...r, unread: false } : r)) }));
      loadAccounts();
    }).catch(() => { /* left unread, as it was */ });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shell]);
  const markReadRef = useRef(markReadAfterReading);
  markReadRef.current = markReadAfterReading;
  const player = () => (playerRef.current ||= createMailReading({
    reader: speaker(),
    load: (row) => shell.mail.message({ accountId: row.accountId, folder: row.folder, id: row.id }).then((m) => {
      // Handed to the pane, which opens this one as it is (one load per message read).
      readLoaded.current.clear();
      if (m) readLoaded.current.set(keyOf(row), m);
      return m;
    }),
    keyOf,
    onMessage: (row) => {
      if (alone) return;
      quietOpen.mark(keyOf(row));
      setSelected((s) => {
        if (s && keyOf(s) === keyOf(row)) {
          quietOpen.unmark(keyOf(row)); // already open: nothing to show
          return s;
        }
        return row;
      });
    },
    // Play all and Play unread only: the open message already follows the reading-pane setting (the PM).
    onRead: (row, m) => markReadRef.current(row, m),
    onState: (s) => setReading((r) => (s.playing ? { ...(r || {}), n: s.n, count: s.count } : null)),
    onEnd: () => {
      quietOpen.clear();
      setReading(null);
    },
    onStall: () => {
      quietOpen.clear();
      setReading(null);
      const said = speakerRef.current?.state().said;
      toast(said ? t(said) : t('mail.reading-stopped-voice'), { ms: 5000 });
    },
  }));
  const stopReading = useCallback(() => {
    playerRef.current?.stop();
    quietOpen.clear();
    setReading(null);
  }, [quietOpen]);
  stopReadingRef.current = stopReading;
  readingNow.current = Boolean(reading);
  // A stop the person asked for (Play again, the bar's close): the form the reading kept out of sight takes the keys.
  const stopAsked = useRef(false);
  const stopReadingAsked = useCallback(() => {
    stopAsked.current = true;
    stopReading();
  }, [stopReading]);
  // What Read aloud reads: the message the reading pane shows, as the ribbon takes it (openMessage); never one
  // clicked that the pane does not show yet, nor one that could not be opened. With the pane off, the row selected.
  const readable = alone || prefs.layout === 'off' ? selected : (openError ? null : openMessage(selected, shownFor));
  const readAloud = useCallback((mode) => {
    // Read aloud again on the message being read stops it; on another message, reads that one.
    if (mode === 'open' && buttons.open.pressed) return stopReadingAsked();
    stopAsked.current = false;
    const list = mode === 'open' ? [readable].filter(Boolean) : queueOf({ threads, rowsIn: allRowsIn, mode });
    if (!list.length) {
      toast(mode === 'unread' ? t('mail.no-unread-messages') : t('mail.no-message-read'));
      return undefined;
    }
    setReading({ mode, key: mode === 'open' ? keyOf(list[0]) : null, n: 1, count: list.length });
    readingModeRef.current = mode;
    handMarked.current = new Set();
    if (!player().play(list)) setReading(null);
    // The added voices there are, so the offer of one shows when there are none.
    if (shell.speech?.voices) shell.speech.voices().then((v) => setAddedVoices(Array.isArray(v) ? v : [])).catch(() => setAddedVoices(null));
    return undefined;
  }, [reading, readable, threads, allRowsIn, toast, stopReadingAsked, shell]);
  // Another folder, account, folder of All accounts, view or search scope: what was being read is not the list any more.
  const listPlace = placeOfList({ accountId, folder, role, view, scope, unified });
  useEffect(() => stopReading(), [listPlace, stopReading]);
  const buttons = readButtons({ reading, rows, alone: Boolean(alone), hasMessage: Boolean(message) && Boolean(readable), openKey: readable ? keyOf(readable) : null });
  useEffect(() => () => { playerRef.current?.stop(); speakerRef.current?.stop(); }, []);

  const checkedRows = useMemo(() => {
    if (!checked.size) return [];
    const out = [];
    for (const thread of threads) for (const row of allRowsIn(thread)) if (checked.has(keyOf(row))) out.push(row);
    return out;
  }, [checked, threads, allRowsIn]);

  const toggleCheck = useCallback(
    (thread, event) => {
      const members = allRowsIn(thread);
      setChecked((current) => {
        const next = new Set(current);
        const on = members.every((r) => next.has(keyOf(r)));

        // Shift extends from the last one touched, the way every list does.
        if (event?.shiftKey && lastClicked.current != null) {
          const from = threads.findIndex((t) => t.threadKey === lastClicked.current);
          const to = threads.indexOf(thread);
          if (from >= 0 && to >= 0) {
            for (let i = Math.min(from, to); i <= Math.max(from, to); i++) {
              for (const r of allRowsIn(threads[i])) next.add(keyOf(r));
            }
            return next;
          }
        }

        for (const r of members) {
          if (on) next.delete(keyOf(r));
          else next.add(keyOf(r));
        }
        return next;
      });
      lastClicked.current = thread.threadKey ?? keyOf(thread);
    },
    [threads, allRowsIn]
  );

  /* ── actions ──────────────────────────────────────────────────────────── */

  /**
   * Do one thing to many messages that may live in different folders and
   * different accounts. Grouping is not an optimisation here — the store is
   * addressed per folder, so it is the only way the call can be made at all.
   */
  const overRows = useCallback(
    async (targets, run) => {
      const groups = new Map();
      for (const row of targets) {
        const key = `${row.accountId}|${row.folder}`;
        if (!groups.has(key)) groups.set(key, { accountId: row.accountId, folder: row.folder, ids: [] });
        groups.get(key).ids.push(row.id);
      }
      for (const group of groups.values()) await run(group);
    },
    []
  );

  const targets = useCallback(() => {
    // An archive opened in place is read only: nothing in it is moved,
    // flagged or deleted, whatever is selected.
    const writable = (rows) => rows.filter((r) => !isArchiveId(r.accountId));
    // A message as the list has it, star and pin with it, so Star and Pin turn off as well as on.
    const listed = (row) => rowAsListed(row, { rows: threads.flatMap(allRowsIn), message, shownFor });
    // In a message's own window, that message and nothing else.
    if (alone) return selected ? writable([listed(selected)]) : [];
    if (checkedRows.length) return writable(checkedRows);
    // The message open in the reading pane, never one clicked that it does not show yet.
    // With the reading pane off, the row selected is the message open: there is no pane to wait for.
    const open = prefs.layout === 'off' ? selected : openMessage(selected, shownFor);
    if (!open) return [];
    const thread = threads.find((t) => allRowsIn(t).some((r) => keyOf(r) === keyOf(open)));
    return writable(thread && prefs.threaded ? thread.messages : [listed(open)]);
  }, [alone, checkedRows, selected, shownFor, threads, allRowsIn, prefs.threaded, prefs.layout, message]);
  // Nothing checked and no message open: the ribbon's message buttons are greyed until the pane shows the one clicked.
  const noMessage = !checked.size && !(alone || prefs.layout === 'off' ? selected : openMessage(selected, shownFor));

  const foldersFor = useCallback(
    (accId) => (accId === accountId && !unified ? folders : []),
    [accountId, unified, folders]
  );

  const act = useCallback(
    async (what, extra, rows = null) => {
      // `rows`: the messages dragged onto a folder, whatever is selected.
      const chosen = rows || targets();
      if (!chosen.length) return;
      // Where the list goes on once these leave it: the message after the
      // last one removed, read at once, as Outlook does, rather than an
      // empty pane and a list that starts again from the top.
      const gone = chosen.map(keyOf);
      const after = nextAfterRemoved(threads, gone, allRowsIn);
      const stays = selected && !gone.includes(keyOf(selected));
      try {
        if (what === 'delete') {
          await overRows(chosen, async (g) => {
            const trash = foldersFor(g.accountId).find((f) => f.role === 'trash');
            if (trash && g.folder !== trash.path) await shell.mail.move({ ...g, to: trash.path });
            else await shell.mail.delete(g);
          });
        } else if (what === 'archive') {
          await overRows(chosen, (g) => shell.mail.move({ ...g, to: foldersFor(g.accountId).find((f) => f.role === 'archive')?.path || 'Archive' }));
        } else if (what === 'junk') {
          await overRows(chosen, (g) => shell.mail.move({ ...g, to: foldersFor(g.accountId).find((f) => f.role === 'junk')?.path || 'Junk' }));
        } else if (what === 'notJunk') {
          await overRows(chosen, (g) => shell.mail.notJunk(g));
        } else if (what === 'block' || what === 'safe' || what === 'safeDomain') {
          // Block Sender, Never Block Sender, Never Block Sender's Domain: the
          // list changed, then the message put where the list now says.
          const senders = [...new Set(chosen.map((r) => String(r.from?.address || '').toLowerCase()).filter((a) => a.includes('@')))];
          const entries = senders.map((a) => (what === 'safeDomain' ? '@' + a.split('@').pop() : a));
          // On the list of the account each message came to: each keeps its own.
          const pairs = new Map();
          for (const r of chosen) {
            const a = String(r.from?.address || '').toLowerCase();
            if (!a.includes('@')) continue;
            const entry = what === 'safeDomain' ? '@' + a.split('@').pop() : a;
            pairs.set(`${r.accountId}|${entry}`, { accountId: r.accountId, address: entry });
          }
          for (const p of pairs.values()) await shell.mail.listSender({ ...p, list: what === 'block' ? 'blocked' : 'safe' });
          const inJunk = (r) => foldersFor(r.accountId).find((f) => f.path === r.folder)?.role === 'junk' || (unified && role === 'junk');
          if (what === 'block') await overRows(chosen.filter((r) => !inJunk(r)), (g) => shell.mail.move({ ...g, to: foldersFor(g.accountId).find((f) => f.role === 'junk')?.path || 'Junk' }));
          else await overRows(chosen.filter(inJunk), (g) => shell.mail.notJunk(g));
          toast(
            entries.length === 1
              ? what === 'block' ? t('mail.sender-added-blocked', { sender: entries[0] }) : t('mail.sender-added-safe', { sender: entries[0] })
              : what === 'block' ? tn(entries.length, 'mail.count-sender-added-blocked') : tn(entries.length, 'mail.count-sender-added'),
            { tone: 'good' }
          );
        } else if (what === 'safeGroup') {
          // Never block this group or mailing list: the address the mail went
          // to kept as a Safe Recipient, and anything of it in Junk put back.
          const added = new Set();
          let refused = null;
          await overRows(chosen, async (g) => {
            try {
              const r = await shell.mail.listRecipient(g);
              for (const a of r.added || []) added.add(a);
            } catch (err) {
              refused = err.message;
            }
          });
          if (!added.size) { toast(refused || t('mail.nothing-to-keep'), { tone: 'bad' }); return; }
          const inJunk = (r) => foldersFor(r.accountId).find((f) => f.path === r.folder)?.role === 'junk' || (unified && role === 'junk');
          await overRows(chosen.filter(inJunk), (g) => shell.mail.notJunk(g));
          toast(added.size === 1 ? t('mail.recipient-added-safe', { recipient: [...added][0] }) : tn(added.size, 'mail.count-group-added'), { tone: 'good' });
        } else if (what === 'move') {
          await overRows(chosen, (g) => shell.mail.move({ ...g, to: extra }));
        } else if (what === 'read') {
          for (const r of chosen) handMarked.current.add(keyOf(r));
          await overRows(chosen, (g) => shell.mail.flag({ ...g, patch: { unread: false } }));
        } else if (what === 'unread') {
          for (const r of chosen) handMarked.current.add(keyOf(r));
          await overRows(chosen, (g) => shell.mail.flag({ ...g, patch: { unread: true } }));
        } else if (what === 'flag') {
          const on = !chosen.every((r) => r.flagged);
          await overRows(chosen, (g) => shell.mail.flag({ ...g, patch: { flagged: on } }));
          // The message open says so at once too, before the list is fetched again.
          if (chosen.some((r) => keyOf(r) === shownForRef.current)) setMessage((m) => (m ? { ...m, flagged: on } : m));
        } else if (what === 'pin') {
          const on = !chosen.every((r) => r.pinned);
          await overRows(chosen, (g) => shell.mail.flag({ ...g, patch: { pinned: on } }));
          if (chosen.some((r) => keyOf(r) === shownForRef.current)) setMessage((m) => (m ? { ...m, pinned: on } : m));
        }
        setChecked(new Set());
        if (!alone && !stays && (LEAVES.includes(what) || (what === 'safeGroup' && chosen.some((r) => foldersFor(r.accountId).find((f) => f.path === r.folder)?.role === 'junk' || (unified && role === 'junk'))))) setSelected(after);
        // Moved out of its folder, the message's own window has nothing left to show.
        if (alone && LEAVES.includes(what)) {
          shell.win.close({}).catch(() => {});
          return;
        }
        if (LEAVES.includes(what)) {
          const gone = new Set(chosen.map(keyOf));
          setList((current) => ({ ...current, rows: current.rows.filter((r) => !gone.has(keyOf(r))), total: Math.max(0, (current.total ?? current.rows.length) - gone.size) }));
          quietSearch.current = true;
        }
        await refreshList();
        await loadAccounts();
        setFolderTick((n) => n + 1);
      } catch (err) {
        toast(err.message, { tone: 'bad' });
        // Refused part way: the folder as it now is.
        await refreshList().catch(() => {});
      }
    },
    [alone, targets, overRows, foldersFor, unified, role, shell, refreshList, loadAccounts, toast, threads, allRowsIn, selected]
  );

  /** A double-click in the list: the message in a window of its own. */
  const popOut = useCallback(
    (row) => shell.win.create({ app: 'mail', query: { message: JSON.stringify({ accountId: row.accountId, folder: row.folder, id: row.id }) } }).catch((err) => toast(err.message, { tone: 'bad' })),
    [shell, toast]
  );

  /*
   * Copy, Cut and Paste of messages (Mail 2.25): Ctrl+C or Ctrl+X on what is
   * selected in the list, then Ctrl+V on a folder in the tree, or with that
   * folder open; the same on the right-click menus. What is copied is held in
   * this window until it is pasted (a cut, once) or something else is copied.
   * Within an account the server copies or moves them; into another account
   * they are added there, and a cut then takes them from where they were.
   */
  const [clip, setClip] = useState(null);
  const copyMessages = useCallback((mode, given = null) => {
    const rows = (given || targets()).map((r) => ({ accountId: r.accountId, folder: r.folder, id: r.id }));
    if (!rows.length) return;
    setClip({ mode, rows });
    toast(mode === 'cut'
      ? tn(rows.length, 'mail.count-message-cut')
      : tn(rows.length, 'mail.count-message-copied'));
  }, [targets, toast]);
  const pasteInto = useCallback(async (to) => {
    if (!clip || !to || unified || archiveId || !accountId) return;
    const groups = new Map();
    for (const r of clip.rows) {
      const k = `${r.accountId}\u0000${r.folder}`;
      if (!groups.has(k)) groups.set(k, { accountId: r.accountId, folder: r.folder, ids: [] });
      groups.get(k).ids.push(r.id);
    }
    let copied = 0;
    let moved = 0;
    try {
      for (const g of groups.values()) {
        const done = await shell.mail.paste({ ...g, to: { accountId, folder: to }, cut: clip.mode === 'cut' });
        copied += done.copied || 0;
        moved += done.moved || 0;
      }
      if (clip.mode === 'cut') setClip(null);
      const name = folders.find((f) => f.path === to)?.name || to.split('/').pop();
      const count = copied + (copied ? 0 : moved);
      toast(count
        ? tn(count, 'mail.count-message-pasted', { folder: name })
        : t('mail.nothing-paste-folder', { folder: name }), { tone: count ? 'good' : undefined });
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
    await refreshList().catch(() => {});
    await loadAccounts().catch(() => {});
    setFolderTick((n) => n + 1);
  }, [clip, unified, archiveId, accountId, shell, folders, toast, refreshList, loadAccounts]);
  /** The message list or the folder tree has the keys, and no words are selected to copy instead. */
  const listHasKeys = () => {
    const el = document.activeElement;
    if (!el?.closest?.('.ml-list, .ml-folder-tree')) return false;
    const words = window.getSelection?.();
    return !words || words.isCollapsed;
  };
  /** The folder Ctrl+V pastes into: the one in the tree with the keys, else the one open. */
  const pasteTarget = () => document.activeElement?.closest?.('.ml-drop')?.dataset.folder || folder;
  // Files copied in Explorer after messages were copied here are what Ctrl+V
  // pastes: the window was left to copy them. Until it is left, the messages.
  const clipAt = useRef(0);
  const leftAt = useRef(0);
  useEffect(() => {
    if (clip) clipAt.current = Date.now();
  }, [clip]);
  useEffect(() => {
    const left = () => { leftAt.current = Date.now(); };
    window.addEventListener('blur', left);
    return () => window.removeEventListener('blur', left);
  }, []);
  const clipIsNewest = () => Boolean(clip) && leftAt.current <= clipAt.current;

  /* ── .eml files: brought in, saved out (Mail 2.25) ──────────────────── */

  /** .eml files dropped, pasted or chosen, brought into `to` of the account open. Answers false when none of them is one. */
  const importFiles = useCallback(async (given, to, { paths: chosen = null } = {}) => {
    const sources = chosen ? { paths: chosen, files: [], count: chosen.length } : await emlSources(given);
    if (!sources) return false;
    if (!to || unified || archiveId || !accountId) {
      toast(t('mail.open-folder-account'), { tone: 'bad' });
      return true;
    }
    try {
      const r = await shell.mail.importEml({ accountId, folder: to, paths: sources.paths, files: sources.files });
      const name = folders.find((f) => f.path === to)?.name || String(to).split('/').pop();
      if (r.added) toast(tn(r.added, 'mail.count-message-added', { folder: name }), { tone: 'good' });
      if (r.skipped) toast(tn(r.skipped, 'mail.count-file-not'), { tone: 'bad', ms: 6000 });
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
    await refreshList().catch(() => {});
    await loadAccounts().catch(() => {});
    setFolderTick((n) => n + 1);
    return true;
  }, [unified, archiveId, accountId, shell, folders, toast, refreshList, loadAccounts]);

  /** File, Open: .eml files chosen, brought into the folder open. */
  const openEmlFiles = useCallback(async () => {
    const paths = await shell.dialog.open({
      title: t('mail.open-messages'),
      multiple: true,
      filters: [{ name: t('mail.email-messages-eml'), extensions: ['eml'] }, { name: t('common.all-files'), extensions: ['*'] }],
    });
    if (paths?.length) await importFiles(null, folder, { paths });
  }, [shell, importFiles, folder]);

  /** Save as .eml: one message to a file named by its subject, several each to a file in a folder chosen. */
  const saveAsEml = useCallback(async (given = null) => {
    const rows = (given || targets()).filter((r) => r && r.id != null);
    if (!rows.length) return;
    let target = null;
    if (rows.length === 1) {
      target = await shell.dialog.save({ title: t('mail.save-as-eml'), defaultPath: emlNameFor(rows[0].subject != null ? rows[0] : message), filters: [{ name: t('mail.email-messages-eml'), extensions: ['eml'] }] });
    } else {
      target = (await shell.dialog.open({ title: t('mail.save-messages-eml'), directory: true }))?.[0] || null;
    }
    if (!target) return;
    const groups = new Map();
    for (const r of rows) {
      const k = `${r.accountId}\u0000${r.folder}`;
      if (!groups.has(k)) groups.set(k, { accountId: r.accountId, folder: r.folder, ids: [] });
      groups.get(k).ids.push(r.id);
    }
    let written = 0;
    try {
      for (const g of groups.values()) written += (await shell.mail.saveEml({ ...g, path: target })).written.length;
      toast(tn(written, 'mail.count-message-saved'), { tone: 'good' });
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
  }, [targets, shell, toast, message]);

  fileMenu.current = { open: openEmlFiles, saveAs: saveAsEml };
  const chooseMailApp = useCallback(async () => {
    try {
      const r = await shell.defaults.mailApp({});
      toast(tFilled(r.message), { tone: r.opened ? 'plain' : r.ours ? 'good' : 'plain', ms: 8000 });
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
  }, [shell, toast]);
  chooseMailAppRef.current = chooseMailApp;

  // A message opened from a file: Add to a folder, any folder of any account, brings the file in.
  const [openedTargets, setOpenedTargets] = useState({});
  useEffect(() => {
    if (!opened) return undefined;
    let live = true;
    Promise.all(accounts.map((a) => shell.mail.folders({ accountId: a.id }).then((f) => [a.id, f || []], () => [a.id, []])))
      .then((pairs) => live && setOpenedTargets(Object.fromEntries(pairs)));
    return () => { live = false; };
  }, [opened, accounts, shell]);
  const addOpenedTo = useCallback(async (account, f) => {
    try {
      const r = await shell.mail.importEml({ accountId: account.id, folder: f.path, paths: [opened.path] });
      if (r.added) toast(tn(r.added, 'mail.count-message-added', { folder: `${account.name || account.email} · ${f.name || f.path}` }), { tone: 'good' });
      else toast(t('common.file-not-mail'), { tone: 'bad' });
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
  }, [shell, opened, toast]);
  const addOpenedMenu = (e) => menu.open(e, accounts.length
    ? accounts.map((a) => ({ label: a.name || a.email, icon: a.local ? 'archive' : 'mail', items: (openedTargets[a.id] || []).slice(0, 40).map((f) => ({ label: f.name || f.path, icon: f.icon || 'folder', run: () => addOpenedTo(a, f) })) }))
    : [{ label: t('mail.add-account-first'), disabled: true }]);

  // Files copied in Explorer and pasted with the list or a folder holding the keys.
  const pasteFilesRef = useRef(null);
  pasteFilesRef.current = (files) => importFiles(files, pasteTarget());
  useEffect(() => {
    const onPaste = (e) => {
      const el = document.activeElement;
      if (el && (el.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(el.tagName))) return;
      const files = [...(e.clipboardData?.files || [])];
      if (files.some((f) => /\.eml$/i.test(f.name))) {
        e.preventDefault();
        pasteFilesRef.current?.(files);
      } else if (clipRef.current) {
        // Left the window and copied something else that is not mail: the messages are still the ones to paste.
        e.preventDefault();
        clipRef.current();
      }
    };
    document.addEventListener('paste', onPaste);
    return () => document.removeEventListener('paste', onPaste);
  }, []);
  const clipRef = useRef(null);
  clipRef.current = clip ? () => (listHasKeys() ? pasteInto(pasteTarget()) : null) : null;

  const clipItems = () => [
    { label: t('common.copy'), icon: 'copy', key: prettyKey('Mod+C'), disabled: !targets().length, run: () => copyMessages('copy') },
    { label: t('common.cut'), icon: 'cut', key: prettyKey('Mod+X'), disabled: !targets().length || Boolean(archiveId), run: () => copyMessages('cut') },
    { label: t('common.paste'), icon: 'paste', key: prettyKey('Mod+V'), disabled: !clip || unified || Boolean(archiveId), run: () => pasteInto(folder) },
  ];

  /** Star one message from the list without opening it. */
  const star = useCallback(
    async (row, event) => {
      event.stopPropagation();
      if (isArchiveId(row.accountId)) return;
      await shell.mail.flag({ accountId: row.accountId, folder: row.folder, ids: [row.id], patch: { flagged: !row.flagged } });
      setList((current) => ({
        ...current,
        rows: current.rows.map((r) => (r.id === row.id && r.folder === row.folder ? { ...r, flagged: !r.flagged } : r)),
      }));
    },
    [shell]
  );

  const sync = useCallback(async () => {
    const ids = unified ? accounts.map((a) => a.id) : [accountId].filter(Boolean);
    if (!ids.length) return;
    setBusy(true);
    let added = 0;
    let localOnly = 0;
    try {
      for (const id of ids) {
        const result = await shell.mail.sync({ accountId: id, folder: id === accountId ? folder : undefined });
        if (result.local) localOnly++;
        else added += result.added || 0;
      }
      toast(
        localOnly === ids.length
          ? t('mail.imported-archives')
          : tn(added, 'mail.count-new-message'),
        { tone: localOnly === ids.length ? 'plain' : 'good' }
      );
      await refreshList();
      await loadAccounts();
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 6000 });
    } finally {
      setBusy(false);
      setProgress(null);
    }
  }, [unified, accounts, accountId, folder, shell, refreshList, loadAccounts, toast]);

  /* ── writing ──────────────────────────────────────────────────────────── */

  // Each account keeps its own signature; the account a message is going out
  // from is the one whose text belongs in it, not whichever account happens
  // to be selected in the sidebar.
  const signatureFor = useCallback((id) => accounts.find((a) => a.id === id)?.signature || '', [accounts]);
  // The rich one for the rich editor: an account with only a plain
  // signature gets that, its lines kept.
  const signatureHtmlFor = useCallback((id) => {
    const a = accounts.find((x) => x.id === id);
    return a?.signatureHtml || signatureTextToHtml(a?.signature || '');
  }, [accounts]);

  const quoted = useCallback((m) => {
    const body = m.text || stripTags(m.html || '');
    return `\n\nOn ${formatWhen(m.date, { long: true })}, ${displayName(m.from?.[0])} wrote:\n${body
      .split('\n')
      .map((l) => `> ${l}`)
      .join('\n')}`;
  }, []);

  // What Reply and Forward take: the message the reading pane shows. With the pane off nothing is shown, so it is
  // the row selected, once it has loaded: never the message loaded before it.
  const answerable = Boolean(message) && (alone || prefs.layout !== 'off' || Boolean(openMessage(selected, shownFor)));
  const reply = useCallback(
    (all = false) => {
      if (!answerable) return;
      const fromId = sendingFrom(shownAt?.accountId || accountId);
      const me = accounts.find((a) => a.id === fromId)?.email?.toLowerCase();
      const to = [message.replyTo?.[0]?.address || message.from?.[0]?.address].filter(Boolean);
      const cc = all
        ? [...(message.to || []), ...(message.cc || [])]
            .map((p) => p.address)
            .filter((a) => a && a.toLowerCase() !== me && !to.includes(a))
        : [];
      openForm({
        to: to.join(', '),
        cc: [...new Set(cc)].join(', '),
        subject: /^re:/i.test(message.subject || '') ? message.subject : `Re: ${message.subject || ''}`,
        text: withSignature(quoted(message), signatureFor(fromId), { reply: true }),
        html: withSignatureHtml(textToHtml(quoted(message)), signatureHtmlFor(fromId), { reply: true }),
        inReplyTo: message.messageId,
        accountId: fromId || undefined,
      });
    },
    [message, answerable, accounts, shownAt, accountId, signatureFor, signatureHtmlFor, quoted, newMailFrom]
  );

  const forward = useCallback(() => {
    if (!answerable) return;
    const fromId = sendingFrom(shownAt?.accountId || accountId);
    const body = `\n\n---------- Forwarded message ----------\nFrom: ${displayName(message.from?.[0])}\nDate: ${formatWhen(message.date, { long: true })}\nSubject: ${message.subject || ''}\nTo: ${(message.to || []).map((t) => t.address).join(', ')}\n\n${message.text || stripTags(message.html || '')}`;
    openForm({
      to: '',
      subject: /^fwd:/i.test(message.subject || '') ? message.subject : `Fwd: ${message.subject || ''}`,
      text: withSignature(body, signatureFor(fromId), { reply: true }),
      html: withSignatureHtml(textToHtml(body), signatureHtmlFor(fromId), { reply: true }),
      accountId: fromId || undefined,
      // Forwarding carries the files. They are already on disk here, so this
      // costs nothing until the message is actually sent.
      attachments: (message.attachments || []).filter((a) => !a.inline && a.stored).map((a) => ({ filename: a.filename, size: a.size, fromMessage: { ...shownAt, index: message.attachments.indexOf(a) } })),
    });
  }, [message, answerable, shownAt, accountId, signatureFor, signatureHtmlFor, accounts, newMailFrom]);

  const doSend = useCallback(
    async (draft, at) => {
      const from = draft.accountId || accountId;
      if (!from || from === EVERYTHING) return toast(t('mail.choose-account-send'), { tone: 'bad' });
      if (draft.smime?.encrypt && String(draft.bcc || '').trim()) {
        return toast(t('mail.encrypted-mail-not'), { tone: 'bad', ms: 9000 });
      }
      if (draft.smime?.encrypt) {
        const addresses = [draft.to, draft.cc].join(',').split(',').map((s) => (/<([^>]+)>/.exec(s)?.[1] || s).trim()).filter(Boolean);
        const can = await shell.mail.canEncryptFor({ addresses });
        if (!can.ok) return toast(t('mail.cannot-encrypted-names', { names: can.missing.join(', ') }), { tone: 'bad' });
      }
      try {
        // Attachments forwarded from another message are held by the backend;
        // they are written to a temporary file so SMTP can stream them.
        const attachments = [];
        for (const a of draft.attachments || []) {
          if (a.path) {
            attachments.push(a);
            continue;
          }
          if (!a.fromMessage) continue;
          const held = await shell.mail.attachment({
            accountId: a.fromMessage.accountId,
            folder: a.fromMessage.folder,
            id: a.fromMessage.id,
            index: a.fromMessage.index,
          });
          if (!held) continue;
          const bytes = new Uint8Array(await (await fetch(held.url)).arrayBuffer());
          const { path } = await shell.fs.temp({ ext: attachmentExtension(a.filename), bytes });
          attachments.push({ filename: a.filename, path });
        }

        const hold = await shell.store.get({ key: 'mail.undoSeconds', fallback: 8 });
        const item = await shell.mail.queue({
          accountId: from,
          draft: { ...draft, attachments },
          at: at || null,
          holdSeconds: at ? 0 : Number(hold) || 0,
        });
        setOutbox(await shell.mail.outbox());
        setCompose(null);
        // A draft sent is no longer a draft.
        if (draft.replaces?.id) {
          await shell.mail.delete({ ...draft.replaces, ids: [draft.replaces.id] }).catch(() => {});
          refreshList();
          setSelected(null);
        }
        if (at) toast(t('mail.scheduled-for-when', { when: new Date(at).toLocaleString() }), { tone: 'good', ms: 6000 });
        return item;
      } catch (err) {
        toast(err.message, { tone: 'bad', ms: 7000 });
        return null;
      }
    },
    [accountId, shell, toast, refreshList]
  );

  const unsend = useCallback(
    async (id) => {
      const item = await shell.mail.unsend({ id });
      setOutbox(await shell.mail.outbox());
      if (item?.sending) toast(t('mail.message-being-sent'), { tone: 'bad', ms: 7000 });
      else if (item) openForm({ ...item.draft, accountId: item.accountId });
    },
    [shell, toast]
  );

  /** Send now: take it off the schedule and hand it to SMTP immediately. */
  const sendOutboxNow = useCallback(
    async (item) => {
      try {
        const done = await shell.mail.sendNow({ id: item.id });
        if (!done.sent) {
          toast(done.error ? t('mail.not-sent-yet-still', { error: done.error }) : t('mail.not-sent-yet'), { tone: 'bad', ms: 7000 });
          return;
        }
        toast(t('mail.sent-to-to', { to: item.draft?.to || t('mail.recipient') }), { tone: 'good' });
        refreshList();
        setFolderTick((n) => n + 1);
      } catch (err) {
        toast(err.message, { tone: 'bad', ms: 7000 });
      } finally {
        setOutbox(await shell.mail.outbox());
      }
    },
    [shell, toast, refreshList]
  );

  /** Cancel a scheduled send. The words are not thrown away — they go to Drafts. */
  const cancelOutboxItem = useCallback(
    async (item) => {
      const back = await shell.mail.unsend({ id: item.id, toDrafts: true });
      setOutbox(await shell.mail.outbox());
      if (back?.sending) {
        toast(t('mail.message-being-sent'), { tone: 'bad', ms: 7000 });
        return;
      }
      toast(t('mail.moved-to-drafts'), { tone: 'good' });
      refreshList();
    },
    [shell, toast, refreshList]
  );

  /* ── attachments ──────────────────────────────────────────────────────── */

  const saveAttachment = useCallback(
    async (index) => {
      if (!shownAt || !message) return;
      const list = index === 'all' ? (message.attachments || []).map((a, i) => i).filter((i) => !message.attachments[i].inline) : [index];
      if (index === 'all') {
        const dirs = await shell.dialog.open({ title: t('mail.save-all-attachments'), directory: true });
        if (!dirs?.[0]) return;
        for (const i of list) {
          const held = await shell.mail.attachment({ ...shownAt, index: i });
          if (!held) continue;
          const bytes = new Uint8Array(await (await fetch(held.url)).arrayBuffer());
          await shell.fs.write({ path: await freeNameIn(dirs[0], held.name, (p) => shell.fs.exists({ path: p })), bytes, fromInternet: true });
        }
        toast(tn(list.length, 'mail.saved-count-file'), { tone: 'good' });
        return;
      }
      const held = await shell.mail.attachment({ ...shownAt, index });
      if (!held) return toast(t('mail.attachment-not-stored'), { tone: 'bad' });
      const target = await shell.dialog.save({ title: t('mail.save-attachment'), defaultPath: held.name });
      if (!target) return;
      const bytes = new Uint8Array(await (await fetch(held.url)).arrayBuffer());
      await shell.fs.write({ path: target, bytes, fromInternet: true });
      toast(t('common.saved-name', { name: held.name }), { tone: 'good' });
    },
    [shownAt, message, shell, toast]
  );

  /** Open an attachment in whichever Rutba Office app owns it. */
  /** Whether a folder holds drafts: its role, the unified Drafts, or its name where a server gave no role. */
  const isDraftsFolder = (acc, path) => (unified && role === 'drafts') || foldersFor(acc).find((x) => x.path === path)?.role === 'drafts' || /(^|[/.])drafts$/i.test(String(path || ''));

  /**
   * A draft opened in Compose, whole: its recipients, subject, words and
   * attachments (written to temporary files, as a forward's are), with the
   * draft it came from, which saving or sending replaces.
   */
  /**
   * Leaving what is being written for something else: what was changed is kept
   * in Drafts, as Save draft keeps it, and so is a new message holding
   * attachments, which nothing else would bring back. False if it could not be kept.
   */
  const keepWriting = useCallback(async (writing) => {
    if (!writing?.changed && !(writing && !writing.draft.replaces && writing.draft.attachments?.length)) return true;
    try {
      await shell.mail.saveDraft({ accountId: writing.draft.accountId || accountId, draft: writing.draft, replaces: writing.draft.replaces || null });
      toast(t('mail.saved-to-drafts'), { tone: 'good' });
      refreshList();
      return true;
    } catch (err) {
      toast(err.message, { tone: 'bad' });
      return false;
    }
  }, [shell, accountId, refreshList, toast]);
  keepWritingNow.current = keepWriting;

  const leaveForm = useCallback(async () => {
    const writing = composeOpen.current?.();
    if (!writing) return;
    // A form Read Aloud keeps out of sight is shown again before it is saved, closed or asked about.
    if (readingStopsFor({ reading: readingNow.current, leaving: true, hidden: formHidden.current })) stopReadingRef.current?.();
    const asked = (editAsked.current += 1);
    if (!(await keepFormOnce(composeKey.current, writing))) return false;
    if (asked === editAsked.current) setCompose(null);
    return true;
  }, [keepFormOnce, setCompose]);
  const leaveFormRef = useRef(null);
  leaveFormRef.current = leaveForm;

  const editDraft = useCallback(async (where, m) => {
    const writing = composeOpen.current?.();
    const from = writing?.draft.replaces;
    // Already in the form (the message was read again): what is being typed stays.
    if (from && from.accountId === where.accountId && from.folder === where.folder && from.id === where.id) return;
    const asked = (editAsked.current += 1);
    if (!(await keepFormOnce(composeKey.current, writing))) return;
    // Another message was clicked while that was being kept.
    if (asked !== editAsked.current) return;
    const list = (v) => (Array.isArray(v) ? v : v ? [v] : []).map((x) => (typeof x === 'string' ? x : x?.name ? `${x.name} <${x.address}>` : x?.address || '')).filter(Boolean).join(', ');
    const attachments = [];
    let html = m.html || null;
    for (const [index, a] of (m.attachments || []).entries()) {
      if (a.inline) {
        // A picture in the words: its place in the page goes on holding it.
        html = html && a.contentId ? await pictureInline(html, a, () => shell.mail.attachment({ ...where, index })) : html;
        continue;
      }
      try {
        const held = await shell.mail.attachment({ ...where, index });
        if (!held) continue;
        const bytes = new Uint8Array(await (await fetch(held.url)).arrayBuffer());
        const { path } = await shell.fs.temp({ ext: attachmentExtension(a.filename || held.name), bytes });
        attachments.push({ filename: a.filename || held.name, path, size: bytes.length });
      } catch {
        /* an attachment that cannot be read is left out, and the rest kept */
      }
    }
    // Another draft was clicked while this one's attachments were being read.
    if (asked !== editAsked.current) return;
    const id = String(m.messageId || '').startsWith('draft-') ? String(m.messageId).slice(6) : undefined;
    setCompose({
      id,
      to: list(m.to),
      cc: list(m.cc),
      bcc: m.draftBcc || '',
      subject: m.subject || '',
      text: m.text || '',
      html: html ? safeDraftHtml(html) : null,
      attachments,
      accountId: isArchiveId(where.accountId) ? undefined : where.accountId,
      replaces: isArchiveId(where.accountId) ? null : where,
    });
  }, [shell, keepFormOnce, setCompose]);
  const editDraftRef = useRef(null);
  editDraftRef.current = editDraft;

  const openAttachment = useCallback(
    async (index, meta, where = shownAt) => {
      const held = await shell.mail.attachment({ ...where, index });
      if (!held) return toast(t('mail.attachment-not-stored'), { tone: 'bad' });
      const bytes = new Uint8Array(await (await fetch(held.url)).arrayBuffer());
      const name = String(meta?.filename || held.name);
      const { path } = await shell.fs.temp({ ext: attachmentExtension(name), bytes, fromInternet: true });
      const owner = appFor(kindFromExtension(name));
      // Anything the suite understands opens in the app that owns it; anything
      // else goes to whatever this computer uses for that kind of file.
      if (owner) return shell.win.create({ app: owner, file: path });
      // For a program, "whatever this computer uses" is the computer itself:
      // the main process warns before it runs one (office-shell's openPath).
      await shell.shell.openPath({ path });
    },
    [shownAt, shell, toast]
  );

  /* ── importing ────────────────────────────────────────────────────────── */

  const importFrom = useCallback(
    async (target) => {
      setDialog({ kind: 'importing', path: target });
      try {
        const found = await shell.mail.importScan({ path: target });
        setDialog({ kind: 'import-preview', path: target, found });
      } catch (err) {
        toast(err.message, { tone: 'bad' });
        setDialog(null);
      }
    },
    [shell, toast]
  );

  const chooseImport = useCallback(async () => {
    const paths = await shell.dialog.open({
      title: t('common.import-mail'),
      filters: [
        { name: t('mail.mail-archives'), extensions: ['pst', 'ost', 'olm', 'mbox', 'mbx', 'eml', 'emlx', 'msg'] },
        { name: t('common.all-files'), extensions: ['*'] },
      ],
    });
    if (paths[0]) importFrom(paths[0]);
  }, [shell, importFrom]);

  /* ── archives opened in place, and refining a search ────────────────── */

  const openArchiveAt = useCallback(
    async (target) => {
      setDialog(null);
      try {
        const a = await shell.mail.openArchive({ path: target });
        await loadArchives();
        setAccountId(a.id);
        setView('mail');
        toast(t('mail.name-open-lies', { name: a.name }), { tone: 'good', ms: 6000 });
      } catch (err) {
        toast(err.message, { tone: 'bad', ms: 8000 });
      }
    },
    [shell, loadArchives, toast]
  );

  const chooseArchive = useCallback(async () => {
    const paths = await shell.dialog.open({
      title: t('mail.open-archive'),
      filters: [
        { name: t('mail.mail-archives'), extensions: ['pst', 'ost', 'olm', 'mbox', 'mbx'] },
        { name: t('common.all-files'), extensions: ['*'] },
      ],
    });
    if (paths[0]) openArchiveAt(paths[0]);
  }, [shell, openArchiveAt]);

  const closeArchive = useCallback(
    async (id) => {
      await shell.mail.closeArchive({ id });
      await loadArchives();
      setAccountId((current) => (current === id ? accounts[0]?.id || null : current));
    },
    [shell, loadArchives, accounts]
  );

  /** Add an operator to the search box and put the cursor after it, as Outlook's Refine buttons do. */
  const refine = useCallback((op) => {
    setQuery((q) => `${q.trim() ? `${q.trim()} ` : ''}${op}`);
    setTimeout(() => {
      const input = document.querySelector('.rw-search input');
      if (!input) return;
      input.focus();
      input.setSelectionRange(input.value.length, input.value.length);
    }, 0);
  }, []);

  const dateMenu = (event) => {
    const year = new Date().getFullYear();
    menu.open(event, [
      { label: t('common.today'), run: () => refine('received:today') },
      { label: t('mail.yesterday'), run: () => refine('received:yesterday') },
      { label: t('mail.this-week'), run: () => refine('newer_than:7d') },
      { label: t('mail.this-month'), run: () => refine('newer_than:1m') },
      { label: t('mail.this-year'), run: () => refine(`received:${year}`) },
      { label: t('mail.last-year'), run: () => refine(`received:${year - 1}`) },
      { label: t('mail.older-year'), run: () => refine('older_than:1y') },
      '-',
      { label: t('mail.after-a-day'), run: () => refine('after:') },
      { label: t('mail.before-a-day'), run: () => refine('before:') },
    ]);
  };

  const sizeMenu = (event) =>
    menu.open(event, [
      { label: t('mail.larger-1-mb'), run: () => refine('size>1mb') },
      { label: t('mail.larger-5-mb'), run: () => refine('size>5mb') },
      { label: t('mail.larger-10-mb'), run: () => refine('size>10mb') },
      { label: t('mail.smaller-100-kb'), run: () => refine('size<100kb') },
    ]);

  // A file of accounts: the picker, the service, and what came of it.
  const importAccountsFrom = useCallback(async (chosen = null) => {
    const paths = typeof chosen === 'string' ? [chosen] : await shell.dialog.open({
      title: t('mail.set-accounts-file'),
      filters: [
        { name: t('mail.accounts-json'), extensions: ['json'] },
        { name: t('common.all-files'), extensions: ['*'] },
      ],
    });
    if (!paths[0]) return;
    setDialog(null);
    try {
      const r = await shell.mail.importAccounts({ path: paths[0], test: true });
      await loadAccounts();
      if (r.added.length) setAccountId(r.added[0].id);
      const parts = [];
      if (r.added.length) parts.push(tn(r.added.length, 'mail.count-set-up'));
      if (r.skipped.length) parts.push(tn(r.skipped.length, 'mail.count-left'));
      if (r.failed.length) parts.push(tn(r.failed.length, 'mail.count-could-not', { addresses: r.failed.map((f) => f.email).join(', ') }));
      toast(parts.length ? `${parts.join('; ')}.` : t('mail.no-accounts-file'), { tone: r.failed.length ? 'bad' : r.added.length ? 'good' : 'plain', ms: 8000 });
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 8000 });
    }
  }, [shell, loadAccounts, toast]);

  const runImport = useCallback(
    async (target, chosen) => {
      setDialog({ kind: 'importing', path: target });
      try {
        const result = await shell.mail.import({ path: target, folders: chosen });
        await loadAccounts();
        setAccountId(result.accountId);
        // A file imported before goes into the same mailbox, and only what it did not hold is added.
        if (result.again) toast(t('mail.file-imported-before', { messages: tn(result.messages, 'mail.count-new-message') }), { tone: 'good', ms: 6000 });
        else toast(t('mail.imported-messages', { messages: tn(result.messages, 'mail.count-message'), folders: tn(result.folders, 'common.count-folder') }), { tone: 'good', ms: 6000 });
      } catch (err) {
        toast(err.message, { tone: 'bad', ms: 7000 });
      } finally {
        setDialog(null);
        setProgress(null);
      }
    },
    [shell, loadAccounts, toast]
  );

  // .eml files go into the folder open (Mail 2.25); anything else is a mailbox to import.
  useFileDrop(useCallback((files) => {
    if (files.length && files.every((p) => /.eml$/i.test(p)) && folder && !unified && !archiveId && accountId) importFiles(null, folder, { paths: files });
    else if (files[0]) importFrom(files[0]);
  }, [importFrom, importFiles, folder, unified, archiveId, accountId]));

  // A calendar part in the open message is the calendar's card (calendar-invite-card.js), mounted in the reader.

  const keepSender = useCallback(
    async (from) => {
      if (!from?.address) return;
      try {
        const kept = await shell.contacts.fromMail({ name: from.name || '', email: from.address });
        toast(t('mail.kept-name-contacts', { name: kept?.display || from.address }), { tone: 'good' });
      } catch (err) {
        toast(err.message, { tone: 'bad' });
      }
    },
    [shell, toast]
  );

  /**
   * An account taken off (D-2026-10-22-09): asked first, its mail on this
   * computer deleted with it or kept, from the ribbon, the account's own menu
   * or the account manager alike.
   */
  const removeAccount = useCallback(async (a) => {
    if (!a) return;
    const { response } = await shell.dialog.message({
      type: 'warning',
      message: t('mail.remove-address', { address: a.email }),
      detail: t('mail.mail-stored-computer'),
      buttons: [t('common.remove'), t('mail.keep-the-mail'), t('common.cancel')],
      cancelId: 2,
    });
    if (response === 2) return;
    try {
      await shell.mail.removeAccount({ id: a.id, keepMail: response === 1 });
      if (a.id === accountId) setAccountId(null);
      loadAccounts();
      toast(t('mail.address-is-removed', { address: a.email }), { tone: 'good' });
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 6000 });
    }
  }, [shell, accountId, loadAccounts, toast]);
  const accountItems = (a) => [
    { label: t('mail.signature'), run: () => setDialog({ kind: 'signature', accountId: a.id }) },
    { label: t('mail.out-of-office'), run: () => setDialog({ kind: 'ooo', accountId: a.id }) },
    '-',
    { label: t('mail.remove-account'), run: () => removeAccount(a) },
  ];

  /** The message being written, as it stands, handed to a Mail window of its own. */
  const popOutCompose = useCallback(async (draft) => {
    const key = `mail.handoff.${Date.now()}.${Math.random().toString(36).slice(2, 8)}`;
    try {
      await shell.store.set({ key, value: draft });
      await shell.win.create({ app: 'mail', query: { compose: key } });
      setCompose(null);
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
  }, [shell, toast]);

  // Opened by another app with a message to write: Contacts' Send mail, or
  // the calendar's reply or invitation with its file attached.
  useEffect(() => {
    const q = new URLSearchParams(window.location.search);
    if (!q.get('to') && !q.get('attach')) return;
    const attach = q.get('attach');
    setCompose({ to: q.get('to') || '', subject: q.get('subject') || '', text: '', attachments: attach ? [{ filename: attach.split(/[\\/]/).pop(), path: attach, size: Number(q.get('size')) || 0 }] : [] });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // The scan runs when the import dialog opens, not at start-up: it touches the
  // disk, and an application that rummages through your profile before you have
  // asked it for anything is not one to be trusted with your mail.
  useEffect(() => {
    if (dialog?.kind !== 'import' || scan) return;
    shell.discover.scan().then(setScan).catch(() => setScan({ accounts: [], files: [], scanned: [] }));
  }, [dialog, scan, shell]);

  const unsubscribe = useCallback(
    async (offer) => {
      const { response } = await shell.dialog.message({
        type: 'question',
        message: t('mail.unsubscribe-from-list', { list: offer.list || t('mail.this-list') }),
        detail:
          offer.method === 'one-click'
            ? t('mail.sender-supports-one')
            : offer.method === 'web'
              ? t('mail.opens-host-browser', { host: new URL(offer.http).hostname })
              : t('mail.sends-message-address', { address: offer.address }),
        buttons: [t('mail.unsubscribe'), t('common.cancel')],
        cancelId: 1,
      });
      if (response !== 0) return;

      if (offer.method === 'one-click') {
        // The one request the question named, sent from here: not the browser.
        const host = new URL(offer.http).hostname;
        const r = shownAt ? await shell.mail.unsubscribe({ accountId: shownAt.accountId, folder: shownAt.folder, id: shownAt.id }).catch(() => ({ done: false, reason: 'unreachable' })) : { done: false, reason: 'not-offered' };
        if (r.done) toast(t('mail.unsubscribe-request-sent', { host }), { tone: 'good' });
        else if (r.reason === 'refused') toast(t('mail.host-did-not', { host }), { tone: 'bad', ms: 8000 });
        else if (r.reason === 'unreachable') toast(t('mail.host-could-not', { host }), { tone: 'bad', ms: 8000 });
        else toast(t('mail.message-does-not'), { tone: 'bad' });
      } else if (offer.method === 'email' || (!offer.http && offer.address)) {
        await doSend({ to: offer.address, subject: offer.subject, text: 'Please remove this address from the list.' }, null);
        toast(t('mail.unsubscribe-request'), { tone: 'good' });
      } else {
        await shell.shell.openExternal({ url: offer.http });
      }
    },
    [shell, doSend, toast, shownAt]
  );

  /* ── commands ─────────────────────────────────────────────────────────── */

  const commands = useMemo(
    () => ({
      // F5, the key every mail client fetches on; Ctrl+R is Reply, and one
      // key cannot do both.
      'mail.sync': { label: t('mail.get-mail'), icon: 'refresh', key: 'F5', run: sync },
      'mail.compose': {
        label: t('mail.new-message'),
        icon: 'new',
        key: 'Mod+N',
        run: () => {
          const fromId = newMailFrom;
          openForm({ to: '', subject: '', text: withSignature('', signatureFor(fromId)), html: withSignatureHtml('', signatureHtmlFor(fromId)) || null, accountId: fromId || undefined });
        },
      },
      'mail.reply': { label: t('common.reply'), icon: 'reply', key: 'Mod+R', run: () => reply(false) },
      'mail.replyAll': { label: t('mail.reply-all'), icon: 'replyAll', run: () => reply(true) },
      'mail.forward': { label: t('common.forward'), icon: 'forward', run: forward },
      'mail.archive': { label: t('mail.archive'), icon: 'archive', key: 'e', run: () => act('archive') },
      'mail.delete': { label: t('common.delete'), icon: 'trash', key: 'Delete', run: () => act('delete') },
      // Only while the list or the folder tree has the keys: anywhere else Ctrl+C, X and V copy and paste words.
      'mail.copyMessages': { label: t('common.copy'), icon: 'copy', key: 'Mod+C', when: listHasKeys, run: () => copyMessages('copy') },
      'mail.cutMessages': { label: t('common.cut'), icon: 'cut', key: 'Mod+X', when: () => listHasKeys() && !archiveId, run: () => copyMessages('cut') },
      'mail.pasteMessages': { label: t('common.paste'), icon: 'paste', key: 'Mod+V', when: () => listHasKeys() && clipIsNewest() && !unified && !archiveId, run: () => pasteInto(pasteTarget()) },
      'mail.openEml': { label: t('mail.open-eml-files'), icon: 'folderOpen', key: 'Mod+O', run: () => openEmlFiles() },
      'mail.saveEml': { label: t('mail.save-as-eml-2'), icon: 'save', key: 'Mod+Shift+S', when: () => !compose, run: () => saveAsEml() },
      'mail.junk': { label: t('mail.move-to-junk'), icon: 'spam', run: () => act('junk') },
      'mail.flag': { label: t('common.star'), icon: 'star', run: () => act('flag') },
      'mail.pin': { label: t('common.pin-top'), icon: 'flag', run: () => act('pin') },
      'mail.read': { label: t('mail.mark-read'), icon: 'check', run: () => act('read') },
      'mail.unread': { label: t('mail.mark-unread'), icon: 'eye', run: () => act('unread') },
      'mail.import': { label: t('mail.import-mail'), icon: 'import', run: () => setDialog({ kind: 'import' }) },
      'mail.account': { label: t('mail.add-account'), icon: 'plus', run: () => setDialog({ kind: 'account' }) },
      'mail.search': { label: t('common.search'), icon: 'find', key: 'Mod+F', run: () => document.querySelector('.rw-search input')?.focus() },
    }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [sync, reply, forward, act, accountId, accounts, signatureFor, signatureHtmlFor, copyMessages, pasteInto, clip, unified, archiveId, folder, openEmlFiles, saveAsEml, compose]
  );

  useCommands(commands, [selected, message, folder, accountId, checked]);
  const commandsRef = useRef(commands);
  commandsRef.current = commands;

  // Reply, Reply all or Forward from a row's right-click menu, for a row that was not open: the form being
  // written is left first (kept in Drafts, as a click on a row does), the row is opened, and the command runs
  // once the reading pane shows it, on that message (afterOpenStep). Another message chosen, something else
  // started in the form meanwhile, or the row would not open: nothing runs.
  const afterOpen = useRef(null);
  useEffect(() => {
    const step = afterOpenStep(afterOpen.current, { selectedKey: selected ? rowKey(selected) : null, shownFor, loaded: Boolean(message), form: composeKey.current });
    if (step === 'drop') afterOpen.current = null;
    if (step !== 'run') return;
    const { id } = afterOpen.current;
    afterOpen.current = null;
    commandsRef.current[id]?.run();
  }, [selected, shownFor, message, compose]);
  const openThen = async (open, id) => {
    afterOpen.current = null;
    // Refused (the draft could not be kept): the form stays, and nothing else opens.
    if ((await leaveForm()) === false) return;
    afterOpen.current = { key: rowKey(open), id, form: composeKey.current };
    setSelected({ accountId: open.accountId, folder: open.folder, id: open.id });
  };

  /** A row's right-click menu: on the message open as the ribbon is, or on the row pointed at (menuRows). */
  const rowMenu = (thread, e) => {
    const items = (ids) => menuItems(commands, ids);
    const pointed = menuRows(thread, { open: openMessage(selected, shownFor), checked, threaded: prefs.threaded, alone, readOnly: isArchiveId, rowsIn: allRowsIn });
    if (!pointed) {
      menu.open(e, [...items(['mail.reply', 'mail.replyAll', 'mail.forward', '-', 'mail.archive', 'mail.flag', 'mail.pin', 'mail.unread', '-', 'mail.junk', 'mail.delete']), '-', ...clipItems(), '-', ...items(['mail.saveEml'])]);
      return;
    }
    const { rows, all, open } = pointed;
    const none = !rows.length;
    const on = (id, run, disabled = false) => ({ ...items([id])[0], disabled, run });
    // A draft opens to be finished, not answered: no Reply or Forward on a row in Drafts.
    const draft = isDraftsFolder(open.accountId, open.folder);
    const opening = (id) => on(id, () => openThen(open, id), draft);
    menu.open(e, [
      opening('mail.reply'), opening('mail.replyAll'), opening('mail.forward'), '-',
      on('mail.archive', () => act('archive', null, rows), none || Boolean(archiveId)),
      on('mail.flag', () => act('flag', null, rows), none),
      on('mail.pin', () => act('pin', null, rows), none),
      on('mail.unread', () => act('unread', null, rows), none), '-',
      on('mail.junk', () => act('junk', null, rows), none || Boolean(archiveId)),
      on('mail.delete', () => act('delete', null, rows), none || Boolean(archiveId)), '-',
      { label: t('common.copy'), icon: 'copy', key: prettyKey('Mod+C'), run: () => copyMessages('copy', all) },
      { label: t('common.cut'), icon: 'cut', key: prettyKey('Mod+X'), disabled: none || Boolean(archiveId), run: () => copyMessages('cut', rows) },
      clipItems()[2], '-',
      on('mail.saveEml', () => saveAsEml(all), Boolean(compose)),
    ]);
  };
  // Ctrl and the wheel over the message, Ctrl+Plus, Ctrl+Minus and Ctrl+0:
  // the reading pane's own zoom, the one View's Out and In buttons set.
  useWheelZoom({ shell, zoom: prefs.zoom ?? 1, setZoom: (z) => setPref({ zoom: z }), min: 0.5, max: 3, area: '.ml-reader' });

  // Single-key navigation, the way every mail client that respects your hands
  // does it. Suppressed whenever a field has focus, so typing "j" in a search
  // box searches for j.
  useEffect(() => {
    const onKey = (e) => {
      const el = document.activeElement;
      if (e.metaKey || e.ctrlKey || e.altKey) return;
      if (el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.isContentEditable)) return;
      // A message's own window has no list to move through: the arrows scroll the message.
      if (alone && ['j', 'k', 'u', 'ArrowDown', 'ArrowUp'].includes(e.key)) return;
      const at = threads.findIndex((t) => selected && allRowsIn(t).some((r) => keyOf(r) === keyOf(selected)));
      const go = (delta) => {
        const next = threads[Math.max(0, Math.min(threads.length - 1, (at < 0 ? -1 : at) + delta))];
        if (next) setSelected({ accountId: next.accountId, folder: next.folder, id: next.id });
      };
      if (e.key === 'j' || e.key === 'ArrowDown') { e.preventDefault(); go(1); }
      else if (e.key === 'k' || e.key === 'ArrowUp') { e.preventDefault(); go(-1); }
      else if (e.key === 'u') setSelected(null);
      else if (e.key === 's') act('flag');
      else if (e.key === 'r') reply(false);
      else if (e.key === 'a') reply(true);
      else if (e.key === 'f') forward();
      else if (e.key === '!') act('junk');
      else if (e.key === '/') { e.preventDefault(); document.querySelector('.rw-search input')?.focus(); }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [alone, threads, selected, allRowsIn, act, reply, forward]);

  /* ── chrome ───────────────────────────────────────────────────────────── */

  const account = accounts.find((a) => a.id === accountId);
  const current = folders.find((f) => f.path === folder);
  const unreadTotal = accounts.reduce((n, a) => n + (a.counts?.unread || 0), 0);
  const title = unified ? t('mail.all-accounts') : account ? `${current?.name || t('common.mail')} \u2014 ${account.name || account.email}` : t('common.mail');

  const moveMenu = useCallback(
    (event) =>
      menu.open(
        event,
        folders
          .filter((f) => f.path !== folder)
          .slice(0, 24)
          .map((f) => ({ label: f.name || f.path, icon: f.icon || 'folder', run: () => act('move', f.path) }))
      ),
    [menu, folders, folder, act]
  );

  // Folders made, renamed and deleted here are so at the server too, as
  // Outlook's are. The account's own folders (the Inbox, Sent, Drafts,
  // Trash and Junk) keep their names.
  const canChangeFolders = Boolean(account) && !unified && !archiveId;
  const keptFolder = (f) => ['inbox', 'sent', 'drafts', 'trash', 'junk'].includes(f?.role) || /^inbox$/i.test(f?.path || '');
  const saveFolderName = useCallback(
    async (name) => {
      const asked = dialog;
      const delimiter = folders.find((f) => f.delimiter)?.delimiter || '/';
      if (name.includes(delimiter)) {
        toast(t('mail.folder-name-cannot', { character: delimiter }), { tone: 'bad' });
        return;
      }
      const parent = asked.folder ? asked.folder.path.split(delimiter).slice(0, -1).join(delimiter) : asked.parent?.path || '';
      const path = parent ? `${parent}${delimiter}${name}` : name;
      if (path !== asked.folder?.path && folders.some((f) => f.path.toLowerCase() === path.toLowerCase())) {
        toast(t('mail.folder-called-name', { name }), { tone: 'bad' });
        return;
      }
      try {
        if (asked.folder) {
          const r = await shell.mail.renameFolder({ accountId, folder: asked.folder.path, name });
          if (folder === asked.folder.path || String(folder || '').startsWith(asked.folder.path + delimiter)) setFolder(r.path + String(folder).slice(asked.folder.path.length));
        } else {
          await shell.mail.createFolder({ accountId, parent: asked.parent?.path || null, name });
        }
        setDialog(null);
        setFolderTick((n) => n + 1);
      } catch (err) {
        toast(err.message, { tone: 'bad' });
      }
    },
    [dialog, folders, folder, accountId, shell, toast]
  );
  const deleteFolderAsk = useCallback(
    async (f) => {
      if (!f || keptFolder(f)) return;
      const delimiter = f.delimiter || '/';
      const inside = folders.filter((x) => x.path === f.path || x.path.startsWith(f.path + delimiter));
      const messages = inside.reduce((n, x) => n + (x.total || 0), 0);
      const { response } = await shell.dialog.message({
        type: 'warning',
        message: t('mail.delete-folder-folder', { folder: f.name || f.path }),
        detail: [
          inside.length > 1 ? tn(inside.length - 1, 'mail.count-folder-inside') : '',
          messages ? tn(messages, 'mail.count-message-deleted') : t('mail.deleted-server'),
        ].filter(Boolean).join(' '),
        buttons: [t('mail.delete'), t('common.cancel')],
        cancelId: 1,
      });
      if (response !== 0) return;
      try {
        await shell.mail.deleteFolder({ accountId, folder: f.path });
        toast(t('mail.folder-is-deleted', { folder: f.name || f.path }), { tone: 'good' });
        setFolderTick((n) => n + 1);
        loadAccounts();
      } catch (err) {
        toast(err.message, { tone: 'bad' });
      }
    },
    [folders, shell, accountId, toast, loadAccounts]
  );
  const folderItems = (f) => [
    { label: t('mail.new-folder-2'), icon: 'plus', run: () => setDialog({ kind: 'folderName', parent: f }) },
    { label: t('mail.rename-folder-2'), disabled: keptFolder(f), run: () => setDialog({ kind: 'folderName', folder: f }) },
    { label: t('mail.delete-folder'), icon: 'trash', disabled: keptFolder(f), run: () => deleteFolderAsk(f) },
    '-',
    { label: t('mail.paste-here'), icon: 'paste', key: prettyKey('Mod+V'), disabled: !clip, run: () => pasteInto(f.path) },
  ];

  // Messages dragged from the list onto a folder are moved there, as Move to
  // does: at the server too. Only within the account shown.
  const dragged = useRef(null);
  const [dropOn, setDropOn] = useState(null);
  const startDrag = useCallback(
    (thread, e) => {
      const rows = allRowsIn(thread);
      const inChecked = checkedRows.length && rows.every((r) => checked.has(keyOf(r)));
      const moving = (inChecked ? checkedRows : rows).filter((r) => !isArchiveId(r.accountId));
      if (!canChangeFolders || !moving.length) {
        e.preventDefault();
        return;
      }
      dragged.current = moving;
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', tn(moving.length, 'mail.count-message'));
    },
    [allRowsIn, checkedRows, checked, canChangeFolders]
  );
  const dropProps = (f) => ({
    onDragOver: (e) => {
      const rows = dragged.current;
      // .eml files from Explorer, brought into this folder (Mail 2.25).
      if (!rows?.length && carriesFiles(e)) {
        e.preventDefault();
        e.stopPropagation();
        e.dataTransfer.dropEffect = 'copy';
        if (dropOn !== f.path) setDropOn(f.path);
        return;
      }
      if (!rows?.length || !rows.some((r) => r.folder !== f.path)) return;
      e.preventDefault();
      // Not the window's own file drop, which would take it for a mailbox to import.
      e.stopPropagation();
      e.dataTransfer.dropEffect = 'move';
      if (dropOn !== f.path) setDropOn(f.path);
    },
    onDragLeave: (e) => {
      if (!e.currentTarget.contains(e.relatedTarget)) setDropOn((p) => (p === f.path ? null : p));
    },
    onDrop: (e) => {
      const rows = dragged.current;
      dragged.current = null;
      setDropOn(null);
      if (!rows?.length && carriesFiles(e)) {
        e.preventDefault();
        e.stopPropagation();
        const files = [...e.dataTransfer.files];
        // Not a message among them: the window's own drop, a mailbox to import.
        importFiles(files, f.path).then((took) => !took && files[0] && pathOf(files[0]) && importFrom(pathOf(files[0])));
        return;
      }
      if (!rows?.length) return;
      e.preventDefault();
      e.stopPropagation();
      act('move', f.path, rows.filter((r) => r.folder !== f.path));
    },
  });

  // Home → Junk, Outlook's menu: mark, unmark, the two lists and the options.
  const inJunk = unified ? role === 'junk' : current?.role === 'junk';

  // A link followed in the message, which the frame hands to the window
  // (owner, 2026-10-16: "it should open browser instead while spam should
  // have a protection on links"): a web page opens in the browser, with a
  // warning first when the message is junk, failed its sender's checks, or
  // the link's words name another site than it goes to; a mailto: address
  // is written to here; anything else is not opened.
  // A mailto: link: a message to write, filled from it; nothing is sent by itself.
  // From the message read, or from Windows, Office Mail being the mail app it
  // opens such links with (Mail 2.25): the window opened on it, or the Mail
  // window already open told of it.
  const writeToLink = useCallback((url) => {
    const f = mailtoFields(url);
    openForm({ to: f.to, cc: f.cc, bcc: f.bcc, subject: f.subject, text: f.text });
  }, [openForm]);
  useEffect(() => {
    const asked = new URLSearchParams(location.search).get('mailto');
    if (asked) writeToLink(asked);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  useEffect(() => shell.on('app:link', (q) => q?.mailto && writeToLink(q.mailto)), [shell, writeToLink]);

  const linkFacts = useRef({});
  linkFacts.current = { message, inJunk, auth: insight?.auth };
  const followLink = useCallback(async (asked) => {
    const url = fromFrameUrl(asked);
    const { message: m, inJunk: junk, auth } = linkFacts.current;
    const kind = linkKind(url);
    if (kind === 'refused') {
      toast(t('mail.link-not-opened'), { tone: 'bad', ms: 5000 });
      return;
    }
    if (kind === 'mail') {
      writeToLink(url);
      return;
    }
    const reasons = linkRisks(url, { html: m?.html || '', inJunk: junk, authFailed: Boolean(auth?.known && !auth.ok) });
    if (reasons.length) {
      const { response } = await shell.dialog.message({
        type: 'warning',
        message: t('mail.open-link-browser'),
        detail: `${reasons.join('\n')}\n\n${t('mail.it-goes-to')}\n${url}`,
        buttons: [t('mail.open-in-browser'), t('mail.copy-link'), t('common.cancel')],
        defaultId: 2,
        cancelId: 2,
      });
      if (response === 1) {
        await shell.clipboard.writeText({ text: url }).catch(() => {});
        toast(t('mail.link-copied'), { ms: 2500 });
      }
      if (response !== 0) return;
    }
    shell.shell.openExternal({ url }).catch((err) => toast(t('mail.link-could-not', { reason: err.message }), { tone: 'bad', ms: 6000 }));
  }, [shell, toast, writeToLink]);
  useEffect(() => shell.on('win:link', ({ url } = {}) => followLink(url)), [shell, followLink]);

  // Safety (safety.js): on for every message as it opens. Turned off, it is off for the message the
  // reading pane shows, held against that message: one clicked meanwhile arrives with Safety on.
  const [safetyOffFor, setSafetyOffFor] = useState(null);
  const safety = !(safetyOffFor && safetyOffFor === shownFor);
  // Stable, over a ref: the message's own right-click menu keeps the one it was given when the window opened.
  const shownForRef = useRef(shownFor);
  shownForRef.current = shownFor;
  const setSafety = useCallback((on) => setSafetyOffFor(on ? null : shownForRef.current), []);
  const safetyRef = useRef(true);
  safetyRef.current = safety;
  useEffect(() => { setSafetyOffFor((off) => (off === shownFor ? off : null)); }, [shownFor]);
  const safetyFirst = useCallback(() => toast(t('mail.safety-message-turn'), { ms: 4500 }), [toast]);
  const [sourceView, setSourceView] = useState(null);
  const messageRef = useRef(null);
  messageRef.current = message;

  // A right-click in the message: Mail's own menu, as Safety allows.
  useEffect(() => shell.on('win:frameMenu', (at = {}) => {
    const items = frameMenuItems({
      at: { ...at, link: at.link ? fromFrameUrl(at.link) : '' },
      safe: safetyRef.current,
      onCopy: () => shell.win.frameAction({ action: 'copy' }).catch(() => {}),
      onOpenLink: (url) => followLink(url),
      onCopyLink: async (url) => { await shell.clipboard.writeText({ text: url }); toast(t('mail.link-copied'), { ms: 2500 }); },
      onCopyPicture: async (p) => {
        try {
          const { bytes } = await pictureBytes({ shell, src: p.src, message: messageRef.current });
          await shell.clipboard.writeImage({ bytes });
          toast(t('mail.picture-copied'), { ms: 2500 });
        } catch (err) {
          toast(err?.message || String(err), { tone: 'bad', ms: 6000 });
        }
      },
      onSavePicture: async (src) => {
        try {
          const saved = await savePicture({ shell, src, message: messageRef.current });
          if (saved) toast(t('common.saved-name', { name: saved.split(/[\\/]/).pop() }), { tone: 'good', ms: 3000 });
        } catch (err) {
          toast(err?.message || String(err), { tone: 'bad', ms: 6000 });
        }
      },
      onSafetyOff: () => setSafety(false),
    });
    if (!items.length) return;
    menu.open({ preventDefault() {}, stopPropagation() {}, currentTarget: null, target: null, type: 'contextmenu', clientX: at.x || 1, clientY: at.y || 1 }, items);
  }), [shell, toast, followLink, menu.open]);

  const junkMenu = (event) => {
    const none = noMessage;
    menu.open(event, [
      { label: t('mail.mark-as-junk'), icon: 'spam', disabled: none || inJunk, run: () => act('junk') },
      { label: t('mail.not-junk'), icon: 'inbox', disabled: none || !inJunk, run: () => act('notJunk') },
      '-',
      { label: t('mail.block-sender'), disabled: none, run: () => act('block') },
      { label: t('mail.never-block-sender'), disabled: none, run: () => act('safe') },
      { label: t('mail.never-block-senders'), disabled: none, run: () => act('safeDomain') },
      { label: t('mail.never-block-group'), disabled: none, run: () => act('safeGroup') },
      '-',
      { label: t('mail.junk-email-options'), icon: 'settings', run: () => setDialog({ kind: 'junk' }) },
    ]);
  };

  // Where the message being written goes: in a window of its own, the whole
  // window; docked where there is a reading pane (or the message read alone in
  // its window), else the whole window.
  const placeNow = composeAlone ? 'window' : composePlace === 'docked' && (alone || prefs.layout !== 'off') ? 'docked' : 'full';
  const composeView = compose ? (
    <Compose
      key={composeKey.current}
      handle={composeOpen}
      draft={compose}
      accounts={accounts}
      accountId={compose.accountId || newMailFrom}
      onAccount={(id) => updateCompose((d) => ({ ...d, accountId: id }))}
      onChange={updateCompose}
      onClose={() => {
        setCompose(null);
        // A draft's own place goes with it. A reply or new message let go goes back to what was being read.
        if (compose.replaces && !alone) setSelected(null);
      }}
      onShut={async () => {
        if ((await leaveForm()) && !alone) setSelected(null);
      }}
      onSend={doSend}
      place={placeNow}
      onPlace={composeAlone ? null : (next) => { setComposePlace(next); shell.store.set({ key: 'mail.composePlace', value: next }).catch(() => {}); }}
      onPopOut={composeAlone ? null : popOutCompose}
      onSaveDraft={async (draft) => {
        await shell.mail.saveDraft({ accountId: draft.accountId || accountId, draft, replaces: draft.replaces || null });
        toast(t('mail.saved-to-drafts'), { tone: 'good' });
        setCompose(null);
        // Saving a draft replaces it, so the place does not go back to the one it was.
        if (draft.replaces && !alone) setSelected(null);
        refreshList();
      }}
      shell={shell}
      toast={toast}
      identities={identities}
      onCertificates={() => setDialog({ kind: 'certificates', start: 'import' })}
    />
  ) : null;

  // The reading pane's body: the message, or a docked form. While Read Aloud reads, the message being read is
  // shown and a docked form is kept as it was, out of sight, in the same place so nothing typed in it is lost.
  const docked = placeNow === 'docked' && composeView;
  const hasDocked = Boolean(docked);
  const hiddenForReading = Boolean(reading) && hasDocked;
  formHidden.current = hiddenForReading;
  // The reading over, the docked form it kept out of sight takes the keys again, to go on writing.
  const wasHidden = useRef(false);
  useEffect(() => {
    if (wasHidden.current && !hiddenForReading && hasDocked) {
      const form = document.querySelector('.ml-docked-form');
      const active = document.activeElement;
      const idle = !active || active === document.body;
      if (refocusAfterReading({ asked: stopAsked.current, idle, inPane: Boolean(form?.parentElement?.contains(active)) })) (form?.querySelector('.ml-rich') || form?.querySelector('.ml-compose-body'))?.focus();
    }
    if (wasHidden.current !== hiddenForReading) stopAsked.current = false;
    wasHidden.current = hiddenForReading;
  }, [hiddenForReading, hasDocked]);
  const paneBody = (fallback) => (docked ? (
    <>
      <div className="ml-docked-form" style={{ display: reading ? 'none' : 'contents' }}>{composeView}</div>
      {reading ? readerView || notOpened || fallback : null}
    </>
  ) : readerView || notOpened || fallback);

  // The message selected that would not open: said so where it would have been.
  const notOpened = openError && selected && shownFor === rowKey(selected) && !message ? (
    <Empty icon="mail" title={t('mail.message-could-not')}>{openError}</Empty>
  ) : null;

  // The message open: in the reading pane, or alone in its own window.
  const readingBar = reading ? (
    <div className="ml-read-aloud">
      {reading.count > 1 ? (
        <div className="ml-read-aloud-list" role="toolbar" aria-label={t('mail.messages-being-read')}>
          <Button ghost icon="skipBack" title={t('mail.the-message-before')} disabled={reading.n <= 1} onClick={() => playerRef.current?.previous()} />
          <span>{t('mail.message-n-count', { n: reading.n, count: reading.count })}</span>
          <Button ghost icon="skipForward" title={t('mail.the-next-message')} disabled={reading.n >= reading.count} onClick={() => playerRef.current?.next()} />
          <span style={{ flex: 1 }} />
          <VoiceOffer shell={shell} voices={addedVoices} />
        </div>
      ) : null}
      <ReadAloudBar reader={speaker()} state={readState} onClose={stopReadingAsked} />
      {reading.count > 1 ? null : <span className="ml-read-aloud-offer"><VoiceOffer shell={shell} voices={addedVoices} /></span>}
    </div>
  ) : null;

  const readerView = message ? (
    <Reader
      message={message}
      zoom={prefs.zoom ?? 1}
      insight={insight}
      inviteCard={shell.calendar ? <InviteCard shell={shell} message={message} where={shownAt} onCompose={setCompose} /> : null}
      onKeepSender={keepSender}
      remote={remote}
      plain={plain}
      dark={dark}
      onRemote={() => setRemote(true)}
      onSaveAttachment={(index) => (safety ? safetyFirst() : saveAttachment(index))}
      onOpenAttachment={(index, meta) => {
        // Handed to another program only with Safety off; Office's own apps open it either way.
        if (safety && !appFor(kindFromExtension(String(meta?.filename || '')))) return safetyFirst();
        return openAttachment(index, meta);
      }}
      safety={safety}
      onSafety={setSafety}
      onBlocked={safetyFirst}
      onClose={alone ? null : () => setSelected(null)}
      onSource={async () => {
        setSourceView({ loading: true });
        try {
          setSourceView(await shell.mail.source({ accountId: shownAt.accountId, folder: shownAt.folder, id: shownAt.id }));
        } catch (err) {
          setSourceView(null);
          toast(err?.message || String(err), { tone: 'bad', ms: 6000 });
        }
      }}
      onReply={() => reply(false)}
      onReplyAll={() => reply(true)}
      onForward={forward}
      onUnsubscribe={unsubscribe}
      onSender={(from) => from?.address && setQuery(from.address)}
      inJunk={inJunk}
      onNotJunk={() => act('notJunk')}
      onSmimeImport={() => setDialog({ kind: 'certificates', start: 'import' })}
      onCertificates={() => setDialog({ kind: 'certificates' })}
      notes={<AttachmentNotes shell={shell} where={shownAt} notes={attachmentNotes} />}
    />
  ) : null;

  return (
    <AppFrame
      app={app}
      shell={shell}
      // A message's own window is named after the message.
      title={alone ? message?.subject || t('common.mail') : title}
      menu={appMenu}
      ribbon={
        <Ribbon
          tabs={[
            { id: 'home', label: t('common.home') },
            { id: 'search', label: t('common.search') },
            { id: 'send', label: t('mail.send-receive') },
            { id: 'folder', label: t('common.folder') },
            { id: 'view', label: t('common.view') },
            { id: 'tools', label: t('common.tools') },
          ]}
          active={tab}
          onTab={setTab}
          quick={
            <>
              <Button icon="refresh" title={t('mail.get-mail')} onClick={sync} disabled={!accountId} />
              <Button icon="new" title={t('mail.new-message')} onClick={() => commands['mail.compose'].run()} disabled={!accounts.length} />
              <Button icon="trash" title={t('common.delete')} onClick={() => act('delete')} disabled={noMessage} />
              <Button icon="star" title={t('common.star')} onClick={() => act('flag')} disabled={noMessage} />
            </>
          }
        >
          {tab === 'home' ? (
            <>
              <Group label={t('common.new')}>
                <Button tall icon="new" label={t('mail.message')} onClick={() => commands['mail.compose'].run()} disabled={!accounts.length} />
              </Group>
              <Group label={t('common.delete')}>
                <Button tall icon="trash" label={t('common.delete')} disabled={(noMessage) || Boolean(archiveId)} onClick={() => act('delete')} />
                <Button icon="archive" label={t('mail.archive')} disabled={(noMessage) || Boolean(archiveId)} onClick={() => act('archive')} />
                <Button icon="spam" label={t('mail.junk')} title={t('mail.junk-mark-block')} onClick={junkMenu} />
              </Group>
              <Group label={t('mail.respond')}>
                <Button tall icon="reply" label={t('common.reply')} disabled={!answerable} onClick={() => reply(false)} />
                <Button tall icon="replyAll" label={t('mail.reply-all')} disabled={!answerable} onClick={() => reply(true)} />
                <Button tall icon="forward" label={t('common.forward')} disabled={!answerable} onClick={forward} />
              </Group>
              <Group label={t('common.move')}>
                <Button tall icon="folder" label={t('mail.move-to')} disabled={(noMessage) || unified || Boolean(archiveId)} onClick={moveMenu} />
              </Group>
              <Group label={t('mail.tags')}>
                <Button icon="star" label={t('common.star')} disabled={noMessage} onClick={() => act('flag')} />
                <Button icon="flag" label={t('mail.pin')} disabled={noMessage} onClick={() => act('pin')} />
                <Button icon="check" label={t('common.read')} disabled={noMessage} onClick={() => act('read')} />
                <Button icon="eye" label={t('common.unread')} disabled={noMessage} onClick={() => act('unread')} />
              </Group>
              <Group label={t('word.ribbon.read-aloud-2')}>
                <Button tall icon="volume" label={t('word.ribbon.read-aloud-2')} pressed={buttons.open.pressed} disabled={buttons.open.disabled} title={t('mail.read-open-message')} onClick={() => readAloud('open')} />
                {buttons.all ? <Button icon="play" label={t('mail.play-all')} pressed={buttons.all.pressed} disabled={buttons.all.disabled} title={t('mail.read-every-message')} onClick={() => (reading?.mode === 'all' ? stopReadingAsked() : readAloud('all'))} /> : null}
                {buttons.unread ? <Button icon="mail" label={t('mail.play-unread')} pressed={buttons.unread.pressed} disabled={buttons.unread.disabled} title={t('mail.read-unread-messages')} onClick={() => (reading?.mode === 'unread' ? stopReadingAsked() : readAloud('unread'))} /> : null}
              </Group>
              <Group label={t('common.find')}>
                <Button tall icon="find" label={t('common.search')} onClick={() => commands['mail.search'].run()} />
                <Button icon="attach" label={t('common.attachments')} pressed={view === 'files'} onClick={() => setView(view === 'files' ? 'mail' : 'files')} disabled={unified || Boolean(archiveId)} />
                <Button icon="contacts" label={t('common.people')} pressed={view === 'people'} onClick={() => setView(view === 'people' ? 'mail' : 'people')} disabled={unified || Boolean(archiveId)} />
                <Button icon="filter" label={t('common.rules')} onClick={() => setDialog({ kind: 'rules' })} />
              </Group>
            </>
          ) : tab === 'search' ? (
            <>
              <Group label={t('mail.scope')}>
                <Button tall icon="folder" label={t('mail.this-folder')} title={t('mail.folder-search-only')} pressed={scope === 'folder'} onClick={() => setScope('folder')} />
                <Button tall icon="grid" label={t('mail.all-mail')} title={t('mail.all-mail-every')} pressed={scope === 'all'} onClick={() => setScope('all')} />
              </Group>
              <Group label={t('mail.refine')}>
                <Button icon="contacts" label={t('common.from')} title={t('mail.from-who-sent-it')} onClick={() => refine('from:')} />
                <Button icon="send" label={t('mail.sent-to')} title={t('mail.sent-went')} onClick={() => refine('to:')} />
                <Button icon="file" label={t('common.subject')} title={t('mail.subject-word-subject')} onClick={() => refine('subject:')} />
                <Button icon="attach" label={t('mail.has')} onClick={() => refine('has:attachment')} />
                <Button icon="find" label={t('mail.attachment')} title={t('mail.attachment-name-file')} onClick={() => refine('filename:')} />
                <Button icon="eye" label={t('common.unread')} onClick={() => refine('is:unread')} />
                <Button icon="calendar" label={t('common.date')} title={t('mail.date-received')} onClick={dateMenu} />
                <Button icon="sliders" label={t('common.size')} title={t('mail.size-larger-smaller')} onClick={sizeMenu} />
              </Group>
              <Group label={t('mail.archives')}>
                <Button tall icon="folderOpen" label={t('mail.open-archive')} title={t('mail.open-archive-pst')} onClick={chooseArchive} />
                <Button tall icon="close" label={t('mail.close')} disabled={!openArchive} onClick={() => openArchive && closeArchive(openArchive.id)} />
              </Group>
              <Group label={t('common.export')}>
                <Button tall icon="export" label={t('mail.export-results')} title={t('mail.export-results-messages')} disabled={!allMail || !list.total} onClick={() => setDialog({ kind: 'hold' })} />
                <Button tall icon="pdf" label={t('mail.export')} title={t('mail.export-folders-mbox')} disabled={!folder || unified} onClick={() => setDialog({ kind: 'hold' })} />
              </Group>
              <Group label={t('common.help')}>
                <Button tall icon="info" label={t('mail.search-syntax')} title={t('mail.search-syntax-every')} onClick={() => setDialog({ kind: 'search-help' })} />
              </Group>
            </>
          ) : tab === 'send' ? (
            <>
              <Group label={t('mail.send-and')}>
                <Button tall icon="refresh" label={unified ? t('mail.all-accounts') : t('mail.this-account')} onClick={sync} disabled={!accountId} />
                <Button icon="download" label={t('mail.this-folder')} disabled={!folder || unified} onClick={sync} />
              </Group>
              <Group label={t('mail.fetch-by')}>
                <Button
                  dropdown
                  icon="clock"
                  className="ml-fetch-every"
                  label={fetchWhen(fetching.every)}
                  title={t('mail.how-often-mail')}
                  onClick={(e) =>
                    menu.open(
                      e,
                      [1, 5, 10, 15, 30, 60, 0].map((every) => ({ label: fetchWhen(every), icon: fetching.every === every ? 'check' : undefined, run: () => changeFetching({ every }) }))
                    )
                  }
                />
                <Button
                  icon="mail"
                  className="ml-fetch-push"
                  label={t('mail.as-mail-arrives')}
                  pressed={fetching.push}
                  title={t('mail.mail-arrives-mail')}
                  onClick={() => changeFetching({ push: !fetching.push })}
                />
              </Group>
              <Group label={t('mail.outbox')}>
                <Button
                  tall
                  icon="clock"
                  label={outbox.length ? t('mail.waiting-count', { count: outbox.length }) : t('mail.nothing-waiting')}
                  disabled={!outbox.length}
                  onClick={() => setDialog({ kind: 'outbox' })}
                />
              </Group>
              <Group label={t('common.import')}>
                <Button tall icon="import" label={t('common.import-mail')} onClick={() => setDialog({ kind: 'import' })} />
              </Group>
              <Group label={t('common.export')}>
                <Button
                  tall
                  icon="export"
                  label="Mbox" // words-ok: a file format's name
                  disabled={!folder || unified || Boolean(archiveId)}
                  onClick={async () => {
                    const target = await shell.dialog.save({ title: t('mail.export-folder'), defaultPath: `${current?.name || t('mail.folder')}.mbox` });
                    if (!target) return;
                    const r = await shell.mail.export({ accountId, folder, path: target, format: 'mbox' });
                    if (r.missing) toast(t('mail.count-exported-missing', { count: r.messages, missing: r.missing }), { tone: 'warn' });
                    else toast(tn(r.messages, 'mail.exported-count-message'), { tone: 'good' });
                  }}
                />
                <Button tall icon="pdf" label={t('mail.for-legal')} title={t('mail.legal-hold-mbox')} disabled={!folder || unified} onClick={() => setDialog({ kind: 'hold' })} />
              </Group>
              <Group label={t('mail.archives')}>
                <Button tall icon="folderOpen" label={t('mail.open-archive')} title={t('mail.open-archive-pst')} onClick={chooseArchive} />
              </Group>
            </>
          ) : tab === 'folder' ? (
            <>
              <Group label={t('common.accounts')}>
                <Button tall icon="plus" label={t('common.add')} onClick={() => setDialog({ kind: 'account' })} />
                <Button
                  tall
                  icon="word"
                  label={t('common.signature')}
                  disabled={!account}
                  // Greyed in All accounts: said why, rather than left to be guessed.
                  title={account ? undefined : t('mail.signature-choose-one')}
                  onClick={() => setDialog({ kind: 'signature', accountId: account.id })}
                />
                <Button
                  tall
                  icon="clock"
                  label={account?.autoReply?.enabled ? t('mail.out-office') : t('mail.out-of')}
                  pressed={Boolean(account?.autoReply?.enabled)}
                  disabled={!account}
                  title={account ? undefined : t('mail.out-office-choose')}
                  onClick={() => setDialog({ kind: 'ooo', accountId: account.id })}
                />
                <Button
                  tall
                  icon="trash"
                  label={t('common.remove')}
                  disabled={!account}
                  title={account ? undefined : t('mail.remove-choose-account')}
                  onClick={() => removeAccount(account)}
                />
              </Group>
              <Group label={t('common.folders')}>
                <Button tall icon="plus" label={t('mail.new-folder')} disabled={!canChangeFolders} onClick={() => setDialog({ kind: 'folderName', parent: null })} />
                <Button icon="pen" label={t('mail.rename-folder')} disabled={!canChangeFolders || !current || keptFolder(current)} onClick={() => setDialog({ kind: 'folderName', folder: current })} />
                <Button icon="trash" label={t('mail.delete')} disabled={!canChangeFolders || !current || keptFolder(current)} onClick={() => deleteFolderAsk(current)} />
              </Group>
              <Group label={t('mail.this-folder')}>
                <Button
                  tall
                  icon="check"
                  label={t('mail.mark-all-read')}
                  disabled={!folder || unified}
                  onClick={async () => {
                    const r = await shell.mail.markAllRead({ accountId, folder });
                    toast(tn(r.changed, 'mail.count-marked-read'), { tone: 'good' });
                    refreshList();
                    loadAccounts();
                  }}
                />
                <Button
                  icon="trash"
                  label={t('mail.empty-folder')}
                  disabled={!current || !['trash', 'junk'].includes(current.role)}
                  onClick={async () => {
                    const { response } = await shell.dialog.message({
                      type: 'warning',
                      message: t('mail.permanently-delete', { folder: current.name }),
                      detail: t('mail.cannot-undone'),
                      buttons: [t('common.delete'), t('common.cancel')],
                      cancelId: 1,
                    });
                    if (response !== 0) return;
                    const r = await shell.mail.emptyFolder({ accountId, folder });
                    toast(tn(r.removed, 'mail.count-deleted'), { tone: 'good' });
                    refreshList();
                  }}
                />
              </Group>
            </>
          ) : tab === 'view' ? (
            <>
              <Group label={t('mail.arrangement')}>
                <Button icon="list" label={t('mail.conversations')} pressed={prefs.threaded} onClick={() => setPref({ threaded: !prefs.threaded })} />
                <Button dropdown
                  icon="sort"
                  label={t('mail.sort-order', { order: SORTS.find((s) => s.id === prefs.sort)?.label })}
                  onClick={(e) =>
                    menu.open(
                      e,
                      SORTS.map((s) => ({ label: s.label, icon: prefs.sort === s.id ? 'check' : undefined, run: () => setPref({ sort: s.id }) }))
                        .concat(['-', { label: prefs.ascending ? t('common.newest-first') : t('common.oldest-first'), icon: 'refresh', run: () => setPref({ ascending: !prefs.ascending }) }])
                    )
                  }
                />
                <Button icon="grid" label={prefs.density === 'compact' ? t('common.compact') : t('mail.comfortable')} onClick={() => setPref({ density: prefs.density === 'compact' ? 'cosy' : 'compact' })} />
              </Group>
              <Group label={t('mail.reading-pane')}>
                <Button icon="grid" label={t('common.right')} pressed={prefs.layout === 'right'} onClick={() => setPref({ layout: 'right' })} />
                <Button icon="list" label={t('common.bottom')} pressed={prefs.layout === 'bottom'} onClick={() => setPref({ layout: 'bottom' })} />
                <Button icon="close" label={t('common.off')} pressed={prefs.layout === 'off'} onClick={() => setPref({ layout: 'off' })} />
              </Group>
              <Group label={t('mail.message')}>
                <Button icon="eye" label={t('mail.remote')} pressed={remote} onClick={() => setRemote((r) => !r)} />
                <Button icon="file" label={t('common.plain-text')} pressed={plain} onClick={() => setPlain((p) => !p)} />
              </Group>
              <Group label={t('common.zoom')}>
                {/* The message alone is scaled; zooming the window took the list and the ribbon with it. */}
                <Button icon="zoomOut" label={t('common.out')} title={t('mail.zoom-out-message', { percent: Math.round((prefs.zoom ?? 1) * 100) })} onClick={() => setPref({ zoom: Math.max(0.5, Math.round(((prefs.zoom ?? 1) - 0.1) * 100) / 100) })} />
                <Button icon="zoomIn" label={t('common.in')} title={t('mail.zoom-message-percent', { percent: Math.round((prefs.zoom ?? 1) * 100) })} onClick={() => setPref({ zoom: Math.min(3, Math.round(((prefs.zoom ?? 1) + 0.1) * 100) / 100) })} />
                <Button icon="check" label="100%" disabled={Math.abs((prefs.zoom ?? 1) - 1) < 0.001} title={t('mail.100-message-own')} onClick={() => setPref({ zoom: 1 })} />
              </Group>
            </>
          ) : (
            <>
              <Group label={t('common.this')}>
                <Button tall icon="settings" label={t('mail.default-apps')} onClick={() => setDialog({ kind: 'defaults' })} />
                <Button tall icon="mail" label={t('mail.mail-app')} className="ml-mail-app" title={t('mail.make-office-mail-app')} onClick={chooseMailApp} />
                <Button tall icon="import" label={t('mail.find-my-mail')} onClick={() => setDialog({ kind: 'import' })} />
              </Group>
              <Group label={t('common.mail')}>
                <Button tall icon="filter" label={t('common.rules')} onClick={() => setDialog({ kind: 'rules' })} />
              </Group>
              <Group label={t('mail.security')}>
                <Button tall icon="shield" label={t('mail.certificates')} title={t('mail.certificates-sign-read')} onClick={() => setDialog({ kind: 'certificates' })} />
              </Group>
              <Group label={t('mail.sending')}>
                <Button
                  tall
                  icon="clock"
                  label={t('mail.undo-window')}
                  onClick={async () => {
                    const now = await shell.store.get({ key: 'mail.undoSeconds', fallback: 8 });
                    setDialog({ kind: 'undo', seconds: now });
                  }}
                />
                <Button
                  tall
                  icon="word"
                  label={t('common.signature')}
                  disabled={!accounts.length}
                  onClick={() => setDialog({ kind: 'signature', accountId: accountId && accountId !== EVERYTHING ? accountId : accounts[0]?.id })}
                />
              </Group>
              <Group label={t('common.privacy')}>
                <Button icon="shield" label={t('mail.always-block')} pressed disabled title={t('mail.remote-content-blocked')} />
              </Group>
            </>
          )}
        </Ribbon>
      }
      status={
        <>
          <span>{unified ? tn(accounts.length, 'mail.count-account') : openArchive ? openArchive.name : account ? account.email : t('mail.no-account-yet')}</span>
          {archives
            .filter((a) => a.index?.state === 'indexing')
            .slice(0, 1)
            .map((a) => (
              <span key={a.id} className="ml-status-indexing">
                <Spinner />
                {t('mail.indexing-name-percent', { name: a.name, percent: Math.floor((100 * (a.index.done || 0)) / Math.max(1, a.index.total || 1)) })}
              </span>
            ))}
          {progress ? (
            <>
              <Spinner />
              <span>
                {progress.phase === 'import' ? t('mail.importing-folder-done', { folder: progress.folder, done: progress.done.toLocaleString() }) : t('mail.fetching-folder-done', { folder: progress.folder, done: progress.done.toLocaleString() })}
              </span>
            </>
          ) : null}
          <Spacer />
          {checked.size ? <Chip>{tn(checked.size, 'mail.count-selected')}</Chip> : null}
          <Chip>{tn(list.total, 'mail.count-message')}</Chip>
          {unreadTotal ? <Chip>{tn(unreadTotal, 'common.count-unread')}</Chip> : null}
        </>
      }
    >
      {!unified && account?.autoReply?.enabled ? (
        <div className="ml-ooo-banner">
          <Icon name="info" size={14} />
          <span>{account.autoReply.where === 'server' ? t('mail.automatic-replies-mail', { address: account.email }) : t('mail.automatic-replies', { address: account.email })}</span>
          <button
            type="button"
            className="ml-ooo-off"
            onClick={async () => {
              try {
                if (account.autoReply.where === 'server') {
                  const r = await shell.mail.serverAutoReply({ accountId: account.id, action: 'off' });
                  if (!r.supported) throw new Error(t('mail.mail-server-could'));
                }
                await shell.mail.updateAccount({ id: account.id, patch: { autoReply: { ...account.autoReply, enabled: false } } });
                loadAccounts();
              } catch (err) {
                toast(err.message, { tone: 'bad', ms: 6000 });
              }
            }}
          >
            {t('common.turn-off')}
          </button>
        </div>
      ) : null}

      {composeAlone ? (
        <Content />
      ) : alone ? (
        <Content>
          {opened ? <OpenedBar name={opened.name} onAdd={addOpenedMenu} /> : null}
          {readingBar}
          {paneBody(
            <div style={{ padding: 24, display: 'grid', placeItems: 'center' }}>
              <Spinner />
            </div>
          )}
        </Content>
      ) : !accounts.length && !archives.length ? (
        <Empty icon="mail" title={t('mail.no-mail-yet')}>
          {t('mail.add-account-fetch')}
          <div style={{ display: 'flex', gap: 8, marginTop: 14 }}>
            <Button primary icon="plus" label={t('common.add-account')} onClick={() => setDialog({ kind: 'account' })} />
            <Button icon="import" label={t('mail.find-my-mail')} onClick={() => setDialog({ kind: 'import' })} />
            <Button icon="folderOpen" label={t('mail.open-archive')} onClick={chooseArchive} />
          </div>
        </Empty>
      ) : (
        <>
          <Panel width={222} resizable name="folders">
            <div className="ml-compose-cta">
              <button type="button" onClick={() => commands['mail.compose'].run()}>
                <Icon name="new" size={15} />
                <span>{t('mail.compose')}</span>
              </button>
            </div>

            <div className="ml-accounts">
              {accounts.length > 1 ? (
                <Item
                  icon="inbox"
                  label={t('mail.all-accounts')}
                  current={unified}
                  count={unreadTotal || undefined}
                  onClick={() => {
                    setAccountId(EVERYTHING);
                    setView('mail');
                  }}
                />
              ) : null}
              {accounts.map((a) => (
                <Item
                  key={a.id}
                  icon={a.local ? 'archive' : 'mail'}
                  label={a.name || a.email}
                  current={a.id === accountId}
                  count={a.counts?.unread || undefined}
                  onClick={() => setAccountId(a.id)}
                  onContextMenu={(e) => { e.preventDefault(); menu.open(e, accountItems(a)); }}
                  title={a.email}
                />
              ))}
            </div>

            <div className="rw-panel-head">{t('mail.archives')}</div>
            <div className="ml-accounts ml-archives">
              {archives.map((a) => (
                <Item
                  key={a.id}
                  icon="archive"
                  label={a.name}
                  current={a.id === accountId}
                  onClick={() => {
                    setAccountId(a.id);
                    setView('mail');
                  }}
                  title={a.missing ? t('mail.path-no-longer', { path: a.path }) : `${a.path}`}
                />
              ))}
              <Item icon="folderOpen" label={t('mail.open-archive-2')} onClick={chooseArchive} title={t('mail.open-archive-pst')} />
            </div>

            <div className="rw-panel-head">{unified ? t('mail.across-every-account') : t('common.folders')}</div>
            <div className="ml-folder-tree">
            <List>
              {unified
                ? [
                    { path: 'inbox', name: t('mail.inbox'), icon: 'inbox' },
                    { path: 'sent', name: t('mail.sent'), icon: 'send' },
                    { path: 'drafts', name: t('mail.drafts'), icon: 'file' },
                    { path: 'archive', name: t('mail.archive'), icon: 'archive' },
                    { path: 'junk', name: t('mail.junk'), icon: 'spam' },
                    { path: 'trash', name: t('common.trash'), icon: 'trash' },
                    { path: 'all', name: t('mail.everything'), icon: 'grid' },
                  ].map((f) => (
                    <Item key={f.path} icon={f.icon} label={f.name} current={role === f.path} onClick={() => setRole(f.path)} />
                  ))
                : folders.map((f) => (
                    <div key={f.path} className={`ml-drop${dropOn === f.path ? ' over' : ''}`} data-folder={f.path} {...(canChangeFolders ? dropProps(f) : {})}>
                      <Item
                        icon={f.icon || 'folder'}
                        label={f.name || f.path}
                        indent={Math.min(f.depth || 0, 4)}
                        current={f.path === folder && view === 'mail'}
                        count={f.unread || undefined}
                        onClick={() => {
                          setFolder(f.path);
                          setView('mail');
                        }}
                        onContextMenu={canChangeFolders ? (e) => { e.preventDefault(); menu.open(e, folderItems(f)); } : undefined}
                        title={f.path}
                      />
                    </div>
                  ))}
            </List>
            </div>

            {!unified && !archiveId ? (
              <>
                <div className="rw-panel-head">{t('common.views')}</div>
                <List>
                  <Item icon="attach" label={t('common.attachments')} current={view === 'files'} onClick={() => setView('files')} />
                  <Item icon="contacts" label={t('common.people')} current={view === 'people'} onClick={() => setView('people')} />
                  <Item
                    icon="clock"
                    label={t('mail.outbox')}
                    current={dialog?.kind === 'outbox'}
                    count={outbox.filter((o) => o.accountId === accountId).length || undefined}
                    onClick={() => setDialog({ kind: 'outbox' })}
                    title={t('mail.scheduled-waiting-send')}
                  />
                </List>
              </>
            ) : unified ? (
              <List>
                <Item
                  icon="clock"
                  label={t('mail.outbox')}
                  current={dialog?.kind === 'outbox'}
                  count={outbox.length || undefined}
                  onClick={() => setDialog({ kind: 'outbox' })}
                  title={t('mail.every-accounts-scheduled')}
                />
              </List>
            ) : null}
            {/* Add account at the foot of the pane, after the folders and views (owner, 2026-10-21): under
                the accounts it stood in the way of the folders once the first account was added. */}
            <div className="ml-add-account">
              <List>
                <Item icon="plus" label={t('common.add-account')} onClick={() => setDialog({ kind: 'account' })} />
                {accounts.length ? <Item icon="settings" label={t('mail.manage-accounts')} onClick={() => setDialog({ kind: 'accounts' })} /> : null}
              </List>
            </div>
          </Panel>

          {view === 'files' ? (
            <Content>
              <FilesView
                shell={shell}
                accountId={accountId}
                onOpen={(f) => openAttachment(f.index, f, { accountId, folder: f.folder, id: f.id })}
              />
            </Content>
          ) : view === 'people' ? (
            <Content>
              <PeopleView
                shell={shell}
                accountId={accountId}
                onPerson={(p) => {
                  setView('mail');
                  setQuery(p.address);
                }}
              />
            </Content>
          ) : (
            <>
              <Panel
                name="messages"
                width={prefs.layout === 'bottom' ? undefined : 380}
                resizable={prefs.layout !== 'bottom'}
                style={{
                  background: 'var(--surface)',
                  ...(prefs.layout === 'bottom' ? { flex: 1 } : {}),
                  ...(prefs.layout === 'off' ? { flex: 1 } : {}),
                }}
              >
                {openArchive ? (
                  <ArchiveBanner
                    archive={openArchive}
                    onClose={() => closeArchive(openArchive.id)}
                    onExport={() => setDialog({ kind: 'hold' })}
                    onStopIndexing={() => shell.mail.stopIndexing({ id: openArchive.id }).catch(() => {})}
                    onResumeIndexing={() => shell.mail.resumeIndexing({ id: openArchive.id }).catch(() => {})}
                  />
                ) : null}
                {/* With the reading pane off, the Read Aloud bar sits above the list. */}
                {prefs.layout === 'off' ? readingBar : null}
                <div className="ml-tools">
                  <Search value={query} onChange={setQuery} placeholder={scope === 'all' ? t('mail.search-all-mail') : unified ? t('mail.search-every-account') : t('mail.search-folder', { folder: current?.name || t('mail.mail') })} />
                  <div className="ml-filters">
                    {FILTERS.filter((f) => f.id !== 'people').map((f) => (
                      <button key={f.id} type="button" className={`ml-filter${prefs.filter === f.id ? ' on' : ''}`} onClick={() => setPref({ filter: f.id })}>
                        {f.icon ? <Icon name={f.icon} size={12} /> : null}
                        <span>{f.label}</span>
                      </button>
                    ))}
                  </div>
                </div>

                {query.trim() ? (
                  <div className="ml-scope">
                    <div className="ml-scope-seg" role="group" aria-label={t('mail.search-in')}>
                      <button type="button" className={scope === 'folder' ? 'on' : ''} onClick={() => setScope('folder')}>
                        {unified ? t('mail.these-accounts') : t('mail.this-folder')}
                      </button>
                      <button type="button" className={scope === 'all' ? 'on' : ''} onClick={() => setScope('all')}>
                        {t('mail.all-mail')}
                      </button>
                    </div>
                    {allMail && found ? (
                      <span className={`ml-scope-found${found.exact ? '' : ' about'}`} data-total={found.total} data-query={found.query} data-sources={new Set(found.rows.map((r) => r.accountId)).size}>
                        {tn(found.total, 'common.count-found')}
                        {found.rows.length ? ` · ${tn(new Set(found.rows.map((r) => r.accountId)).size, 'mail.count-mailbox-archive')}` : ''}
                      </span>
                    ) : null}
                    {allMail && found?.indexing?.length ? (
                      <span className="ml-scope-indexing">
                        {found.indexing
                          .map((a) =>
                            a.state === 'paused'
                              ? t('mail.name-count-indexed-indexing', { name: a.name, count: indexedCount(a) })
                              : t('mail.name-count-indexed', { name: a.name, count: indexedCount(a) })
                          )
                          .join(' ')}
                      </span>
                    ) : null}
                  </div>
                ) : null}

                {checked.size ? (
                  <div className="ml-bulk">
                    <input
                      type="checkbox"
                      className="ml-check"
                      checked
                      onChange={() => setChecked(new Set())}
                      title={t('common.clear-the-selection')}
                    />
                    <strong>{checked.size}</strong>
                    <span>{t('mail.selected')}</span>
                    <Spacer />
                    <Button icon="archive" title={t('mail.archive')} onClick={() => act('archive')} />
                    <Button icon="trash" title={t('common.delete')} onClick={() => act('delete')} />
                    <Button icon="spam" title={t('mail.junk')} onClick={() => act('junk')} />
                    <Button icon="check" title={t('mail.mark-read')} onClick={() => act('read')} />
                    <Button icon="eye" title={t('mail.mark-unread')} onClick={() => act('unread')} />
                    <Button icon="star" title={t('common.star')} onClick={() => act('flag')} />
                    {!unified ? <Button icon="folder" title={t('mail.move-to')} onClick={moveMenu} /> : null}
                  </div>
                ) : (
                  <div className="ml-listbar">
                    <input
                      type="checkbox"
                      className="ml-check"
                      checked={false}
                      onChange={() => setChecked(new Set(threads.flatMap((t) => allRowsIn(t)).map(keyOf)))}
                      title={t('mail.select-everything')}
                      aria-label={t('mail.select-everything')}
                    />
                    <span>
                      {prefs.threaded ? tn(threads.length, 'mail.count-conversation') : tn(threads.length, 'mail.count-message')}
                    </span>
                    <Spacer />
                    <Button icon="refresh" title={t('mail.get-mail')} onClick={sync} />
                  </div>
                )}

                {busy ? (
                  <div style={{ padding: 24, display: 'grid', placeItems: 'center' }}>
                    <Spinner />
                  </div>
                ) : threads.length ? (
                  <div
                    className={`ml-list${prefs.density === 'compact' ? ' compact' : ''}`}
                    role="grid"
                    aria-label={t('mail.messages')}
                    aria-multiselectable="true"
                    // Below the messages: Paste into the folder open.
                    onContextMenu={(e) => { if (e.target.closest?.('.ml-row')) return; e.preventDefault(); menu.open(e, clipItems().slice(2)); }}
                    // .eml files dropped on the list come into the folder open (Mail 2.25).
                    onDragOver={(e) => {
                      if (dragged.current?.length || !carriesFiles(e)) return;
                      e.preventDefault();
                      e.stopPropagation();
                      e.dataTransfer.dropEffect = 'copy';
                    }}
                    onDrop={(e) => {
                      if (dragged.current?.length || !carriesFiles(e)) return;
                      e.preventDefault();
                      e.stopPropagation();
                      const files = [...e.dataTransfer.files];
                      importFiles(files, folder).then((took) => !took && files[0] && pathOf(files[0]) && importFrom(pathOf(files[0])));
                    }}
                  >
                    {threads.map((thread, ti) => (
                      <Row
                        key={thread.threadKey || keyOf(thread)}
                        first={ti === 0 && !selected}
                        thread={thread}
                        unified={unified}
                        where={allMail}
                        threaded={prefs.threaded}
                        selected={Boolean(selected) && allRowsIn(thread).some((r) => keyOf(r) === keyOf(selected))}
                        checked={allRowsIn(thread).every((r) => checked.has(keyOf(r)))}
                        expanded={expanded.has(thread.threadKey)}
                        onOpen={(row) => setSelected({ accountId: row.accountId, folder: row.folder, id: row.id })}
                        onPopOut={popOut}
                        onCheck={(e) => toggleCheck(thread, e)}
                        onStar={star}
                        // Delete at the row's end: this message, or its conversation where the list shows them so, whatever else is selected.
                        onDelete={rowDeleteTargets(thread, { alone, readOnly: isArchiveId }) ? (only) => act('delete', null, rowDeleteTargets(thread, { threaded: prefs.threaded, alone, readOnly: isArchiveId, rowsIn: allRowsIn, only })) : undefined}
                        onDragStart={canChangeFolders ? (e) => startDrag(thread, e) : undefined}
                        onDragEnd={() => { dragged.current = null; setDropOn(null); }}
                        onExpand={() =>
                          setExpanded((s) => {
                            const next = new Set(s);
                            if (next.has(thread.threadKey)) next.delete(thread.threadKey);
                            else next.add(thread.threadKey);
                            return next;
                          })
                        }
                        onMenu={(e) => rowMenu(thread, e)}
                      />
                    ))}
                  </div>
                ) : (
                  <Empty icon="inbox" title={query ? t('common.nothing-matches') : prefs.filter !== 'all' ? t('mail.nothing-under-filter') : t('mail.folder-empty')}>
                    {query ? t('mail.try-different-search') : null}
                    {query && scope !== 'all' ? (
                      <div style={{ marginTop: 10 }}>
                        <Button icon="grid" label={t('mail.search-all')} onClick={() => setScope('all')} />
                      </div>
                    ) : null}
                  </Empty>
                )}
              </Panel>

              {prefs.layout !== 'off' ? (
                <Content>
                  {readingBar}
                  {paneBody(
                    <Empty icon="mail" title={t('mail.no-message-selected')}>
                      <span className="ml-keys-hint">
                        {t('mail.choose-message-read')}{' '}
                        <kbd>j</kbd> {t('mail.and')} <kbd>k</kbd> {t('mail.move')} <kbd>r</kbd> {t('mail.replies')} <kbd>s</kbd> {t('mail.stars')}
                      </span>
                    </Empty>
                  )}
                </Content>
              ) : null}
            </>
          )}
        </>
      )}

      {menu.node}
      <style>{SAFETY_CSS}</style>
      {sourceView ? <SourceDialog shell={shell} source={sourceView} subject={message?.subject} safe={safety} toast={toast} onClose={() => setSourceView(null)} /> : null}

      {/* the undo window ------------------------------------------------- */}
      {outbox.length ? (
        <div className="ml-outbox">
          <Icon name="send" size={14} />
          <span>
            {outbox.length === 1
              ? t('mail.sending-to-to', { to: outbox[0].draft?.to || t('mail.recipient') })
              : tn(outbox.length, 'mail.count-message-waiting')}
          </span>
          <button type="button" onClick={() => unsend(outbox[outbox.length - 1].id)}>
            {t('common.undo')}
          </button>
        </div>
      ) : null}

      {dialog?.kind === 'account' ? (
        <AccountDialog
          shell={shell}
          seed={dialog.seed}
          onClose={() => setDialog(null)}
          onSaved={() => {
            setDialog(null);
            loadAccounts();
          }}
          onSeveral={() => setDialog({ kind: 'bulk' })}
          toast={toast}
        />
      ) : null}

      {dialog?.kind === 'bulk' ? (
        <BulkAccountsDialog
          shell={shell}
          toast={toast}
          onClose={(added) => {
            setDialog(null);
            if (added) {
              loadAccounts();
              toast(tn(added, 'mail.count-account-added'), { tone: 'good' });
            }
          }}
          onSignIn={(email) => {
            loadAccounts();
            setDialog({ kind: 'account', seed: { email } });
          }}
          onByHand={(email) => {
            loadAccounts();
            setDialog({ kind: 'account', seed: { email } });
          }}
          onImportFile={(file) => importAccountsFrom(file)}
        />
      ) : null}

      {dialog?.kind === 'import' ? (
        <ImportDialog
          scan={scan}
          onClose={() => setDialog(null)}
          onChoose={chooseImport}
          onImportStore={(path) => importFrom(path)}
          onUseAccount={(found) => setDialog({ kind: 'account', seed: found })}
          onImportAccounts={importAccountsFrom}
        />
      ) : null}

      {dialog?.kind === 'importing' ? <ImportingDialog path={dialog.path} progress={progress} /> : null}

      {dialog?.kind === 'certificates' ? (
        <CertificatesDialog
          shell={shell}
          toast={toast}
          start={dialog.start || null}
          onClose={() => {
            setDialog(null);
            setSmimeTick((n) => n + 1);
          }}
        />
      ) : null}

      {dialog?.kind === 'hold' ? (
        <ExportHoldDialog
          shell={shell}
          query={allMail ? query.trim() : null}
          resultCount={allMail ? list.total : 0}
          folderOf={!unified && folder && accountId ? { accountId, folder } : null}
          folders={unified ? [] : folders}
          sourceName={openArchive ? openArchive.name : account ? account.name || account.email : ''}
          onClose={() => setDialog(null)}
        />
      ) : null}

      {dialog?.kind === 'search-help' ? (
        <SearchHelp
          onClose={() => setDialog(null)}
          onInsert={(code) => {
            setDialog(null);
            refine(code);
          }}
        />
      ) : null}

      {dialog?.kind === 'import-preview' ? (
        <ImportPreview found={dialog.found} onCancel={() => setDialog(null)} onImport={(chosen) => runImport(dialog.path, chosen)} onOpenInPlace={['pst', 'ost', 'olm', 'mbox'].includes(dialog.found?.kind) ? () => openArchiveAt(dialog.path) : null} />
      ) : null}

      {dialog?.kind === 'defaults' ? <DefaultsDialog shell={shell} onClose={() => setDialog(null)} toast={toast} /> : null}

      {dialog?.kind === 'rules' ? (
        <RulesDialog
          shell={shell}
          accountId={unified ? accounts[0]?.id : accountId}
          folder={folder}
          folders={folders}
          toast={toast}
          onClose={() => {
            setDialog(null);
            refreshList();
            loadAccounts();
          }}
        />
      ) : null}

      {dialog?.kind === 'outbox' ? (
        <Dialog title={t('mail.outbox')} width={620} onClose={() => setDialog(null)} actions={<Button primary label={t('common.close')} onClick={() => setDialog(null)} />}>
          {outbox.length ? (
            <div className="ml-found">
              {(unified ? outbox : outbox.filter((o) => o.accountId === accountId)).map((item) => (
                <div key={item.id} className="ml-found-item" style={{ cursor: 'default', alignItems: 'flex-start' }}>
                  <span className="ml-found-logo">
                    <Icon name="clock" size={15} />
                  </span>
                  <span className="grow">
                    <div className="who">{item.draft?.subject || t('mail.no-subject')}</div>
                    <div className="what">
                      {t('mail.to-to', { to: item.draft?.to || '' })} · {new Date(item.at) > new Date() ? t('mail.sends-when', { when: new Date(item.at).toLocaleString() }) : t('mail.sending-now')}
                      {unified ? ` · ${accounts.find((a) => a.id === item.accountId)?.email || item.accountId}` : ''}
                      {item.held ? ` · ${t('mail.rutba-office-stopped')}` : ''}
                      {item.error ? ` · ${t('mail.last-try-failed', { error: item.error })}` : ''}
                    </div>
                    <div className="ml-outbox-actions">
                      <Button label={t('common.edit')} title={t('mail.take-back-reopen')} onClick={() => unsend(item.id)} />
                      <Button label={t('mail.send-now')} onClick={() => sendOutboxNow(item)} />
                      <Button label={t('common.cancel')} title={t('mail.move-drafts')} onClick={() => cancelOutboxItem(item)} />
                    </div>
                  </span>
                </div>
              ))}
            </div>
          ) : (
            <p style={{ marginTop: 0 }}>{t('mail.nothing-is-waiting')}</p>
          )}
          <p className="rw-hint" style={{ marginBottom: 0 }}>
            {t('mail.scheduled-held-message')}
          </p>
        </Dialog>
      ) : null}

      {/* An account whose password is not saved here: asked for, not failed at every fetch. */}
      <PasswordAgain shell={shell} accounts={accounts} toast={toast} onSaved={async (a) => { await loadAccounts(); shell.mail.sync({ accountId: a.id }).catch(() => {}); }} />
      {dialog?.kind === 'accounts' ? (
        <AccountsDialog
          shell={shell}
          accounts={accounts}
          archives={archives}
          onClose={() => setDialog(null)}
          onAdd={() => setDialog({ kind: 'account' })}
          onRemove={(a) => removeAccount(a)}
          onSignature={(a) => setDialog({ kind: 'signature', accountId: a.id })}
          onOutOfOffice={(a) => setDialog({ kind: 'ooo', accountId: a.id })}
          onDefault={async (id) => {
            try {
              const chosen = await shell.mail.setDefaultAccount({ id });
              await loadAccounts();
              toast(t('mail.new-mail-goes', { address: chosen.address }), { tone: 'good' });
            } catch (err) {
              toast(err?.message || String(err), { tone: 'bad' });
            }
          }}
        />
      ) : null}
      {dialog?.kind === 'signature' ? (
        <SignatureDialog
          shell={shell}
          accounts={accounts}
          accountId={dialog.accountId}
          onClose={() => setDialog(null)}
          onSaved={() => {
            setDialog(null);
            loadAccounts();
          }}
          toast={toast}
        />
      ) : null}

      {dialog?.kind === 'folderName' ? (
        <PromptDialog
          title={dialog.folder ? t('mail.rename-folder') : t('mail.new-folder')}
          label={
            dialog.folder
              ? t('mail.new-name-folder', { folder: dialog.folder.name || dialog.folder.path })
              : dialog.parent
                ? t('mail.name-folder-make', { folder: dialog.parent.name || dialog.parent.path })
                : t('mail.name-new-folder')
          }
          initial={dialog.folder?.name || ''}
          okLabel={dialog.folder ? t('common.rename') : t('mail.make-folder')}
          className="ml-folder-name"
          onClose={() => setDialog(null)}
          onApply={saveFolderName}
        />
      ) : null}

      {dialog?.kind === 'ooo' ? (
        <OutOfOfficeDialog
          shell={shell}
          accounts={accounts}
          accountId={dialog.accountId}
          onClose={() => setDialog(null)}
          onSaved={() => {
            setDialog(null);
            loadAccounts();
          }}
          toast={toast}
        />
      ) : null}

      {dialog?.kind === 'junk' ? (
        <JunkDialog
          shell={shell}
          accounts={accounts}
          accountId={accountId && accountId !== EVERYTHING ? accountId : accounts[0]?.id}
          onClose={() => setDialog(null)}
          onSaved={() => setDialog(null)}
          toast={toast}
        />
      ) : null}

      {dialog?.kind === 'undo' ? (
        <Dialog
          title={t('mail.undo-window')}
          width={460}
          onClose={() => setDialog(null)}
          actions={
            <>
              <Button label={t('common.cancel')} onClick={() => setDialog(null)} />
              <Button
                primary
                label={t('common.save')}
                onClick={async () => {
                  await shell.store.set({ key: 'mail.undoSeconds', value: Number(dialog.seconds) || 0 });
                  toast(Number(dialog.seconds) ? tn(Number(dialog.seconds), 'mail.count-second-change') : t('mail.messages-now-send'), { tone: 'good' });
                  setDialog(null);
                }}
              />
            </>
          }
        >
          <Field label={t('mail.hold-message')} hint={t('mail.every-message-waits')}>
            <Select value={String(dialog.seconds)} onChange={(e) => setDialog((d) => ({ ...d, seconds: Number(e.target.value) }))}>
              <option value="0">{t('mail.no-delay')}</option>
              <option value="5">{tn(5, 'mail.count-second')}</option>
              <option value="8">{tn(8, 'mail.count-second')}</option>
              <option value="15">{tn(15, 'mail.count-second')}</option>
              <option value="30">{tn(30, 'mail.count-second')}</option>
            </Select>
          </Field>
        </Dialog>
      ) : null}

      {placeNow !== 'docked' ? composeView : null}
    </AppFrame>
  );
}

/* ── one line of the list ────────────────────────────────────────────────── */

/**
 * Memoised because a 500-row list re-renders on every keystroke in the search
 * box otherwise, and the whole point of holding the mail locally is that the
 * list is instant.
 */
const Row = React.memo(function Row({ thread, unified, where, threaded, selected, checked, expanded, first, onOpen, onPopOut, onCheck, onStar, onDelete, onExpand, onMenu, onDragStart, onDragEnd }) {
  const avatar = avatarFor(thread.from);
  const many = threaded && thread.count > 1;

  return (
    <>
      <div
        className={`ml-row${selected ? ' selected' : ''}${thread.unread ? ' unread' : ''}${checked ? ' checked' : ''}${thread.pinned ? ' pinned' : ''}`}
        role="row"
        aria-selected={selected}
        aria-expanded={many ? Boolean(expanded) : undefined}
        tabIndex={selected || first ? 0 : -1}
        data-f6-home={selected || first ? '' : undefined}
        // Ctrl+click adds the message to those checked and Shift+click checks the run to it, as in Outlook: the whole row is the target.
        onClick={(e) => (e.ctrlKey || e.metaKey || e.shiftKey ? onCheck(e.nativeEvent) : onOpen(thread))}
        onDoubleClick={(e) => !(e.ctrlKey || e.metaKey || e.shiftKey) && onPopOut?.(thread)}
        onKeyDown={(e) => (e.key === 'Enter' || e.key === ' ') && (e.preventDefault(), onOpen(thread))}
        onContextMenu={onMenu}
        draggable={Boolean(onDragStart)}
        onDragStart={onDragStart}
        onDragEnd={onDragEnd}
      >
        <span role="gridcell">
          <input
            type="checkbox"
            className="ml-check"
            checked={checked}
            aria-label={t('common.select')}
            style={{ marginTop: 8 }}
            onClick={(e) => e.stopPropagation()}
            onChange={(e) => onCheck(e.nativeEvent)}
          />
        </span>

        <span className="ml-avatar-sm" role="gridcell" aria-hidden="true" style={{ background: avatar.colour }} title={thread.from?.address || ''}>
          {avatar.initial}
        </span>

        <div className="ml-row-main" role="gridcell">
          <div className="ml-row-top">
            <span className="ml-from">
              {many
                ? thread.participants.slice(0, 3).map(displayName).join(', ') + (thread.participants.length > 3 ? `, +${thread.participants.length - 3}` : '')
                : displayName(thread.from)}
            </span>
            {many ? (
              <button
                type="button"
                className="ml-count"
                onClick={(e) => {
                  e.stopPropagation();
                  onExpand();
                }}
                title={tn(thread.count, 'mail.count-message-conversation')}
              >
                {thread.count}
              </button>
            ) : null}
          </div>
          <div className="ml-subject">
            {thread.pinned ? <Icon name="flag" size={12} /> : null}
            {thread.subject || t('mail.no-subject')}
          </div>
          <div className="ml-preview">
            {thread.hasAttachments ? <Icon name="attach" size={12} /> : null}
            {thread.preview}
          </div>
          {where && thread.sourceName ? (
            <div className={`ml-where${thread.readOnly ? ' archive' : ''}`} title={thread.sourceFile || thread.sourceName}>
              <Icon name={thread.readOnly ? 'archive' : 'mail'} size={11} />
              <span>{`${thread.sourceName} · ${thread.folder}`}</span>
            </div>
          ) : null}
        </div>

        <div className="ml-row-side" role="gridcell">
          <span className="ml-when">{formatWhen(thread.date)}</span>
          <span className="ml-marks">
            {unified ? <span className="ml-folder-tag">{String(thread.folder || '').split('/').pop()}</span> : null}
            <button
              type="button"
              className={`ml-star${thread.flagged ? ' on' : ''}`}
              title={thread.flagged ? t('mail.remove-star') : t('common.star')}
              onClick={(e) => onStar(thread, e)}
            >
              <Icon name="star" size={13} />
            </button>
            {onDelete ? (
              // Shown on the row under the pointer and on the selected one; the Delete key does the same.
              <button
                type="button"
                className="ml-row-delete"
                title={many ? t('mail.delete-this') : t('mail.delete-this-message')}
                aria-label={many ? t('mail.delete-this') : t('mail.delete-this-message')}
                tabIndex={-1}
                onClick={(e) => {
                  e.stopPropagation();
                  onDelete(null);
                }}
                onDoubleClick={(e) => e.stopPropagation()}
              >
                <Icon name="trash" size={13} />
              </button>
            ) : null}
          </span>
        </div>
      </div>

      {many && expanded ? (
        <div className="ml-thread-children" role="rowgroup">
          {thread.messages
            .slice()
            .reverse()
            .map((m) => (
              <div key={m.id} className={`ml-row${m.unread ? ' unread' : ''}`} role="row" tabIndex={-1} onClick={() => onOpen(m)} onDoubleClick={() => onPopOut?.(m)} onKeyDown={(e) => e.key === 'Enter' && onOpen(m)}>
                <span role="gridcell" />
                <span className="ml-avatar-sm" role="gridcell" aria-hidden="true" style={{ background: avatarFor(m.from).colour, width: 22, height: 22, fontSize: 10 }}>
                  {avatarFor(m.from).initial}
                </span>
                <div className="ml-row-main" role="gridcell">
                  <div className="ml-row-top">
                    <span className="ml-from">{displayName(m.from)}</span>
                  </div>
                  <div className="ml-preview">{m.preview}</div>
                </div>
                <div className="ml-row-side" role="gridcell">
                  <span className="ml-when">{formatWhen(m.date)}</span>
                  {onDelete ? (
                    <span className="ml-marks">
                      <button
                        type="button"
                        className="ml-row-delete"
                        title={t('mail.delete-this-message')}
                        aria-label={t('mail.delete-this-message')}
                        tabIndex={-1}
                        onClick={(e) => {
                          e.stopPropagation();
                          onDelete([m]);
                        }}
                        onDoubleClick={(e) => e.stopPropagation()}
                      >
                        <Icon name="trash" size={13} />
                      </button>
                    </span>
                  ) : null}
                </div>
              </div>
            ))}
        </div>
      ) : null}
    </>
  );
});

/**
 * Mail's window. Opened on a .eml file (from Explorer, Open with, or a file
 * dropped on the taskbar), it is that message alone, read only, with Add to a
 * folder; the file is read first, and a file that is not mail says so and
 * leaves an ordinary Mail window (Mail 2.25).
 */
export default function Mail({ app, shell, boot }) {
  const file = openedFileOf(boot);
  const [opened, setOpened] = useState(file ? null : false);
  const toast = useToast();
  useEffect(() => {
    if (!file) return;
    shell.mail.openEml({ path: file }).then(setOpened, (err) => {
      toast(err?.message || String(err), { tone: 'bad', ms: 8000 });
      setOpened(false);
    });
  }, [shell, file, toast]);
  if (opened === null) {
    return (
      <div style={{ height: '100vh', display: 'grid', placeItems: 'center' }}>
        <Spinner />
      </div>
    );
  }
  return <MailWindow app={app} shell={shell} opened={opened || null} />;
}
