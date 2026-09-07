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
  Search, Dialog, Field, Input, Select, useToast, useMenu, useCommands, menuItems,
  formatWhen, formatBytes,
} from '@rutba/office-ui';
import { AppFrame, useAppMenu, useFileDrop } from '../shell.js';
import { appFor, kindFromExtension } from '@rutba/office-formats/sniff';
import { DefaultsDialog } from '../defaults.js';
import { avatarFor, displayName, buildThreads, arrange, FILTERS, SORTS, installStyles, stripTags } from './mail/parts.js';
import Reader from './mail/reader.js';
import Compose from './mail/compose.js';
import { AccountDialog, ImportDialog, ImportPreview, ImportingDialog, FilesView, PeopleView } from './mail/dialogs.js';
import { RulesDialog } from './mail/rules.js';

installStyles();

/** Every account at once. Not an id any account can have. */
const EVERYTHING = '*';

const keyOf = (row) => `${row.accountId}|${row.folder}|${row.id}`;

export default function Mail({ app, shell }) {
  const toast = useToast();
  const menu = useMenu();

  const [accounts, setAccounts] = useState([]);
  const [accountId, setAccountId] = useState(null);
  const [folders, setFolders] = useState([]);
  const [folder, setFolder] = useState(null);
  const [role, setRole] = useState('inbox');
  const [list, setList] = useState({ rows: [], total: 0 });
  const [busy, setBusy] = useState(false);

  const [checked, setChecked] = useState(() => new Set());
  const [selected, setSelected] = useState(null); // the row object, so unified lists know where it lives
  const [message, setMessage] = useState(null);
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
  const [compose, setCompose] = useState(null);
  const [scan, setScan] = useState(null);
  const [outbox, setOutbox] = useState([]);

  // How the list looks. Remembered, because nobody wants to choose twice.
  const [prefs, setPrefs] = useState({
    filter: 'all',
    sort: 'date',
    ascending: false,
    threaded: true,
    layout: 'right',
    density: 'cosy',
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

  const appMenu = useAppMenu({
    shell,
    appKey: 'mail',
    extra: [
      { label: 'Import mail…', icon: 'import', run: () => setDialog({ kind: 'import' }) },
      { label: 'Add account…', icon: 'plus', run: () => setDialog({ kind: 'account' }) },
      { label: 'Open files with Rutba Office…', icon: 'settings', run: () => setDialog({ kind: 'defaults' }) },
    ],
  });

  /* ── loading ──────────────────────────────────────────────────────────── */

  useEffect(() => {
    shell.store
      .get({ key: 'mail.view', fallback: null })
      .then((saved) => saved && setPrefs((p) => ({ ...p, ...saved })))
      .catch(() => {});
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

  useEffect(() => {
    if (!accountId || unified) return;
    shell.mail
      .folders({ accountId })
      .then((f) => {
        setFolders(f);
        setFolder((current) =>
          f.some((x) => x.path === current) ? current : f.find((x) => x.role === 'inbox')?.path || f[0]?.path || null
        );
      })
      .catch(() => setFolders([]));
  }, [accountId, unified, shell]);

  const refreshList = useCallback(async () => {
    if (view !== 'mail') return;
    if (unified) {
      setBusy(true);
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
    setBusy(true);
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
  }, [unified, role, accountId, folder, query, view, shell, toast]);

  useEffect(() => {
    refreshList();
    setChecked(new Set());
    setSelected(null);
    setMessage(null);
    setInsight(null);
  }, [refreshList]);

  // Opening a message: the body and the report on it come together, because
  // showing the body first and the warning a moment later is exactly backwards.
  useEffect(() => {
    if (!selected) {
      setMessage(null);
      setInsight(null);
      return;
    }
    const { accountId: a, folder: f, id } = selected;
    setRemote(false);
    Promise.all([shell.mail.message({ accountId: a, folder: f, id }), shell.mail.insight({ accountId: a, folder: f, id })])
      .then(([m, i]) => {
        setMessage(m);
        setInsight(i);
        if (m?.unread) {
          shell.mail.flag({ accountId: a, folder: f, ids: [id], patch: { unread: false } }).then(() => {
            setList((current) => ({
              ...current,
              rows: current.rows.map((r) => (r.id === id && r.folder === f ? { ...r, unread: false } : r)),
            }));
            loadAccounts();
          });
        }
      })
      .catch((err) => toast(err.message, { tone: 'bad' }));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selected]);

  useEffect(() => {
    const offProgress = shell.on('mail:progress', (p) => setProgress(p));
    const offNew = shell.on('mail:new', () => {
      setProgress(null);
      loadAccounts();
      refreshList();
    });
    const offSent = shell.on('mail:sent', ({ to }) => {
      toast(`Sent to ${to}`, { tone: 'good' });
      shell.mail.outbox().then(setOutbox).catch(() => {});
      refreshList();
    });
    const offFailed = shell.on('mail:sendFailed', ({ message: why, gaveUp }) => {
      toast(gaveUp ? `Could not send: ${why}. The message is back in Drafts.` : `Send failed, trying again: ${why}`, {
        tone: 'bad',
        ms: 9000,
      });
      shell.mail.outbox().then(setOutbox).catch(() => {});
    });
    return () => {
      offProgress?.();
      offNew?.();
      offSent?.();
      offFailed?.();
    };
  }, [shell, loadAccounts, refreshList, toast]);

  useEffect(() => {
    shell.mail.outbox().then(setOutbox).catch(() => {});
  }, [shell]);

  /* ── the list, as it is actually drawn ────────────────────────────────── */

  const rows = useMemo(() => arrange(list.rows, prefs), [list.rows, prefs]);
  const threads = useMemo(() => (prefs.threaded ? buildThreads(rows) : rows.map((r) => ({ ...r, messages: [r], count: 1 }))), [rows, prefs.threaded]);

  const allRowsIn = useCallback((thread) => (prefs.threaded ? thread.messages : [thread]), [prefs.threaded]);

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
    if (checkedRows.length) return checkedRows;
    if (!selected) return [];
    const thread = threads.find((t) => allRowsIn(t).some((r) => keyOf(r) === keyOf(selected)));
    return thread && prefs.threaded ? thread.messages : [selected];
  }, [checkedRows, selected, threads, allRowsIn, prefs.threaded]);

  const foldersFor = useCallback(
    (accId) => (accId === accountId && !unified ? folders : []),
    [accountId, unified, folders]
  );

  const act = useCallback(
    async (what, extra) => {
      const chosen = targets();
      if (!chosen.length) return;
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
        } else if (what === 'move') {
          await overRows(chosen, (g) => shell.mail.move({ ...g, to: extra }));
        } else if (what === 'read') {
          await overRows(chosen, (g) => shell.mail.flag({ ...g, patch: { unread: false } }));
        } else if (what === 'unread') {
          await overRows(chosen, (g) => shell.mail.flag({ ...g, patch: { unread: true } }));
        } else if (what === 'flag') {
          const on = !chosen.every((r) => r.flagged);
          await overRows(chosen, (g) => shell.mail.flag({ ...g, patch: { flagged: on } }));
        } else if (what === 'pin') {
          const on = !chosen.every((r) => r.pinned);
          await overRows(chosen, (g) => shell.mail.flag({ ...g, patch: { pinned: on } }));
        }
        setChecked(new Set());
        if (what === 'delete' || what === 'archive' || what === 'junk' || what === 'move') setSelected(null);
        await refreshList();
        await loadAccounts();
      } catch (err) {
        toast(err.message, { tone: 'bad' });
      }
    },
    [targets, overRows, foldersFor, shell, refreshList, loadAccounts, toast]
  );

  /** Star one message from the list without opening it. */
  const star = useCallback(
    async (row, event) => {
      event.stopPropagation();
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
          ? 'These are imported archives — there is nothing to fetch.'
          : `${added} new message${added === 1 ? '' : 's'}`,
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

  const signature = useCallback(async () => {
    const text = await shell.store.get({ key: 'mail.signature', fallback: '' });
    return text ? `\n\n-- \n${text}` : '';
  }, [shell]);

  const quoted = useCallback((m) => {
    const body = m.text || stripTags(m.html || '');
    return `\n\nOn ${formatWhen(m.date, { long: true })}, ${displayName(m.from?.[0])} wrote:\n${body
      .split('\n')
      .map((l) => `> ${l}`)
      .join('\n')}`;
  }, []);

  const reply = useCallback(
    async (all = false) => {
      if (!message) return;
      const me = accounts.find((a) => a.id === (selected?.accountId || accountId))?.email?.toLowerCase();
      const to = [message.replyTo?.[0]?.address || message.from?.[0]?.address].filter(Boolean);
      const cc = all
        ? [...(message.to || []), ...(message.cc || [])]
            .map((p) => p.address)
            .filter((a) => a && a.toLowerCase() !== me && !to.includes(a))
        : [];
      setCompose({
        to: to.join(', '),
        cc: [...new Set(cc)].join(', '),
        subject: /^re:/i.test(message.subject || '') ? message.subject : `Re: ${message.subject || ''}`,
        text: (await signature()) + quoted(message),
        inReplyTo: message.messageId,
      });
    },
    [message, accounts, selected, accountId, signature, quoted]
  );

  const forward = useCallback(async () => {
    if (!message) return;
    setCompose({
      to: '',
      subject: /^fwd:/i.test(message.subject || '') ? message.subject : `Fwd: ${message.subject || ''}`,
      text:
        (await signature()) +
        `\n\n---------- Forwarded message ----------\nFrom: ${displayName(message.from?.[0])}\nDate: ${formatWhen(message.date, { long: true })}\nSubject: ${message.subject || ''}\nTo: ${(message.to || []).map((t) => t.address).join(', ')}\n\n${message.text || stripTags(message.html || '')}`,
      // Forwarding carries the files. They are already on disk here, so this
      // costs nothing until the message is actually sent.
      attachments: (message.attachments || []).filter((a) => !a.inline && a.stored).map((a) => ({ filename: a.filename, size: a.size, fromMessage: { ...selected, index: message.attachments.indexOf(a) } })),
    });
  }, [message, selected, signature]);

  const doSend = useCallback(
    async (draft, at) => {
      const from = draft.accountId || accountId;
      if (!from || from === EVERYTHING) return toast('Choose which account to send from.', { tone: 'bad' });
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
          const { path } = await shell.fs.temp({ ext: a.filename.split('.').pop(), bytes });
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
        if (at) toast(`Scheduled for ${new Date(at).toLocaleString()}`, { tone: 'good', ms: 6000 });
        return item;
      } catch (err) {
        toast(err.message, { tone: 'bad', ms: 7000 });
        return null;
      }
    },
    [accountId, shell, toast]
  );

  const unsend = useCallback(
    async (id) => {
      const item = await shell.mail.unsend({ id });
      setOutbox(await shell.mail.outbox());
      if (item) setCompose({ ...item.draft, accountId: item.accountId });
    },
    [shell]
  );

  /* ── attachments ──────────────────────────────────────────────────────── */

  const saveAttachment = useCallback(
    async (index) => {
      if (!selected) return;
      const list = index === 'all' ? (message.attachments || []).map((a, i) => i).filter((i) => !message.attachments[i].inline) : [index];
      if (index === 'all') {
        const dirs = await shell.dialog.open({ title: 'Save all attachments to…', directory: true });
        if (!dirs?.[0]) return;
        for (const i of list) {
          const held = await shell.mail.attachment({ ...selected, index: i });
          if (!held) continue;
          const bytes = new Uint8Array(await (await fetch(held.url)).arrayBuffer());
          await shell.fs.write({ path: `${dirs[0]}/${held.name}`, bytes });
        }
        toast(`Saved ${list.length} files`, { tone: 'good' });
        return;
      }
      const held = await shell.mail.attachment({ ...selected, index });
      if (!held) return toast('That attachment is not stored.', { tone: 'bad' });
      const target = await shell.dialog.save({ title: 'Save attachment', defaultPath: held.name });
      if (!target) return;
      const bytes = new Uint8Array(await (await fetch(held.url)).arrayBuffer());
      await shell.fs.write({ path: target, bytes });
      toast(`Saved ${held.name}`, { tone: 'good' });
    },
    [selected, message, shell, toast]
  );

  /** Open an attachment in whichever Rutba Office app owns it. */
  const openAttachment = useCallback(
    async (index, meta, where = selected) => {
      const held = await shell.mail.attachment({ ...where, index });
      if (!held) return toast('That attachment is not stored.', { tone: 'bad' });
      const bytes = new Uint8Array(await (await fetch(held.url)).arrayBuffer());
      const name = String(meta?.filename || held.name);
      const { path } = await shell.fs.temp({ ext: name.split('.').pop(), bytes });
      const owner = appFor(kindFromExtension(name));
      // Anything the suite understands opens in the app that owns it; anything
      // else goes to whatever this computer uses for that kind of file.
      if (owner) await shell.win.create({ app: owner, file: path });
      else await shell.shell.openPath({ path });
    },
    [selected, shell, toast]
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
      title: 'Import mail',
      filters: [
        { name: 'Mail archives', extensions: ['pst', 'ost', 'olm', 'mbox', 'mbx', 'eml', 'emlx', 'msg'] },
        { name: 'All files', extensions: ['*'] },
      ],
    });
    if (paths[0]) importFrom(paths[0]);
  }, [shell, importFrom]);

  const runImport = useCallback(
    async (target, chosen) => {
      setDialog({ kind: 'importing', path: target });
      try {
        const result = await shell.mail.import({ path: target, folders: chosen });
        await loadAccounts();
        setAccountId(result.accountId);
        toast(`Imported ${result.messages.toLocaleString()} messages into ${result.folders} folders`, { tone: 'good', ms: 6000 });
      } catch (err) {
        toast(err.message, { tone: 'bad', ms: 7000 });
      } finally {
        setDialog(null);
        setProgress(null);
      }
    },
    [shell, loadAccounts, toast]
  );

  useFileDrop(useCallback((files) => files[0] && importFrom(files[0]), [importFrom]));

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
        message: `Unsubscribe from ${offer.list || 'this list'}?`,
        detail:
          offer.method === 'one-click'
            ? 'The sender supports one-click unsubscribe. A single request is sent to them and nothing else leaves this computer.'
            : offer.method === 'web'
              ? `This opens ${new URL(offer.http).hostname} in your browser, where the sender handles it.`
              : `This sends a message to ${offer.address}.`,
        buttons: ['Unsubscribe', 'Cancel'],
        cancelId: 1,
      });
      if (response !== 0) return;

      if (offer.method === 'email' || (!offer.http && offer.address)) {
        await doSend({ to: offer.address, subject: offer.subject, text: 'Please remove this address from the list.' }, null);
        toast('Unsubscribe request queued.', { tone: 'good' });
      } else {
        await shell.shell.openExternal({ url: offer.http });
      }
    },
    [shell, doSend, toast]
  );

  /* ── commands ─────────────────────────────────────────────────────────── */

  const commands = useMemo(
    () => ({
      'mail.sync': { label: 'Get mail', icon: 'refresh', key: 'Mod+R', run: sync },
      'mail.compose': { label: 'New message', icon: 'new', key: 'Mod+N', run: () => setCompose({ to: '', subject: '', text: '' }) },
      'mail.reply': { label: 'Reply', icon: 'reply', key: 'Mod+R', run: () => reply(false) },
      'mail.replyAll': { label: 'Reply all', icon: 'replyAll', run: () => reply(true) },
      'mail.forward': { label: 'Forward', icon: 'forward', run: forward },
      'mail.archive': { label: 'Archive', icon: 'archive', key: 'e', run: () => act('archive') },
      'mail.delete': { label: 'Delete', icon: 'trash', key: 'Delete', run: () => act('delete') },
      'mail.junk': { label: 'Move to Junk', icon: 'spam', run: () => act('junk') },
      'mail.flag': { label: 'Star', icon: 'star', run: () => act('flag') },
      'mail.pin': { label: 'Pin to the top', icon: 'flag', run: () => act('pin') },
      'mail.read': { label: 'Mark read', icon: 'check', run: () => act('read') },
      'mail.unread': { label: 'Mark unread', icon: 'eye', run: () => act('unread') },
      'mail.import': { label: 'Import mail…', icon: 'import', run: () => setDialog({ kind: 'import' }) },
      'mail.account': { label: 'Add account…', icon: 'plus', run: () => setDialog({ kind: 'account' }) },
      'mail.search': { label: 'Search', icon: 'find', key: 'Mod+F', run: () => document.querySelector('.rw-search input')?.focus() },
    }),
    [sync, reply, forward, act]
  );

  useCommands(commands, [selected, message, folder, accountId, checked]);

  // Single-key navigation, the way every mail client that respects your hands
  // does it. Suppressed whenever a field has focus, so typing "j" in a search
  // box searches for j.
  useEffect(() => {
    const onKey = (e) => {
      const el = document.activeElement;
      if (e.metaKey || e.ctrlKey || e.altKey) return;
      if (el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.isContentEditable)) return;
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
  }, [threads, selected, allRowsIn, act, reply, forward]);

  /* ── chrome ───────────────────────────────────────────────────────────── */

  const account = accounts.find((a) => a.id === accountId);
  const current = folders.find((f) => f.path === folder);
  const unreadTotal = accounts.reduce((n, a) => n + (a.counts?.unread || 0), 0);
  const title = unified ? 'All accounts' : account ? `${current?.name || 'Mail'} — ${account.name || account.email}` : 'Mail';

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

  return (
    <AppFrame
      app={app}
      shell={shell}
      title={title}
      menu={appMenu}
      ribbon={
        <Ribbon
          tabs={[
            { id: 'home', label: 'Home' },
            { id: 'send', label: 'Send / Receive' },
            { id: 'folder', label: 'Folder' },
            { id: 'view', label: 'View' },
            { id: 'tools', label: 'Tools' },
          ]}
          active={tab}
          onTab={setTab}
          quick={
            <>
              <Button icon="refresh" title="Get mail" onClick={sync} disabled={!accountId} />
              <Button icon="new" title="New message" onClick={() => commands['mail.compose'].run()} disabled={!accounts.length} />
              <Button icon="trash" title="Delete" onClick={() => act('delete')} disabled={!selected && !checked.size} />
              <Button icon="star" title="Star" onClick={() => act('flag')} disabled={!selected && !checked.size} />
            </>
          }
        >
          {tab === 'home' ? (
            <>
              <Group label="New">
                <Button tall icon="new" label="Message" onClick={() => commands['mail.compose'].run()} disabled={!accounts.length} />
              </Group>
              <Group label="Delete">
                <Button tall icon="trash" label="Delete" disabled={!selected && !checked.size} onClick={() => act('delete')} />
                <Button icon="archive" label="Archive" disabled={!selected && !checked.size} onClick={() => act('archive')} />
                <Button icon="spam" label="Junk" disabled={!selected && !checked.size} onClick={() => act('junk')} />
              </Group>
              <Group label="Respond">
                <Button tall icon="reply" label="Reply" disabled={!message} onClick={() => reply(false)} />
                <Button tall icon="replyAll" label="Reply all" disabled={!message} onClick={() => reply(true)} />
                <Button tall icon="forward" label="Forward" disabled={!message} onClick={forward} />
              </Group>
              <Group label="Move">
                <Button tall icon="folder" label="Move to" disabled={(!selected && !checked.size) || unified} onClick={moveMenu} />
              </Group>
              <Group label="Tags">
                <Button icon="star" label="Star" disabled={!selected && !checked.size} onClick={() => act('flag')} />
                <Button icon="flag" label="Pin" disabled={!selected && !checked.size} onClick={() => act('pin')} />
                <Button icon="check" label="Read" disabled={!selected && !checked.size} onClick={() => act('read')} />
                <Button icon="eye" label="Unread" disabled={!selected && !checked.size} onClick={() => act('unread')} />
              </Group>
              <Group label="Find">
                <Button tall icon="find" label="Search" onClick={() => commands['mail.search'].run()} />
                <Button icon="attach" label="Attachments" pressed={view === 'files'} onClick={() => setView(view === 'files' ? 'mail' : 'files')} disabled={unified} />
                <Button icon="reply" label="People" pressed={view === 'people'} onClick={() => setView(view === 'people' ? 'mail' : 'people')} disabled={unified} />
                <Button icon="filter" label="Rules" onClick={() => setDialog({ kind: 'rules' })} />
              </Group>
            </>
          ) : tab === 'send' ? (
            <>
              <Group label="Send and receive">
                <Button tall icon="refresh" label={unified ? 'All accounts' : 'This account'} onClick={sync} disabled={!accountId} />
                <Button icon="download" label="This folder" disabled={!folder || unified} onClick={sync} />
              </Group>
              <Group label="Outbox">
                <Button
                  tall
                  icon="clock"
                  label={outbox.length ? `Waiting (${outbox.length})` : 'Nothing waiting'}
                  disabled={!outbox.length}
                  onClick={() => setDialog({ kind: 'outbox' })}
                />
              </Group>
              <Group label="Import">
                <Button tall icon="import" label="Import mail" onClick={() => setDialog({ kind: 'import' })} />
              </Group>
              <Group label="Export">
                <Button
                  tall
                  icon="export"
                  label="Mbox"
                  disabled={!folder || unified}
                  onClick={async () => {
                    const target = await shell.dialog.save({ title: 'Export folder', defaultPath: `${current?.name || 'folder'}.mbox` });
                    if (!target) return;
                    const r = await shell.mail.export({ accountId, folder, path: target, format: 'mbox' });
                    toast(`Exported ${r.messages} messages`, { tone: 'good' });
                  }}
                />
              </Group>
            </>
          ) : tab === 'folder' ? (
            <>
              <Group label="Accounts">
                <Button tall icon="plus" label="Add" onClick={() => setDialog({ kind: 'account' })} />
                <Button
                  tall
                  icon="trash"
                  label="Remove"
                  disabled={!account}
                  onClick={async () => {
                    const { response } = await shell.dialog.message({
                      type: 'warning',
                      message: `Remove ${account.email}?`,
                      detail: 'The mail stored on this computer for that account is deleted too.',
                      buttons: ['Remove', 'Keep the mail', 'Cancel'],
                      cancelId: 2,
                    });
                    if (response === 2) return;
                    await shell.mail.removeAccount({ id: account.id, keepMail: response === 1 });
                    setAccountId(null);
                    loadAccounts();
                  }}
                />
              </Group>
              <Group label="This folder">
                <Button
                  tall
                  icon="check"
                  label="Mark all read"
                  disabled={!folder || unified}
                  onClick={async () => {
                    const r = await shell.mail.markAllRead({ accountId, folder });
                    toast(`${r.changed} marked read`, { tone: 'good' });
                    refreshList();
                    loadAccounts();
                  }}
                />
                <Button
                  icon="trash"
                  label="Empty folder"
                  disabled={!current || !['trash', 'junk'].includes(current.role)}
                  onClick={async () => {
                    const { response } = await shell.dialog.message({
                      type: 'warning',
                      message: `Permanently delete everything in ${current.name}?`,
                      detail: 'This cannot be undone.',
                      buttons: ['Delete', 'Cancel'],
                      cancelId: 1,
                    });
                    if (response !== 0) return;
                    const r = await shell.mail.emptyFolder({ accountId, folder });
                    toast(`${r.removed} deleted`, { tone: 'good' });
                    refreshList();
                  }}
                />
              </Group>
            </>
          ) : tab === 'view' ? (
            <>
              <Group label="Arrangement">
                <Button icon="list" label="Conversations" pressed={prefs.threaded} onClick={() => setPref({ threaded: !prefs.threaded })} />
                <Button
                  icon="sort"
                  label={`Sort: ${SORTS.find((s) => s.id === prefs.sort)?.label}`}
                  onClick={(e) =>
                    menu.open(
                      e,
                      SORTS.map((s) => ({ label: s.label, icon: prefs.sort === s.id ? 'check' : undefined, run: () => setPref({ sort: s.id }) }))
                        .concat(['-', { label: prefs.ascending ? 'Newest first' : 'Oldest first', icon: 'refresh', run: () => setPref({ ascending: !prefs.ascending }) }])
                    )
                  }
                />
                <Button icon="grid" label={prefs.density === 'compact' ? 'Compact' : 'Comfortable'} onClick={() => setPref({ density: prefs.density === 'compact' ? 'cosy' : 'compact' })} />
              </Group>
              <Group label="Reading pane">
                <Button icon="grid" label="Right" pressed={prefs.layout === 'right'} onClick={() => setPref({ layout: 'right' })} />
                <Button icon="list" label="Bottom" pressed={prefs.layout === 'bottom'} onClick={() => setPref({ layout: 'bottom' })} />
                <Button icon="close" label="Off" pressed={prefs.layout === 'off'} onClick={() => setPref({ layout: 'off' })} />
              </Group>
              <Group label="Message">
                <Button icon="eye" label="Remote images" pressed={remote} onClick={() => setRemote((r) => !r)} />
                <Button icon="file" label="Plain text" pressed={plain} onClick={() => setPlain((p) => !p)} />
              </Group>
              <Group label="Zoom">
                <Button icon="zoomOut" label="Out" onClick={() => shell.win.zoom({ delta: -0.1 })} />
                <Button icon="zoomIn" label="In" onClick={() => shell.win.zoom({ delta: 0.1 })} />
              </Group>
            </>
          ) : (
            <>
              <Group label="This computer">
                <Button tall icon="settings" label="Default apps" onClick={() => setDialog({ kind: 'defaults' })} />
                <Button tall icon="import" label="Find my mail" onClick={() => setDialog({ kind: 'import' })} />
              </Group>
              <Group label="Mail">
                <Button tall icon="filter" label="Rules" onClick={() => setDialog({ kind: 'rules' })} />
              </Group>
              <Group label="Sending">
                <Button
                  tall
                  icon="clock"
                  label="Undo window"
                  onClick={async () => {
                    const now = await shell.store.get({ key: 'mail.undoSeconds', fallback: 8 });
                    setDialog({ kind: 'undo', seconds: now });
                  }}
                />
                <Button
                  tall
                  icon="word"
                  label="Signature"
                  onClick={async () => {
                    const text = await shell.store.get({ key: 'mail.signature', fallback: '' });
                    setDialog({ kind: 'signature', text });
                  }}
                />
              </Group>
              <Group label="Privacy">
                <Button icon="shield" label="Always block" pressed disabled title="Remote content is blocked in every message until you ask for it." />
              </Group>
            </>
          )}
        </Ribbon>
      }
      status={
        <>
          <span>{unified ? `${accounts.length} accounts` : account ? account.email : 'No account yet'}</span>
          {progress ? (
            <>
              <Spinner />
              <span>
                {progress.phase === 'import' ? 'Importing' : 'Fetching'} {progress.folder} — {progress.done.toLocaleString()}
              </span>
            </>
          ) : null}
          <Spacer />
          {checked.size ? <Chip>{checked.size} selected</Chip> : null}
          <Chip>{list.total.toLocaleString()} messages</Chip>
          {unreadTotal ? <Chip>{unreadTotal.toLocaleString()} unread</Chip> : null}
        </>
      }
    >
      {!accounts.length ? (
        <Empty icon="mail" title="No mail here yet">
          Add an account to fetch mail, or import what you already have — an Outlook .pst or .ost, an mbox, or a
          folder of messages. Rutba Office can find them for you.
          <div style={{ display: 'flex', gap: 8, marginTop: 14 }}>
            <Button primary icon="plus" label="Add account" onClick={() => setDialog({ kind: 'account' })} />
            <Button icon="import" label="Find my mail" onClick={() => setDialog({ kind: 'import' })} />
          </div>
        </Empty>
      ) : (
        <>
          <Panel width={222} resizable>
            <div className="ml-compose-cta">
              <button type="button" onClick={() => commands['mail.compose'].run()}>
                <Icon name="new" size={15} />
                <span>Compose</span>
              </button>
            </div>

            <div className="ml-accounts">
              {accounts.length > 1 ? (
                <Item
                  icon="inbox"
                  label="All accounts"
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
                  title={a.email}
                />
              ))}
              <Item icon="plus" label="Add account" onClick={() => setDialog({ kind: 'account' })} />
            </div>

            <div className="rw-panel-head">{unified ? 'Across every account' : 'Folders'}</div>
            <List>
              {unified
                ? [
                    { path: 'inbox', name: 'Inbox', icon: 'inbox' },
                    { path: 'sent', name: 'Sent', icon: 'send' },
                    { path: 'drafts', name: 'Drafts', icon: 'file' },
                    { path: 'archive', name: 'Archive', icon: 'archive' },
                    { path: 'junk', name: 'Junk', icon: 'spam' },
                    { path: 'trash', name: 'Trash', icon: 'trash' },
                    { path: 'all', name: 'Everything', icon: 'grid' },
                  ].map((f) => (
                    <Item key={f.path} icon={f.icon} label={f.name} current={role === f.path} onClick={() => setRole(f.path)} />
                  ))
                : folders.map((f) => (
                    <Item
                      key={f.path}
                      icon={f.icon || 'folder'}
                      label={f.name || f.path}
                      indent={Math.min(f.depth || 0, 4)}
                      current={f.path === folder && view === 'mail'}
                      count={f.unread || undefined}
                      onClick={() => {
                        setFolder(f.path);
                        setView('mail');
                      }}
                      title={f.path}
                    />
                  ))}
            </List>

            {!unified ? (
              <>
                <div className="rw-panel-head">Views</div>
                <List>
                  <Item icon="attach" label="Attachments" current={view === 'files'} onClick={() => setView('files')} />
                  <Item icon="reply" label="People" current={view === 'people'} onClick={() => setView('people')} />
                </List>
              </>
            ) : null}
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
                width={prefs.layout === 'bottom' ? undefined : 380}
                resizable={prefs.layout !== 'bottom'}
                style={{
                  background: 'var(--surface)',
                  ...(prefs.layout === 'bottom' ? { flex: 1 } : {}),
                  ...(prefs.layout === 'off' ? { flex: 1 } : {}),
                }}
              >
                <div className="ml-tools">
                  <Search value={query} onChange={setQuery} placeholder={unified ? 'Search every account' : `Search ${current?.name || 'mail'}`} />
                  <div className="ml-filters">
                    {FILTERS.filter((f) => f.id !== 'people').map((f) => (
                      <button key={f.id} type="button" className={`ml-filter${prefs.filter === f.id ? ' on' : ''}`} onClick={() => setPref({ filter: f.id })}>
                        {f.icon ? <Icon name={f.icon} size={12} /> : null}
                        <span>{f.label}</span>
                      </button>
                    ))}
                  </div>
                </div>

                {checked.size ? (
                  <div className="ml-bulk">
                    <input
                      type="checkbox"
                      className="ml-check"
                      checked
                      onChange={() => setChecked(new Set())}
                      title="Clear the selection"
                    />
                    <strong>{checked.size}</strong>
                    <span>selected</span>
                    <Spacer />
                    <Button icon="archive" title="Archive" onClick={() => act('archive')} />
                    <Button icon="trash" title="Delete" onClick={() => act('delete')} />
                    <Button icon="spam" title="Junk" onClick={() => act('junk')} />
                    <Button icon="check" title="Mark read" onClick={() => act('read')} />
                    <Button icon="eye" title="Mark unread" onClick={() => act('unread')} />
                    <Button icon="star" title="Star" onClick={() => act('flag')} />
                    {!unified ? <Button icon="folder" title="Move to" onClick={moveMenu} /> : null}
                  </div>
                ) : (
                  <div className="ml-listbar">
                    <input
                      type="checkbox"
                      className="ml-check"
                      checked={false}
                      onChange={() => setChecked(new Set(threads.flatMap((t) => allRowsIn(t)).map(keyOf)))}
                      title="Select everything here"
                    />
                    <span>
                      {threads.length.toLocaleString()} {prefs.threaded ? 'conversations' : 'messages'}
                    </span>
                    <Spacer />
                    <Button icon="refresh" title="Get mail" onClick={sync} />
                  </div>
                )}

                {busy ? (
                  <div style={{ padding: 24, display: 'grid', placeItems: 'center' }}>
                    <Spinner />
                  </div>
                ) : threads.length ? (
                  <div className={`ml-list${prefs.density === 'compact' ? ' compact' : ''}`}>
                    {threads.map((thread) => (
                      <Row
                        key={thread.threadKey || keyOf(thread)}
                        thread={thread}
                        unified={unified}
                        threaded={prefs.threaded}
                        selected={Boolean(selected) && allRowsIn(thread).some((r) => keyOf(r) === keyOf(selected))}
                        checked={allRowsIn(thread).every((r) => checked.has(keyOf(r)))}
                        expanded={expanded.has(thread.threadKey)}
                        onOpen={(row) => setSelected({ accountId: row.accountId, folder: row.folder, id: row.id })}
                        onCheck={(e) => toggleCheck(thread, e)}
                        onStar={star}
                        onExpand={() =>
                          setExpanded((s) => {
                            const next = new Set(s);
                            if (next.has(thread.threadKey)) next.delete(thread.threadKey);
                            else next.add(thread.threadKey);
                            return next;
                          })
                        }
                        onMenu={(e) =>
                          menu.open(
                            e,
                            menuItems(commands, ['mail.reply', 'mail.replyAll', 'mail.forward', '-', 'mail.archive', 'mail.flag', 'mail.pin', 'mail.unread', '-', 'mail.junk', 'mail.delete'])
                          )
                        }
                      />
                    ))}
                  </div>
                ) : (
                  <Empty icon="inbox" title={query ? 'Nothing matches' : prefs.filter !== 'all' ? 'Nothing here under this filter' : 'This folder is empty'}>
                    {query ? 'Try a different search, or search every account.' : null}
                  </Empty>
                )}
              </Panel>

              {prefs.layout !== 'off' ? (
                <Content>
                  {message ? (
                    <Reader
                      message={message}
                      insight={insight}
                      remote={remote}
                      plain={plain}
                      dark={dark}
                      onRemote={() => setRemote(true)}
                      onSaveAttachment={saveAttachment}
                      onOpenAttachment={openAttachment}
                      onReply={() => reply(false)}
                      onReplyAll={() => reply(true)}
                      onForward={forward}
                      onUnsubscribe={unsubscribe}
                      onSender={(from) => from?.address && setQuery(from.address)}
                    />
                  ) : (
                    <Empty icon="mail" title="No message selected">
                      Choose a message to read it here. <kbd>j</kbd> and <kbd>k</kbd> move, <kbd>r</kbd> replies,
                      <kbd>s</kbd> stars.
                    </Empty>
                  )}
                </Content>
              ) : null}
            </>
          )}
        </>
      )}

      {menu.node}

      {/* the undo window ------------------------------------------------- */}
      {outbox.length ? (
        <div className="ml-outbox">
          <Icon name="send" size={14} />
          <span>
            {outbox.length === 1
              ? `Sending to ${outbox[0].draft?.to || 'recipient'}`
              : `${outbox.length} messages waiting to go`}
          </span>
          <button type="button" onClick={() => unsend(outbox[outbox.length - 1].id)}>
            Undo
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
          toast={toast}
        />
      ) : null}

      {dialog?.kind === 'import' ? (
        <ImportDialog
          scan={scan}
          onClose={() => setDialog(null)}
          onChoose={chooseImport}
          onImportStore={(path) => importFrom(path)}
          onUseAccount={(found) => setDialog({ kind: 'account', seed: found })}
        />
      ) : null}

      {dialog?.kind === 'importing' ? <ImportingDialog path={dialog.path} progress={progress} /> : null}

      {dialog?.kind === 'import-preview' ? (
        <ImportPreview found={dialog.found} onCancel={() => setDialog(null)} onImport={(chosen) => runImport(dialog.path, chosen)} />
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
        <Dialog title="Waiting to go out" width={520} onClose={() => setDialog(null)} actions={<Button primary label="Close" onClick={() => setDialog(null)} />}>
          {outbox.length ? (
            <div className="ml-found">
              {outbox.map((item) => (
                <div key={item.id} className="ml-found-item" style={{ cursor: 'default' }}>
                  <span className="ml-found-logo">
                    <Icon name="clock" size={15} />
                  </span>
                  <span className="grow">
                    <div className="who">{item.draft?.subject || '(no subject)'}</div>
                    <div className="what">
                      To {item.draft?.to} · {new Date(item.at) > new Date() ? `goes out ${new Date(item.at).toLocaleString()}` : 'going now'}
                      {item.error ? ` · last error: ${item.error}` : ''}
                    </div>
                  </span>
                  <Button label="Take back" onClick={() => unsend(item.id)} />
                </div>
              ))}
            </div>
          ) : (
            <p style={{ marginTop: 0 }}>Nothing is waiting.</p>
          )}
        </Dialog>
      ) : null}

      {dialog?.kind === 'signature' ? (
        <Dialog
          title="Signature"
          width={520}
          onClose={() => setDialog(null)}
          actions={
            <>
              <Button label="Cancel" onClick={() => setDialog(null)} />
              <Button
                primary
                label="Save"
                onClick={async () => {
                  await shell.store.set({ key: 'mail.signature', value: dialog.text });
                  toast('Signature saved', { tone: 'good' });
                  setDialog(null);
                }}
              />
            </>
          }
        >
          <Field label="Added to the bottom of everything you write">
            <textarea
              className="rw-input ml-compose-body"
              rows={7}
              value={dialog.text}
              onChange={(e) => setDialog((d) => ({ ...d, text: e.target.value }))}
            />
          </Field>
        </Dialog>
      ) : null}

      {dialog?.kind === 'undo' ? (
        <Dialog
          title="Undo window"
          width={460}
          onClose={() => setDialog(null)}
          actions={
            <>
              <Button label="Cancel" onClick={() => setDialog(null)} />
              <Button
                primary
                label="Save"
                onClick={async () => {
                  await shell.store.set({ key: 'mail.undoSeconds', value: Number(dialog.seconds) || 0 });
                  toast(Number(dialog.seconds) ? `${dialog.seconds} seconds to change your mind` : 'Messages now send immediately', { tone: 'good' });
                  setDialog(null);
                }}
              />
            </>
          }
        >
          <Field label="Hold a message for" hint="Every message waits this long before it leaves, and can be taken back until then.">
            <Select value={String(dialog.seconds)} onChange={(e) => setDialog((d) => ({ ...d, seconds: Number(e.target.value) }))}>
              <option value="0">No delay</option>
              <option value="5">5 seconds</option>
              <option value="8">8 seconds</option>
              <option value="15">15 seconds</option>
              <option value="30">30 seconds</option>
            </Select>
          </Field>
        </Dialog>
      ) : null}

      {compose ? (
        <Compose
          draft={compose}
          accounts={accounts}
          accountId={compose.accountId || (unified ? accounts[0]?.id : accountId)}
          onAccount={(id) => setCompose((d) => ({ ...d, accountId: id }))}
          onChange={setCompose}
          onClose={() => setCompose(null)}
          onSend={doSend}
          onSaveDraft={async (draft) => {
            await shell.mail.saveDraft({ accountId: draft.accountId || accountId, draft });
            toast('Saved to Drafts', { tone: 'good' });
            setCompose(null);
            refreshList();
          }}
          shell={shell}
          toast={toast}
        />
      ) : null}
    </AppFrame>
  );
}

/* ── one line of the list ────────────────────────────────────────────────── */

/**
 * Memoised because a 500-row list re-renders on every keystroke in the search
 * box otherwise, and the whole point of holding the mail locally is that the
 * list is instant.
 */
const Row = React.memo(function Row({ thread, unified, threaded, selected, checked, expanded, onOpen, onCheck, onStar, onExpand, onMenu }) {
  const avatar = avatarFor(thread.from);
  const many = threaded && thread.count > 1;

  return (
    <>
      <div
        className={`ml-row${selected ? ' selected' : ''}${thread.unread ? ' unread' : ''}${checked ? ' checked' : ''}${thread.pinned ? ' pinned' : ''}`}
        role="button"
        tabIndex={0}
        onClick={() => onOpen(thread)}
        onKeyDown={(e) => (e.key === 'Enter' || e.key === ' ') && (e.preventDefault(), onOpen(thread))}
        onContextMenu={onMenu}
      >
        <input
          type="checkbox"
          className="ml-check"
          checked={checked}
          style={{ marginTop: 8 }}
          onClick={(e) => e.stopPropagation()}
          onChange={(e) => onCheck(e.nativeEvent)}
        />

        <span className="ml-avatar-sm" style={{ background: avatar.colour }} title={thread.from?.address || ''}>
          {avatar.initial}
        </span>

        <div className="ml-row-main">
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
                title={`${thread.count} messages in this conversation`}
              >
                {thread.count}
              </button>
            ) : null}
          </div>
          <div className="ml-subject">
            {thread.pinned ? <Icon name="flag" size={12} /> : null}
            {thread.subject || '(no subject)'}
          </div>
          <div className="ml-preview">
            {thread.hasAttachments ? <Icon name="attach" size={12} /> : null}
            {thread.preview}
          </div>
        </div>

        <div className="ml-row-side">
          <span className="ml-when">{formatWhen(thread.date)}</span>
          <span className="ml-marks">
            {unified ? <span className="ml-folder-tag">{String(thread.folder || '').split('/').pop()}</span> : null}
            <button
              type="button"
              className={`ml-star${thread.flagged ? ' on' : ''}`}
              title={thread.flagged ? 'Remove star' : 'Star'}
              onClick={(e) => onStar(thread, e)}
            >
              <Icon name="star" size={13} />
            </button>
          </span>
        </div>
      </div>

      {many && expanded ? (
        <div className="ml-thread-children">
          {thread.messages
            .slice()
            .reverse()
            .map((m) => (
              <div key={m.id} className={`ml-row${m.unread ? ' unread' : ''}`} role="button" tabIndex={0} onClick={() => onOpen(m)} onKeyDown={(e) => e.key === 'Enter' && onOpen(m)}>
                <span />
                <span className="ml-avatar-sm" style={{ background: avatarFor(m.from).colour, width: 22, height: 22, fontSize: 10 }}>
                  {avatarFor(m.from).initial}
                </span>
                <div className="ml-row-main">
                  <div className="ml-row-top">
                    <span className="ml-from">{displayName(m.from)}</span>
                  </div>
                  <div className="ml-preview">{m.preview}</div>
                </div>
                <div className="ml-row-side">
                  <span className="ml-when">{formatWhen(m.date)}</span>
                </div>
              </div>
            ))}
        </div>
      ) : null}
    </>
  );
});
