// Rutba Mail.
//
// Three panes and one rule: a message body is never trusted. It is rendered in
// a sandboxed frame with no scripts, no same-origin access and a policy that
// blocks every remote fetch, so a tracking pixel cannot report that you opened
// the mail and a script cannot reach anything at all. Remote images load only
// when you ask, per message.
//
// The other half of this app is the part no free client does well: opening what
// you already have. A .pst, a .ost, an Outlook for Mac archive, an mbox, a
// folder of .eml files — scanned first so you can see what is in there, then
// imported into the same local store live accounts synchronise into.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  Ribbon, Group, Button, Separator, Icon, Spacer, Chip, Empty, Spinner, Panel, Content, List, Item,
  Search, Dialog, Field, Input, Select, Progress, useToast, useMenu, useCommands, menuItems,
  formatWhen, formatBytes,
} from '@rutba/office-ui';
import { AppFrame, useAppMenu, useFileDrop } from '../shell.js';

/* ── the reading frame ───────────────────────────────────────────────────── */

const BLOCKED_NOTE = 'blocked-remote';

/**
 * Build the document shown in the sandboxed frame.
 *
 * Everything remote is stripped unless the reader asked for it; `cid:` images
 * are rewritten to the URLs the backend is holding for this message, so inline
 * pictures work with no network at all.
 */
function bodyDocument(message, { remote = false, dark = false } = {}) {
  let html = message.html || '';
  if (!html && message.text) {
    html = `<pre class="plain">${message.text.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]))}</pre>`;
  }

  // Inline images first — these are ours and are safe.
  for (const [cid, url] of Object.entries(message.inlineUrls || {})) {
    html = html.split(`cid:${cid}`).join(url);
  }

  if (!remote) {
    html = html.replace(/(<img\b[^>]*?\bsrc=)(["'])(https?:[^"']*)\2/gi, `$1$2$2 data-${BLOCKED_NOTE}=$2$3$2`);
  }

  const csp = remote
    ? "default-src 'none'; img-src https: data: rutba: blob:; style-src 'unsafe-inline'; font-src data:"
    : "default-src 'none'; img-src data: rutba: blob:; style-src 'unsafe-inline'; font-src data:";

  return `<!doctype html><html><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="${csp}">
<style>
  :root { color-scheme: ${dark ? 'dark' : 'light'}; }
  body {
    margin: 0; padding: 18px 22px;
    font: 14px/1.55 "Segoe UI", system-ui, -apple-system, Roboto, Arial, sans-serif;
    color: ${dark ? '#e6e9ee' : '#1a1c20'}; background: ${dark ? '#1e2127' : '#ffffff'};
    word-wrap: break-word; overflow-wrap: anywhere;
  }
  img { max-width: 100%; height: auto; }
  pre.plain { white-space: pre-wrap; font: inherit; margin: 0; }
  blockquote { margin: 0 0 0 12px; padding-left: 12px; border-left: 3px solid ${dark ? '#3b4048' : '#d5d9e0'}; color: ${dark ? '#a9b0bb' : '#5b626d'}; }
  a { color: ${dark ? '#7fb0ff' : '#1a56c4'}; }
  table { max-width: 100%; }
</style></head><body>${html}</body></html>`;
}

/* ── the app ─────────────────────────────────────────────────────────────── */

export default function Mail({ app, shell }) {
  const toast = useToast();
  const [accounts, setAccounts] = useState([]);
  const [accountId, setAccountId] = useState(null);
  const [folders, setFolders] = useState([]);
  const [folder, setFolder] = useState(null);
  const [list, setList] = useState({ rows: [], total: 0 });
  const [selected, setSelected] = useState(null);
  const [message, setMessage] = useState(null);
  const [query, setQuery] = useState('');
  const [tab, setTab] = useState('home');
  const [remote, setRemote] = useState(false);
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState(null);
  const [dialog, setDialog] = useState(() => (new URLSearchParams(location.search).has('import') ? { kind: 'import' } : null));
  const [compose, setCompose] = useState(null);
  const menu = useMenu();
  const appMenu = useAppMenu({
    shell,
    appKey: 'mail',
    extra: [
      { label: 'Import mail…', icon: 'import', run: () => setDialog({ kind: 'import' }) },
      { label: 'Add account…', icon: 'plus', run: () => setDialog({ kind: 'account' }) },
    ],
  });

  const dark = document.documentElement.dataset.theme === 'dark';

  /* ── loading ──────────────────────────────────────────────────────────── */

  const loadAccounts = useCallback(async () => {
    const list = await shell.mail.accounts();
    setAccounts(list);
    setAccountId((current) => current || list[0]?.id || null);
    return list;
  }, [shell]);

  useEffect(() => {
    loadAccounts().catch(() => {});
  }, [loadAccounts]);

  useEffect(() => {
    if (!accountId) return;
    shell.mail
      .folders({ accountId })
      .then((f) => {
        setFolders(f);
        setFolder((current) => (f.some((x) => x.path === current) ? current : f.find((x) => x.role === 'inbox')?.path || f[0]?.path || null));
      })
      .catch(() => setFolders([]));
  }, [accountId, shell]);

  const refreshList = useCallback(async () => {
    if (!accountId || !folder) {
      setList({ rows: [], total: 0 });
      return;
    }
    setBusy(true);
    try {
      const next = await shell.mail.messages({ accountId, folder, query, limit: 300 });
      setList(next);
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    } finally {
      setBusy(false);
    }
  }, [accountId, folder, query, shell, toast]);

  useEffect(() => {
    refreshList();
    setSelected(null);
    setMessage(null);
  }, [refreshList]);

  useEffect(() => {
    if (!selected || !accountId || !folder) {
      setMessage(null);
      return;
    }
    setRemote(false);
    shell.mail
      .message({ accountId, folder, id: selected })
      .then((m) => {
        setMessage(m);
        if (m?.unread) {
          shell.mail.flag({ accountId, folder, ids: [selected], patch: { unread: false } }).then(() => refreshList());
        }
      })
      .catch((err) => toast(err.message, { tone: 'bad' }));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selected, accountId, folder]);

  useEffect(() => {
    const off = shell.on('mail:progress', (p) => setProgress(p));
    const offNew = shell.on('mail:new', () => {
      setProgress(null);
      loadAccounts();
      refreshList();
    });
    return () => {
      off?.();
      offNew?.();
    };
  }, [shell, loadAccounts, refreshList]);

  /* ── actions ──────────────────────────────────────────────────────────── */

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
    async (target, folders) => {
      setDialog({ kind: 'importing', path: target });
      try {
        const result = await shell.mail.import({ path: target, folders });
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

  const sync = useCallback(async () => {
    if (!accountId) return;
    setBusy(true);
    try {
      const result = await shell.mail.sync({ accountId, folder });
      toast(result.local ? 'This is an imported archive — there is nothing to fetch.' : `${result.added} new`, { tone: result.local ? 'plain' : 'good' });
      await refreshList();
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 6000 });
    } finally {
      setBusy(false);
      setProgress(null);
    }
  }, [accountId, folder, shell, refreshList, toast]);

  const act = useCallback(
    async (what) => {
      if (!selected) return;
      const ids = [selected];
      try {
        if (what === 'delete') {
          const trash = folders.find((f) => f.role === 'trash');
          if (trash && folder !== trash.path) await shell.mail.move({ accountId, folder, ids, to: trash.path });
          else await shell.mail.delete({ accountId, folder, ids });
        } else if (what === 'archive') {
          await shell.mail.move({ accountId, folder, ids, to: folders.find((f) => f.role === 'archive')?.path || 'Archive' });
        } else if (what === 'flag') {
          await shell.mail.flag({ accountId, folder, ids, patch: { flagged: !message?.flagged } });
        } else if (what === 'unread') {
          await shell.mail.flag({ accountId, folder, ids, patch: { unread: true } });
        }
        setSelected(null);
        await refreshList();
      } catch (err) {
        toast(err.message, { tone: 'bad' });
      }
    },
    [selected, folder, folders, accountId, message, shell, refreshList, toast]
  );

  const reply = useCallback(
    (all = false) => {
      if (!message) return;
      const to = [message.from?.[0]?.address, ...(all ? (message.to || []).map((t) => t.address) : [])].filter(Boolean).join(', ');
      setCompose({
        to,
        subject: message.subject?.startsWith('Re:') ? message.subject : `Re: ${message.subject || ''}`,
        text: `\n\n---\nOn ${formatWhen(message.date, { long: true })}, ${message.from?.[0]?.display || 'someone'} wrote:\n${(message.text || '').split('\n').map((l) => `> ${l}`).join('\n')}`,
      });
    },
    [message]
  );

  const saveAttachment = useCallback(
    async (index) => {
      const held = await shell.mail.attachment({ accountId, folder, id: selected, index });
      if (!held) return toast('That attachment is not stored.', { tone: 'bad' });
      const target = await shell.dialog.save({ title: 'Save attachment', defaultPath: held.name });
      if (!target) return undefined;
      const response = await fetch(held.url);
      const bytes = new Uint8Array(await response.arrayBuffer());
      await shell.fs.write({ path: target, bytes });
      toast(`Saved ${held.name}`, { tone: 'good' });
      return undefined;
    },
    [accountId, folder, selected, shell, toast]
  );

  const commands = useMemo(
    () => ({
      'mail.sync': { label: 'Get mail', icon: 'refresh', key: 'Mod+R', run: sync },
      'mail.compose': { label: 'New message', icon: 'new', key: 'Mod+N', run: () => setCompose({ to: '', subject: '', text: '' }) },
      'mail.reply': { label: 'Reply', icon: 'reply', key: 'Mod+Shift+R', run: () => reply(false) },
      'mail.replyAll': { label: 'Reply all', icon: 'replyAll', run: () => reply(true) },
      'mail.forward': { label: 'Forward', icon: 'forward', run: () => message && setCompose({ to: '', subject: `Fwd: ${message.subject || ''}`, text: message.text || '' }) },
      'mail.archive': { label: 'Archive', icon: 'archive', key: 'e', run: () => act('archive') },
      'mail.delete': { label: 'Delete', icon: 'trash', key: 'Delete', run: () => act('delete') },
      'mail.flag': { label: 'Flag', icon: 'flag', run: () => act('flag') },
      'mail.unread': { label: 'Mark unread', icon: 'eye', run: () => act('unread') },
      'mail.import': { label: 'Import mail…', icon: 'import', run: chooseImport },
      'mail.account': { label: 'Add account…', icon: 'plus', run: () => setDialog({ kind: 'account' }) },
    }),
    [sync, reply, act, chooseImport, message]
  );

  useCommands(commands, [selected, message, folder, accountId]);

  const account = accounts.find((a) => a.id === accountId);
  const current = folders.find((f) => f.path === folder);

  return (
    <AppFrame
      app={app}
      shell={shell}
      title={account ? `${current?.name || 'Mail'} — ${account.name || account.email}` : 'Mail'}
      menu={appMenu}
      ribbon={
        <Ribbon
          tabs={[
            { id: 'home', label: 'Home' },
            { id: 'folder', label: 'Folder' },
            { id: 'view', label: 'View' },
          ]}
          active={tab}
          onTab={setTab}
          quick={
            <>
              <Button icon="refresh" title="Get mail" onClick={sync} disabled={!accountId} />
              <Button icon="new" title="New message" onClick={() => commands['mail.compose'].run()} disabled={!accountId} />
            </>
          }
        >
          {tab === 'home' ? (
            <>
              <Group label="New">
                <Button tall icon="new" label="Message" onClick={() => commands['mail.compose'].run()} disabled={!accountId} />
              </Group>
              <Group label="Get">
                <Button tall icon="refresh" label="Get mail" onClick={sync} disabled={!accountId} />
                <Button tall icon="import" label="Import" onClick={chooseImport} />
              </Group>
              <Group label="Respond">
                <Button tall icon="reply" label="Reply" disabled={!message} onClick={() => reply(false)} />
                <Button tall icon="replyAll" label="Reply all" disabled={!message} onClick={() => reply(true)} />
                <Button tall icon="forward" label="Forward" disabled={!message} onClick={() => commands['mail.forward'].run()} />
              </Group>
              <Group label="Manage">
                <Button icon="archive" label="Archive" disabled={!selected} onClick={() => act('archive')} />
                <Button icon="trash" label="Delete" disabled={!selected} onClick={() => act('delete')} />
                <Button icon="flag" label="Flag" disabled={!selected} onClick={() => act('flag')} />
                <Button icon="eye" label="Unread" disabled={!selected} onClick={() => act('unread')} />
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
              <Group label="Export">
                <Button
                  tall
                  icon="export"
                  label="Mbox"
                  disabled={!folder}
                  onClick={async () => {
                    const target = await shell.dialog.save({ title: 'Export folder', defaultPath: `${current?.name || 'folder'}.mbox` });
                    if (!target) return;
                    const r = await shell.mail.export({ accountId, folder, path: target, format: 'mbox' });
                    toast(`Exported ${r.messages} messages`, { tone: 'good' });
                  }}
                />
              </Group>
            </>
          ) : (
            <>
              <Group label="Message">
                <Button icon="eye" label="Remote images" pressed={remote} onClick={() => setRemote((r) => !r)} />
              </Group>
              <Group label="Zoom">
                <Button icon="zoomOut" label="Out" onClick={() => shell.win.zoom({ delta: -0.1 })} />
                <Button icon="zoomIn" label="In" onClick={() => shell.win.zoom({ delta: 0.1 })} />
              </Group>
            </>
          )}
        </Ribbon>
      }
      status={
        <>
          <span>{account ? account.email : 'No account yet'}</span>
          {progress ? (
            <>
              <Spinner />
              <span>
                {progress.phase === 'import' ? 'Importing' : 'Fetching'} {progress.folder} — {progress.done.toLocaleString()}
              </span>
            </>
          ) : null}
          <Spacer />
          <Chip>{list.total.toLocaleString()} messages</Chip>
          {current?.unread ? <Chip>{current.unread} unread</Chip> : null}
        </>
      }
    >
      {!accounts.length ? (
        <Empty icon="mail" title="No mail here yet">
          Add an account to fetch mail, or import what you already have — an Outlook .pst or .ost, an mbox, or a
          folder of messages.
          <div style={{ display: 'flex', gap: 8, marginTop: 14 }}>
            <Button primary icon="plus" label="Add account" onClick={() => setDialog({ kind: 'account' })} />
            <Button icon="import" label="Import mail" onClick={chooseImport} />
          </div>
        </Empty>
      ) : (
        <>
          <Panel width={214} resizable>
            <div className="ml-accounts">
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
            <div className="rw-panel-head">Folders</div>
            <List>
              {folders.map((f) => (
                <Item
                  key={f.path}
                  icon={f.icon || 'folder'}
                  label={f.name || f.path}
                  indent={Math.min(f.depth || 0, 4)}
                  current={f.path === folder}
                  count={f.unread || undefined}
                  onClick={() => setFolder(f.path)}
                  title={f.path}
                />
              ))}
            </List>
          </Panel>

          <Panel width={352} resizable style={{ background: 'var(--surface)' }}>
            <div style={{ padding: '8px 10px' }}>
              <Search value={query} onChange={setQuery} placeholder={`Search ${current?.name || 'mail'}`} />
            </div>
            {busy ? (
              <div style={{ padding: 24, display: 'grid', placeItems: 'center' }}>
                <Spinner />
              </div>
            ) : list.rows.length ? (
              <div className="ml-list">
                {list.rows.map((row) => (
                  <button
                    key={row.id}
                    type="button"
                    className={`ml-row${row.id === selected ? ' selected' : ''}${row.unread ? ' unread' : ''}`}
                    onClick={() => setSelected(row.id)}
                    onContextMenu={(e) => menu.open(e, menuItems(commands, ['mail.reply', 'mail.forward', '-', 'mail.archive', 'mail.flag', 'mail.unread', '-', 'mail.delete']))}
                  >
                    <div className="ml-row-top">
                      <span className="ml-from">{row.from?.name || row.from?.address || '(unknown sender)'}</span>
                      <span className="ml-when">{formatWhen(row.date)}</span>
                    </div>
                    <div className="ml-subject">
                      {row.flagged ? <Icon name="flag" size={12} /> : null}
                      {row.subject || '(no subject)'}
                    </div>
                    <div className="ml-preview">
                      {row.hasAttachments ? <Icon name="attach" size={12} /> : null}
                      {row.preview}
                    </div>
                  </button>
                ))}
              </div>
            ) : (
              <Empty icon="inbox" title={query ? 'Nothing matches' : 'This folder is empty'}>
                {query ? 'Try a different search.' : null}
              </Empty>
            )}
          </Panel>

          <Content>
            {message ? (
              <Reader
                message={message}
                remote={remote}
                dark={dark}
                onRemote={() => setRemote(true)}
                onSaveAttachment={saveAttachment}
                onReply={() => reply(false)}
                onOpenLink={(url) => shell.shell.openExternal({ url })}
              />
            ) : (
              <Empty icon="mail" title="No message selected">
                Choose a message to read it here.
              </Empty>
            )}
          </Content>
        </>
      )}

      {menu.node}

      {dialog?.kind === 'account' ? (
        <AccountDialog shell={shell} onClose={() => setDialog(null)} onSaved={() => { setDialog(null); loadAccounts(); }} toast={toast} />
      ) : null}

      {dialog?.kind === 'import' ? (
        <Dialog
          title="Import mail"
          onClose={() => setDialog(null)}
          actions={<><Button label="Cancel" onClick={() => setDialog(null)} /><Button primary label="Choose a file…" onClick={chooseImport} /></>}
        >
          <p style={{ marginTop: 0 }}>Rutba Office reads the archives other mail clients leave behind.</p>
          <ul className="ml-formats">
            <li><strong>.pst</strong> and <strong>.ost</strong> — Outlook data files, with folders, attachments and read state</li>
            <li><strong>.olm</strong> — Outlook for Mac archives</li>
            <li><strong>.mbox</strong> — Thunderbird, Apple Mail, Google Takeout</li>
            <li><strong>.eml</strong>, <strong>.msg</strong> — single messages, or a folder full of them</li>
          </ul>
          <p className="rw-hint">Nothing is sent anywhere. The import reads the file and writes to this computer only.</p>
        </Dialog>
      ) : null}

      {dialog?.kind === 'importing' ? (
        <Dialog title="Importing" onClose={undefined}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10, padding: '6px 0 14px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
              <Spinner />
              <span>{progress ? `${progress.done.toLocaleString()} messages — ${progress.folder}` : 'Reading the archive…'}</span>
            </div>
            <Progress value={progress?.done ? (progress.done % 500) / 500 : 0.15} />
            <div className="rw-hint">{dialog.path}</div>
          </div>
        </Dialog>
      ) : null}

      {dialog?.kind === 'import-preview' ? (
        <ImportPreview
          found={dialog.found}
          onCancel={() => setDialog(null)}
          onImport={(folders) => runImport(dialog.path, folders)}
        />
      ) : null}

      {compose ? (
        <Compose
          draft={compose}
          account={account}
          onChange={setCompose}
          onClose={() => setCompose(null)}
          onSend={async (draft) => {
            try {
              await shell.mail.send({ accountId, draft });
              toast('Sent', { tone: 'good' });
              setCompose(null);
            } catch (err) {
              toast(err.message, { tone: 'bad', ms: 7000 });
            }
          }}
          onSaveDraft={async (draft) => {
            await shell.mail.saveDraft({ accountId, draft });
            toast('Saved to Drafts', { tone: 'good' });
            setCompose(null);
          }}
        />
      ) : null}
    </AppFrame>
  );
}

/* ── reading pane ────────────────────────────────────────────────────────── */

function Reader({ message, remote, dark, onRemote, onSaveAttachment, onReply, onOpenLink }) {
  const frameRef = useRef(null);
  const [height, setHeight] = useState(600);
  const doc = useMemo(() => bodyDocument(message, { remote, dark }), [message, remote, dark]);
  const blocked = !remote && /data-blocked-remote=/.test(doc);

  // The frame has no scripts, so its height is measured from outside once it
  // has laid out. A same-origin read is impossible under the sandbox, so the
  // frame is simply given room and allowed to scroll itself.
  useEffect(() => {
    setHeight(600);
    const el = frameRef.current;
    if (!el) return undefined;
    const ro = new ResizeObserver(() => setHeight(el.parentElement?.clientHeight || 600));
    ro.observe(el.parentElement);
    return () => ro.disconnect();
  }, [message]);

  return (
    <div className="ml-reader">
      <header className="ml-head">
        <h2>{message.subject || '(no subject)'}</h2>
        <div className="ml-meta">
          <span className="ml-avatar">{(message.from?.[0]?.name || message.from?.[0]?.address || '?').trim().charAt(0).toUpperCase()}</span>
          <div className="ml-meta-text">
            <div>
              <strong>{message.from?.[0]?.name || message.from?.[0]?.address || 'Unknown sender'}</strong>
              {message.from?.[0]?.name && message.from?.[0]?.address ? <span className="ml-addr"> &lt;{message.from[0].address}&gt;</span> : null}
            </div>
            <div className="ml-to">
              to {(message.to || []).map((t) => t.name || t.address).join(', ') || 'undisclosed recipients'}
              {message.cc?.length ? ` · cc ${message.cc.map((c) => c.name || c.address).join(', ')}` : ''}
            </div>
          </div>
          <span className="ml-date">{formatWhen(message.date, { long: true })}</span>
          <Button icon="reply" title="Reply" onClick={onReply} />
        </div>

        {message.attachments?.filter((a) => !a.inline).length ? (
          <div className="ml-attachments">
            {message.attachments.filter((a) => !a.inline).map((a, i) => (
              <button key={`${a.filename}-${i}`} type="button" className="ml-attachment" onClick={() => onSaveAttachment(message.attachments.indexOf(a))}>
                <Icon name="attach" size={14} />
                <span className="name">{a.filename}</span>
                <span className="size">{formatBytes(a.size)}</span>
                <Icon name="download" size={13} />
              </button>
            ))}
          </div>
        ) : null}

        {blocked ? (
          <div className="ml-blocked">
            <Icon name="shield" size={14} />
            <span>Remote images are blocked, so the sender cannot tell you opened this.</span>
            <Button label="Show images" onClick={onRemote} />
          </div>
        ) : null}
      </header>

      <div className="ml-body">
        {/*
          sandbox with no tokens: no scripts, no forms, no same-origin, no
          top-level navigation. The only thing this frame can do is draw.
        */}
        <iframe
          ref={frameRef}
          title="Message"
          sandbox=""
          referrerPolicy="no-referrer"
          srcDoc={doc}
          style={{ height }}
        />
      </div>

      {message.internetHeaders ? (
        <details className="ml-headers">
          <summary>Message headers</summary>
          <pre className="selectable">{message.internetHeaders}</pre>
        </details>
      ) : null}
      <div style={{ display: 'none' }} onClick={() => onOpenLink('')} />
    </div>
  );
}

/* ── import preview ──────────────────────────────────────────────────────── */

function ImportPreview({ found, onCancel, onImport }) {
  const [chosen, setChosen] = useState(() => new Set((found.folders || []).filter((f) => f.messages > 0).map((f) => f.path || f.name)));
  const total = (found.folders || []).filter((f) => chosen.has(f.path || f.name)).reduce((n, f) => n + f.messages, 0);

  if (!found.readable) {
    return (
      <Dialog title="This file cannot be read" onClose={onCancel} actions={<Button primary label="Close" onClick={onCancel} />}>
        <p style={{ marginTop: 0 }}>{found.label}</p>
        <p className="rw-hint">{found.encoding ? `It uses the ${found.encoding} encoding, which this build cannot decode.` : 'The format was not recognised.'}</p>
      </Dialog>
    );
  }

  return (
    <Dialog
      title={`Import from ${found.label}`}
      width={560}
      onClose={onCancel}
      actions={
        <>
          <Button label="Cancel" onClick={onCancel} />
          <Button primary label={`Import ${total.toLocaleString()} messages`} disabled={!total} onClick={() => onImport([...chosen])} />
        </>
      }
    >
      <div className="ml-import-summary">
        <Chip>{formatBytes(found.bytes || 0)}</Chip>
        <Chip>{(found.folders || []).length} folders</Chip>
        <Chip>{(found.messages || 0).toLocaleString()} messages</Chip>
        {found.encoding ? <Chip>{found.encoding}</Chip> : null}
      </div>
      <div className="ml-import-folders">
        {(found.folders || []).map((f) => {
          const key = f.path || f.name;
          const on = chosen.has(key);
          return (
            <label key={key} className="ml-import-folder" style={{ paddingLeft: 8 + Math.min(f.depth || 0, 5) * 13 }}>
              <input
                type="checkbox"
                checked={on}
                onChange={() =>
                  setChosen((s) => {
                    const next = new Set(s);
                    if (on) next.delete(key);
                    else next.add(key);
                    return next;
                  })
                }
              />
              <Icon name="folder" size={14} />
              <span className="name">{f.name}</span>
              <span className="count">{f.messages.toLocaleString()}</span>
            </label>
          );
        })}
      </div>
      <p className="rw-hint">Importing the same archive twice adds nothing: a message is identified by its own contents.</p>
    </Dialog>
  );
}

/* ── account setup ───────────────────────────────────────────────────────── */

function AccountDialog({ shell, onClose, onSaved, toast }) {
  const [form, setForm] = useState({ email: '', name: '', password: '', imapHost: '', imapPort: 993, imapSecure: true, smtpHost: '', smtpPort: 465, smtpSecure: true });
  const [note, setNote] = useState(null);
  const [testing, setTesting] = useState(false);
  const [result, setResult] = useState(null);
  const set = (patch) => setForm((f) => ({ ...f, ...patch }));

  const discover = useCallback(
    async (email) => {
      if (!email.includes('@')) return;
      const found = await shell.mail.autodiscover({ email });
      if (found.imap) {
        set({
          imapHost: found.imap.host,
          imapPort: found.imap.port,
          imapSecure: found.imap.secure,
          smtpHost: found.smtp.host,
          smtpPort: found.smtp.port,
          smtpSecure: found.smtp.secure,
        });
      }
      setNote(found.note);
    },
    [shell]
  );

  const account = () => ({
    email: form.email,
    name: form.name || form.email,
    imap: { host: form.imapHost, port: Number(form.imapPort), secure: Boolean(form.imapSecure) },
    smtp: { host: form.smtpHost, port: Number(form.smtpPort), secure: Boolean(form.smtpSecure) },
  });

  return (
    <Dialog
      title="Add a mail account"
      width={520}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button
            label={testing ? 'Testing…' : 'Test'}
            disabled={testing || !form.email || !form.password}
            onClick={async () => {
              setTesting(true);
              try {
                setResult(await shell.mail.testAccount({ account: account(), password: form.password }));
              } finally {
                setTesting(false);
              }
            }}
          />
          <Button
            primary
            label="Add"
            disabled={!form.email || !form.imapHost}
            onClick={async () => {
              try {
                await shell.mail.addAccount({ account: account(), password: form.password });
                onSaved();
              } catch (err) {
                toast(err.message, { tone: 'bad', ms: 7000 });
              }
            }}
          />
        </>
      }
    >
      <div className="ml-form">
        <Field label="Email address">
          <Input value={form.email} onChange={(e) => set({ email: e.target.value })} onBlur={(e) => discover(e.target.value)} placeholder="you@example.com" />
        </Field>
        <Field label="Your name">
          <Input value={form.name} onChange={(e) => set({ name: e.target.value })} placeholder="How your name appears on messages you send" />
        </Field>
        <Field label="Password" hint="Kept in this computer's keystore, never in a file you can read.">
          <Input type="password" value={form.password} onChange={(e) => set({ password: e.target.value })} />
        </Field>
        {note ? <div className="ml-note"><Icon name="info" size={14} />{note}</div> : null}
        <div className="ml-servers">
          <Field label="IMAP server">
            <Input value={form.imapHost} onChange={(e) => set({ imapHost: e.target.value })} placeholder="imap.example.com" />
          </Field>
          <Field label="Port">
            <Input type="number" value={form.imapPort} onChange={(e) => set({ imapPort: e.target.value })} />
          </Field>
          <Field label="SMTP server">
            <Input value={form.smtpHost} onChange={(e) => set({ smtpHost: e.target.value })} placeholder="smtp.example.com" />
          </Field>
          <Field label="Port">
            <Input type="number" value={form.smtpPort} onChange={(e) => set({ smtpPort: e.target.value })} />
          </Field>
        </div>
        {result ? (
          <div className={`ml-result${result.error ? ' bad' : ' good'}`}>
            <Icon name={result.error ? 'info' : 'check'} size={14} />
            <span>
              IMAP {result.imap?.ok ? `connected — ${result.imap.folders} folders` : `failed: ${result.imap?.message}`}
              {' · '}
              SMTP {result.smtp?.ok ? 'connected' : `failed: ${result.smtp?.message}`}
            </span>
          </div>
        ) : null}
      </div>
    </Dialog>
  );
}

/* ── compose ─────────────────────────────────────────────────────────────── */

function Compose({ draft, account, onChange, onClose, onSend, onSaveDraft }) {
  return (
    <Dialog
      title="New message"
      width={640}
      onClose={onClose}
      actions={
        <>
          <Button label="Discard" onClick={onClose} />
          <Button label="Save draft" onClick={() => onSaveDraft(draft)} />
          <Button primary icon="send" label="Send" disabled={!draft.to} onClick={() => onSend(draft)} />
        </>
      }
    >
      <div className="ml-form">
        <Field label="From">
          <Input value={account ? `${account.name} <${account.email}>` : ''} disabled />
        </Field>
        <Field label="To">
          <Input autoFocus value={draft.to} onChange={(e) => onChange({ ...draft, to: e.target.value })} placeholder="someone@example.com" />
        </Field>
        <Field label="Subject">
          <Input value={draft.subject} onChange={(e) => onChange({ ...draft, subject: e.target.value })} />
        </Field>
        <Field label="Message">
          <textarea
            className="rw-input ml-compose-body"
            value={draft.text}
            onChange={(e) => onChange({ ...draft, text: e.target.value })}
            rows={12}
          />
        </Field>
      </div>
    </Dialog>
  );
}

const CSS = `
.ml-accounts { padding: 6px 6px 2px; border-bottom: 1px solid var(--line-soft); }
.ml-list { display: flex; flex-direction: column; }
.ml-row {
  display: flex; flex-direction: column; gap: 2px; padding: 8px 12px; text-align: left;
  border: 0; border-bottom: 1px solid var(--line-soft); background: transparent; color: var(--ink);
  font: inherit; width: 100%; position: relative; transition: background var(--fast);
}
.ml-row:hover { background: var(--hover); }
.ml-row.selected { background: var(--selected); }
.ml-row.unread::before {
  content: ''; position: absolute; left: 4px; top: 15px; width: 5px; height: 5px;
  border-radius: 50%; background: var(--accent);
}
.ml-row-top { display: flex; align-items: baseline; gap: 8px; }
.ml-from { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 12.5px; }
.ml-row.unread .ml-from { font-weight: 700; }
.ml-when { font-size: 11px; color: var(--ink-3); flex: none; }
.ml-subject {
  font-size: 12.5px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  display: flex; align-items: center; gap: 5px;
}
.ml-row.unread .ml-subject { font-weight: 600; }
.ml-preview {
  font-size: 11.5px; color: var(--ink-3); overflow: hidden; text-overflow: ellipsis;
  white-space: nowrap; display: flex; align-items: center; gap: 5px;
}

.ml-reader { flex: 1; display: flex; flex-direction: column; min-height: 0; background: var(--surface); }
.ml-head { padding: 16px 22px 10px; border-bottom: 1px solid var(--line); }
.ml-head h2 { margin: 0 0 12px; font-family: var(--font-display); font-size: 18px; font-weight: 600; }
.ml-meta { display: flex; align-items: center; gap: 11px; }
.ml-avatar {
  width: 34px; height: 34px; border-radius: 50%; background: var(--accent); color: #fff;
  display: grid; place-items: center; font-weight: 600; flex: none;
}
.ml-meta-text { flex: 1; min-width: 0; font-size: 12.5px; }
.ml-addr { color: var(--ink-3); }
.ml-to { color: var(--ink-3); font-size: 11.5px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ml-date { color: var(--ink-3); font-size: 11.5px; flex: none; }
.ml-attachments { display: flex; flex-wrap: wrap; gap: 7px; margin-top: 12px; }
.ml-attachment {
  display: inline-flex; align-items: center; gap: 7px; padding: 5px 10px;
  border: 1px solid var(--line); border-radius: 20px; background: var(--chrome);
  color: var(--ink); font: inherit; font-size: 12px;
}
.ml-attachment:hover { border-color: var(--accent-line); background: var(--selected); }
.ml-attachment .size { color: var(--ink-3); font-size: 11px; }
.ml-blocked {
  display: flex; align-items: center; gap: 9px; margin-top: 11px; padding: 7px 11px;
  background: var(--sunken); border-radius: var(--r-2); font-size: 12px; color: var(--ink-2);
}
.ml-blocked .rw-btn { margin-left: auto; }
.ml-body { flex: 1; min-height: 0; }
.ml-body iframe { width: 100%; border: 0; display: block; }
.ml-headers { border-top: 1px solid var(--line); padding: 8px 22px; font-size: 11.5px; color: var(--ink-3); }
.ml-headers pre { white-space: pre-wrap; max-height: 220px; overflow: auto; font-family: var(--mono); font-size: 11px; }

.ml-form { display: flex; flex-direction: column; gap: 11px; padding: 4px 0 10px; }
.ml-servers { display: grid; grid-template-columns: 1fr 90px; gap: 10px; }
.ml-note, .ml-result {
  display: flex; align-items: center; gap: 8px; padding: 7px 11px; border-radius: var(--r-2);
  background: var(--sunken); font-size: 12px; color: var(--ink-2);
}
.ml-result.good { color: var(--good); }
.ml-result.bad { color: var(--bad); }
.ml-compose-body { font-family: var(--font); resize: vertical; min-height: 180px; }
.ml-formats { margin: 10px 0; padding-left: 20px; font-size: 12.5px; line-height: 1.9; }
.ml-import-summary { display: flex; gap: 7px; margin: 4px 0 12px; }
.ml-import-summary .chip { background: var(--sunken); padding: 2px 9px; border-radius: 20px; font-size: 11.5px; }
.ml-import-folders {
  max-height: 300px; overflow: auto; border: 1px solid var(--line); border-radius: var(--r-2);
  background: var(--chrome);
}
.ml-import-folder {
  display: flex; align-items: center; gap: 8px; padding: 5px 10px; font-size: 12.5px;
  border-bottom: 1px solid var(--line-soft);
}
.ml-import-folder:last-child { border-bottom: 0; }
.ml-import-folder .name { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ml-import-folder .count { color: var(--ink-3); font-size: 11.5px; font-variant-numeric: tabular-nums; }
`;

// The stylesheet is injected once, rather than per render of a large tree.
if (typeof document !== 'undefined' && !document.getElementById('rutba-mail-css')) {
  const style = document.createElement('style');
  style.id = 'rutba-mail-css';
  style.textContent = CSS;
  document.head.appendChild(style);
}
