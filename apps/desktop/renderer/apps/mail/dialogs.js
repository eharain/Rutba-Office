// The dialogs and secondary views.
//
// Account setup, importing, and the two views that treat the mailbox as
// something other than a list of messages: everything you have been sent as
// files, and everyone who has ever written to you. Both are derived from the
// index that is already there, which is why a client that owns its own store
// can offer them and a thin IMAP front end cannot.

import React, { useCallback, useEffect, useState } from 'react';
import {
  Button, Dialog, Field, Input, Icon, Chip, Empty, Spinner, Search, Progress,
  formatBytes, formatWhen,
} from '@rutba/office-ui';
import { avatarFor, displayName } from './parts.js';

/* ── found on this computer ──────────────────────────────────────────────── */

const SOURCE_ICON = { Outlook: 'mail', Thunderbird: 'globe', 'Apple Mail': 'mail', 'Windows Live Mail': 'mail' };

/**
 * What the other mail clients on this machine have left behind.
 *
 * The offer every competitor makes is "import a file, if you can find it". The
 * offer here is a list of what is actually on the disk, with the account
 * settings alongside the mail, so switching is one press rather than an
 * afternoon of exporting.
 */
export function Discovered({ scan, onImportStore, onUseAccount }) {
  if (!scan) {
    return (
      <div className="ml-found">
        <div className="ml-found-item" style={{ cursor: 'default' }}>
          <Spinner />
          <span className="grow">
            <div className="who">Looking for mail already on this computer…</div>
            <div className="what">Outlook, Thunderbird, Apple Mail and Windows Mail</div>
          </span>
        </div>
      </div>
    );
  }

  if (!scan.accounts?.length) {
    return (
      <p className="rw-hint" style={{ marginTop: 0 }}>
        No other mail client was found on this computer. Choose a file below if you have an archive elsewhere.
      </p>
    );
  }

  return (
    <div className="ml-found">
      {scan.accounts.map((account, i) => (
        <div key={`${account.source}-${account.email}-${i}`} style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
          <button
            type="button"
            className="ml-found-item"
            disabled={!account.stores?.length}
            onClick={() => account.stores?.[0] && onImportStore(account.stores[0].path)}
          >
            <span className="ml-found-logo">
              <Icon name={SOURCE_ICON[account.source] || 'mail'} size={16} />
            </span>
            <span className="grow">
              <div className="who">{account.email}</div>
              <div className="what">
                {account.source}
                {account.storeCount ? ` · ${account.storeCount} data file${account.storeCount === 1 ? '' : 's'}` : ''}
                {account.messagesHint ? ` · ${formatBytes(account.messagesHint)}` : ''}
                {account.incoming?.host ? ` · ${account.incoming.host}` : ''}
              </div>
            </span>
            {account.stores?.length ? <Chip>Import</Chip> : null}
          </button>

          {/* More than one data file: name them, because "4 stores" is not a choice. */}
          {account.stores?.length > 1
            ? account.stores.map((store) => (
                <button
                  key={store.path}
                  type="button"
                  className="ml-found-item"
                  style={{ marginLeft: 24, padding: '6px 11px' }}
                  onClick={() => onImportStore(store.path)}
                >
                  <Icon name="file" size={14} />
                  <span className="grow">
                    <div className="what" style={{ fontFamily: 'var(--mono)', fontSize: 11 }}>{store.path}</div>
                  </span>
                  <span className="what">{formatBytes(store.bytes)}</span>
                </button>
              ))
            : null}

          {account.incoming?.host ? (
            <button type="button" className="ml-found-item" style={{ marginLeft: 24, padding: '6px 11px' }} onClick={() => onUseAccount(account)}>
              <Icon name="settings" size={14} />
              <span className="grow">
                <div className="what">
                  Set up this account here — {account.incoming.protocol.toUpperCase()} {account.incoming.host}
                  {account.outgoing?.host ? `, SMTP ${account.outgoing.host}` : ''}
                </div>
              </span>
              <Chip>Use settings</Chip>
            </button>
          ) : null}
        </div>
      ))}
      <p className="rw-hint" style={{ margin: 0 }}>
        Nothing was read — only the file names and the settings those clients keep in plain text. Your password is
        never among them, so you will be asked for it once.
      </p>
    </div>
  );
}

/* ── import ──────────────────────────────────────────────────────────────── */

export function ImportDialog({ scan, onClose, onChoose, onImportStore, onUseAccount }) {
  return (
    <Dialog
      title="Import mail"
      width={600}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button primary label="Choose a file…" onClick={onChoose} />
        </>
      }
    >
      <Discovered scan={scan} onImportStore={onImportStore} onUseAccount={onUseAccount} />
      <p style={{ marginTop: 4, marginBottom: 6, fontWeight: 600, fontSize: 12.5 }}>Or open an archive directly</p>
      <ul className="ml-formats">
        <li><strong>.pst</strong> and <strong>.ost</strong> — Outlook data files, with folders, attachments and read state</li>
        <li><strong>.olm</strong> — Outlook for Mac archives</li>
        <li><strong>.mbox</strong> — Thunderbird, Apple Mail, Google Takeout</li>
        <li><strong>.eml</strong>, <strong>.msg</strong> — single messages, or a folder full of them</li>
      </ul>
      <p className="rw-hint">Nothing is sent anywhere. The import reads the file and writes to this computer only.</p>
    </Dialog>
  );
}

export function ImportPreview({ found, onCancel, onImport }) {
  const [chosen, setChosen] = useState(() => new Set((found.folders || []).filter((f) => f.messages > 0).map((f) => f.path || f.name)));
  const all = (found.folders || []).filter((f) => f.messages > 0);
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
      <div className="ml-listbar" style={{ borderTop: '1px solid var(--line-soft)' }}>
        <input
          type="checkbox"
          className="ml-check"
          checked={chosen.size === all.length && all.length > 0}
          onChange={(e) => setChosen(e.target.checked ? new Set(all.map((f) => f.path || f.name)) : new Set())}
        />
        <span>{chosen.size} of {all.length} folders</span>
      </div>
      <div className="ml-import-folders">
        {(found.folders || []).map((f) => {
          const key = f.path || f.name;
          const on = chosen.has(key);
          return (
            <label key={key} className="ml-import-folder" style={{ paddingLeft: 8 + Math.min(f.depth || 0, 5) * 13 }}>
              <input
                type="checkbox"
                className="ml-check"
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

export function AccountDialog({ shell, seed, onClose, onSaved, toast }) {
  const [form, setForm] = useState(() => ({
    email: seed?.email && seed.email.includes('@') ? seed.email : '',
    name: seed?.name || '',
    password: '',
    imapHost: seed?.incoming?.host || '',
    imapPort: seed?.incoming?.port || 993,
    imapSecure: seed?.incoming?.secure !== false,
    smtpHost: seed?.outgoing?.host || '',
    smtpPort: seed?.outgoing?.port || 465,
    smtpSecure: seed?.outgoing?.secure !== false,
  }));
  const [note, setNote] = useState(seed ? `These settings came from ${seed.source} on this computer.` : null);
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

/* ── everything you have been sent ───────────────────────────────────────── */

/**
 * The attachment view.
 *
 * People remember the file, not the message it arrived in. Outlook has no such
 * view at all; Gmail makes you write a search query. Here it is a place.
 */
export function FilesView({ shell, accountId, onOpen, onReveal }) {
  const [query, setQuery] = useState('');
  const [files, setFiles] = useState(null);

  useEffect(() => {
    let live = true;
    setFiles(null);
    shell.mail
      .files({ accountId, query })
      .then((r) => live && setFiles(r))
      .catch(() => live && setFiles({ rows: [], total: 0 }));
    return () => {
      live = false;
    };
  }, [shell, accountId, query]);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0 }}>
      <div className="ml-tools">
        <Search value={query} onChange={setQuery} placeholder="Search attachments by name or subject" />
      </div>
      {!files ? (
        <div style={{ padding: 40, display: 'grid', placeItems: 'center' }}><Spinner /></div>
      ) : files.rows.length ? (
        <div className="ml-grid">
          {files.rows.map((f) => (
            <button key={`${f.folder}-${f.id}-${f.index}`} type="button" className="ml-card" onClick={() => onOpen(f)}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 7 }}>
                <Icon name={iconForFile(f.filename)} size={15} />
                <span className="name">{f.filename}</span>
              </div>
              <div className="sub">{formatBytes(f.size)} · {displayName(f.from)}</div>
              <div className="sub">{f.subject || '(no subject)'}</div>
              <div className="sub">{formatWhen(f.date)}</div>
            </button>
          ))}
        </div>
      ) : (
        <Empty icon="attach" title={query ? 'Nothing matches' : 'No attachments yet'}>
          Every file anyone has sent you appears here, whatever folder it landed in.
        </Empty>
      )}
    </div>
  );
}

const FILE_ICONS = {
  pdf: 'pdf', doc: 'word', docx: 'word', rtf: 'word', odt: 'word', txt: 'file',
  xls: 'sheets', xlsx: 'sheets', csv: 'sheets', ods: 'sheets',
  ppt: 'slides', pptx: 'slides', odp: 'slides',
  png: 'image', jpg: 'image', jpeg: 'image', gif: 'image', webp: 'image', bmp: 'image', svg: 'image',
  mp4: 'video', mkv: 'video', mov: 'video', webm: 'video', avi: 'video',
  zip: 'archive', '7z': 'archive', rar: 'archive', tar: 'archive', gz: 'archive',
};
const iconForFile = (name) => FILE_ICONS[String(name || '').split('.').pop().toLowerCase()] || 'file';

/* ── everyone who writes to you ──────────────────────────────────────────── */

export function PeopleView({ shell, accountId, onPerson }) {
  const [query, setQuery] = useState('');
  const [people, setPeople] = useState(null);

  useEffect(() => {
    let live = true;
    setPeople(null);
    shell.mail
      .people({ accountId })
      .then((r) => live && setPeople(r))
      .catch(() => live && setPeople({ rows: [], total: 0 }));
    return () => {
      live = false;
    };
  }, [shell, accountId]);

  const rows = (people?.rows || []).filter(
    (p) => !query || p.address.includes(query.toLowerCase()) || String(p.name || '').toLowerCase().includes(query.toLowerCase())
  );

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0 }}>
      <div className="ml-tools">
        <Search value={query} onChange={setQuery} placeholder="Search people" />
      </div>
      {!people ? (
        <div style={{ padding: 40, display: 'grid', placeItems: 'center' }}><Spinner /></div>
      ) : rows.length ? (
        <div className="ml-people">
          {rows.map((p) => {
            const avatar = avatarFor(p);
            return (
              <button key={p.address} type="button" className="ml-person" onClick={() => onPerson(p)}>
                <span className="ml-avatar-sm" style={{ background: avatar.colour }}>{avatar.initial}</span>
                <span className="grow">
                  <div className="nm">{p.name || p.address}</div>
                  <div className="ad">{p.address}</div>
                </span>
                <span className="n">
                  {p.received ? `${p.received.toLocaleString()} received` : ''}
                  {p.received && p.sent ? ' · ' : ''}
                  {p.sent ? `${p.sent.toLocaleString()} sent` : ''}
                </span>
              </button>
            );
          })}
        </div>
      ) : (
        <Empty icon="reply" title="Nobody yet">This fills in as mail arrives, and after an import.</Empty>
      )}
    </div>
  );
}

/* ── importing progress ──────────────────────────────────────────────────── */

export function ImportingDialog({ path, progress }) {
  return (
    <Dialog title="Importing" onClose={undefined}>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 10, padding: '6px 0 14px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <Spinner />
          <span>{progress ? `${progress.done.toLocaleString()} messages — ${progress.folder}` : 'Reading the archive…'}</span>
        </div>
        <Progress value={progress?.done ? (progress.done % 500) / 500 : 0.15} />
        <div className="rw-hint">{path}</div>
      </div>
    </Dialog>
  );
}
