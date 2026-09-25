// The dialogs and secondary views.
//
// Account setup, importing, and the two views that treat the mailbox as
// something other than a list of messages: everything you have been sent as
// files, and everyone who has ever written to you. Both are derived from the
// index that is already there, which is why a client that owns its own store
// can offer them and a thin IMAP front end cannot.

import React, { useCallback, useEffect, useRef, useState } from 'react';
import {
  Button, Dialog, Field, Input, Icon, Chip, Empty, Spinner, Search, Progress,
  formatBytes, formatWhen,
} from '@rutba/office-ui';
import { avatarFor, displayName } from './parts.js';

/* ── signature ────────────────────────────────────────────────────────────── */

/**
 * A signature lives on the account, not on the person, because the one thing
 * every mail client gets wrong at least once is sending a work reply carrying
 * a personal sign-off. `accountId` picks which one opens; a mailbox with more
 * than one account gets a picker so a wrong account is a correction, not a
 * surprise the next time that account is used to write something.
 */
export function SignatureDialog({ shell, accounts, accountId, onClose, onSaved, toast }) {
  const [id, setId] = useState(accountId || accounts[0]?.id || '');
  const account = accounts.find((a) => a.id === id) || null;
  const [text, setText] = useState(account?.signature || '');
  const [saving, setSaving] = useState(false);

  // Switching the picker loads that account's own text — it does not carry
  // over unsaved edits from the one before it.
  useEffect(() => {
    setText((accounts.find((a) => a.id === id) || {}).signature || '');
  }, [id, accounts]);

  const save = useCallback(async () => {
    if (!id) return;
    setSaving(true);
    try {
      await shell.mail.updateAccount({ id, patch: { signature: text } });
      toast('Signature saved', { tone: 'good' });
      onSaved();
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    } finally {
      setSaving(false);
    }
  }, [shell, id, text, onSaved, toast]);

  return (
    <Dialog
      title="Signature"
      width={520}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button primary label={saving ? 'Saving…' : 'Save'} disabled={saving || !id} onClick={save} />
        </>
      }
    >
      {accounts.length > 1 ? (
        <Field label="Account">
          <select className="rw-input ml-signature-account" value={id} onChange={(e) => setId(e.target.value)}>
            {accounts.map((a) => (
              <option key={a.id} value={a.id}>{a.name ? `${a.name} <${a.email}>` : a.email}</option>
            ))}
          </select>
        </Field>
      ) : null}
      <Field label="Added to the end of new messages, and above the quote when you reply or forward">
        <textarea
          className="rw-input ml-compose-body ml-signature-text"
          rows={7}
          value={text}
          onChange={(e) => setText(e.target.value)}
          placeholder="Left blank, nothing is added."
        />
      </Field>
    </Dialog>
  );
}

/* ── out of office ────────────────────────────────────────────────────────── */

/**
 * An ISO instant as a `<input type="datetime-local">` wants it — local time,
 * minutes only. Built from the Date object's own local getters rather than
 * `toISOString()`, which is UTC: on a machine whose zone is not UTC, showing
 * those digits in a field the browser reads as local time would tell the
 * account's own start or end time back wrong by the zone's offset.
 */
const forDateInput = (iso) => {
  if (!iso) return '';
  const d = new Date(iso);
  const pad = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
};
/** The other way: the input's local-time string back to an ISO instant, or null when empty. */
const fromDateInput = (value) => {
  if (!value) return null;
  const at = new Date(value);
  return Number.isNaN(at.getTime()) ? null : at.toISOString();
};

/**
 * Automatic replies, per account, the same shape as the Signature dialog:
 * opened from that account's own settings, not a preference every account
 * in the mailbox shares. On sends one reply per sender for as long as it
 * stays on — turning it off and on again is what starts that over, the same
 * as Outlook's own "once per sender while away".
 */
export function OutOfOfficeDialog({ shell, accounts, accountId, onClose, onSaved, toast }) {
  const [id, setId] = useState(accountId || accounts[0]?.id || '');
  const account = accounts.find((a) => a.id === id) || null;
  const blank = { enabled: false, start: '', end: '', subject: '', message: '', contactsOnly: false, sentTo: [] };
  const [form, setForm] = useState({ ...blank, ...(account?.autoReply || {}) });
  const [saving, setSaving] = useState(false);
  const set = (patch) => setForm((f) => ({ ...f, ...patch }));

  useEffect(() => {
    setForm({ ...blank, ...((accounts.find((a) => a.id === id) || {}).autoReply || {}) });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [id, accounts]);

  const save = useCallback(async () => {
    if (!id) return;
    setSaving(true);
    try {
      // Turning it on from off starts a fresh "answered" set — the same
      // sender who already heard from it before is owed a fresh reply now
      // that it is on again, not silence because a set from last time is
      // still sitting there.
      const wasOn = Boolean(account?.autoReply?.enabled);
      const sentTo = form.enabled && !wasOn ? [] : form.sentTo || [];
      const patch = { enabled: form.enabled, start: form.start || null, end: form.end || null, subject: form.subject, message: form.message, contactsOnly: form.contactsOnly, sentTo };
      await shell.mail.updateAccount({ id, patch: { autoReply: patch } });
      toast(patch.enabled ? 'Automatic replies are on' : 'Automatic replies are off', { tone: 'good' });
      onSaved();
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    } finally {
      setSaving(false);
    }
  }, [shell, id, form, account, onSaved, toast]);

  return (
    <Dialog
      title="Automatic replies"
      width={520}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button primary label={saving ? 'Saving…' : 'Save'} disabled={saving || !id} onClick={save} />
        </>
      }
    >
      {accounts.length > 1 ? (
        <Field label="Account">
          <select className="rw-input ml-signature-account" value={id} onChange={(e) => setId(e.target.value)}>
            {accounts.map((a) => (
              <option key={a.id} value={a.id}>{a.name ? `${a.name} <${a.email}>` : a.email}</option>
            ))}
          </select>
        </Field>
      ) : null}

      <label className="ml-ooo-toggle ml-ooo-enabled">
        <input type="checkbox" checked={form.enabled} onChange={(e) => set({ enabled: e.target.checked })} />
        <span>Send automatic replies</span>
      </label>

      <div className="ml-servers3" style={{ gridTemplateColumns: '1fr 1fr' }}>
        <Field label="Start (optional)">
          <input type="datetime-local" className="rw-input" value={forDateInput(form.start)} onChange={(e) => set({ start: fromDateInput(e.target.value) })} />
        </Field>
        <Field label="End (optional)">
          <input type="datetime-local" className="rw-input" value={forDateInput(form.end)} onChange={(e) => set({ end: fromDateInput(e.target.value) })} />
        </Field>
      </div>
      <p className="rw-hint" style={{ marginTop: -4 }}>
        Left blank, it starts the moment you save it and runs until you turn it off. Replies go out while Rutba
        Office is running — one due while it is closed goes out the next time mail is fetched.
      </p>

      <Field label="Subject" hint="Leave blank for “Automatic reply: ” and the original subject.">
        <Input value={form.subject} onChange={(e) => set({ subject: e.target.value })} placeholder="Automatic reply: <original subject>" />
      </Field>
      <Field label="Message">
        <textarea
          className="rw-input ml-compose-body ml-ooo-message"
          rows={6}
          value={form.message}
          onChange={(e) => set({ message: e.target.value })}
          placeholder="I'm away until … and will answer when I'm back."
        />
      </Field>

      <label className="ml-ooo-toggle ml-ooo-contacts">
        <input type="checkbox" checked={form.contactsOnly} onChange={(e) => set({ contactsOnly: e.target.checked })} />
        <span>Only reply to people in my contacts</span>
      </label>

      <p className="rw-hint" style={{ marginBottom: 0 }}>
        Each sender hears from this once while it stays on — a second message from the same person is not answered
        again until it is turned off and on. A mailing list, a no-reply address and a message that already says it
        is automatic are never answered.
      </p>
    </Dialog>
  );
}

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

export function ImportDialog({ scan, onClose, onChoose, onImportStore, onUseAccount, onImportAccounts }) {
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
      <p style={{ marginTop: 10, marginBottom: 6, fontWeight: 600, fontSize: 12.5 }}>Or set up accounts from a file</p>
      <p className="rw-hint" style={{ marginBottom: 6 }}>
        A JSON file of accounts — address, password, servers — from another client or from whoever runs your mail. Each is tried before it is kept; one already set up is left alone.
      </p>
      <Button label="Choose an accounts file…" onClick={onImportAccounts} className="ml-import-accounts" />
      <p className="rw-hint" style={{ marginTop: 10 }}>Nothing is sent anywhere. The import reads the file and writes to this computer only.</p>
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

/* ── choosing a provider ─────────────────────────────────────────────────── */

/**
 * The tile row above the address field: Gmail, Outlook, Yahoo and the rest,
 * each filling its own servers with no network and saying exactly what it
 * wants — a browser sign-in, or an app password with a link to make one.
 * "Other" asks for nothing; it just gets out of the way of the address field.
 */
function ProviderTiles({ providers, onChoose, onOther }) {
  return (
    <div className="ml-providers">
      <div className="ml-providers-head">Choose your provider</div>
      <div className="ml-providers-row">
        {providers.map((p) => (
          <button key={p.id} type="button" className="ml-provider" data-provider={p.id} onClick={() => onChoose(p)}>
            <span className="label">{p.tileLabel}</span>
            <span className="domain">@{p.domain}</span>
          </button>
        ))}
        <button type="button" className="ml-provider" data-provider="other" onClick={onOther}>
          <span className="label">Other</span>
          <span className="domain">Enter your address and the servers are found for you</span>
        </button>
      </div>
    </div>
  );
}

/* ── account setup ───────────────────────────────────────────────────────── */

export function AccountDialog({ shell, seed, onClose, onSaved, onImportAccounts, toast }) {
  // An address and a password. Everything else is found: the provider table,
  // the domain's MX and SRV records, autoconfig, Microsoft autodiscover, and a
  // knock on the conventional names (main/mail-discover.js). What was found is
  // shown, what was checked can be opened, and the advanced fields — always one
  // click away, never the only way in — are filled with the answer so that a
  // wrong guess is a correction, not a form from scratch.
  const security = (server) => (!server ? 'tls' : server.secure ? 'tls' : server.starttls === false ? 'none' : 'starttls');
  const [form, setForm] = useState(() => ({
    email: seed?.email && seed.email.includes('@') ? seed.email : '',
    name: seed?.name || '',
    password: '',
    user: seed?.incoming?.user || '',
    imapHost: seed?.incoming?.host || '',
    imapPort: seed?.incoming?.port || 993,
    imapSecurity: security(seed?.incoming ? { secure: seed.incoming.secure !== false, starttls: seed.incoming.secure === false } : null),
    smtpHost: seed?.outgoing?.host || '',
    smtpPort: seed?.outgoing?.port || 465,
    smtpSecurity: security(seed?.outgoing ? { secure: seed.outgoing.secure !== false, starttls: seed.outgoing.secure === false } : null),
  }));
  const [advanced, setAdvanced] = useState(false);
  const [looking, setLooking] = useState(false);
  const [found, setFound] = useState(null);
  const [showSteps, setShowSteps] = useState(false);
  const [provider, setProvider] = useState(null);
  const [testing, setTesting] = useState(false);
  const [adding, setAdding] = useState(false);
  const [result, setResult] = useState(null);
  const [signingIn, setSigningIn] = useState(false);
  const [providers, setProviders] = useState([]);
  // null lets the tiles show or hide by themselves — address empty, or its
  // provider not known yet; 'open'/'closed' is the person overriding that,
  // with Change provider or by picking a tile (or Other).
  const [tilesOverride, setTilesOverride] = useState(null);
  const lookedUp = useRef('');
  const set = (patch) => setForm((f) => ({ ...f, ...patch }));

  useEffect(() => {
    let live = true;
    shell.mail
      .providers()
      .then((list) => live && setProviders(list || []))
      .catch(() => {});
    return () => {
      live = false;
    };
  }, [shell]);

  const discover = useCallback(
    async (email) => {
      const domain = (email.split('@')[1] || '').toLowerCase();
      if (!domain.includes('.')) return;
      lookedUp.current = domain;
      setLooking(true);
      setResult(null);
      setTilesOverride(null); // a fresh search decides for itself whether the tiles are still wanted
      try {
        const r = await shell.mail.autodiscover({
          email,
          seed: seed?.incoming?.host ? { imap: seed.incoming, smtp: seed.outgoing } : null,
        });
        if (lookedUp.current !== domain) return; // the address changed while this ran
        setFound(r);
        const patch = {};
        if (r.imap) Object.assign(patch, { imapHost: r.imap.host, imapPort: r.imap.port, imapSecurity: security(r.imap), user: r.imap.user && r.imap.user !== email ? r.imap.user : '' });
        if (r.smtp) Object.assign(patch, { smtpHost: r.smtp.host, smtpPort: r.smtp.port, smtpSecurity: security(r.smtp) });
        set(patch);
        const p = r.oauth ? await shell.oauth.provider({ email, id: r.oauth }).catch(() => null) : null;
        if (p) setProvider(p);
      } catch (err) {
        setFound({ imap: null, smtp: null, note: err.message, steps: [{ name: 'the search', status: 'failed', detail: err.message }] });
      } finally {
        if (lookedUp.current === domain) setLooking(false);
      }
    },
    [shell, seed]
  );

  // A tile fills exactly what a matching typed address would have found — the
  // same servers, the same note, the same sign-in card — but with no network,
  // because the provider is already known.
  const chooseTile = useCallback(
    async (entry) => {
      lookedUp.current = entry.domain || '';
      setFound({
        imap: { ...entry.imap, verified: false, source: 'provider' },
        smtp: { ...entry.smtp, verified: false, source: 'provider' },
        provider: { id: entry.id, label: entry.tileLabel || entry.label, domain: entry.domain },
        oauth: entry.oauth || null,
        appPassword: entry.appPassword || null,
        note: entry.note || `${entry.tileLabel || entry.label}'s servers, filled in — nothing was searched.`,
        steps: [],
      });
      set({
        imapHost: entry.imap.host, imapPort: entry.imap.port, imapSecurity: security(entry.imap),
        smtpHost: entry.smtp.host, smtpPort: entry.smtp.port, smtpSecurity: security(entry.smtp),
      });
      setTilesOverride('closed');
      if (entry.oauth) {
        const p = await shell.oauth.provider({ email: form.email.trim(), id: entry.oauth }).catch(() => null);
        setProvider(p || null);
      } else {
        setProvider(null);
      }
    },
    [shell, form.email]
  );

  const otherTile = useCallback(() => {
    setTilesOverride('closed');
    document.getElementById('ml-address')?.focus();
  }, []);

  // A provider the address alone names — Gmail, Outlook.com — is offered its
  // sign-in the moment the address is typed, before the search confirms it;
  // the search can still add one the address cannot name, such as a company
  // domain hosted at Microsoft 365.
  useEffect(() => {
    const email = form.email.trim();
    if (!email.includes('@')) {
      setProvider(null);
      return undefined;
    }
    let live = true;
    shell.oauth
      .provider({ email })
      .then((p) => {
        if (live) setProvider(p);
      })
      .catch(() => {});
    return () => {
      live = false;
    };
  }, [form.email, shell]);

  // The search runs once the domain looks complete and the typing has paused,
  // and again only when the domain changes.
  useEffect(() => {
    const email = form.email.trim();
    const domain = (email.split('@')[1] || '').toLowerCase();
    if (!domain.includes('.') || domain === lookedUp.current) return undefined;
    const timer = setTimeout(() => discover(email), 700);
    return () => clearTimeout(timer);
  }, [form.email, discover]);

  const account = () => {
    const sec = (s) => ({ secure: s === 'tls', starttls: s === 'starttls' });
    return {
      email: form.email.trim(),
      name: form.name || form.email.trim(),
      imap: { host: form.imapHost.trim(), port: Number(form.imapPort), ...sec(form.imapSecurity), user: form.user.trim() || undefined },
      smtp: { host: form.smtpHost.trim(), port: Number(form.smtpPort), ...sec(form.smtpSecurity), user: form.user.trim() || undefined },
    };
  };

  const signIn = useCallback(async () => {
    setSigningIn(true);
    try {
      const signed = await shell.oauth.signIn({ email: form.email.trim(), provider: provider.id });
      await shell.mail.addAccount({
        account: { email: signed.email, name: signed.name || form.name || signed.email, imap: signed.imap, smtp: signed.smtp, auth: 'oauth', provider: signed.provider },
      });
      onSaved();
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 12000 });
    } finally {
      setSigningIn(false);
    }
  }, [shell, form, provider, onSaved, toast]);

  const test = useCallback(async () => {
    setTesting(true);
    try {
      const r = await shell.mail.testAccount({ account: account(), password: form.password });
      setResult(r);
      return r;
    } finally {
      setTesting(false);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shell, form]);

  const add = useCallback(async () => {
    setAdding(true);
    try {
      // Tried before it is kept: an account that cannot sign in is not added,
      // and the failure opens the fields it needs.
      const r = await test();
      if (r.error) {
        setAdvanced(true);
        return;
      }
      await shell.mail.addAccount({ account: account(), password: form.password });
      onSaved();
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 7000 });
    } finally {
      setAdding(false);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shell, form, test, onSaved, toast]);

  const ready = form.email.includes('@') && form.imapHost && form.smtpHost;
  const showTiles = tilesOverride ? tilesOverride === 'open' : !form.email.trim() || !found?.provider;
  const addressPlaceholder = found?.provider?.domain ? `you@${found.provider.domain}` : 'you@example.com';
  const label = (s) => ({ tls: 'TLS', starttls: 'STARTTLS', none: 'None' }[s] || s);
  const server = (host, port, sec, verified) => (
    <>
      <b>{host}</b> · {port} · {label(sec)}
      {verified ? <Icon name="check" size={13} /> : null}
    </>
  );

  return (
    <Dialog
      title="Add a mail account"
      width={560}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          {onImportAccounts ? <Button label="From a file…" title="Set up the accounts in a file, all at once" className="ml-import-accounts" onClick={onImportAccounts} /> : null}
          <Button label={advanced ? 'Simple' : 'Advanced…'} onClick={() => setAdvanced((a) => !a)} />
          <Button label={testing ? 'Testing…' : 'Test'} disabled={testing || adding || !ready || !form.password} onClick={test} />
          <Button primary label={adding ? 'Adding…' : 'Add account'} disabled={adding || testing || !ready || !form.password} onClick={add} />
        </>
      }
    >
      <div className="ml-form">
        {showTiles ? (
          <ProviderTiles providers={providers} onChoose={chooseTile} onOther={otherTile} />
        ) : (
          <button type="button" className="ml-provider-change" onClick={() => setTilesOverride('open')}>
            Change provider
          </button>
        )}

        <Field label="Email address">
          <Input
            id="ml-address"
            value={form.email}
            autoFocus
            onChange={(e) => set({ email: e.target.value })}
            onBlur={(e) => {
              const domain = (e.target.value.split('@')[1] || '').toLowerCase();
              if (domain.includes('.') && domain !== lookedUp.current) discover(e.target.value.trim());
            }}
            placeholder={addressPlaceholder}
          />
        </Field>

        {provider ? (
          <div className="ml-found">
            <button type="button" className="ml-found-item" disabled={signingIn || !provider.configured} onClick={signIn}>
              <span className="ml-found-logo">{signingIn ? <Spinner /> : <Icon name="lock" size={15} />}</span>
              <span className="grow">
                <div className="who">{signingIn ? `Waiting for ${provider.label}…` : `Sign in with ${provider.label}`}</div>
                <div className="what">
                  {provider.configured
                    ? 'Opens your own browser. Your password is typed into their page and never reaches this application.'
                    : `This build has no ${provider.label} client id, so signing in is not set up. See docs/OAUTH.md.`}
                </div>
              </span>
              {provider.configured ? <Chip>Recommended</Chip> : null}
            </button>
            <p className="rw-hint" style={{ margin: 0 }}>
              {provider.label} no longer accepts an ordinary password for mail programs. An app password below still works if you have one.
            </p>
          </div>
        ) : null}

        <Field label="Password" hint="Kept in this computer's keystore, never in a file you can read.">
          <Input type="password" value={form.password} onChange={(e) => set({ password: e.target.value })} onKeyDown={(e) => { if (e.key === 'Enter' && ready && form.password && !adding) add(); }} />
        </Field>

        <div className="ml-search" data-state={looking ? 'looking' : found ? (found.imap ? 'found' : 'nothing') : 'idle'}>
          {looking ? (
            <div className="ml-search-line"><Spinner /> Looking up where {form.email.split('@')[1]} keeps its mail…</div>
          ) : found ? (
            <>
              {found.imap ? (
                <div className="ml-search-line">
                  <Icon name={found.imap.verified ? 'check' : 'info'} size={14} />
                  <span>Incoming {server(found.imap.host, found.imap.port, security(found.imap), found.imap.verified)}</span>
                </div>
              ) : null}
              {found.smtp ? (
                <div className="ml-search-line">
                  <Icon name={found.smtp.verified ? 'check' : 'info'} size={14} />
                  <span>Outgoing {server(found.smtp.host, found.smtp.port, security(found.smtp), found.smtp.verified)}</span>
                </div>
              ) : null}
              {found.note ? (
                <div className="ml-note">
                  <Icon name="info" size={14} />
                  <span className="grow">{found.note}</span>
                  {found.appPassword ? (
                    <Button
                      className="ml-app-password"
                      label="Make an app password"
                      title={`Make an app password for ${found.provider?.label || 'this account'} — ${found.appPassword.url}`}
                      onClick={() => shell.shell.openExternal({ url: found.appPassword.url })}
                    />
                  ) : null}
                </div>
              ) : null}
              {found.steps?.length ? (
                <button type="button" className="ml-steps-toggle" onClick={() => setShowSteps((s) => !s)}>
                  {showSteps ? 'Hide what was checked' : 'What was checked'}
                </button>
              ) : null}
              {showSteps ? (
                <ul className="ml-steps">
                  {found.steps.map((s) => (
                    <li key={s.name} className={`ml-step ${s.status}`}>
                      <Icon name={s.status === 'ok' ? 'check' : s.status === 'failed' ? 'close' : 'minus'} size={12} />
                      <span className="name">{s.name}</span>
                      <span className="detail">{s.detail}</span>
                    </li>
                  ))}
                </ul>
              ) : null}
            </>
          ) : seed ? (
            <div className="ml-note"><Icon name="info" size={14} />These settings came from {seed.source} on this computer; they are checked when the address is complete.</div>
          ) : (
            <div className="ml-search-line quiet">The mail server is found from the address. If it is not, Advanced takes the names from your provider.</div>
          )}
        </div>

        {advanced ? (
          <div className="ml-advanced">
            <Field label="Your name">
              <Input value={form.name} onChange={(e) => set({ name: e.target.value })} placeholder="How your name appears on messages you send" />
            </Field>
            <Field label="Username" hint="Only when it is not the address itself.">
              <Input value={form.user} onChange={(e) => set({ user: e.target.value })} placeholder={form.email || 'you@example.com'} />
            </Field>
            <div className="ml-servers3">
              <Field label="Incoming server (IMAP)">
                <Input value={form.imapHost} onChange={(e) => set({ imapHost: e.target.value })} placeholder="imap.example.com" />
              </Field>
              <Field label="Port">
                <Input type="number" value={form.imapPort} onChange={(e) => set({ imapPort: e.target.value })} />
              </Field>
              <Field label="Security">
                <select className="rw-input" value={form.imapSecurity} onChange={(e) => set({ imapSecurity: e.target.value, imapPort: e.target.value === 'tls' ? 993 : 143 })}>
                  <option value="tls">TLS</option>
                  <option value="starttls">STARTTLS</option>
                  <option value="none">None</option>
                </select>
              </Field>
              <Field label="Outgoing server (SMTP)">
                <Input value={form.smtpHost} onChange={(e) => set({ smtpHost: e.target.value })} placeholder="smtp.example.com" />
              </Field>
              <Field label="Port">
                <Input type="number" value={form.smtpPort} onChange={(e) => set({ smtpPort: e.target.value })} />
              </Field>
              <Field label="Security">
                <select className="rw-input" value={form.smtpSecurity} onChange={(e) => set({ smtpSecurity: e.target.value, smtpPort: e.target.value === 'tls' ? 465 : 587 })}>
                  <option value="tls">TLS</option>
                  <option value="starttls">STARTTLS</option>
                  <option value="none">None</option>
                </select>
              </Field>
            </div>
          </div>
        ) : null}

        {result ? (
          <div className={`ml-result${result.error ? ' bad' : ' good'}`}>
            <Icon name={result.error ? 'info' : 'check'} size={14} />
            <span>
              Incoming {result.imap?.ok ? `signed in — ${result.imap.folders} folders` : `failed: ${result.imap?.message}`}
              {' · '}
              Outgoing {result.smtp?.ok ? 'signed in' : `failed: ${result.smtp?.message}`}
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
