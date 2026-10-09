// The dialogs and secondary views.
//
// Account setup, importing, and the two views that treat the mailbox as
// something other than a list of messages: everything you have been sent as
// files, and everyone who has ever written to you. Both are derived from the
// index that is already there, which is why a client that owns its own store
// can offer them and a thin IMAP front end cannot.

import React, { useCallback, useEffect, useRef, useState } from 'react';
import {
  Button, Dialog, Field, Input, Icon, Chip, Empty, Spinner, Search, Progress, Separator,
  formatBytes, formatWhen, t, tn,
} from '@rutba/office-ui';
import { avatarFor, displayName } from './parts.js';
import { cleanHtml, htmlToText, LinkRow, selectionIn } from './richtext.js';
import { signatureTextToHtml } from '@rutba/mailbox/signature';

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
  const [saving, setSaving] = useState(false);
  const [linking, setLinking] = useState(null);
  const editor = useRef(null);

  // Switching the picker loads that account's own signature — it does not
  // carry over unsaved edits from the one before it. One with only a plain
  // signature opens as that, its lines kept. Only the pick does this: a list
  // refreshed by arriving mail must not wipe what is being typed.
  useEffect(() => {
    const a = accounts.find((x) => x.id === id) || {};
    if (editor.current) editor.current.innerHTML = cleanHtml(a.signatureHtml || signatureTextToHtml(a.signature || ''));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [id]);

  const run = (command, value) => {
    editor.current?.focus();
    try {
      document.execCommand(command, false, value);
    } catch {
      /* a command this engine does not know; the words are unharmed */
    }
  };
  const keep = (e) => e.preventDefault();

  // A picture — a logo — kept in the signature itself, so it goes with every
  // message without a fetch from anywhere; small, because it goes with every
  // message.
  const picture = useCallback(async () => {
    const [file] = (await shell.dialog.open({ title: t('Insert a picture'), filters: [{ name: t('Pictures'), extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp'] }] })) || [];
    if (!file) return;
    const { bytes } = await shell.fs.read({ path: file });
    if (bytes.length > 256 * 1024) {
      toast(t('That picture is over 256 KB — a signature goes with every message, so use a smaller one.'), { tone: 'bad' });
      return;
    }
    const ext = String(file).split('.').pop().toLowerCase();
    const type = { jpg: 'jpeg', jpeg: 'jpeg', png: 'png', gif: 'gif', webp: 'webp' }[ext] || 'png';
    let binary = '';
    const view = new Uint8Array(bytes);
    for (let i = 0; i < view.length; i += 0x8000) binary += String.fromCharCode(...view.subarray(i, i + 0x8000));
    run('insertImage', `data:image/${type};base64,${btoa(binary)}`);
  }, [shell, toast]);

  const save = useCallback(async () => {
    if (!id) return;
    setSaving(true);
    try {
      // What is kept is what a signature can carry, and its plain words go
      // with it for a message written as plain text.
      const html = cleanHtml(editor.current?.innerHTML || '');
      const text = htmlToText(html).replace(/\s+$/, '');
      const plainOnly = !text || html === signatureTextToHtml(text);
      await shell.mail.updateAccount({ id, patch: { signature: text, signatureHtml: plainOnly ? null : html } });
      toast(t('Signature saved'), { tone: 'good' });
      onSaved();
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    } finally {
      setSaving(false);
    }
  }, [shell, id, onSaved, toast]);

  return (
    <Dialog
      title={t('Signature')}
      width={560}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={saving ? t('Saving…') : t('Save')} disabled={saving || !id} onClick={save} />
        </>
      }
    >
      {accounts.length > 1 ? (
        <Field label={t('Account')}>
          <select className="rw-input ml-signature-account" value={id} onChange={(e) => setId(e.target.value)}>
            {accounts.map((a) => (
              <option key={a.id} value={a.id}>{a.name ? `${a.name} <${a.email}>` : a.email}</option>
            ))}
          </select>
        </Field>
      ) : null}
      <Field label={t('Added to the end of new messages, and above the quote when you reply or forward')}>
        <div className="ml-toolbar ml-signature-tools">
          <Button icon="bold" title={t('Bold')} onMouseDown={keep} onClick={() => run('bold')} />
          <Button icon="italic" title={t('Italic')} onMouseDown={keep} onClick={() => run('italic')} />
          <Button icon="underline" title={t('Underline')} onMouseDown={keep} onClick={() => run('underline')} />
          <Separator />
          {SIGNATURE_COLOURS.map(([colour, name]) => (
            <button key={colour} type="button" className="ml-signature-colour" title={t('Colour: {name}', { name })} aria-label={t('Colour: {name}', { name })} style={{ background: colour }} onMouseDown={keep} onClick={() => run('foreColor', colour)} />
          ))}
          <Separator />
          <Button label="A−" title={t('Smaller')} onMouseDown={keep} onClick={() => run('fontSize', '2')} />
          <Button label="A+" title={t('Larger')} onMouseDown={keep} onClick={() => run('fontSize', '4')} />
          <Separator />
          <Button icon="link" title={t('Link')} onMouseDown={keep} onClick={() => setLinking({ range: selectionIn(editor.current) })} />
          <Button icon="picture" title={t('Picture')} onMouseDown={keep} onClick={picture} />
          <Button icon="undo" title={t('Clear formatting')} onMouseDown={keep} onClick={() => run('removeFormat')} />
        </div>
        {linking ? <LinkRow editor={editor.current} range={linking.range} onDone={() => setLinking(null)} /> : null}
        <div
          ref={editor}
          className="rw-input ml-signature-editor"
          contentEditable
          suppressContentEditableWarning
          spellCheck
          role="textbox"
          aria-multiline="true"
          aria-label={t('Signature')}
          data-placeholder={t('Left blank, nothing is added.')}
        />
      </Field>
    </Dialog>
  );
}

/** The colours the signature editor offers: the house blue, and the ones signatures use. */
const SIGNATURE_COLOURS = [['#1a1c20', t('Black')], ['#5f6368', t('Grey')], ['#2b5fd9', t('Blue')], ['#0f9d58', t('Green')], ['#c00000', t('Red')], ['#7b5cd6', t('Purple')]];

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
      toast(patch.enabled ? t('Automatic replies are on') : t('Automatic replies are off'), { tone: 'good' });
      onSaved();
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    } finally {
      setSaving(false);
    }
  }, [shell, id, form, account, onSaved, toast]);

  return (
    <Dialog
      title={t('Automatic replies')}
      width={520}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={saving ? t('Saving…') : t('Save')} disabled={saving || !id} onClick={save} />
        </>
      }
    >
      {accounts.length > 1 ? (
        <Field label={t('Account')}>
          <select className="rw-input ml-signature-account" value={id} onChange={(e) => setId(e.target.value)}>
            {accounts.map((a) => (
              <option key={a.id} value={a.id}>{a.name ? `${a.name} <${a.email}>` : a.email}</option>
            ))}
          </select>
        </Field>
      ) : null}

      <label className="ml-ooo-toggle ml-ooo-enabled">
        <input type="checkbox" checked={form.enabled} onChange={(e) => set({ enabled: e.target.checked })} />
        <span>{t('Send automatic replies')}</span>
      </label>

      <div className="ml-servers3" style={{ gridTemplateColumns: '1fr 1fr' }}>
        <Field label={t('Start (optional)')}>
          <input type="datetime-local" className="rw-input" value={forDateInput(form.start)} onChange={(e) => set({ start: fromDateInput(e.target.value) })} />
        </Field>
        <Field label={t('End (optional)')}>
          <input type="datetime-local" className="rw-input" value={forDateInput(form.end)} onChange={(e) => set({ end: fromDateInput(e.target.value) })} />
        </Field>
      </div>
      <p className="rw-hint" style={{ marginTop: -4 }}>
        {t('Left blank, it starts the moment you save it and runs until you turn it off. Replies go out while Rutba Office is running — one due while it is closed goes out the next time mail is fetched.')}
      </p>

      <Field label={t('Subject')} hint={t('Leave blank for “Automatic reply: ” and the original subject.')}>
        <Input value={form.subject} onChange={(e) => set({ subject: e.target.value })} placeholder={t('Automatic reply: <original subject>')} />
      </Field>
      <Field label={t('Message')}>
        <textarea
          className="rw-input ml-compose-body ml-ooo-message"
          rows={6}
          value={form.message}
          onChange={(e) => set({ message: e.target.value })}
          placeholder={t("I'm away until … and will answer when I'm back.")}
        />
      </Field>

      <label className="ml-ooo-toggle ml-ooo-contacts">
        <input type="checkbox" checked={form.contactsOnly} onChange={(e) => set({ contactsOnly: e.target.checked })} />
        <span>{t('Only reply to people in my contacts')}</span>
      </label>

      <p className="rw-hint" style={{ marginBottom: 0 }}>
        {t('Each sender hears from this once while it stays on — a second message from the same person is not answered again until it is turned off and on. A mailing list, a no-reply address and a message that already says it is automatic are never answered.')}
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
            <div className="who">{t('Looking for mail already on this computer…')}</div>
            <div className="what">{t('Outlook, Thunderbird, Apple Mail and Windows Mail')}</div>
          </span>
        </div>
      </div>
    );
  }

  if (!scan.accounts?.length) {
    return (
      <p className="rw-hint" style={{ marginTop: 0 }}>
        {t('No other mail client was found on this computer. Choose a file below if you have an archive elsewhere.')}
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
                {account.storeCount ? ` · ${tn(account.storeCount, '{count} data file', '{count} data files')}` : ''}
                {account.messagesHint ? ` · ${formatBytes(account.messagesHint)}` : ''}
                {account.incoming?.host ? ` · ${account.incoming.host}` : ''}
              </div>
            </span>
            {account.stores?.length ? <Chip>{t('Import')}</Chip> : null}
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
                  {t('Set up this account here — {protocol} {host}', { protocol: account.incoming.protocol.toUpperCase(), host: account.incoming.host })}
                  {account.outgoing?.host ? `, SMTP ${account.outgoing.host}` : ''}
                </div>
              </span>
              <Chip>{t('Use settings')}</Chip>
            </button>
          ) : null}
        </div>
      ))}
      <p className="rw-hint" style={{ margin: 0 }}>
        {t('Nothing was read — only the file names and the settings those clients keep in plain text. Your password is never among them, so you will be asked for it once.')}
      </p>
    </div>
  );
}

/* ── import ──────────────────────────────────────────────────────────────── */

export function ImportDialog({ scan, onClose, onChoose, onImportStore, onUseAccount, onImportAccounts }) {
  return (
    <Dialog
      title={t('Import mail')}
      width={600}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('Choose a file…')} onClick={onChoose} />
        </>
      }
    >
      <Discovered scan={scan} onImportStore={onImportStore} onUseAccount={onUseAccount} />
      <p style={{ marginTop: 4, marginBottom: 6, fontWeight: 600, fontSize: 12.5 }}>{t('Or open an archive directly')}</p>
      <ul className="ml-formats">
        <li><strong>.pst</strong> {t('and')} <strong>.ost</strong> — {t('Outlook data files, with folders, attachments and read state')}</li>{/* words-ok: file extensions */}
        <li><strong>.olm</strong> — {t('Outlook for Mac archives')}</li>
        <li><strong>.mbox</strong> — Thunderbird, Apple Mail, Google Takeout{/* words-ok: product names */}</li>
        <li><strong>.eml</strong>, <strong>.msg</strong> — {t('single messages, or a folder full of them')}</li>
      </ul>
      <p style={{ marginTop: 10, marginBottom: 6, fontWeight: 600, fontSize: 12.5 }}>{t('Or set up accounts from a file')}</p>
      <p className="rw-hint" style={{ marginBottom: 6 }}>
        {t('A JSON file of accounts — address, password, servers — from another client or from whoever runs your mail. Each is tried before it is kept; one already set up is left alone.')}
      </p>
      <Button label={t('Choose an accounts file…')} onClick={onImportAccounts} className="ml-import-accounts" />
      <p className="rw-hint" style={{ marginTop: 10 }}>{t('Nothing is sent anywhere. The import reads the file and writes to this computer only.')}</p>
    </Dialog>
  );
}

export function ImportPreview({ found, onCancel, onImport }) {
  const [chosen, setChosen] = useState(() => new Set((found.folders || []).filter((f) => f.messages > 0).map((f) => f.path || f.name)));
  const all = (found.folders || []).filter((f) => f.messages > 0);
  const total = (found.folders || []).filter((f) => chosen.has(f.path || f.name)).reduce((n, f) => n + f.messages, 0);

  if (!found.readable) {
    return (
      <Dialog title={t('This file cannot be read')} onClose={onCancel} actions={<Button primary label={t('Close')} onClick={onCancel} />}>
        <p style={{ marginTop: 0 }}>{found.label}</p>
        <p className="rw-hint">{found.encoding ? t('It uses the {encoding} encoding, which this build cannot decode.', { encoding: found.encoding }) : t('The format was not recognised.')}</p>
      </Dialog>
    );
  }

  return (
    <Dialog
      title={t('Import from {source}', { source: found.label })}
      width={560}
      onClose={onCancel}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onCancel} />
          <Button primary label={tn(total, 'Import {count} message', 'Import {count} messages')} disabled={!total} onClick={() => onImport([...chosen])} />
        </>
      }
    >
      <div className="ml-import-summary">
        <Chip>{formatBytes(found.bytes || 0)}</Chip>
        <Chip>{tn((found.folders || []).length, '{count} folder', '{count} folders')}</Chip>
        <Chip>{tn(found.messages || 0, '{count} message', '{count} messages')}</Chip>
        {found.encoding ? <Chip>{found.encoding}</Chip> : null}
      </div>
      <div className="ml-listbar" style={{ borderTop: '1px solid var(--line-soft)' }}>
        <input
          type="checkbox"
          className="ml-check"
          checked={chosen.size === all.length && all.length > 0}
          onChange={(e) => setChosen(e.target.checked ? new Set(all.map((f) => f.path || f.name)) : new Set())}
        />
        <span>{tn(all.length, '{chosen} of {count} folder', '{chosen} of {count} folders', { chosen: chosen.size })}</span>
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
      <p className="rw-hint">{t('Importing the same archive twice adds nothing: a message is identified by its own contents.')}</p>
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
      <div className="ml-providers-head">{t('Choose your provider')}</div>
      <div className="ml-providers-row">
        {providers.map((p) => (
          <button key={p.id} type="button" className="ml-provider" data-provider={p.id} onClick={() => onChoose(p)}>
            <span className="label">{p.tileLabel}</span>
            <span className="domain">@{p.domain}</span>
          </button>
        ))}
        <button type="button" className="ml-provider" data-provider="other" onClick={onOther}>
          <span className="label">{t('Other')}</span>
          <span className="domain">{t('Enter your address and the servers are found for you')}</span>
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
        setFound({ imap: null, smtp: null, note: err.message, steps: [{ name: t('the search'), status: 'failed', detail: err.message }] });
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
        note: entry.note || t("{provider}'s servers, filled in — nothing was searched.", { provider: entry.tileLabel || entry.label }),
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
  const label = (s) => ({ tls: 'TLS', starttls: 'STARTTLS', none: t('None') }[s] || s);
  const server = (host, port, sec, verified) => (
    <>
      <b>{host}</b> · {port} · {label(sec)}
      {verified ? <Icon name="check" size={13} /> : null}
    </>
  );

  return (
    <Dialog
      title={t('Add a mail account')}
      width={560}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          {onImportAccounts ? <Button label={t('From a file…')} title={t('Set up the accounts in a file, all at once')} className="ml-import-accounts" onClick={onImportAccounts} /> : null}
          <Button label={advanced ? t('Simple') : t('Advanced…')} onClick={() => setAdvanced((a) => !a)} />
          <Button label={testing ? t('Testing…') : t('Test')} disabled={testing || adding || !ready || !form.password} onClick={test} />
          <Button primary label={adding ? t('Adding…') : t('Add account')} disabled={adding || testing || !ready || !form.password} onClick={add} />
        </>
      }
    >
      <div className="ml-form">
        {showTiles ? (
          <ProviderTiles providers={providers} onChoose={chooseTile} onOther={otherTile} />
        ) : (
          <button type="button" className="ml-provider-change" onClick={() => setTilesOverride('open')}>
            {t('Change provider')}
          </button>
        )}

        <Field label={t('Email address')}>
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
                <div className="who">{signingIn ? t('Waiting for {provider}…', { provider: provider.label }) : t('Sign in with {provider}', { provider: provider.label })}</div>
                <div className="what">
                  {provider.configured
                    ? t('Opens your own browser. Your password is typed into their page and never reaches this application.')
                    : t('This build has no {provider} client id, so signing in is not set up. See docs/OAUTH.md.', { provider: provider.label })}
                </div>
              </span>
              {provider.configured ? <Chip>{t('Recommended')}</Chip> : null}
            </button>
            <p className="rw-hint" style={{ margin: 0 }}>
              {t('{provider} no longer accepts an ordinary password for mail programs. An app password below still works if you have one.', { provider: provider.label })}
            </p>
          </div>
        ) : null}

        <Field label={t('Password')} hint={t("Kept in this computer's keystore, never in a file you can read.")}>
          <Input type="password" value={form.password} onChange={(e) => set({ password: e.target.value })} onKeyDown={(e) => { if (e.key === 'Enter' && ready && form.password && !adding) add(); }} />
        </Field>

        <div className="ml-search" data-state={looking ? 'looking' : found ? (found.imap ? 'found' : 'nothing') : 'idle'}>
          {looking ? (
            <div className="ml-search-line"><Spinner /> {t('Looking up where {domain} keeps its mail…', { domain: form.email.split('@')[1] })}</div>
          ) : found ? (
            <>
              {found.imap ? (
                <div className="ml-search-line">
                  <Icon name={found.imap.verified ? 'check' : 'info'} size={14} />
                  <span>{t('Incoming')} {server(found.imap.host, found.imap.port, security(found.imap), found.imap.verified)}</span>
                </div>
              ) : null}
              {found.smtp ? (
                <div className="ml-search-line">
                  <Icon name={found.smtp.verified ? 'check' : 'info'} size={14} />
                  <span>{t('Outgoing')} {server(found.smtp.host, found.smtp.port, security(found.smtp), found.smtp.verified)}</span>
                </div>
              ) : null}
              {found.note ? (
                <div className="ml-note">
                  <Icon name="info" size={14} />
                  <span className="grow">{found.note}</span>
                  {found.appPassword ? (
                    <Button
                      className="ml-app-password"
                      label={t('Make an app password')}
                      title={t('Make an app password for {provider} — {url}', { provider: found.provider?.label || t('this account'), url: found.appPassword.url })}
                      onClick={() => shell.shell.openExternal({ url: found.appPassword.url })}
                    />
                  ) : null}
                </div>
              ) : null}
              {found.steps?.length ? (
                <button type="button" className="ml-steps-toggle" onClick={() => setShowSteps((s) => !s)}>
                  {showSteps ? t('Hide what was checked') : t('What was checked')}
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
            <div className="ml-note"><Icon name="info" size={14} />{t('These settings came from {source} on this computer; they are checked when the address is complete.', { source: seed.source })}</div>
          ) : (
            <div className="ml-search-line quiet">{t('The mail server is found from the address. If it is not, Advanced takes the names from your provider.')}</div>
          )}
        </div>

        {advanced ? (
          <div className="ml-advanced">
            <Field label={t('Your name')}>
              <Input value={form.name} onChange={(e) => set({ name: e.target.value })} placeholder={t('How your name appears on messages you send')} />
            </Field>
            <Field label={t('Username')} hint={t('Only when it is not the address itself.')}>
              <Input value={form.user} onChange={(e) => set({ user: e.target.value })} placeholder={form.email || 'you@example.com'} />
            </Field>
            <div className="ml-servers3">
              <Field label={t('Incoming server (IMAP)')}>
                <Input value={form.imapHost} onChange={(e) => set({ imapHost: e.target.value })} placeholder="imap.example.com" />
              </Field>
              <Field label={t('Port')}>
                <Input type="number" value={form.imapPort} onChange={(e) => set({ imapPort: e.target.value })} />
              </Field>
              <Field label={t('Security')}>
                <select className="rw-input" value={form.imapSecurity} onChange={(e) => set({ imapSecurity: e.target.value, imapPort: e.target.value === 'tls' ? 993 : 143 })}>
                  <option value="tls">TLS</option>{/* words-ok: protocol names */}
                  <option value="starttls">STARTTLS</option>
                  <option value="none">{t('None')}</option>
                </select>
              </Field>
              <Field label={t('Outgoing server (SMTP)')}>
                <Input value={form.smtpHost} onChange={(e) => set({ smtpHost: e.target.value })} placeholder="smtp.example.com" />
              </Field>
              <Field label={t('Port')}>
                <Input type="number" value={form.smtpPort} onChange={(e) => set({ smtpPort: e.target.value })} />
              </Field>
              <Field label={t('Security')}>
                <select className="rw-input" value={form.smtpSecurity} onChange={(e) => set({ smtpSecurity: e.target.value, smtpPort: e.target.value === 'tls' ? 465 : 587 })}>
                  <option value="tls">TLS</option>{/* words-ok: protocol names */}
                  <option value="starttls">STARTTLS</option>
                  <option value="none">{t('None')}</option>
                </select>
              </Field>
            </div>
          </div>
        ) : null}

        {result ? (
          <div className={`ml-result${result.error ? ' bad' : ' good'}`}>
            <Icon name={result.error ? 'info' : 'check'} size={14} />
            <span>
              {t('Incoming')} {result.imap?.ok ? tn(result.imap.folders, 'signed in — {count} folder', 'signed in — {count} folders') : t('failed: {reason}', { reason: result.imap?.message })}
              {' · '}
              {t('Outgoing')} {result.smtp?.ok ? t('signed in') : t('failed: {reason}', { reason: result.smtp?.message })}
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
        <Search value={query} onChange={setQuery} placeholder={t('Search attachments by name or subject')} />
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
              <div className="sub">{f.subject || t('(no subject)')}</div>
              <div className="sub">{formatWhen(f.date)}</div>
            </button>
          ))}
        </div>
      ) : (
        <Empty icon="attach" title={query ? t('Nothing matches') : t('No attachments yet')}>
          {t('Every file anyone has sent you appears here, whatever folder it landed in.')}
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
        <Search value={query} onChange={setQuery} placeholder={t('Search people')} />
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
                  {p.received ? tn(p.received, '{count} received', '{count} received') : ''}
                  {p.received && p.sent ? ' · ' : ''}
                  {p.sent ? tn(p.sent, '{count} sent', '{count} sent') : ''}
                </span>
              </button>
            );
          })}
        </div>
      ) : (
        <Empty icon="reply" title={t('Nobody yet')}>{t('This fills in as mail arrives, and after an import.')}</Empty>
      )}
    </div>
  );
}

/* ── importing progress ──────────────────────────────────────────────────── */

export function ImportingDialog({ path, progress }) {
  return (
    <Dialog title={t('Importing')} onClose={undefined}>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 10, padding: '6px 0 14px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <Spinner />
          <span>{progress ? tn(progress.done, '{count} message — {folder}', '{count} messages — {folder}', { folder: progress.folder }) : t('Reading the archive…')}</span>
        </div>
        <Progress value={progress?.done ? (progress.done % 500) / 500 : 0.15} />
        <div className="rw-hint">{path}</div>
      </div>
    </Dialog>
  );
}

/* ── junk email options ───────────────────────────────────────────────────── */

const JUNK_CHOICES = [
  ['off', t('No automatic filtering'), t('Only senders on your Blocked Senders list go to Junk.')],
  ['low', t('Low'), t('The most obvious junk goes to Junk.')],
  ['high', t('High'), t('More junk is caught, and now and then a good message with it — look in Junk once in a while.')],
  ['safeOnly', t('Safe Lists Only'), t('Anything not from a Safe Sender or a contact goes to Junk.')],
];

/**
 * Junk Email Options, as Outlook lays them out: how hard the filter looks,
 * the Safe Senders, Safe Recipients and Blocked Senders lists, the
 * International lists — top-level domains and encodings blocked — and,
 * what Outlook does not show, what the filter has learned so far, with a
 * way to teach it from the mail already here. Each account keeps its own,
 * as in Outlook; the learning is one for all of them.
 */
export function JunkDialog({ shell, accountId, accounts = [], onClose, onSaved, toast }) {
  const [form, setForm] = useState(null);
  const [who, setWho] = useState(accountId || accounts[0]?.id || null);
  const [saving, setSaving] = useState(false);
  const [learning, setLearning] = useState(false);
  const set = (patch) => setForm((f) => ({ ...f, ...patch }));
  const load = (j) => setForm({
    level: j.level, safe: j.safe.join('\n'), blocked: j.blocked.join('\n'), trustContacts: j.trustContacts,
    safeRecipients: (j.safeRecipients || []).join('\n'), blockedTlds: (j.blockedTlds || []).join('\n'), blockedEncodings: j.blockedEncodings || [], encodings: j.encodings || [],
    learned: j.learned, ready: j.ready, minimum: j.minimum,
  });

  useEffect(() => {
    setForm(null);
    shell.mail.junk({ accountId: who }).then(load).catch((err) => toast(err.message, { tone: 'bad' }));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [who]);

  const lines = (text) => String(text || '').split(/[\n,;]+/).map((s) => s.trim()).filter(Boolean);

  const save = useCallback(async () => {
    if (!form) return;
    setSaving(true);
    try {
      await shell.mail.setJunk({
        accountId: who,
        patch: {
          level: form.level, safe: lines(form.safe), blocked: lines(form.blocked), trustContacts: form.trustContacts,
          safeRecipients: lines(form.safeRecipients), blockedTlds: lines(form.blockedTlds), blockedEncodings: form.blockedEncodings,
        },
      });
      toast(t('Junk email options saved'), { tone: 'good' });
      onSaved?.();
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    } finally {
      setSaving(false);
    }
  }, [shell, form, who, onSaved, toast]);

  const learn = useCallback(async () => {
    if (!who) return;
    setLearning(true);
    try {
      const j = await shell.mail.learnJunk({ accountId: who });
      setForm((f) => ({ ...f, learned: j.learned, ready: j.ready }));
      toast(tn(j.added.good, 'Learned from {junk} junk and {count} good message', 'Learned from {junk} junk and {count} good messages', { junk: j.added.junk }), { tone: 'good' });
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    } finally {
      setLearning(false);
    }
  }, [shell, who, toast]);

  const account = accounts.find((a) => a.id === who);
  return (
    <Dialog
      title={t('Junk email options')}
      width={600}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={saving ? t('Saving…') : t('Save')} disabled={saving || !form} onClick={save} />
        </>
      }
    >
      {accounts.length > 1 ? (
        <Field label={t('Account')} hint={t('Each account keeps its own options, as in Outlook.')}>
          <select className="rw-input ml-junk-account" value={who || ''} onChange={(e) => setWho(e.target.value)}>
            {accounts.map((a) => (
              <option key={a.id} value={a.id}>{a.name ? `${a.name} <${a.email}>` : a.email}</option>
            ))}
          </select>
        </Field>
      ) : account ? <div className="rw-hint ml-junk-for">{t('For {address}', { address: account.email })}</div> : null}
      {!form ? <Spinner /> : (
        <>
          <div className="ml-junk-levels" role="radiogroup" aria-label={t('How hard the filter looks')}>
            {JUNK_CHOICES.map(([value, label, hint]) => (
              <label key={value} className={`ml-junk-level${form.level === value ? ' on' : ''}`}>
                <input type="radio" name="ml-junk-level" value={value} checked={form.level === value} onChange={() => set({ level: value })} />
                <span><strong>{label}</strong><span className="rw-hint">{hint}</span></span>
              </label>
            ))}
          </div>

          <div className="ml-junk-learned">
            <span>
              {form.ready
                ? t('The filter has learned from {junk} junk and {good} good messages.', { junk: form.learned.junk, good: form.learned.good })
                : t('The filter has learned from {junk} junk and {good} good messages, and judges nothing until it has seen {minimum} of each. Mark messages Junk and Not junk, or learn from what is already here.', { junk: form.learned.junk, good: form.learned.good, minimum: form.minimum })}
            </span>
            <Button label={learning ? t('Learning…') : t('Learn from my folders')} title={t('What is in Junk is learned as junk, the Inbox as good')} disabled={learning || !who} onClick={learn} />
          </div>

          <label className="ml-ooo-toggle ml-junk-contacts">
            <input type="checkbox" checked={form.trustContacts} onChange={(e) => set({ trustContacts: e.target.checked })} />
            <span>{t('Trust email from my contacts')}</span>
          </label>

          <div className="ml-servers3" style={{ gridTemplateColumns: '1fr 1fr' }}>
            <Field label={t('Safe Senders')} hint={t('Never junk. One address or domain a line.')}>
              <textarea className="rw-input ml-junk-safe" rows={6} value={form.safe} onChange={(e) => set({ safe: e.target.value })} placeholder={'friend@example.com\nexample.org'} />
            </Field>
            <Field label={t('Blocked Senders')} hint={t('Always junk.')}>
              <textarea className="rw-input ml-junk-blocked" rows={6} value={form.blocked} onChange={(e) => set({ blocked: e.target.value })} placeholder="@offers.example" /* words-ok: an example address */ />
            </Field>
            <Field label={t('Safe Recipients')} hint={t('Mail sent to these — a group or a mailing list — is never junk.')}>
              <textarea className="rw-input ml-junk-recipients" rows={4} value={form.safeRecipients} onChange={(e) => set({ safeRecipients: e.target.value })} placeholder="team@lists.example.org" /* words-ok: an example address */ />
            </Field>
            <Field label={t('Blocked Top-Level Domains')} hint={t('Mail from addresses ending in these is junk, such as ru or cn.')}>
              <textarea className="rw-input ml-junk-tlds" rows={4} value={form.blockedTlds} onChange={(e) => set({ blockedTlds: e.target.value })} placeholder={'ru\ncn'} />
            </Field>
          </div>

          <Field label={t('Blocked Encodings')} hint={t("Mail written in these languages' character sets is junk.")}>
            <div className="ml-junk-encodings">
              {form.encodings.map((e) => (
                <label key={e.key} className="ml-junk-encoding">
                  <input
                    type="checkbox"
                    value={e.key}
                    checked={form.blockedEncodings.includes(e.key)}
                    onChange={(ev) => set({ blockedEncodings: ev.target.checked ? [...form.blockedEncodings, e.key] : form.blockedEncodings.filter((k) => k !== e.key) })}
                  />
                  <span>{e.label}</span>
                </label>
              ))}
            </div>
          </Field>
        </>
      )}
    </Dialog>
  );
}
