// The dialogs and secondary views.
//
// Account setup, importing, and the two views that treat the mailbox as
// something other than a list of messages: everything you have been sent as
// files, and everyone who has ever written to you. Both are derived from the
// index that is already there, which is why a client that owns its own store
// can offer them and a thin IMAP front end cannot.

import React, { useCallback, useEffect, useRef, useState } from 'react';
import {
  Button, Dialog, Field, Input, Icon, Chip, Empty, Spinner, Search, Progress, Separator, Select,
  formatBytes, formatWhen, t, tn,
} from '@rutba/office-ui';
import { avatarFor, displayName, fileNameOf } from './parts.js';
import { cleanHtml, htmlToText, LinkRow, selectionIn } from './richtext.js';
import { signatureTextToHtml } from '@rutba/mailbox/signature';
import { AiStepPicker } from '../../ai-tools.js';

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
    const [file] = (await shell.dialog.open({ title: t('common.insert-a-picture'), filters: [{ name: t('common.pictures'), extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp'] }] })) || [];
    if (!file) return;
    const { bytes } = await shell.fs.read({ path: file });
    if (bytes.length > 256 * 1024) {
      toast(t('mail.dialogs.picture-over-256'), { tone: 'bad' });
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
      toast(t('mail.dialogs.signature-saved'), { tone: 'good' });
      onSaved();
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    } finally {
      setSaving(false);
    }
  }, [shell, id, onSaved, toast]);

  return (
    <Dialog
      title={t('common.signature')}
      width={560}
      onClose={onClose}
      actions={
        <>
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary label={saving ? t('common.saving') : t('common.save')} disabled={saving || !id} onClick={save} />
        </>
      }
    >
      {accounts.length > 1 ? (
        <Field label={t('common.account')}>
          <select className="rw-input ml-signature-account" value={id} onChange={(e) => setId(e.target.value)}>
            {accounts.map((a) => (
              <option key={a.id} value={a.id}>{a.name ? `${a.name} <${a.email}>` : a.email}</option>
            ))}
          </select>
        </Field>
      ) : null}
      <Field label={t('mail.dialogs.added-end-new')}>
        <div className="ml-toolbar ml-signature-tools">
          <Button icon="bold" title={t('common.bold')} onMouseDown={keep} onClick={() => run('bold')} />
          <Button icon="italic" title={t('common.italic')} onMouseDown={keep} onClick={() => run('italic')} />
          <Button icon="underline" title={t('common.underline')} onMouseDown={keep} onClick={() => run('underline')} />
          <Separator />
          {SIGNATURE_COLOURS.map(([colour, name]) => (
            <button key={colour} type="button" className="ml-signature-colour" title={t('mail.dialogs.colour-name', { name })} aria-label={t('mail.dialogs.colour-name', { name })} style={{ background: colour }} onMouseDown={keep} onClick={() => run('foreColor', colour)} />
          ))}
          <Separator />
          <Button label="A−" title={t('mail.smaller')} onMouseDown={keep} onClick={() => run('fontSize', '2')} />
          <Button label="A+" title={t('mail.dialogs.larger')} onMouseDown={keep} onClick={() => run('fontSize', '4')} />
          <Separator />
          <Button icon="link" title={t('common.link')} onMouseDown={keep} onClick={() => setLinking({ range: selectionIn(editor.current) })} />
          <Button icon="picture" title={t('common.picture')} onMouseDown={keep} onClick={picture} />
          <Button icon="undo" title={t('common.clear-formatting')} onMouseDown={keep} onClick={() => run('removeFormat')} />
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
          aria-label={t('common.signature')}
          data-placeholder={t('mail.dialogs.left-blank-nothing')}
        />
      </Field>
    </Dialog>
  );
}

/** The colours the signature editor offers: the house blue, and the ones signatures use. */
const SIGNATURE_COLOURS = [['#1a1c20', t('common.black')], ['#5f6368', t('common.grey')], ['#2b5fd9', t('common.blue')], ['#0f9d58', t('common.green')], ['#c00000', t('common.red')], ['#7b5cd6', t('common.purple')]];

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
/** Why an account's mail server cannot send the replies itself, in words. */
export function sieveReason(code, answer = null) {
  switch (code) {
    case 'provider': return t('mail.dialogs.provider-does-not', { provider: answer?.provider || t('mail.dialogs.this-mail-service') });
    case 'no-vacation': return t('mail.dialogs.mail-server-cannot');
    case 'no-tls': return t('mail.dialogs.mail-server-offers');
    case 'auth': return t('mail.dialogs.mail-server-did');
    case 'oauth': return t('mail.dialogs.not-yet-accounts');
    case 'local': return t('mail.dialogs.imported-mailbox-no');
    default: return t('mail.mail-server-could');
  }
}

export function OutOfOfficeDialog({ shell, accounts, accountId, onClose, onSaved, toast }) {
  const [id, setId] = useState(accountId || accounts[0]?.id || '');
  const account = accounts.find((a) => a.id === id) || null;
  const blank = { enabled: false, start: '', end: '', subject: '', message: '', contactsOnly: false, sentTo: [], where: null };
  const [form, setForm] = useState({ ...blank, ...(account?.autoReply || {}) });
  const [saving, setSaving] = useState(false);
  // What the account's mail server says: whether it can send the replies itself, and whether it is.
  const [server, setServer] = useState(null);
  // Changed by the person: the server's answer, arriving after, does not undo it.
  const touched = useRef(false);
  const set = (patch) => { touched.current = true; setForm((f) => ({ ...f, ...patch })); };

  useEffect(() => {
    setForm({ ...blank, ...((accounts.find((a) => a.id === id) || {}).autoReply || {}) });
    setServer(null);
    touched.current = false;
    if (!id) return undefined;
    let gone = false;
    shell.mail.serverAutoReply({ accountId: id, action: 'state' }).then((r) => {
      if (gone) return;
      setServer(r);
      // The server's, where it can: replies go out with this computer off.
      setForm((f) => {
        const where = f.where || (r.supported ? 'server' : 'office');
        // Turned on from another computer: what the server holds is what is shown.
        const fromServer = r.on && r.settings && !touched.current ? { enabled: true, subject: r.settings.subject || '', message: r.settings.message || '', start: r.settings.start || '', end: r.settings.end || '' } : {};
        return { ...f, ...(where === 'server' ? fromServer : {}), where: r.supported ? where : 'office' };
      });
    }, () => { if (!gone) { setServer({ supported: false, reason: 'unreachable' }); setForm((f) => ({ ...f, where: 'office' })); } });
    return () => { gone = true; };
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
      const where = form.where === 'server' && server?.supported ? 'server' : 'office';
      const patch = { enabled: form.enabled, start: form.start || null, end: form.end || null, subject: form.subject, message: form.message, contactsOnly: form.contactsOnly, sentTo, where };
      if (where === 'server') {
        const r = await shell.mail.serverAutoReply({ accountId: id, action: form.enabled ? 'on' : 'off', settings: { subject: form.subject, message: form.message, start: form.start || null, end: form.end || null } });
        if (!r.supported) throw new Error(sieveReason(r.reason, r));
      } else if (server?.on) {
        // Moved to Rutba Office's own: the server's is turned off, so nobody hears twice.
        await shell.mail.serverAutoReply({ accountId: id, action: 'off' });
      }
      await shell.mail.updateAccount({ id, patch: { autoReply: patch } });
      toast(patch.enabled ? (where === 'server' ? t('mail.dialogs.automatic-replies-mail') : t('mail.dialogs.automatic-replies-are-on')) : t('mail.dialogs.automatic-replies-off'), { tone: 'good' });
      onSaved();
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    } finally {
      setSaving(false);
    }
  }, [shell, id, form, account, server, onSaved, toast]);

  return (
    <Dialog
      title={t('mail.dialogs.automatic-replies')}
      width={520}
      onClose={onClose}
      actions={
        <>
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary label={saving ? t('common.saving') : t('common.save')} disabled={saving || !id} onClick={save} />
        </>
      }
    >
      {accounts.length > 1 ? (
        <Field label={t('common.account')}>
          <select className="rw-input ml-signature-account" value={id} onChange={(e) => setId(e.target.value)}>
            {accounts.map((a) => (
              <option key={a.id} value={a.id}>{a.name ? `${a.name} <${a.email}>` : a.email}</option>
            ))}
          </select>
        </Field>
      ) : null}

      <label className="ml-ooo-toggle ml-ooo-enabled">
        <input type="checkbox" checked={form.enabled} onChange={(e) => set({ enabled: e.target.checked })} />
        <span>{t('mail.dialogs.send-automatic')}</span>
      </label>

      <fieldset className="ml-ooo-where" aria-label={t('mail.dialogs.sent-by')}>
        <legend>{t('mail.dialogs.sent-by')}</legend>
        <label className="ml-ooo-way" data-where="server">
          <input type="radio" name="ml-ooo-where" checked={form.where === 'server'} disabled={!server?.supported} onChange={() => set({ where: 'server' })} />
          <span>
            <strong>{t('mail.dialogs.your-mail-server')}</strong>
            <small>{server === null ? t('mail.dialogs.asking-mail-server') : server.supported ? t('mail.dialogs.replies-go-out') : sieveReason(server.reason, server)}</small>
          </span>
        </label>
        <label className="ml-ooo-way" data-where="office">
          <input type="radio" name="ml-ooo-where" checked={form.where !== 'server'} onChange={() => set({ where: 'office' })} />
          <span><strong>{t('mail.dialogs.rutba-office')}</strong><small>{t('mail.dialogs.while-rutba-office')}</small></span>
        </label>
      </fieldset>

      <div className="ml-servers3" style={{ gridTemplateColumns: '1fr 1fr' }}>
        <Field label={t('mail.dialogs.start-optional')}>
          <input type="datetime-local" className="rw-input" value={forDateInput(form.start)} onChange={(e) => set({ start: fromDateInput(e.target.value) })} />
        </Field>
        <Field label={t('mail.dialogs.end-optional')}>
          <input type="datetime-local" className="rw-input" value={forDateInput(form.end)} onChange={(e) => set({ end: fromDateInput(e.target.value) })} />
        </Field>
      </div>
      {form.where === 'server' ? (
        <p className="rw-hint" style={{ marginTop: -4 }}>{t('mail.dialogs.left-blank-starts')}</p>
      ) : (
      <p className="rw-hint" style={{ marginTop: -4 }}>
        {t('mail.dialogs.left-blank-starts-moment')}
      </p>
      )}

      <Field label={t('common.subject')} hint={t('mail.dialogs.leave-blank-automatic')}>
        <Input value={form.subject} onChange={(e) => set({ subject: e.target.value })} placeholder={t('mail.dialogs.automatic-reply-original')} />
      </Field>
      <Field label={t('mail.message')}>
        <textarea
          className="rw-input ml-compose-body ml-ooo-message"
          rows={6}
          value={form.message}
          onChange={(e) => set({ message: e.target.value })}
          placeholder={t('mail.dialogs.im-away-until')}
        />
      </Field>

      <label className="ml-ooo-toggle ml-ooo-contacts">
        <input type="checkbox" checked={form.contactsOnly && form.where !== 'server'} disabled={form.where === 'server'} onChange={(e) => set({ contactsOnly: e.target.checked })} />
        <span>{form.where === 'server' ? t('mail.dialogs.only-reply-people-contacts') : t('mail.dialogs.only-reply-people')}</span>
      </label>

      <p className="rw-hint" style={{ marginBottom: 0 }}>
        {t('mail.dialogs.each-sender-hears')}
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
            <div className="who">{t('mail.dialogs.looking-mail-already')}</div>
            <div className="what">{t('mail.dialogs.outlook-thunderbird')}</div>
          </span>
        </div>
      </div>
    );
  }

  if (!scan.accounts?.length) {
    return (
      <p className="rw-hint" style={{ marginTop: 0 }}>
        {t('mail.dialogs.no-other-mail')}
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
                {account.storeCount ? ` · ${tn(account.storeCount, 'mail.dialogs.count-data-file')}` : ''}
                {account.messagesHint ? ` · ${formatBytes(account.messagesHint)}` : ''}
                {account.incoming?.host ? ` · ${account.incoming.host}` : ''}
              </div>
            </span>
            {account.stores?.length ? <Chip>{t('common.import')}</Chip> : null}
          </button>

          {/* More than one data file: name them, because "4 stores" is not a choice. */}
          {account.stores?.length > 1
            ? account.stores.map((store) => (
                <button
                  key={store.path}
                  type="button"
                  className="ml-found-item indent"
                  style={{ padding: '6px 11px' }}
                  onClick={() => onImportStore(store.path)}
                >
                  <Icon name="file" size={14} />
                  <span className="grow">
                    <div className="what" style={{ fontFamily: 'var(--mono)', fontSize: 11 }}>{fileNameOf(store.path)}</div>
                  </span>
                  <span className="what">{formatBytes(store.bytes)}</span>
                </button>
              ))
            : null}

          {account.incoming?.host ? (
            <button type="button" className="ml-found-item indent" style={{ padding: '6px 11px' }} onClick={() => onUseAccount(account)}>
              <Icon name="settings" size={14} />
              <span className="grow">
                <div className="what">
                  {t('mail.dialogs.set-account-protocol', { protocol: account.incoming.protocol.toUpperCase(), host: account.incoming.host })}
                  {account.outgoing?.host ? `, SMTP ${account.outgoing.host}` : ''}
                </div>
              </span>
              <Chip>{t('mail.dialogs.use-settings')}</Chip>
            </button>
          ) : null}
        </div>
      ))}
      <p className="rw-hint" style={{ margin: 0 }}>
        {t('mail.dialogs.nothing-read-only')}
      </p>
    </div>
  );
}

/* ── import ──────────────────────────────────────────────────────────────── */

export function ImportDialog({ scan, onClose, onChoose, onImportStore, onUseAccount, onImportAccounts }) {
  return (
    <Dialog
      title={t('common.import-mail')}
      width={600}
      onClose={onClose}
      actions={
        <>
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary label={t('mail.dialogs.choose-a')} onClick={onChoose} />
        </>
      }
    >
      <Discovered scan={scan} onImportStore={onImportStore} onUseAccount={onUseAccount} />
      <p style={{ marginTop: 4, marginBottom: 6, fontWeight: 600, fontSize: 12.5 }}>{t('mail.dialogs.open-archive-directly')}</p>
      <ul className="ml-formats">
        <li><strong>.pst</strong> {t('mail.and')} <strong>.ost</strong> {'\u2014'} {t('mail.dialogs.outlook-data-files')}</li>{/* words-ok: file extensions */}
        <li><strong>.olm</strong> {'\u2014'} {t('mail.dialogs.outlook-mac-archives')}</li>
        <li><strong>.mbox</strong> — Thunderbird, Apple Mail, Google Takeout{/* words-ok: product names */}</li>
        <li><strong>.eml</strong>, <strong>.msg</strong> {'\u2014'} {t('mail.dialogs.single-messages-folder')}</li>
      </ul>
      <p style={{ marginTop: 10, marginBottom: 6, fontWeight: 600, fontSize: 12.5 }}>{t('mail.dialogs.set-accounts-file')}</p>
      <p className="rw-hint" style={{ marginBottom: 6 }}>
        {t('mail.dialogs.json-file-accounts')}
      </p>
      <Button label={t('mail.dialogs.choose')} onClick={onImportAccounts} className="ml-import-accounts" />
      <p className="rw-hint" style={{ marginTop: 10 }}>{t('mail.dialogs.nothing-sent-anywhere')}</p>
    </Dialog>
  );
}

export function ImportPreview({ found, onCancel, onImport, onOpenInPlace = null }) {
  const [chosen, setChosen] = useState(() => new Set((found.folders || []).filter((f) => f.messages > 0).map((f) => f.path || f.name)));
  const all = (found.folders || []).filter((f) => f.messages > 0);
  const total = (found.folders || []).filter((f) => chosen.has(f.path || f.name)).reduce((n, f) => n + f.messages, 0);

  if (!found.readable) {
    return (
      <Dialog title={t('mail.dialogs.file-cannot-read')} onClose={onCancel} actions={<Button primary label={t('common.close')} onClick={onCancel} />}>
        <p style={{ marginTop: 0 }}>{found.label}</p>
        <p className="rw-hint">{found.encoding ? t('mail.dialogs.uses-encoding-encoding', { encoding: found.encoding }) : t('mail.dialogs.format-not-recognised')}</p>
      </Dialog>
    );
  }

  return (
    <Dialog
      title={t('mail.dialogs.import-from-source', { source: found.label })}
      width={560}
      onClose={onCancel}
      actions={
        <>
          <Button label={t('common.cancel')} onClick={onCancel} />
          {onOpenInPlace ? (
            <Button icon="folderOpen" label={t('mail.dialogs.open')} title={t('mail.dialogs.open-read-place')} onClick={onOpenInPlace} />
          ) : null}
          <Button primary label={tn(total, 'mail.dialogs.import-count')} disabled={!total} onClick={() => onImport([...chosen])} />
        </>
      }
    >
      <div className="ml-import-summary">
        <Chip>{formatBytes(found.bytes || 0)}</Chip>
        <Chip>{tn((found.folders || []).length, 'common.count-folder')}</Chip>
        <Chip>{tn(found.messages || 0, 'mail.count-message')}</Chip>
        {found.encoding ? <Chip>{found.encoding}</Chip> : null}
      </div>
      <div className="ml-listbar" style={{ borderTop: '1px solid var(--line-soft)' }}>
        <input
          type="checkbox"
          className="ml-check"
          checked={chosen.size === all.length && all.length > 0}
          onChange={(e) => setChosen(e.target.checked ? new Set(all.map((f) => f.path || f.name)) : new Set())}
        />
        <span>{tn(all.length, 'mail.dialogs.chosen-count-folder', { chosen: chosen.size })}</span>
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
      <p className="rw-hint">{t('mail.dialogs.importing-same-archive')}</p>
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
      <div className="ml-providers-head">{t('mail.dialogs.choose-your-provider')}</div>
      <div className="ml-providers-row">
        {providers.map((p) => (
          <button key={p.id} type="button" className="ml-provider" data-provider={p.id} onClick={() => onChoose(p)}>
            <span className="label">{p.tileLabel}</span>
            <span className="domain">@{p.domain}</span>
          </button>
        ))}
        <button type="button" className="ml-provider" data-provider="other" onClick={onOther}>
          <span className="label">{t('common.other')}</span>
          <span className="domain">{t('mail.dialogs.enter-address-servers')}</span>
        </button>
      </div>
    </div>
  );
}

/* ── account setup ───────────────────────────────────────────────────────── */

export function AccountDialog({ shell, seed, onClose, onSaved, onSeveral, toast }) {
  // An address and a password. Everything else is found: while the address
  // is typed, from the built-in table of providers alone, offline; when Add
  // is pressed, for a domain the table does not know, from the domain's own
  // records - its DNS (SRV and MX), its own autoconfig and autodiscover
  // addresses - and a knock on its conventional names (main/mail-discover.js),
  // the window saying what it is looking up and where. What was found is
  // shown, what was checked can be opened, and the advanced fields — always one
  // click away, never the only way in — are filled with the answer so that a
  // wrong guess is a correction, not a form from scratch.
  // TLS from the first byte, or an upgrade the connection insists on: there is no unencrypted choice.
  const security = (server) => (!server ? 'tls' : server.secure ? 'tls' : 'starttls');
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
  // The domain being looked up on the network, while Add does so.
  const [lookingUp, setLookingUp] = useState(null);
  // Servers the domain's records named outside it, waiting for the person to agree before the password goes there.
  const [agree, setAgree] = useState(null);
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
        // Offline: nothing leaves the computer until Add is pressed.
        const r = await shell.mail.autodiscover({
          email,
          offline: true,
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
        setFound({ imap: null, smtp: null, note: err.message, steps: [{ name: t('mail.dialogs.the-search'), status: 'failed', detail: err.message }] });
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
        // "The servers for": a tile named for several services ("Outlook, Hotmail and Microsoft 365") reads as one sentence.
        note: entry.note || t('mail.dialogs.servers-provider-filled', { provider: entry.tileLabel || entry.label }),
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

  // The browser a sign-in opens in: the person's choice, remembered. The
  // picker is there from the start with the default browser in it, and the
  // others join it once Windows has listed them, which on a busy computer can
  // take many seconds. A list that came back empty is asked for again the
  // next time the dialog opens (the main process keeps only a whole one).
  const [browsers, setBrowsers] = useState([]);
  const [browser, setBrowser] = useState('');
  const chosen = useRef(false);
  useEffect(() => {
    if (!provider) return undefined;
    let live = true;
    Promise.all([shell.oauth.browsers().catch(() => []), shell.store.get({ key: 'mail.signInBrowser', fallback: '' }).catch(() => '')]).then(([list, saved]) => {
      if (!live) return;
      setBrowsers(list || []);
      if (!chosen.current) setBrowser(list?.some((b) => b.id === saved) ? saved : '');
    });
    return () => {
      live = false;
    };
  }, [provider, shell]);
  const chooseBrowser = (id) => {
    chosen.current = true;
    setBrowser(id);
    shell.store.set({ key: 'mail.signInBrowser', value: id }).catch(() => {});
  };

  const signIn = useCallback(async () => {
    setSigningIn(true);
    try {
      const signed = await shell.oauth.signIn({ email: form.email.trim(), provider: provider.id, browser: browser || null });
      await shell.mail.addAccount({
        account: { email: signed.email, name: signed.name || form.name || signed.email, imap: signed.imap, smtp: signed.smtp, auth: 'oauth', provider: signed.provider },
      });
      onSaved();
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 12000 });
    } finally {
      setSigningIn(false);
    }
  }, [shell, form, provider, browser, onSaved, toast]);

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

  /**
   * The network, at last, because Add was pressed: the domain's own records
   * looked up and every server they name tried until one takes the password
   * (main/mail-setup.js). True when the account was added.
   */
  const lookUpAndAdd = async (confirmed = []) => {
    const email = form.email.trim();
    setLookingUp(email.split('@')[1] || '');
    setAgree(null);
    try {
      const tried = await shell.mail.setupAccount({ email, password: form.password, name: form.name, user: form.user.trim(), confirmed });
      if (tried?.outcome === 'confirm') {
        setAgree({ domain: email.split('@')[1] || '', foreign: tried.foreign || [] });
        return false;
      }
      if (tried?.outcome === 'not-an-address') {
        toast(t('mail.not-email-address'), { tone: 'bad', ms: 7000 });
        return false;
      }
      if (tried?.account && (tried.outcome === 'ready' || tried.outcome === 'outgoing')) {
        await shell.mail.addAccount({ account: { ...tried.account, name: form.name || tried.account.name }, password: form.password });
        onSaved();
        return true;
      }
      if (tried?.outcome === 'app-password') toast(t('mail.dialogs.provider-refused', { provider: tried.provider?.label || t('mail.the-provider') }), { tone: 'bad', ms: 9000 });
      else if (tried?.outcome === 'not-found') toast(t('mail.no-mail-server', { domain: email.split('@')[1] || '' }), { tone: 'bad', ms: 7000 });
      else if (tried?.outcome === 'password') toast(t('mail.password-refused'), { tone: 'bad', ms: 7000 });
      else if (tried?.outcome === 'unreachable') toast(t('mail.mail-server-did'), { tone: 'bad', ms: 7000 });
      setAdvanced(true);
      return false;
    } finally {
      setLookingUp(null);
    }
  };

  const add = useCallback(async () => {
    setAdding(true);
    try {
      // Servers not known yet (a domain the table does not know): looked up now.
      if (!form.imapHost || !form.smtpHost) {
        await lookUpAndAdd();
        return;
      }
      // Tried before it is kept: an account that cannot sign in is not added,
      // and the failure opens the fields it needs.
      const r = await test();
      if (r.error) {
        // The servers first found did not take it: every other server found
        // is tried, with the address and with the name before the @, before
        // the fields are opened for a person to put right.
        await lookUpAndAdd();
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

  // Add looks up what it does not know, so an address is enough; Test tries the servers as they stand.
  const ready = form.email.includes('@') && form.email.split('@')[1]?.includes('.');
  const testable = ready && form.imapHost && form.smtpHost;
  const showTiles = tilesOverride ? tilesOverride === 'open' : !form.email.trim() || !found?.provider;
  const addressPlaceholder = found?.provider?.domain ? `you@${found.provider.domain}` : 'you@example.com';
  const label = (s) => ({ tls: 'TLS', starttls: 'STARTTLS' }[s] || s);
  const server = (host, port, sec, verified) => (
    <>
      <b>{host}</b> · {port} · {label(sec)}
      {verified ? <Icon name="check" size={13} /> : null}
    </>
  );

  return (
    <Dialog
      title={t('mail.dialogs.add-mail-account')}
      width={560}
      onClose={onClose}
      actions={
        <>
          <Button label={t('common.cancel')} onClick={onClose} />
          {onSeveral ? <Button label={t('mail.dialogs.several')} title={t('mail.dialogs.several-accounts-rows')} className="ml-several-accounts" onClick={onSeveral} /> : null}
          <Button label={advanced ? t('mail.dialogs.simple') : t('mail.dialogs.advanced')} onClick={() => setAdvanced((a) => !a)} />
          <Button label={testing ? t('mail.dialogs.testing') : t('mail.dialogs.test')} disabled={testing || adding || !testable || !form.password} onClick={test} />
          <Button primary label={adding ? t('common.adding') : t('common.add-account')} disabled={adding || testing || !ready || !form.password} onClick={add} />
        </>
      }
    >
      <div className="ml-form">
        {showTiles ? (
          <ProviderTiles providers={providers} onChoose={chooseTile} onOther={otherTile} />
        ) : (
          <button type="button" className="ml-provider-change" onClick={() => setTilesOverride('open')}>
            {t('mail.dialogs.change-provider')}
          </button>
        )}

        <Field label={t('common.email-address')}>
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
                <div className="who">{signingIn ? t('mail.dialogs.waiting-for-provider', { provider: provider.label }) : t('mail.dialogs.sign-provider', { provider: provider.label })}</div>
                <div className="what">
                  {provider.configured
                    ? t('mail.dialogs.opens-own-browser')
                    : t('mail.dialogs.build-no-provider', { provider: provider.label })}
                </div>
              </span>
              {provider.configured ? <Chip>{t('mail.dialogs.recommended')}</Chip> : null}
            </button>
            <label className="ml-browser">
              <span>{t('mail.dialogs.open-sign')}</span>
              <Select value={browser} onChange={(e) => chooseBrowser(e.target.value)}>
                <option value="">{t('mail.dialogs.the-default-browser')}</option>
                {browsers.map((b) => (
                  <option key={b.id} value={b.id}>
                    {b.isDefault ? t('mail.dialogs.browser-the-default', { browser: b.name }) : b.name}
                  </option>
                ))}
              </Select>
            </label>
            <p className="rw-hint" style={{ margin: 0 }}>
              {t('mail.dialogs.provider-no-longer', { provider: provider.label })}
            </p>
          </div>
        ) : null}

        <Field label={t('common.password')} hint={t('mail.dialogs.kept-computers-keystore')}>
          <Input type="password" value={form.password} onChange={(e) => set({ password: e.target.value })} onKeyDown={(e) => { if (e.key === 'Enter' && ready && form.password && !adding) add(); }} />
        </Field>

        {agree ? (
          <div className="ml-agree" role="alert">
            <div className="ml-agree-text">
              <Icon name="info" size={15} />
              <span>{t('mail.dialogs.domain-says-mail', { domain: agree.domain })}</span>
            </div>
            <ul className="ml-agree-list" dir="ltr">
              {agree.foreign.map((x) => (
                <li key={`${x.kind}-${x.host}`}>
                  {x.kind === 'imap' ? t('mail.dialogs.incoming') : t('mail.dialogs.outgoing')} <b>{x.host}</b> · {x.port}
                </li>
              ))}
            </ul>
            <div className="ml-agree-actions">
              <Button className="ml-agree-no" label={t('mail.dialogs.do-not-use')} onClick={() => { setAgree(null); setAdvanced(true); }} />
              <Button primary className="ml-agree-yes" label={t('mail.use-these')} disabled={adding} onClick={async () => { const hosts = agree.foreign.map((x) => x.host); setAdding(true); try { await lookUpAndAdd(hosts); } finally { setAdding(false); } }} />
            </div>
          </div>
        ) : null}

        <div className="ml-search" data-state={looking ? 'looking' : found ? (found.imap ? 'found' : 'nothing') : 'idle'}>
          {lookingUp ? (
            <div className="ml-search-line ml-looking-up"><Spinner /> {t('mail.looking-domain-dns', { domain: lookingUp })}</div>
          ) : looking ? (
            <div className="ml-search-line"><Spinner /> {t('mail.dialogs.looking-domain-keeps', { domain: form.email.split('@')[1] })}</div>
          ) : found ? (
            <>
              {found.imap ? (
                <div className="ml-search-line">
                  <Icon name={found.imap.verified ? 'check' : 'info'} size={14} />
                  <span>{t('mail.dialogs.incoming')} {server(found.imap.host, found.imap.port, security(found.imap), found.imap.verified)}</span>
                </div>
              ) : null}
              {found.smtp ? (
                <div className="ml-search-line">
                  <Icon name={found.smtp.verified ? 'check' : 'info'} size={14} />
                  <span>{t('mail.dialogs.outgoing')} {server(found.smtp.host, found.smtp.port, security(found.smtp), found.smtp.verified)}</span>
                </div>
              ) : null}
              {found.auth && (found.imap || found.provider) ? (
                <div className="ml-search-line ml-auth-kind" data-kind={found.auth.kind}>
                  <Icon name={found.auth.kind === 'oauth' ? 'lock' : 'info'} size={14} />
                  <span>
                    {found.auth.kind === 'oauth'
                      ? t('mail.dialogs.signs-provider-s', { provider: found.auth.provider || t('mail.dialogs.the-provider') })
                      : found.auth.kind === 'app-password'
                        ? t('mail.dialogs.provider-takes-app', { provider: found.auth.provider || t('mail.the-provider') })
                        : found.auth.oauthOffered
                          ? t('mail.dialogs.signs-accounts-password-server')
                          : t('mail.dialogs.signs-accounts-password')}
                  </span>
                </div>
              ) : null}
              {found.server ? (
                <div className="ml-search-line ml-search-server">
                  <Icon name="info" size={14} />
                  <span>{t('mail.dialogs.server-says-software', { software: found.server })}</span>
                </div>
              ) : null}
              {found.note ? (
                <div className="ml-note">
                  <Icon name="info" size={14} />
                  <span className="grow">{found.note}</span>
                  {found.appPassword ? (
                    <Button
                      className="ml-app-password"
                      label={t('mail.dialogs.make-app')}
                      title={t('mail.dialogs.make-app-password', { provider: found.provider?.label || t('mail.dialogs.this-account'), url: found.appPassword.url })}
                      onClick={() => shell.shell.openExternal({ url: found.appPassword.url })}
                    />
                  ) : null}
                </div>
              ) : null}
              {found.steps?.length ? (
                <button type="button" className="ml-steps-toggle" onClick={() => setShowSteps((s) => !s)}>
                  {showSteps ? t('mail.dialogs.hide-checked') : t('mail.dialogs.what-was-checked')}
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
            <div className="ml-note"><Icon name="info" size={14} />{t('mail.dialogs.settings-came-source', { source: seed.source })}</div>
          ) : (
            <div className="ml-search-line quiet">{t('mail.dialogs.mail-server-found')}</div>
          )}
        </div>

        {advanced ? (
          <div className="ml-advanced">
            <Field label={t('common.your-name')}>
              <Input value={form.name} onChange={(e) => set({ name: e.target.value })} placeholder={t('mail.dialogs.how-name-appears')} />
            </Field>
            <Field label={t('mail.dialogs.username')} hint={t('mail.dialogs.only-not-address')}>
              <Input value={form.user} onChange={(e) => set({ user: e.target.value })} placeholder={form.email || 'you@example.com'} />
            </Field>
            <div className="ml-servers3">
              <Field label={t('mail.dialogs.incoming-server-imap')}>
                <Input value={form.imapHost} onChange={(e) => set({ imapHost: e.target.value })} placeholder="imap.example.com" />
              </Field>
              <Field label={t('common.port')}>
                <Input type="number" value={form.imapPort} onChange={(e) => set({ imapPort: e.target.value })} />
              </Field>
              <Field label={t('mail.security')}>
                <select className="rw-input" value={form.imapSecurity} onChange={(e) => set({ imapSecurity: e.target.value, imapPort: e.target.value === 'tls' ? 993 : 143 })}>
                  <option value="tls">TLS</option>{/* words-ok: protocol names */}
                  <option value="starttls">STARTTLS</option>
                </select>
              </Field>
              <Field label={t('mail.dialogs.outgoing-server-smtp')}>
                <Input value={form.smtpHost} onChange={(e) => set({ smtpHost: e.target.value })} placeholder="smtp.example.com" />
              </Field>
              <Field label={t('common.port')}>
                <Input type="number" value={form.smtpPort} onChange={(e) => set({ smtpPort: e.target.value })} />
              </Field>
              <Field label={t('mail.security')}>
                <select className="rw-input" value={form.smtpSecurity} onChange={(e) => set({ smtpSecurity: e.target.value, smtpPort: e.target.value === 'tls' ? 465 : 587 })}>
                  <option value="tls">TLS</option>{/* words-ok: protocol names */}
                  <option value="starttls">STARTTLS</option>
                </select>
              </Field>
            </div>
          </div>
        ) : null}

        {result ? (
          <div className={`ml-result${result.error ? ' bad' : ' good'}`}>
            <Icon name={result.error ? 'info' : 'check'} size={14} />
            <span>
              {t('mail.dialogs.incoming')} {result.imap?.ok ? tn(result.imap.folders, 'mail.dialogs.signed-count-folder') : t('mail.dialogs.failed-reason', { reason: result.imap?.message })}
              {' · '}
              {t('mail.dialogs.outgoing')} {result.smtp?.ok ? t('mail.dialogs.signed-in') : t('mail.dialogs.failed-reason', { reason: result.smtp?.message })}
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
        <Search value={query} onChange={setQuery} placeholder={t('mail.dialogs.search-attachments-name')} />
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
              <div className="sub">{f.subject || t('mail.no-subject')}</div>
              <div className="sub">{formatWhen(f.date)}</div>
            </button>
          ))}
        </div>
      ) : (
        <Empty icon="attach" title={query ? t('common.nothing-matches') : t('mail.dialogs.no-attachments-yet')}>
          {t('mail.dialogs.every-file-anyone')}
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
        <Search value={query} onChange={setQuery} placeholder={t('mail.dialogs.search-people')} />
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
                  {p.received ? tn(p.received, 'mail.dialogs.count-received') : ''}
                  {p.received && p.sent ? ' · ' : ''}
                  {p.sent ? tn(p.sent, 'mail.dialogs.count-sent') : ''}
                </span>
              </button>
            );
          })}
        </div>
      ) : (
        <Empty icon="reply" title={t('mail.dialogs.nobody-yet')}>{t('mail.dialogs.fills-mail-arrives')}</Empty>
      )}
    </div>
  );
}

/* ── importing progress ──────────────────────────────────────────────────── */

export function ImportingDialog({ path, progress }) {
  return (
    <Dialog title={t('mail.dialogs.importing')} onClose={undefined}>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 10, padding: '6px 0 14px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <Spinner />
          <span>{progress ? tn(progress.done, 'mail.dialogs.count-message-folder', { folder: progress.folder }) : t('mail.dialogs.reading-the-archive')}</span>
        </div>
        <Progress value={progress?.done ? (progress.done % 500) / 500 : 0.15} />
        <div className="rw-hint">{fileNameOf(path)}</div>
      </div>
    </Dialog>
  );
}

/* ── junk email options ───────────────────────────────────────────────────── */

const JUNK_CHOICES = [
  ['off', t('mail.dialogs.no-automatic'), t('mail.dialogs.only-senders-blocked')],
  ['low', t('common.low'), t('mail.dialogs.most-obvious-junk')],
  ['high', t('common.high'), t('mail.dialogs.more-junk-caught')],
  ['safeOnly', t('mail.dialogs.safe-lists-only'), t('mail.dialogs.anything-not-safe')],
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
      toast(t('mail.dialogs.junk-email-options-saved'), { tone: 'good' });
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
      toast(tn(j.added.good, 'mail.dialogs.learned-junk-junk', { junk: j.added.junk }), { tone: 'good' });
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    } finally {
      setLearning(false);
    }
  }, [shell, who, toast]);

  const account = accounts.find((a) => a.id === who);
  return (
    <Dialog
      title={t('mail.dialogs.junk-email-options')}
      width={600}
      onClose={onClose}
      actions={
        <>
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary label={saving ? t('common.saving') : t('common.save')} disabled={saving || !form} onClick={save} />
        </>
      }
    >
      {accounts.length > 1 ? (
        <Field label={t('common.account')} hint={t('mail.dialogs.each-account-keeps')}>
          <select className="rw-input ml-junk-account" value={who || ''} onChange={(e) => setWho(e.target.value)}>
            {accounts.map((a) => (
              <option key={a.id} value={a.id}>{a.name ? `${a.name} <${a.email}>` : a.email}</option>
            ))}
          </select>
        </Field>
      ) : account ? <div className="rw-hint ml-junk-for">{t('mail.dialogs.for-address', { address: account.email })}</div> : null}
      {!form ? <Spinner /> : (
        <>
          <div className="ml-junk-levels" role="radiogroup" aria-label={t('mail.dialogs.how-hard-filter')}>
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
                ? t('mail.dialogs.filter-learned-junk', { junk: form.learned.junk, good: form.learned.good })
                : t('mail.dialogs.filter-learned-junk-junk', { junk: form.learned.junk, good: form.learned.good, minimum: form.minimum })}
            </span>
            <Button label={learning ? t('mail.dialogs.learning') : t('mail.dialogs.learn')} title={t('mail.dialogs.junk-learned-junk')} disabled={learning || !who} onClick={learn} />
          </div>

          <label className="ml-ooo-toggle ml-junk-contacts">
            <input type="checkbox" checked={form.trustContacts} onChange={(e) => set({ trustContacts: e.target.checked })} />
            <span>{t('mail.dialogs.trust-email-contacts')}</span>
          </label>

          <div className="ml-servers3" style={{ gridTemplateColumns: '1fr 1fr' }}>
            <Field label={t('mail.dialogs.safe-senders')} hint={t('mail.dialogs.never-junk-one')}>
              <textarea className="rw-input ml-junk-safe" rows={6} value={form.safe} onChange={(e) => set({ safe: e.target.value })} placeholder={'friend@example.com\nexample.org'} />
            </Field>
            <Field label={t('mail.dialogs.blocked-senders')} hint={t('mail.dialogs.always-junk')}>
              <textarea className="rw-input ml-junk-blocked" rows={6} value={form.blocked} onChange={(e) => set({ blocked: e.target.value })} placeholder="@offers.example" /* words-ok: an example address */ />
            </Field>
            <Field label={t('mail.dialogs.safe-recipients')} hint={t('mail.dialogs.mail-sent-group')}>
              <textarea className="rw-input ml-junk-recipients" rows={4} value={form.safeRecipients} onChange={(e) => set({ safeRecipients: e.target.value })} placeholder="team@lists.example.org" /* words-ok: an example address */ />
            </Field>
            <Field label={t('mail.dialogs.blocked-top-level')} hint={t('mail.dialogs.mail-addresses-ending')}>
              <textarea className="rw-input ml-junk-tlds" rows={4} value={form.blockedTlds} onChange={(e) => set({ blockedTlds: e.target.value })} placeholder={'ru\ncn'} />
            </Field>
          </div>

          <Field label={t('mail.dialogs.blocked-encodings')} hint={t('mail.dialogs.mail-written-languages')}>
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

/* ── the account manager (D-2026-10-22-09) ─────────────────────────── */

/**
 * Every account, as Add account's counterpart: each with what it is (an
 * account on a server, or a mailbox brought in from a file), its Signature,
 * its Out of office and Remove; and Add account for another. Each, and each
 * archive opened in place, with what the AI tools may do with it (ai-tools.js:
 * the same setting AI Tools lists).
 */
export function AccountsDialog({ shell, accounts, archives = [], onClose, onAdd, onRemove, onSignature, onOutOfOffice, onDefault }) {
  // "Send new mail from" (D-2026-11-11-02): new mail, and Calendar's and Talk's invitations, go from this one.
  const senders = accounts.filter((a) => a.canSend);
  const sender = senders.find((a) => a.isDefault) || senders[0] || null;
  return (
    <Dialog
      title={t('common.accounts')}
      width={620}
      onClose={onClose}
      actions={(
        <>
          <Button icon="plus" label={t('common.add-account')} className="ml-accounts-add" onClick={onAdd} />
          <span style={{ flex: 1 }} />
          <Button primary label={t('common.done')} onClick={onClose} />
        </>
      )}
    >
      {sender ? (
        <div className="ml-default-sender">
          <Field label={t('mail.dialogs.send-new-mail')} hint={t('mail.dialogs.new-messages-invitations')}>
            {senders.length > 1 ? (
              <Select value={sender.id} aria-label={t('mail.dialogs.send-new-mail')} className="rw-select ml-default-account" onChange={(e) => onDefault?.(e.target.value)}>
                {senders.map((a) => <option key={a.id} value={a.id}>{a.name && a.name !== a.email ? `${a.name} <${a.email}>` : a.email}</option>)}
              </Select>
            ) : (
              <span className="ml-default-account" data-account={sender.email}>{sender.email}</span>
            )}
          </Field>
        </div>
      ) : null}
      <ul className="ml-account-list" aria-label={t('common.accounts')}>
        {accounts.map((a) => (
          <li key={a.id} className="ml-account-row" data-account={a.email}>
            <Icon name={a.local ? 'archive' : 'mail'} size={18} />
            <span className="ml-account-words">
              <strong>{a.name || a.email}</strong>
              <small>{a.local ? t('mail.dialogs.address-brought-file', { address: a.email }) : a.autoReply?.enabled ? t('mail.dialogs.address-automatic', { address: a.email }) : a.email}</small>
            </span>
            <AiStepPicker shell={shell} kind="mail" id={a.id} name={a.email} className="ml-account-ai" />
            <Button ghost label={t('common.signature')} className="ml-account-signature" onClick={() => onSignature(a)} />
            {a.local ? null : <Button ghost label={t('mail.out-of')} className="ml-account-ooo" onClick={() => onOutOfOffice(a)} />}
            <Button ghost icon="trash" label={t('common.remove')} className="ml-account-remove" onClick={() => onRemove(a)} />
          </li>
        ))}
        {archives.map((a) => (
          <li key={a.id} className="ml-account-row ml-archive-row" data-archive={a.id}>
            <Icon name="archive" size={18} />
            <span className="ml-account-words">
              <strong>{a.name}</strong>
              <small>{t('mail.dialogs.archive-opened-place')}</small>
            </span>
            <AiStepPicker shell={shell} kind="archive" id={a.id} name={a.name} className="ml-account-ai" />
          </li>
        ))}
      </ul>
    </Dialog>
  );
}
