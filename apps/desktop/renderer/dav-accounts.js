// Accounts, shared by Calendar and Contacts: the CalDAV and CardDAV servers
// kept in step with this computer. Each account says when it last synced
// or why it could not; one is added with the server's address, a user name
// and a password (an app password, for iCloud and the like), and its
// calendars and address books come in at once. Each is called what the
// person names it (the server's name when left blank), and Rename changes
// that later. In Contacts, where a new card goes is chosen here too, and
// said in a sentence: this computer, or which account's server (an account
// added from 2.19 on takes them until chosen; a setup from before keeps
// them here, dav-sync.js). In both, what the AI tools
// may do with each calendar or address book (ai-tools.js: the same settings
// AI Tools lists).

import React, { useCallback, useEffect, useState } from 'react';
import { Button, Dialog, Field, Input, Select, t, tn } from '@rutba/office-ui';
import { AiThingsList } from './ai-tools.js';

const ago = (ms) => {
  if (!ms) return t('not yet');
  const s = Math.round((Date.now() - ms) / 1000);
  if (s < 60) return t('just now');
  if (s < 3600) return tn(Math.round(s / 60), '{count} min ago', '{count} min ago');
  if (s < 86400) return tn(Math.round(s / 3600), '{count} h ago', '{count} h ago');
  return new Date(ms).toLocaleDateString();
};

export function AccountsDialog({ shell, kind = 'calendar', toast, onClose }) {
  const [accounts, setAccounts] = useState([]);
  const [books, setBooks] = useState({ books: [], defaultBook: null });
  const [url, setUrl] = useState('');
  const [user, setUser] = useState('');
  const [password, setPassword] = useState('');
  const [name, setName] = useState('');
  const [busy, setBusy] = useState(null);
  const [error, setError] = useState('');
  // The account being renamed, and the name typed so far.
  const [renaming, setRenaming] = useState(null);
  const [again, setAgain] = useState({});

  const refresh = useCallback(async () => {
    setAccounts((await shell.dav.accounts()) || []);
    if (kind === 'contacts') setBooks((await shell.dav.books()) || { books: [], defaultBook: null });
  }, [shell, kind]);
  useEffect(() => { refresh(); }, [refresh]);
  useEffect(() => shell.on('dav:changed', refresh), [shell, refresh]);

  // An address book named with its account, since two servers may each call theirs "Contacts".
  const accountOf = (b) => accounts.find((a) => a.id === b.account)?.name || '';
  const bookLabel = (b) => (accountOf(b) ? t('{account}: {book}', { account: accountOf(b), book: b.name }) : b.name);
  /** Where a new card goes, said plainly: kept here, or sent to which account's server. */
  const where = ({ books: list, defaultBook }) => {
    const b = list.find((x) => x.id === defaultBook);
    return b
      ? t('New cards are sent to {book} on the server of {account}.', { book: b.name, account: accountOf(b) || b.name })
      : t('New cards stay on this computer and are not sent to any server.');
  };

  const add = async () => {
    if (!url.trim()) return;
    setBusy('add');
    setError('');
    try {
      const acc = await shell.dav.add({ url: url.trim(), user: user.trim(), password, name: name.trim() });
      setUrl(''); setUser(''); setPassword(''); setName('');
      toast?.(t('{name}: {calendars}, {books}', { name: acc.name, calendars: tn(acc.calendars, '{count} calendar', '{count} calendars'), books: tn(acc.books, '{count} address book', '{count} address books') }), { tone: 'good' });
      await refresh();
    } catch (err) {
      setError(String(err?.message || err).replace(/^Error invoking remote method '[^']+': (Error: )?/, ''));
    } finally {
      setBusy(null);
    }
  };
  const sync = async (id) => {
    setBusy(id);
    try {
      const [r] = (await shell.dav.sync({ id })) || [];
      if (r?.error) toast?.(r.error, { tone: 'warn', ms: 5000 });
      else if (r) toast?.(r.conflicts ? t('Synced: {received} in, {sent} out, {conflicts} kept as the server had them', { received: r.received, sent: r.sent, conflicts: r.conflicts }) : t('Synced: {received} in, {sent} out', { received: r.received, sent: r.sent }), { tone: 'good' });
    } finally {
      setBusy(null);
      refresh();
    }
  };
  // A password not saved on this computer any more, typed again here.
  const savePassword = async (id) => {
    setBusy(id);
    try {
      await shell.dav.setPassword({ id, password: again[id] || '' });
      setAgain({ ...again, [id]: '' });
    } catch (err) {
      toast?.(String(err?.message || err).replace(/^Error invoking remote method '[^']+': (Error: )?/, ''), { tone: 'warn', ms: 5000 });
    } finally {
      setBusy(null);
      refresh();
    }
  };
  const remove = async (id) => {
    await shell.dav.remove({ id });
    refresh();
  };
  const rename = async () => {
    if (!renaming?.name.trim()) return;
    try {
      await shell.dav.rename({ id: renaming.id, name: renaming.name.trim() });
      setRenaming(null);
      refresh();
    } catch (err) {
      toast?.(String(err?.message || err).replace(/^Error invoking remote method '[^']+': (Error: )?/, ''), { tone: 'warn' });
    }
  };

  return (
    <Dialog title={t('Accounts')} width={520} onClose={onClose} actions={<Button label={t('Close')} onClick={onClose} />}>
      <div className="dav-accounts" style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        {accounts.length ? accounts.map((a) => (
          <div key={a.id} className="dav-account" style={{ display: 'flex', alignItems: 'center', gap: 8, padding: 8, border: '1px solid var(--line)', borderRadius: 6 }}>
            <div style={{ flex: 1, minWidth: 0 }}>
              {renaming?.id === a.id ? (
                <Input className="dav-rename" aria-label={t('Account name')} value={renaming.name} autoFocus onChange={(e) => setRenaming({ id: a.id, name: e.target.value })}
                  onKeyDown={(e) => { if (e.key === 'Enter') rename(); if (e.key === 'Escape') { e.stopPropagation(); setRenaming(null); } }} />
              ) : <div className="dav-name" style={{ fontWeight: 600 }}>{a.name}</div>}
              <div style={{ fontSize: 12, opacity: 0.75, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{a.user ? `${a.user} · ` : ''}{a.url}</div>
              <div className="dav-status" style={{ fontSize: 12, color: a.lastError ? 'var(--bad, #c62828)' : 'inherit' }}>
                {a.lastError ? a.lastError : t('{calendars}, {books} · synced {when}', { calendars: tn(a.calendars, '{count} calendar', '{count} calendars'), books: tn(a.books, '{count} address book', '{count} address books'), when: ago(a.lastSync) })}
              </div>
              {a.needsPassword ? (
                <div className="dav-password-again" style={{ display: 'flex', gap: 6, alignItems: 'center', marginTop: 6 }}>
                  <Input className="dav-password-again-input" type="password" placeholder={t('Password')} aria-label={t('The password for {name}, not saved on this computer', { name: a.name })} value={again[a.id] || ''} onChange={(e) => setAgain({ ...again, [a.id]: e.target.value })} onKeyDown={(e) => { if (e.key === 'Enter') savePassword(a.id); }} />
                  <Button label={t('Save')} className="dav-password-again-save" disabled={Boolean(busy)} onClick={() => savePassword(a.id)} />
                </div>
              ) : null}
            </div>
            {renaming?.id === a.id ? (
              <>
                <Button primary label={t('Save')} className="dav-rename-save" disabled={!renaming.name.trim()} onClick={rename} />
                <Button label={t('Cancel')} onClick={() => setRenaming(null)} />
              </>
            ) : <Button label={t('Rename')} className="dav-rename-start" title={t('Call this account something else')} onClick={() => setRenaming({ id: a.id, name: a.name })} />}
            <Button label={busy === a.id ? t('Syncing…') : t('Sync now')} className="dav-sync" disabled={Boolean(busy)} onClick={() => sync(a.id)} />
            <Button label={t('Remove')} icon="trash" title={t('Remove the account and the calendars and cards it brought; nothing on the server is touched')} onClick={() => remove(a.id)} />
          </div>
        )) : <div style={{ opacity: 0.75 }}>{t('No accounts yet. Your calendars and contacts stay on this computer until you add one.')}</div>}
        {kind === 'contacts' && books.books.length ? (
          <Field label={t('New cards go to')} hint={<span className="dav-default-where">{where(books)}</span>}>
            <Select className="dav-default-book" value={books.defaultBook || ''} onChange={async (e) => { await shell.dav.setDefaultBook({ id: e.target.value || null }); refresh(); }}>
              <option value="">{t('This computer only')}</option>
              {books.books.map((b) => <option key={b.id} value={b.id}>{bookLabel(b)}</option>)}
            </Select>
          </Field>
        ) : null}
        <AiThingsList shell={shell} kind={kind === 'contacts' ? 'book' : 'calendar'} title={kind === 'contacts' ? t('AI tools: address books') : t('AI tools: calendars')} />
        <div style={{ borderTop: '1px solid var(--line)', paddingTop: 10, display: 'flex', flexDirection: 'column', gap: 8 }}>
          <div style={{ fontWeight: 600 }}>{t('Add an account')}</div>
          <Field label={t('Server')} hint={t('A CalDAV or CardDAV server: iCloud (caldav.icloud.com), Fastmail, Nextcloud, Synology, Radicale and the like.')}>
            <Input className="dav-url" value={url} onChange={(e) => setUrl(e.target.value)} placeholder="https://dav.example.com" autoFocus />
          </Field>
          <Field label={t('User name')}>
            <Input className="dav-user" value={user} onChange={(e) => setUser(e.target.value)} />
          </Field>
          <Field label={t('Password')} hint={t('An app password, where the service gives one.')}>
            <Input className="dav-password" type="password" value={password} onChange={(e) => setPassword(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') add(); }} />
          </Field>
          <Field label={t('Name')} hint={t('What to call it here, such as Work or Family. Left blank, it is named after its server.')}>
            <Input className="dav-name-new" value={name} onChange={(e) => setName(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') add(); }} />
          </Field>
          {error ? <div className="dav-error" style={{ color: 'var(--bad, #c62828)', fontSize: 13 }}>{error}</div> : null}
          <div><Button primary className="dav-add" label={busy === 'add' ? t('Adding…') : t('Add account')} disabled={!url.trim() || Boolean(busy)} onClick={add} /></div>
        </div>
      </div>
    </Dialog>
  );
}
