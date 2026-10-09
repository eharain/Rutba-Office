// Accounts, shared by Calendar and Contacts: the CalDAV and CardDAV servers
// kept in step with this computer. Each account says when it last synced
// or why it could not; one is added with the server's address, a user name
// and a password (an app password, for iCloud and the like), and its
// calendars and address books come in at once. In Contacts, where a new card
// goes is chosen here too.

import React, { useCallback, useEffect, useState } from 'react';
import { Button, Dialog, Field, Input, Select, t, tn } from '@rutba/office-ui';

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
  const [busy, setBusy] = useState(null);
  const [error, setError] = useState('');

  const refresh = useCallback(async () => {
    setAccounts((await shell.dav.accounts()) || []);
    if (kind === 'contacts') setBooks((await shell.dav.books()) || { books: [], defaultBook: null });
  }, [shell, kind]);
  useEffect(() => { refresh(); }, [refresh]);
  useEffect(() => shell.on('dav:changed', refresh), [shell, refresh]);

  const add = async () => {
    if (!url.trim()) return;
    setBusy('add');
    setError('');
    try {
      const acc = await shell.dav.add({ url: url.trim(), user: user.trim(), password });
      setUrl(''); setUser(''); setPassword('');
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
  const remove = async (id) => {
    await shell.dav.remove({ id });
    refresh();
  };

  return (
    <Dialog title={t('Accounts')} width={520} onClose={onClose} actions={<Button label={t('Close')} onClick={onClose} />}>
      <div className="dav-accounts" style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        {accounts.length ? accounts.map((a) => (
          <div key={a.id} className="dav-account" style={{ display: 'flex', alignItems: 'center', gap: 8, padding: 8, border: '1px solid var(--line)', borderRadius: 6 }}>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontWeight: 600 }}>{a.name}</div>
              <div style={{ fontSize: 12, opacity: 0.75, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{a.user ? `${a.user} · ` : ''}{a.url}</div>
              <div className="dav-status" style={{ fontSize: 12, color: a.lastError ? 'var(--bad, #c62828)' : 'inherit' }}>
                {a.lastError ? a.lastError : t('{calendars}, {books} · synced {when}', { calendars: tn(a.calendars, '{count} calendar', '{count} calendars'), books: tn(a.books, '{count} address book', '{count} address books'), when: ago(a.lastSync) })}
              </div>
            </div>
            <Button label={busy === a.id ? t('Syncing…') : t('Sync now')} className="dav-sync" disabled={Boolean(busy)} onClick={() => sync(a.id)} />
            <Button label={t('Remove')} icon="trash" title={t('Remove the account and the calendars and cards it brought; nothing on the server is touched')} onClick={() => remove(a.id)} />
          </div>
        )) : <div style={{ opacity: 0.75 }}>{t('No accounts yet. Your calendars and contacts stay on this computer until you add one.')}</div>}
        {kind === 'contacts' && books.books.length ? (
          <Field label={t('New cards go to')}>
            <Select className="dav-default-book" value={books.defaultBook || ''} onChange={async (e) => { await shell.dav.setDefaultBook({ id: e.target.value || null }); refresh(); }}>
              <option value="">{t('This computer only')}</option>
              {books.books.map((b) => <option key={b.id} value={b.id}>{b.name}</option>)}
            </Select>
          </Field>
        ) : null}
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
          {error ? <div className="dav-error" style={{ color: 'var(--bad, #c62828)', fontSize: 13 }}>{error}</div> : null}
          <div><Button primary className="dav-add" label={busy === 'add' ? t('Adding…') : t('Add account')} disabled={!url.trim() || Boolean(busy)} onClick={add} /></div>
        </div>
      </div>
    </Dialog>
  );
}
