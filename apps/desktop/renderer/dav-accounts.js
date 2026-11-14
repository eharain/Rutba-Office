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
  if (!ms) return t('dav-accounts.not-yet');
  const s = Math.round((Date.now() - ms) / 1000);
  if (s < 60) return t('dav-accounts.just-now');
  if (s < 3600) return tn(Math.round(s / 60), 'common.count-min-ago');
  if (s < 86400) return tn(Math.round(s / 3600), 'common.count-h-ago');
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
  const bookLabel = (b) => (accountOf(b) ? t('dav-accounts.account-book', { account: accountOf(b), book: b.name }) : b.name);
  /** Where a new card goes, said plainly: kept here, or sent to which account's server. */
  const where = ({ books: list, defaultBook }) => {
    const b = list.find((x) => x.id === defaultBook);
    return b
      ? t('dav-accounts.new-cards-sent', { book: b.name, account: accountOf(b) || b.name })
      : t('dav-accounts.new-cards-stay');
  };

  const add = async () => {
    if (!url.trim()) return;
    setBusy('add');
    setError('');
    try {
      const acc = await shell.dav.add({ url: url.trim(), user: user.trim(), password, name: name.trim() });
      setUrl(''); setUser(''); setPassword(''); setName('');
      toast?.(t('dav-accounts.name-calendars-books', { name: acc.name, calendars: tn(acc.calendars, 'dav-accounts.count-calendar'), books: tn(acc.books, 'dav-accounts.count-address-book') }), { tone: 'good' });
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
      else if (r) toast?.(r.conflicts ? t('dav-accounts.synced-received-sent', { received: r.received, sent: r.sent, conflicts: r.conflicts }) : t('common.synced-received-sent', { received: r.received, sent: r.sent }), { tone: 'good' });
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
    <Dialog title={t('common.accounts')} width={520} onClose={onClose} actions={<Button label={t('common.close')} onClick={onClose} />}>
      <div className="dav-accounts" style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        {accounts.length ? accounts.map((a) => (
          <div key={a.id} className="dav-account" style={{ display: 'flex', alignItems: 'center', gap: 8, padding: 8, border: '1px solid var(--line)', borderRadius: 6 }}>
            <div style={{ flex: 1, minWidth: 0 }}>
              {renaming?.id === a.id ? (
                <Input className="dav-rename" aria-label={t('dav-accounts.account-name')} value={renaming.name} autoFocus onChange={(e) => setRenaming({ id: a.id, name: e.target.value })}
                  onKeyDown={(e) => { if (e.key === 'Enter') rename(); if (e.key === 'Escape') { e.stopPropagation(); setRenaming(null); } }} />
              ) : <div className="dav-name" style={{ fontWeight: 600 }}>{a.name}</div>}
              <div style={{ fontSize: 12, opacity: 0.75, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{a.user ? `${a.user} · ` : ''}{a.url}</div>
              <div className="dav-status" style={{ fontSize: 12, color: a.lastError ? 'var(--bad, #c62828)' : 'inherit' }}>
                {a.lastError ? a.lastError : t('dav-accounts.calendars-books-synced', { calendars: tn(a.calendars, 'dav-accounts.count-calendar'), books: tn(a.books, 'dav-accounts.count-address-book'), when: ago(a.lastSync) })}
              </div>
              {a.needsPassword ? (
                <div className="dav-password-again" style={{ display: 'flex', gap: 6, alignItems: 'center', marginTop: 6 }}>
                  <Input className="dav-password-again-input" type="password" placeholder={t('common.password')} aria-label={t('dav-accounts.password-name-not', { name: a.name })} value={again[a.id] || ''} onChange={(e) => setAgain({ ...again, [a.id]: e.target.value })} onKeyDown={(e) => { if (e.key === 'Enter') savePassword(a.id); }} />
                  <Button label={t('common.save')} className="dav-password-again-save" disabled={Boolean(busy)} onClick={() => savePassword(a.id)} />
                </div>
              ) : null}
            </div>
            {renaming?.id === a.id ? (
              <>
                <Button primary label={t('common.save')} className="dav-rename-save" disabled={!renaming.name.trim()} onClick={rename} />
                <Button label={t('common.cancel')} onClick={() => setRenaming(null)} />
              </>
            ) : <Button label={t('common.rename')} className="dav-rename-start" title={t('dav-accounts.call-account-something')} onClick={() => setRenaming({ id: a.id, name: a.name })} />}
            <Button label={busy === a.id ? t('dav-accounts.syncing') : t('common.sync-now')} className="dav-sync" disabled={Boolean(busy)} onClick={() => sync(a.id)} />
            <Button label={t('common.remove')} icon="trash" title={t('dav-accounts.remove-account-calendars')} onClick={() => remove(a.id)} />
          </div>
        )) : <div style={{ opacity: 0.75 }}>{t('dav-accounts.no-accounts-yet')}</div>}
        {kind === 'contacts' && books.books.length ? (
          <Field label={t('dav-accounts.new-cards-go')} hint={<span className="dav-default-where">{where(books)}</span>}>
            <Select className="dav-default-book" value={books.defaultBook || ''} onChange={async (e) => { await shell.dav.setDefaultBook({ id: e.target.value || null }); refresh(); }}>
              <option value="">{t('common.this-computer-only')}</option>
              {books.books.map((b) => <option key={b.id} value={b.id}>{bookLabel(b)}</option>)}
            </Select>
          </Field>
        ) : null}
        <AiThingsList shell={shell} kind={kind === 'contacts' ? 'book' : 'calendar'} title={kind === 'contacts' ? t('dav-accounts.ai-tools-address') : t('dav-accounts.ai-tools-calendars')} />
        <div style={{ borderTop: '1px solid var(--line)', paddingTop: 10, display: 'flex', flexDirection: 'column', gap: 8 }}>
          <div style={{ fontWeight: 600 }}>{t('dav-accounts.add-an-account')}</div>
          <Field label={t('common.server')} hint={t('dav-accounts.cal-dav-card')}>
            <Input className="dav-url" value={url} onChange={(e) => setUrl(e.target.value)} placeholder="https://dav.example.com" autoFocus />
          </Field>
          <Field label={t('common.user-name')}>
            <Input className="dav-user" value={user} onChange={(e) => setUser(e.target.value)} />
          </Field>
          <Field label={t('common.password')} hint={t('dav-accounts.app-password-service')}>
            <Input className="dav-password" type="password" value={password} onChange={(e) => setPassword(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') add(); }} />
          </Field>
          <Field label={t('common.name')} hint={t('dav-accounts.call-such-work')}>
            <Input className="dav-name-new" value={name} onChange={(e) => setName(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') add(); }} />
          </Field>
          {error ? <div className="dav-error" style={{ color: 'var(--bad, #c62828)', fontSize: 13 }}>{error}</div> : null}
          <div><Button primary className="dav-add" label={busy === 'add' ? t('common.adding') : t('common.add-account')} disabled={!url.trim() || Boolean(busy)} onClick={add} /></div>
        </div>
      </div>
    </Dialog>
  );
}
