// Invite by email (Office 2.25, D-2026-11-11-01 and -02; owner: "on talk we
// need to make it easier to connect and invite connections via their own
// emails").
//
// The person picks people from Contacts or types their addresses. Each is
// asked in Talk, and gets the invitation from the person's own sending
// account in Mail, as a note they wrote and can change before it goes: the
// link to accept and the way to get Rutba Office go below it. With no Mail
// account to send from, Rutba emails them instead, from no-reply@rutba.io
// and only so many a day and a week, and the dialog says so.

import React, { useEffect, useMemo, useState } from 'react';
import { Button, Input, Select, Dialog, Spinner } from '@rutba/office-ui';
import { t, tn } from '../talk-words.js';
import { errorText, toldText, leftOf } from './invite.js';
import { parseAddresses, NOTE_CAP, INVITE_BATCH, EMAIL } from './addresses.js';

/** The note an invitation starts with: the person changes it as they like. */
export function startingNote(me, own) {
  if (!own) return t('I would like to talk with you on Rutba Talk, where I keep my calls and messages.');
  return [
    t('Hello {name},'),
    t('I would like to talk with you on Rutba Talk, where I keep my calls and messages. It is free, and the link below is all you need.'),
    me?.name || '',
  ].filter(Boolean).join('\n\n');
}

/** What goes below the note in an email from the person's own account: {accept} and {join} are filled with the links. */
export const linksBlock = () => [
  t('Accept: {accept}'),
  t('No Rutba Office yet? The same link lets you accept with a free Rutba account, and Rutba Office is free to download: {join}'),
].join('\n\n');

/** The people in Contacts with an address, each address once, as { name, email }. */
const withAddresses = (cards) => {
  const out = [];
  const seen = new Set();
  for (const c of cards || []) {
    const name = String(c.display || c.name?.full || '').trim();
    for (const e of c.emails || []) {
      const email = String(e?.value || '').trim();
      if (!EMAIL.test(email) || seen.has(email.toLowerCase())) continue;
      seen.add(email.toLowerCase());
      out.push({ name, email });
    }
  }
  return out;
};

/** How one person's invitation went, as a short line and a key the checks read. */
function outcome(r) {
  if (r.contact) return { key: 'contact', text: t('{name} is a contact already.', { name: r.contact.name || r.email }) };
  if (r.error) return { key: `error:${r.error}`, text: errorText(r.error) };
  return { key: `${r.told?.by}:${r.told?.sent ? 'sent' : r.told?.reason || 'not'}`, text: toldText(r.told) };
}

export function InviteByEmailButton({ onClick }) {
  return <Button ghost icon="mail" label={t('Invite by email…')} className="tk-ie-btn" onClick={onClick} />;
}

export function InviteByEmail({ shell, state, toast, onClose, initial = [] }) {
  const [people, setPeople] = useState(() => parseAddresses(initial.map((p) => (p.name ? `${p.name} <${p.email}>` : p.email)).join(', ')));
  const [typed, setTyped] = useState('');
  const [query, setQuery] = useState('');
  const [cards, setCards] = useState(null);
  const [sender, setSender] = useState(null);
  const [accounts, setAccounts] = useState([]);
  const [accountId, setAccountId] = useState(null);
  const [subject, setSubject] = useState(null);
  const [note, setNote] = useState(null);
  const [busy, setBusy] = useState(false);
  const [results, setResults] = useState(null);

  useEffect(() => {
    let gone = false;
    shell.talk.allowance().catch(() => {});
    Promise.all([
      shell.talk.sender().catch(() => ({ account: null })),
      shell.mail.accounts().then((list) => (list || []).filter((a) => a.email && !a.local).map((a) => ({ id: a.id, address: a.email, name: a.name || '' })), () => []),
    ]).then(([s, list]) => {
      if (gone) return;
      setSender(s || { account: null });
      setAccounts(list);
      setAccountId(s?.account?.id ?? null);
    });
    return () => { gone = true; };
  }, [shell]);

  // The people in Contacts, found as the person types.
  useEffect(() => {
    let gone = false;
    const timer = setTimeout(() => {
      shell.contacts.list({ query: query.trim() }).then((list) => { if (!gone) setCards(withAddresses(list).slice(0, 60)); }, () => { if (!gone) setCards([]); });
    }, query ? 150 : 0);
    return () => { gone = true; clearTimeout(timer); };
  }, [shell, query]);

  const from = accountId == null ? null : accounts.find((a) => a.id === accountId) || (sender?.account?.id === accountId ? sender.account : null);
  const own = Boolean(from);
  // The note it starts with follows the way it goes, until the person writes their own.
  const shownNote = note ?? (sender ? startingNote(state.me, own) : '');
  const shownSubject = subject ?? t('{name} would like to talk with you on Rutba Talk', { name: state.me?.name || state.me?.email || '' });
  const allowance = sender?.allowance || state.mailAllowance;
  const left = leftOf(allowance);
  const chosen = useMemo(() => new Set(people.map((p) => p.email.toLowerCase())), [people]);

  const addTyped = () => {
    const more = parseAddresses(typed).filter((p) => !chosen.has(p.email.toLowerCase()));
    if (more.length) setPeople([...people, ...more]);
    setTyped('');
  };
  const toggle = (p) => setPeople(chosen.has(p.email.toLowerCase()) ? people.filter((x) => x.email.toLowerCase() !== p.email.toLowerCase()) : [...people, p]);
  const pickFrom = (id) => { setAccountId(id === 'rutba' ? null : accounts.find((a) => String(a.id) === id)?.id ?? null); };

  const pending = parseAddresses(typed).filter((p) => !chosen.has(p.email.toLowerCase()));
  const all = [...people, ...pending];
  const tooMany = all.length > INVITE_BATCH;

  const send = async () => {
    setBusy(true);
    try {
      if (pending.length) setPeople(all);
      setTyped('');
      const text = shownNote.trim();
      const r = await shell.talk.inviteByEmail({
        people: all,
        via: own ? 'own' : 'rutba',
        accountId: own ? from.id : null,
        note: own ? '' : text.slice(0, NOTE_CAP),
        letter: own ? { subject: shownSubject.trim(), text: `${text}\n\n${linksBlock()}` } : null,
      });
      if (r.error) { toast(r.error === 'no-account' ? t('That Mail account cannot send now. Choose another, or let Rutba send it.') : errorText(r.error), { tone: 'bad', ms: 5000 }); return; }
      const good = r.results.filter((x) => x.contact || x.told?.sent);
      if (good.length === r.results.length) {
        const asked = r.results.filter((x) => x.told).length;
        toast(r.via === 'own'
          ? tn(asked, 'Invited {count} person from {address}.', 'Invited {count} people from {address}.', { address: r.from })
          : tn(asked, 'Invited {count} person. Rutba emailed the invitation.', 'Invited {count} people. Rutba emailed the invitations.'), { tone: 'good', ms: 5000 });
        onClose();
        return;
      }
      setResults(r);
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 5000 });
    } finally {
      setBusy(false);
    }
  };

  if (results) {
    return (
      <Dialog title={t('Invite by email')} width={520} className="tk-ie" onClose={onClose} actions={<Button primary label={t('Done')} className="tk-ie-done" onClick={onClose} />}>
        <p className="tk-lead">{t('Not every invitation went by email. Those asked can still accept in Talk, and you can send them their link another way from the list of people you asked.')}</p>
        <ul className="tk-ie-results">
          {results.results.map((r) => {
            const o = outcome(r);
            return <li key={r.email} className="tk-ie-result" data-email={r.email} data-result={o.key}><strong>{r.email}</strong><small>{o.text}</small></li>;
          })}
        </ul>
      </Dialog>
    );
  }

  return (
    <Dialog
      title={t('Invite by email')}
      width={600}
      className="tk-ie"
      onClose={onClose}
      actions={<><Button label={t('Cancel')} onClick={onClose} /><Button primary className="tk-ie-send" label={all.length ? tn(all.length, 'Send the invitation', 'Send {count} invitations') : t('Send')} disabled={busy || !sender || !all.length || tooMany || !shownNote.trim()} onClick={send} /></>}
    >
      <p className="tk-lead">{t('Each person gets an email from you with your note and a link to accept. Once they accept, you can write to and call each other.')}</p>

      <div className="tk-field">
        <span>{t('To')}</span>
        <div className="tk-ie-to" role="group" aria-label={t('People to invite')}>
          {people.map((p) => (
            <span key={p.email} className="tk-ie-chip" data-email={p.email} title={p.email}>
              {p.name || p.email}
              <button type="button" aria-label={t('Take {email} off', { email: p.email })} onClick={() => toggle(p)}>×</button>
            </span>
          ))}
          <Input
            className="tk-ie-typed"
            value={typed}
            autoFocus
            placeholder={people.length ? '' : t('Type or paste addresses')}
            aria-label={t('Email addresses')}
            onChange={(e) => setTyped(e.target.value)}
            onBlur={addTyped}
            onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ',' || e.key === ';') { e.preventDefault(); addTyped(); } else if (e.key === 'Backspace' && !typed && people.length) setPeople(people.slice(0, -1)); }}
          />
        </div>
        {tooMany ? <small className="tk-error">{t('At most {max} people at once.', { max: INVITE_BATCH })}</small> : null}
      </div>

      <details className="tk-ie-contacts" open={!people.length}>
        <summary>{t('Pick from Contacts')}</summary>
        <Input className="tk-ie-search" value={query} placeholder={t('Search Contacts')} aria-label={t('Search Contacts')} onChange={(e) => setQuery(e.target.value)} />
        <div className="tk-ie-list" role="list">
          {cards === null ? <Spinner /> : cards.length ? cards.map((p) => (
            <label key={p.email} className="tk-ie-pick" role="listitem" data-email={p.email}>
              <input type="checkbox" checked={chosen.has(p.email.toLowerCase())} onChange={() => toggle(p)} />
              <span className="tk-who"><strong>{p.name || p.email}</strong>{p.name ? <small>{p.email}</small> : null}</span>
            </label>
          )) : <p className="tk-none">{query ? t('Nobody in Contacts matches.') : t('Nobody in Contacts has an email address yet.')}</p>}
        </div>
      </details>

      <div className="tk-field">
        <span>{t('From')}</span>
        {!sender ? <Spinner /> : (
          <Select className="tk-ie-from" value={own ? String(from.id) : 'rutba'} aria-label={t('From')} onChange={(e) => pickFrom(e.target.value)}>
            {accounts.map((a) => <option key={a.id} value={String(a.id)}>{a.name ? `${a.name} <${a.address}>` : a.address}</option>)}
            <option value="rutba">{t('Rutba, from no-reply@rutba.io')}</option>
          </Select>
        )}
        {sender && !own ? (
          <small className="tk-ie-rutba" data-left={left ? `${left.day}/${left.week}` : ''}>
            {accounts.length ? '' : t('You have no Mail account to send from, so Rutba emails them, in your name, from no-reply@rutba.io.')}
            {' '}
            {left ? t('Rutba sends at most {dayMax} a day and {weekMax} a week for you: {day} left today, {week} this week.', { dayMax: allowance.day.limit, weekMax: allowance.week.limit, day: left.day, week: left.week }) : t('Rutba sends only a few a day and a week.')}
            {left && all.length > Math.min(left.day, left.week) ? ` ${t('Those beyond that are asked in Talk only, with no email.')}` : ''}
          </small>
        ) : null}
      </div>

      {own ? (
        <label className="tk-field"><span>{t('Subject')}</span><Input className="tk-ie-subject" value={shownSubject} onChange={(e) => setSubject(e.target.value)} /></label>
      ) : null}
      <label className="tk-field">
        <span>{own ? t('Your note') : t('Your note, in Rutba’s email')}</span>
        <textarea className="rw-input tk-ie-note" rows={own ? 7 : 3} maxLength={own ? 4000 : NOTE_CAP} value={shownNote} onChange={(e) => setNote(e.target.value)} />
        <small>{own ? t('{name} becomes each person’s name. The link to accept and the way to get Rutba Office go below your note.') : t('Up to {max} characters. Rutba’s email adds the link to accept.', { max: NOTE_CAP })}</small>
      </label>
    </Dialog>
  );
}
