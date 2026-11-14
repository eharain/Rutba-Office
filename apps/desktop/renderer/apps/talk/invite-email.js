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
import { ConnectButton, reachText, canConnect } from './connect.js';

/** The note an invitation starts with: the person changes it as they like. */
export function startingNote(me, own) {
  if (!own) return t('talk.invite-email.would-like-talk');
  return [
    t('talk.invite-email.hello-name'),
    t('talk.invite-email.would-like-talk-rutba'),
    me?.name || '',
  ].filter(Boolean).join('\n\n');
}

/** What goes below the note in an email from the person's own account: {accept} and {join} are filled with the links. */
export const linksBlock = () => [
  t('talk.invite.accept-accept'),
  t('talk.invite.no-rutba-office'),
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
  if (r.contact) return { key: 'contact', text: t('talk.invite.name-contact-already', { name: r.contact.name || r.email }) };
  if (r.error) return { key: `error:${r.error}`, text: errorText(r.error) };
  return { key: `${r.told?.by}:${r.told?.sent ? 'sent' : r.told?.reason || 'not'}`, text: toldText(r.told) };
}

export function InviteByEmailButton({ onClick }) {
  return <Button ghost icon="mail" label={t('talk.invite-email.invite-by')} className="tk-ie-btn" onClick={onClick} />;
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
  const [reach, setReach] = useState({});

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

  // Where each person stands in Talk, again whenever Talk's own lists change.
  const known = `${state.contacts.length}/${state.requests.incoming.length}/${state.requests.outgoing.length}`;
  useEffect(() => {
    const emails = [...(cards || []), ...people].map((p) => p.email);
    if (!emails.length) return undefined;
    let gone = false;
    shell.talk.reach({ emails }).then((r) => { if (!gone) setReach(r || {}); }, () => {});
    return () => { gone = true; };
  }, [shell, cards, people, known]);
  const reachOf = (p) => reach[p.email.toLowerCase()] || null;
  const connected = (r) => {
    if (r?.error) toast(errorText(r.error), { tone: 'bad', ms: 5000 });
    // Asked in Talk now: not emailed as well.
    else setPeople((list) => list.filter((x) => x.email.toLowerCase() !== String(r?.request?.email || r?.contact?.email || '').toLowerCase()));
  };

  const from = accountId == null ? null : accounts.find((a) => a.id === accountId) || (sender?.account?.id === accountId ? sender.account : null);
  const own = Boolean(from);
  // The note it starts with follows the way it goes, until the person writes their own.
  const shownNote = note ?? (sender ? startingNote(state.me, own) : '');
  const shownSubject = subject ?? t('talk.invite.name-would-like', { name: state.me?.name || state.me?.email || '' });
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
      if (r.error) { toast(r.error === 'no-account' ? t('talk.invite-email.mail-account-cannot') : errorText(r.error), { tone: 'bad', ms: 5000 }); return; }
      const good = r.results.filter((x) => x.contact || x.told?.sent);
      if (good.length === r.results.length) {
        const asked = r.results.filter((x) => x.told).length;
        toast(r.via === 'own'
          ? tn(asked, 'talk.invite-email.invited-count-person', { address: r.from })
          : tn(asked, 'talk.invite-email.invited-count-person-rutba'), { tone: 'good', ms: 5000 });
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
      <Dialog title={t('talk.invite-email.invite-by-email')} width={520} className="tk-ie" onClose={onClose} actions={<Button primary label={t('common.done')} className="tk-ie-done" onClick={onClose} />}>
        <p className="tk-lead">{t('talk.invite-email.not-every-invitation')}</p>
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
      title={t('talk.invite-email.invite-by-email')}
      width={600}
      className="tk-ie"
      onClose={onClose}
      actions={<><Button label={t('common.cancel')} onClick={onClose} /><Button primary className="tk-ie-send" label={all.length ? tn(all.length, 'talk.invite-email.send-the') : t('common.send')} disabled={busy || !sender || !all.length || tooMany || !shownNote.trim()} onClick={send} /></>}
    >
      <p className="tk-lead">{t('talk.invite-email.each-person-gets')}</p>

      <div className="tk-field">
        <span>{t('common.to')}</span>
        <div className="tk-ie-to" role="group" aria-label={t('talk.invite-email.people-to-invite')}>
          {people.map((p) => (
            <span key={p.email} className="tk-ie-chip" data-email={p.email} data-reach={reachOf(p) || ''} title={[p.email, reachText(reachOf(p))].filter(Boolean).join(' · ')}>
              {p.name || p.email}
              {canConnect(reachOf(p)) ? <ConnectButton shell={shell} person={p} toast={toast} onDone={connected} /> : null}
              <button type="button" aria-label={t('talk.invite-email.take-email-off', { email: p.email })} onClick={() => toggle(p)}>×</button>
            </span>
          ))}
          <Input
            className="tk-ie-typed"
            value={typed}
            autoFocus
            placeholder={people.length ? '' : t('talk.invite-email.type-paste-addresses')}
            aria-label={t('talk.invite-email.email-addresses')}
            onChange={(e) => setTyped(e.target.value)}
            onBlur={addTyped}
            onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ',' || e.key === ';') { e.preventDefault(); addTyped(); } else if (e.key === 'Backspace' && !typed && people.length) setPeople(people.slice(0, -1)); }}
          />
        </div>
        {tooMany ? <small className="tk-error">{t('talk.invite-email.most-max-people', { max: INVITE_BATCH })}</small> : null}
      </div>

      <details className="tk-ie-contacts" open={!people.length}>
        <summary>{t('talk.invite-email.pick-from-contacts')}</summary>
        <Input className="tk-ie-search" value={query} placeholder={t('talk.invite-email.search-contacts')} aria-label={t('talk.invite-email.search-contacts')} onChange={(e) => setQuery(e.target.value)} />
        <div className="tk-ie-list" role="list">
          {cards === null ? <Spinner /> : cards.length ? cards.map((p) => (
            <label key={p.email} className="tk-ie-pick" role="listitem" data-email={p.email} data-reach={reachOf(p) || ''}>
              <input type="checkbox" checked={chosen.has(p.email.toLowerCase())} disabled={reachOf(p) === 'contact' || reachOf(p) === 'asked'} onChange={() => toggle(p)} />
              <span className="tk-who"><strong>{p.name || p.email}</strong><small>{[p.name ? p.email : '', reachText(reachOf(p))].filter(Boolean).join(' · ')}</small></span>
              {canConnect(reachOf(p)) ? <ConnectButton shell={shell} person={p} toast={toast} onDone={connected} /> : null}
            </label>
          )) : <p className="tk-none">{query ? t('talk.invite-email.nobody-contacts-matches') : t('talk.invite-email.nobody-contacts-email')}</p>}
        </div>
      </details>

      <div className="tk-field">
        <span>{t('common.from')}</span>
        {!sender ? <Spinner /> : (
          <Select className="tk-ie-from" value={own ? String(from.id) : 'rutba'} aria-label={t('common.from')} onChange={(e) => pickFrom(e.target.value)}>
            {accounts.map((a) => <option key={a.id} value={String(a.id)}>{a.name ? `${a.name} <${a.address}>` : a.address}</option>)}
            <option value="rutba">{t('talk.invite-email.rutba-no-reply')}</option>
          </Select>
        )}
        {sender && !own ? (
          <small className="tk-ie-rutba" data-left={left ? `${left.day}/${left.week}` : ''}>
            {accounts.length ? '' : t('talk.invite-email.no-mail-account')}
            {' '}
            {left ? t('talk.invite-email.rutba-sends-most', { dayMax: allowance.day.limit, weekMax: allowance.week.limit, day: left.day, week: left.week }) : t('talk.invite-email.rutba-sends-only')}
            {left && all.length > Math.min(left.day, left.week) ? ` ${t('talk.invite-email.beyond-asked-talk')}` : ''}
          </small>
        ) : null}
      </div>

      {own ? (
        <label className="tk-field"><span>{t('common.subject')}</span><Input className="tk-ie-subject" value={shownSubject} onChange={(e) => setSubject(e.target.value)} /></label>
      ) : null}
      <label className="tk-field">
        <span>{own ? t('talk.invite-email.your-note') : t('talk.invite-email.note-rutbas-email')}</span>
        <textarea className="rw-input tk-ie-note" rows={own ? 7 : 3} maxLength={own ? 4000 : NOTE_CAP} value={shownNote} onChange={(e) => setNote(e.target.value)} />
        <small>{own ? t('talk.invite-email.name-becomes-each') : t('talk.invite-email.max-characters-rutbas', { max: NOTE_CAP })}</small>
      </label>
    </Dialog>
  );
}
