// Inviting people to Talk (owner, 2026-10-21, D-2026-10-21-27 and -28).
//
// Adding someone by address also emails them: from the person's own address
// in Mail when they have one, else from Rutba, which sends only so many a day
// and a week, or not at all. Each request says how its person was told. An
// invitation link goes to friends over any channel of the person's own: a
// chat app, their email, a code a phone reads. A friend without Rutba Office
// accepts on the link's page with a Rutba account and gets Office there.

import React, { useEffect, useMemo, useState } from 'react';
import { Button, Input, Dialog, Spinner, formatWhen, t } from '@rutba/office-ui';
import { qrPath } from '../../qrcode.js';
import { askedInTalkText } from './connect.js';

const EMAIL = /^\S+@\S+\.\S+$/;
const enc = encodeURIComponent;

/** A refusal from the introducer, in words. */
export function errorText(code) {
  switch (code) {
    case 'address': return t('talk.invite.not-email-address');
    case 'yourself': return t('talk.invite.own-address');
    case 'too-many': return t('talk.invite.many-talk-takes');
    case 'no-invite': return t('talk.invite.invitation-ended');
    case 'blocked': return t('talk.invite.invitation-ended');
    case 'offline':
    case 'no-answer': return t('talk.invite.talk-not-connected');
    case 'unknown': return t('talk.invite.talk-server-does');
    default: return t('talk.invite.talk-could-not');
  }
}

/** How a request's person was told, as the asker sees it. */
export function toldText(told) {
  if (!told) return '';
  if (told.by === 'talk') return askedInTalkText();
  if (told.by === 'own') return told.sent ? t('talk.invite.emailed-from-address', { address: told.from }) : t('talk.invite.not-emailed-address', { address: told.from || t('talk.invite.your-account') });
  if (told.by === 'rutba') {
    if (told.sent) return t('talk.invite.emailed-by-rutba');
    if (told.reason === 'limit') return t('talk.invite.not-emailed-rutbas');
    if (told.reason === 'recently') return t('talk.invite.not-emailed-rutba-emailed');
    if (told.reason === 'stopped') return t('talk.invite.not-emailed-address-asked');
    return t('talk.invite.not-emailed-rutba');
  }
  return t('talk.invite.not-emailed');
}

/** What Rutba will still send: { day, week } left, or null when not known. */
export const leftOf = (allowance) => (allowance?.day && allowance?.week
  ? { day: Math.max(0, allowance.day.limit - allowance.day.used), week: Math.max(0, allowance.week.limit - allowance.week.used) }
  : null);

/** The email an invitation from the person's own address carries; {accept} and {join} are filled with the links. */
const letterFor = (me, note) => ({
  subject: t('talk.invite.name-would-like', { name: me?.name || me?.email || '' }),
  text: [
    t('talk.invite.hello'),
    t('talk.invite.name-address-would', { name: me?.name || me?.email || '', address: me?.email || '' }),
    note ? `"${note}"` : '',
    t('talk.invite.accept-accept'),
    t('talk.invite.no-rutba-office'),
  ].filter(Boolean).join('\n\n'),
});

/** The letter for an invitation sent again from the person's own address. */
export const againLetter = (me) => ({
  subject: t('talk.invite.name-would-still', { name: me?.name || me?.email || '' }),
  text: [
    t('talk.invite.hello'),
    t('talk.invite.name-address-asked', { name: me?.name || me?.email || '', address: me?.email || '' }),
    t('talk.invite.accept-accept'),
    t('talk.invite.no-rutba-office'),
  ].join('\n\n'),
});

/* ── Add a contact ────────────────────────────────────────────────── */

export function AddContact({ shell, state, toast, onClose }) {
  const [email, setEmail] = useState('');
  const [note, setNote] = useState('');
  const [accounts, setAccounts] = useState(null);
  const [how, setHow] = useState(null);
  const [suggestions, setSuggestions] = useState([]);
  const [busy, setBusy] = useState(false);
  const left = leftOf(state.mailAllowance);
  const rutbaOpen = !left || (left.day > 0 && left.week > 0);

  useEffect(() => {
    shell.talk.allowance().catch(() => {});
    shell.mail.accounts().then((list) => setAccounts((list || []).filter((a) => a.email && !a.local)), () => setAccounts([]));
  }, [shell]);
  // Their own address first, when they have one; Rutba's when they have none; else no email.
  useEffect(() => {
    if (accounts && how === null) setHow(accounts.length ? `own:${accounts[0].id}` : rutbaOpen ? 'rutba' : 'none');
  }, [accounts, how, rutbaOpen]);
  // The people in Contacts and in Mail, as the address is typed.
  useEffect(() => {
    const q = email.trim();
    if (q.length < 2 || EMAIL.test(q)) { setSuggestions([]); return undefined; }
    const timer = setTimeout(() => shell.contacts.suggest({ query: q, limit: 6 }).then((list) => setSuggestions((list || []).filter((s) => s.email)), () => {}), 150);
    return () => clearTimeout(timer);
  }, [email, shell]);

  const ask = async () => {
    setBusy(true);
    try {
      const own = how?.startsWith('own:') ? how.slice(4) : null;
      const r = await shell.talk.invite({ email: email.trim(), note: note.trim(), how: own ? 'own' : how, accountId: own, letter: own ? letterFor(state.me, note.trim()) : null });
      if (r.error) { toast(errorText(r.error), { tone: 'bad', ms: 5000 }); return; }
      if (r.contact) toast(t('talk.invite.name-contact-already', { name: r.contact.name || r.contact.email }), { tone: 'good', ms: 4000 });
      else {
        const sent = r.told?.sent || r.told?.by === 'none';
        toast([t('talk.invite.asked-email-talk', { email: email.trim() }), toldText(r.told)].join(' '), { tone: sent ? 'good' : 'warn', ms: 6000 });
      }
      onClose();
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 5000 });
    } finally {
      setBusy(false);
    }
  };

  return (
    <Dialog title={t('talk.invite.add-a-contact')} width={480} onClose={onClose} actions={<><Button label={t('common.cancel')} onClick={onClose} /><Button primary label={t('talk.invite.ask')} className="tk-ask" disabled={busy || !how || !EMAIL.test(email.trim())} onClick={ask} /></>}>
      <p className="tk-lead">{t('talk.invite.see-name-note')}</p>
      <label className="tk-field"><span>{t('common.email-address')}</span><Input className="tk-email" value={email} autoFocus list="tk-suggest" onChange={(e) => setEmail(e.target.value)} /></label>
      <datalist id="tk-suggest">{suggestions.map((s) => <option key={s.email} value={s.email}>{s.name || s.email}</option>)}</datalist>
      <label className="tk-field"><span>{t('talk.invite.a-note-optional')}</span><Input className="tk-note" value={note} onChange={(e) => setNote(e.target.value)} /></label>
      <fieldset className="tk-how" aria-label={t('talk.invite.how-hear')}>
        <legend>{t('talk.invite.how-hear')}</legend>
        {accounts === null ? <Spinner /> : null}
        {(accounts || []).map((a) => (
          <label key={a.id} className="tk-relay" data-how={`own:${a.email}`}>
            <input type="radio" name="tk-how" checked={how === `own:${a.id}`} onChange={() => setHow(`own:${a.id}`)} />
            <span><strong>{t('talk.invite.email-address', { address: a.email })}</strong><small>{t('talk.invite.sent-own-account')}</small></span>
          </label>
        ))}
        <label className="tk-relay" data-how="rutba">
          <input type="radio" name="tk-how" checked={how === 'rutba'} disabled={!rutbaOpen} onChange={() => setHow('rutba')} />
          <span>
            <strong>{t('talk.invite.email-rutba')}</strong>
            <small>{!rutbaOpen ? t('talk.invite.rutbas-limit-today')
              : left ? t('talk.invite.no-reply-rutba-io', { day: left.day, week: left.week })
                : t('talk.invite.no-reply-rutba')}</small>
          </span>
        </label>
        <label className="tk-relay" data-how="none">
          <input type="radio" name="tk-how" checked={how === 'none'} onChange={() => setHow('none')} />
          <span><strong>{t('talk.invite.no-email')}</strong><small>{t('talk.invite.see-talk-use')}</small></span>
        </label>
      </fieldset>
    </Dialog>
  );
}

/* ── invitation links ─────────────────────────────────────────────── */

/** The places a link can be sent from, each opened in the browser or the person's own email app. */
const channels = ({ url, text, subject }) => [
  { id: 'whatsapp', name: 'WhatsApp', href: `https://wa.me/?text=${enc(text)}` },
  { id: 'telegram', name: 'Telegram', href: `https://t.me/share/url?url=${enc(url)}&text=${enc(subject)}` },
  { id: 'facebook', name: 'Facebook', href: `https://www.facebook.com/sharer/sharer.php?u=${enc(url)}` },
  { id: 'x', name: 'X', href: `https://x.com/intent/post?text=${enc(text)}` },
  { id: 'linkedin', name: 'LinkedIn', href: `https://www.linkedin.com/sharing/share-offsite/?url=${enc(url)}` },
  { id: 'email-app', name: t('talk.invite.my-email-app'), href: `mailto:?subject=${enc(subject)}&body=${enc(text)}` },
];

function QrCode({ text }) {
  const q = useMemo(() => qrPath(text), [text]);
  if (!q) return null;
  return (
    <svg className="tk-qr" viewBox={`0 0 ${q.side} ${q.side}`} role="img" aria-label={t('talk.invite.code-phone-camera')} shapeRendering="crispEdges">
      <rect width={q.side} height={q.side} fill="#fff" />
      <path d={q.d} fill="#000" />
    </svg>
  );
}

function LinkCard({ shell, invite, me, accounts, toast, onRevoke }) {
  const subject = t('talk.invite.talk-rutba-talk');
  const text = t('talk.invite.name-invites-talk-rutba', { name: me?.name || me?.email || '', url: invite.url });
  const [qr, setQr] = useState(false);
  const copy = async (value, said) => {
    await shell.clipboard.writeText({ text: value }).catch(() => {});
    toast(said, { tone: 'good', ms: 2500 });
  };
  const open = (href) => shell.shell.openExternal({ url: href }).catch((err) => toast(err.message, { tone: 'bad' }));
  // Written in Mail, from one of the person's own addresses, to whoever they choose.
  const inMail = async () => {
    const key = `mail.handoff.${Date.now()}.${Math.random().toString(36).slice(2, 8)}`;
    try {
      await shell.store.set({ key, value: { to: '', subject, text, accountId: accounts[0]?.id } });
      await shell.win.create({ app: 'mail', query: { compose: key } });
    } catch (err) {
      toast(err.message, { tone: 'bad' });
    }
  };
  return (
    <div className="tk-link" data-invite={invite.code}>
      <div className="tk-link-row">
        <Input className="tk-link-url" value={invite.url} readOnly aria-label={t('talk.invite.invitation-link')} onFocus={(e) => e.target.select()} />
        <Button label={t('talk.invite.copy-link')} className="tk-link-copy" onClick={() => copy(invite.url, t('talk.link-copied-paste'))} />
      </div>
      <small className="tk-link-info">
        {invite.note ? `"${invite.note}" · ` : ''}
        {t('talk.invite.used-uses-max', { uses: invite.uses || 0, max: invite.maxUses || 20, when: formatWhen(invite.expires) })}
      </small>
      <div className="tk-share" role="group" aria-label={t('talk.invite.send-it-with')}>
        <Button ghost label={t('talk.invite.copy-message')} className="tk-share-copy" onClick={() => copy(text, t('talk.invite.message-copied-link'))} />
        {accounts.length ? <Button ghost label={t('common.mail')} className="tk-share-mail" onClick={inMail} /> : null}
        {channels({ url: invite.url, text, subject }).map((c) => <Button key={c.id} ghost label={c.name} className={`tk-share-${c.id}`} onClick={() => open(c.href)} />)}
        <Button ghost label={qr ? t('talk.invite.hide-the') : t('talk.invite.phone-code')} className="tk-share-qr" aria-pressed={qr} onClick={() => setQr(!qr)} />
        <Button ghost label={t('talk.invite.withdraw')} className="tk-link-revoke" onClick={() => onRevoke(invite)} />
      </div>
      {qr ? <QrCode text={invite.url} /> : null}
    </div>
  );
}

export function InviteLinks({ shell, state, toast, onClose }) {
  const [note, setNote] = useState('');
  const [busy, setBusy] = useState(false);
  const [accounts, setAccounts] = useState([]);
  useEffect(() => { shell.mail.accounts().then((list) => setAccounts((list || []).filter((a) => a.email && !a.local)), () => {}); }, [shell]);
  const make = async () => {
    setBusy(true);
    const r = await shell.talk.inviteLink({ note: note.trim() }).catch((err) => ({ error: err?.message }));
    setBusy(false);
    if (r.error) { toast(errorText(r.error), { tone: 'bad', ms: 5000 }); return; }
    setNote('');
  };
  const revoke = async (invite) => {
    const r = await shell.talk.revokeInvite({ code: invite.code }).catch(() => ({ error: 'server' }));
    if (r.error) toast(errorText(r.error), { tone: 'bad', ms: 5000 });
    else toast(t('talk.invite.link-withdrawn-nobody'), { ms: 3000 });
  };
  const invites = state.invites || [];
  return (
    <Dialog title={t('talk.invite.invite-link')} width={560} onClose={onClose} actions={<Button primary label={t('common.done')} onClick={onClose} />}>
      <p className="tk-lead">{t('talk.invite.send-link-friends')}</p>
      <div className="tk-link-make">
        <Input className="tk-link-note" value={note} placeholder={t('talk.invite.a-note-optional')} aria-label={t('talk.invite.a-note-optional')} onChange={(e) => setNote(e.target.value)} />
        <Button primary label={t('talk.invite.make-a-link')} className="tk-link-new" disabled={busy} onClick={make} />
      </div>
      {invites.length ? invites.map((i) => <LinkCard key={i.code} shell={shell} invite={i} me={state.me} accounts={accounts} toast={toast} onRevoke={revoke} />)
        : <p className="tk-none">{t('talk.invite.no-links-yet')}</p>}
    </Dialog>
  );
}

/* ── an invitation someone sent ───────────────────────────────────── */

/** Paste a link someone sent, when it did not open Rutba Office itself. */
export function OpenInvitation({ onOpen, onClose }) {
  const [link, setLink] = useState('');
  return (
    <Dialog title={t('talk.invite.open-an-invitation')} width={460} onClose={onClose} actions={<><Button label={t('common.cancel')} onClick={onClose} /><Button primary label={t('common.open')} className="tk-open-invite" disabled={!link.trim()} onClick={() => { onOpen(link.trim()); onClose(); }} /></>}>
      <p className="tk-lead">{t('talk.invite.paste-invitation-link')}</p>
      <Input className="tk-invite-link" value={link} autoFocus aria-label={t('talk.invite.invitation-link')} onChange={(e) => setLink(e.target.value)} />
    </Dialog>
  );
}

/** Who the invitation is from, and Accept; looked up once Talk is connected. */
export function Invitation({ shell, link, online, toast, onDone, onAccepted }) {
  const [seen, setSeen] = useState(null);
  const [busy, setBusy] = useState(false);
  useEffect(() => {
    if (!online) return;
    let gone = false;
    shell.talk.peekInvite({ link }).then((r) => {
      if (gone) return;
      if (r.error) { toast(errorText(r.error), { tone: 'bad', ms: 5000 }); onDone(); return; }
      setSeen(r);
    }, () => {});
    return () => { gone = true; };
  }, [shell, link, online, toast, onDone]);
  const accept = async () => {
    setBusy(true);
    const r = await shell.talk.acceptInvite({ link }).catch(() => ({ error: 'server' }));
    setBusy(false);
    if (r.error) { toast(errorText(r.error), { tone: 'bad', ms: 5000 }); onDone(); return; }
    toast(t('talk.invite.name-contacts-now', { name: r.contact?.name || seen?.from?.name || '' }), { tone: 'good', ms: 4000 });
    onAccepted?.(r.contact);
    onDone();
  };
  if (!online) return <div className="tk-invitation" role="status">{t('talk.invite.sign-talk-accept')}</div>;
  if (!seen) return <div className="tk-invitation" role="status"><Spinner /></div>;
  return (
    <div className="tk-invitation" role="alertdialog" aria-label={t('talk.invite.invitation')} data-from={seen.from?.email}>
      <span className="tk-who">
        <strong>{t('talk.invite.name-invites-talk', { name: seen.from?.name || seen.from?.email })}</strong>
        <small>{seen.note ? `"${seen.note}"` : seen.from?.email}</small>
      </span>
      <Button primary label={t('common.accept')} className="tk-invite-accept" disabled={busy} onClick={accept} />
      <Button label={t('common.not-now')} className="tk-invite-later" onClick={onDone} />
    </div>
  );
}
