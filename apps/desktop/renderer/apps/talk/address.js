// My address in Talk (Office 2.25, D-2026-11-11-02: the third sense of a
// default email). Of the addresses the person has proved, one is shown to
// the people they talk with, and each can find them or not when someone asks
// by address. Proved means the sign-in address, or one confirmed with a code
// emailed to it.
//
// The introducer keeps only the sign-in address today. A Talk server that
// keeps more says so in its welcome (me.addresses); until then this shows
// the sign-in address as the one shown and the one that finds you, and says
// that more addresses come with a later Talk server.

import React, { useEffect, useState } from 'react';
import { Button, Input, Dialog, Spinner } from '@rutba/office-ui';
import { t, tn } from '../talk-words.js';
import { EMAIL } from './addresses.js';

/** A refusal about an address, in words. */
export function addressErrorText(code, left = null) {
  switch (code) {
    case 'address': return t('That is not an email address.');
    case 'taken': return t('That address is proved by someone else in Talk.');
    case 'too-many': return t('That is as many addresses as Talk keeps for one person.');
    case 'limit': return t('Too many codes were sent to that address today. Try again tomorrow.');
    case 'bad-code': return left ? tn(left, 'That code is not right. {count} try left.', 'That code is not right. {count} tries left.') : t('That code is not right.');
    case 'expired': return t('That code has run out. Send a new one.');
    case 'sign-in': return t('Your sign-in address stays: it is how Rutba knows you.');
    case 'unsupported': return t('Talk’s server keeps only your sign-in address for now.');
    default: return t('Talk could not do that now. Try again in a moment.');
  }
}

export function MyAddressButton({ onClick }) {
  return <Button ghost icon="contacts" label={t('My address…')} className="tk-addr-btn" onClick={onClick} />;
}

export function MyAddress({ shell, state, toast, onClose }) {
  const [info, setInfo] = useState(null);
  const [sender, setSender] = useState(null);
  const [adding, setAdding] = useState('');
  const [waiting, setWaiting] = useState(null);
  const [code, setCode] = useState('');
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    shell.talk.addresses().then(setInfo, () => setInfo({ supported: false, list: [] }));
    shell.talk.sender().then(setSender, () => setSender({ account: null }));
  }, [shell, state.me]);

  const run = async (what, then) => {
    setBusy(true);
    try {
      const r = await what();
      if (r?.error) { toast(addressErrorText(r.error, r.left), { tone: 'bad', ms: 5000 }); return; }
      then?.(r);
      setInfo(await shell.talk.addresses());
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 5000 });
    } finally {
      setBusy(false);
    }
  };
  const sendCode = () => run(() => shell.talk.addAddress({ address: adding.trim() }), () => {
    setWaiting(adding.trim());
    setCode('');
    toast(t('A code is on its way to {address}.', { address: adding.trim() }), { ms: 4000 });
  });
  const confirm = () => run(() => shell.talk.confirmAddress({ address: waiting, code: code.trim() }), () => {
    toast(t('{address} is proved. It can find you now.', { address: waiting }), { tone: 'good', ms: 4000 });
    setWaiting(null);
    setAdding('');
    setCode('');
  });

  const supported = Boolean(info?.supported);
  return (
    <Dialog title={t('My address in Talk')} width={540} className="tk-addr" onClose={onClose} actions={<Button primary label={t('Done')} onClick={onClose} />}>
      <p className="tk-lead">{t('Choose which of your own addresses the people you talk with see, and which ones find you when someone asks for you by address. Each is one you have proved: your sign-in address, or one you confirmed with a code sent to it.')}</p>
      {!info ? <Spinner /> : (
        <div className="tk-addr-list" role="list" data-supported={supported ? 'yes' : 'no'}>
          {info.list.map((a) => (
            <div key={a.address} className="tk-addr-row" role="listitem" data-address={a.address} data-shown={a.shown ? 'yes' : 'no'} data-findable={a.findable ? 'yes' : 'no'}>
              <label className="tk-addr-shown">
                <input type="radio" name="tk-addr-shown" checked={Boolean(a.shown)} disabled={!supported || busy} onChange={() => run(() => shell.talk.setAddress({ address: a.address, shown: true }))} />
                <span className="tk-who"><strong>{a.address}</strong><small>{[a.signIn ? t('Your sign-in address') : t('Proved with a code'), a.shown ? t('Shown to the people you talk with') : ''].filter(Boolean).join(' · ')}</small></span>
              </label>
              <label className="tk-addr-find">
                <input type="checkbox" checked={Boolean(a.findable)} disabled={!supported || busy} onChange={(e) => run(() => shell.talk.setAddress({ address: a.address, findable: e.target.checked }))} />
                <span>{t('Finds me')}</span>
              </label>
              {supported && !a.signIn ? <Button ghost icon="close" title={t('Take {address} off', { address: a.address })} aria-label={t('Take {address} off', { address: a.address })} className="tk-addr-remove" disabled={busy} onClick={() => run(() => shell.talk.removeAddress({ address: a.address }))} /> : null}
            </div>
          ))}
        </div>
      )}
      {info && !supported ? <p className="tk-none tk-addr-later">{t('Talk’s server keeps only your sign-in address for now, so it is the one people see and the one that finds you. Adding more addresses comes with a later Talk server.')}</p> : null}
      {info && supported ? (
        waiting ? (
          <div className="tk-addr-add">
            <span className="tk-lead">{t('Type the code sent to {address}.', { address: waiting })}</span>
            <Input className="tk-addr-code" value={code} inputMode="numeric" autoFocus aria-label={t('Code')} onChange={(e) => setCode(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && code.trim()) confirm(); }} />
            <Button primary label={t('Confirm')} className="tk-addr-confirm" disabled={busy || !code.trim()} onClick={confirm} />
            <Button ghost label={t('Cancel')} onClick={() => setWaiting(null)} />
          </div>
        ) : (
          <div className="tk-addr-add">
            <Input className="tk-addr-new" value={adding} placeholder={t('Another address of yours')} aria-label={t('Another address of yours')} onChange={(e) => setAdding(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && EMAIL.test(adding.trim())) sendCode(); }} />
            <Button label={t('Send a code')} className="tk-addr-send" disabled={busy || !EMAIL.test(adding.trim())} onClick={sendCode} />
          </div>
        )
      ) : null}
      {sender ? (
        <p className="tk-none tk-addr-sender">{sender.account
          ? t('Invitations you send by email go from your sending account in Mail, {address}.', { address: sender.account.address })
          : t('With no Mail account to send from, Rutba sends your invitations by email from no-reply@rutba.io.')}</p>
      ) : null}
    </Dialog>
  );
}
