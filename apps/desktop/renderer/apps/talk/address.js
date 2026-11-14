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
    case 'address': return t('talk.invite.not-email-address');
    case 'taken': return t('talk.address.address-proved-someone');
    case 'too-many': return t('talk.address.many-addresses-talk');
    case 'limit': return t('talk.address.too-many-codes');
    case 'bad-code': return left ? tn(left, 'talk.address.code-not-right-count') : t('talk.address.code-not-right');
    case 'expired': return t('talk.address.code-run-out');
    case 'sign-in': return t('talk.address.sign-address-stays');
    case 'unsupported': return t('talk.address.talks-server-keeps');
    default: return t('talk.address.talk-could-not');
  }
}

export function MyAddressButton({ onClick }) {
  return <Button ghost icon="contacts" label={t('talk.address.my-address')} className="tk-addr-btn" onClick={onClick} />;
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
    toast(t('talk.address.code-way-address', { address: adding.trim() }), { ms: 4000 });
  });
  const confirm = () => run(() => shell.talk.confirmAddress({ address: waiting, code: code.trim() }), () => {
    toast(t('talk.address.address-proved-find', { address: waiting }), { tone: 'good', ms: 4000 });
    setWaiting(null);
    setAdding('');
    setCode('');
  });

  const supported = Boolean(info?.supported);
  return (
    <Dialog title={t('talk.address.address-talk')} width={540} className="tk-addr" onClose={onClose} actions={<Button primary label={t('common.done')} onClick={onClose} />}>
      <p className="tk-lead">{t('talk.address.choose-own-addresses')}</p>
      {!info ? <Spinner /> : (
        <div className="tk-addr-list" role="list" data-supported={supported ? 'yes' : 'no'}>
          {info.list.map((a) => (
            <div key={a.address} className="tk-addr-row" role="listitem" data-address={a.address} data-shown={a.shown ? 'yes' : 'no'} data-findable={a.findable ? 'yes' : 'no'}>
              <label className="tk-addr-shown">
                <input type="radio" name="tk-addr-shown" checked={Boolean(a.shown)} disabled={!supported || busy} onChange={() => run(() => shell.talk.setAddress({ address: a.address, shown: true }))} />
                <span className="tk-who"><strong>{a.address}</strong><small>{[a.signIn ? t('talk.address.sign-address') : t('talk.address.proved-code'), a.shown ? t('talk.address.shown-people-talk') : ''].filter(Boolean).join(' · ')}</small></span>
              </label>
              <label className="tk-addr-find">
                <input type="checkbox" checked={Boolean(a.findable)} disabled={!supported || busy} onChange={(e) => run(() => shell.talk.setAddress({ address: a.address, findable: e.target.checked }))} />
                <span>{t('talk.address.finds-me')}</span>
              </label>
              {supported && !a.signIn ? <Button ghost icon="close" title={t('talk.address.take-address-off', { address: a.address })} aria-label={t('talk.address.take-address-off', { address: a.address })} className="tk-addr-remove" disabled={busy} onClick={() => run(() => shell.talk.removeAddress({ address: a.address }))} /> : null}
            </div>
          ))}
        </div>
      )}
      {info && !supported ? <p className="tk-none tk-addr-later">{t('talk.address.talks-server-keeps-only')}</p> : null}
      {info && supported ? (
        waiting ? (
          <div className="tk-addr-add">
            <span className="tk-lead">{t('talk.address.type-code-sent', { address: waiting })}</span>
            <Input className="tk-addr-code" value={code} inputMode="numeric" autoFocus aria-label={t('talk.address.code')} onChange={(e) => setCode(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && code.trim()) confirm(); }} />
            <Button primary label={t('talk.address.confirm')} className="tk-addr-confirm" disabled={busy || !code.trim()} onClick={confirm} />
            <Button ghost label={t('common.cancel')} onClick={() => setWaiting(null)} />
          </div>
        ) : (
          <div className="tk-addr-add">
            <Input className="tk-addr-new" value={adding} placeholder={t('talk.address.another-address')} aria-label={t('talk.address.another-address')} onChange={(e) => setAdding(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && EMAIL.test(adding.trim())) sendCode(); }} />
            <Button label={t('talk.address.send-a-code')} className="tk-addr-send" disabled={busy || !EMAIL.test(adding.trim())} onClick={sendCode} />
          </div>
        )
      ) : null}
      {sender ? (
        <p className="tk-none tk-addr-sender">{sender.account
          ? t('talk.address.invitations-send-email', { address: sender.account.address })
          : t('talk.address.no-mail-account')}</p>
      ) : null}
    </Dialog>
  );
}
