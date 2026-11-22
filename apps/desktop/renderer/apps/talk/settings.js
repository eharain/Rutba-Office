// Talk's settings (Office 2.27.0): whether Talk connects when Office starts,
// whether a call makes a sound, and the people blocked, each with a button
// to let them back. One short dialog; nothing here is a place to type a
// server address.

import React, { useState } from 'react';
import { Button, Dialog } from '@rutba/office-ui';
import { t } from '../talk-words.js';

/** A person blocked, as the list names them: someone blocked from another computer before names were kept has none. */
export const blockedName = (p) => p?.name || p?.email || t('talk.settings.someone-you-blocked');

export function TalkSettings({ shell, state, toast, onClose }) {
  const [busy, setBusy] = useState(null);
  const online = state.phase === 'online';
  const blocked = state.blockedPeople || (state.blocked || []).map((id) => ({ id, name: '', email: '' }));
  const set = (options) => shell.talk.setOptions(options).catch((err) => toast(err?.message || String(err), { tone: 'bad', ms: 5000 }));
  const unblock = async (p) => {
    setBusy(p.id);
    try {
      const r = await shell.talk.unblock({ contactId: p.id });
      if (r?.error) toast(t('talk.settings.name-could-not', { name: blockedName(p) }), { tone: 'bad', ms: 5000 });
      else toast(t('talk.settings.name-no-longer', { name: blockedName(p) }), { tone: 'good', ms: 4000 });
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 5000 });
    } finally {
      setBusy(null);
    }
  };
  return (
    <Dialog title={t('talk.settings.talk-settings')} width={480} className="tk-settings" onClose={onClose} actions={<Button primary label={t('common.done')} onClick={onClose} />}>
      <label className="tk-relay tk-option">
        <input type="checkbox" className="tk-start-with-office" checked={state.startWithOffice !== false} onChange={(e) => set({ startWithOffice: e.target.checked })} />
        <span><strong>{t('talk.settings.connect-rutba-office')}</strong><small>{t('talk.settings.contacts-call-write')}</small></span>
      </label>
      <label className="tk-relay tk-option">
        <input type="checkbox" className="tk-ring-sound" checked={state.ring !== false} onChange={(e) => set({ ring: e.target.checked })} />
        <span><strong>{t('talk.settings.play-sound-calls')}</strong><small>{t('talk.settings.ring-someone-calls')}</small></span>
      </label>
      <h3 className="tk-settings-head">{t('talk.settings.blocked-people')}</h3>
      {!blocked.length ? <p className="tk-none">{t('talk.settings.nobody-is-blocked')}</p> : (
        <>
          <div className="tk-member-list tk-blocked-list">
            {blocked.map((p) => (
              <div key={p.id} className="tk-member" data-blocked={p.id}>
                <span className="tk-who"><strong>{blockedName(p)}</strong>{p.name && p.email ? <small>{p.email}</small> : null}</span>
                <Button className="tk-unblock" label={t('talk.settings.unblock')} disabled={!online || busy === p.id} title={online ? '' : t('talk.settings.talk-not-connected')} onClick={() => unblock(p)} />
              </div>
            ))}
          </div>
          <p className="tk-lead tk-blocked-note">{t('talk.settings.someone-unblock-ask')}</p>
        </>
      )}
    </Dialog>
  );
}
