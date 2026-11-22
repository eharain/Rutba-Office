// The calls that were missed (Office 2.28), newest first: who, when, and one
// button to call back. Opening the list is looking at it, so its count goes.

import React, { useEffect, useState } from 'react';
import { Button, Dialog, Spinner, formatWhen } from '@rutba/office-ui';
import { t } from '../talk-words.js';
import { missedText, missedWhy } from './lines.js';

export function MissedCalls({ shell, state, onClose, onCallBack, onOpen }) {
  const [calls, setCalls] = useState(null);
  useEffect(() => { shell.talk.calls({ seen: true }).then(setCalls, () => setCalls([])); }, [shell]);
  return (
    <Dialog title={t('talk.missed-calls')} width={460} className="tk-missed" onClose={onClose} actions={<Button primary label={t('common.done')} onClick={onClose} />}>
      {!calls ? <Spinner /> : !calls.length ? <p className="tk-none">{t('talk.missed.no-missed-calls')}</p> : (
        <div className="tk-member-list">
          {calls.map((c) => {
            const contact = state.contacts.find((x) => x.id === c.with) || null;
            const reachable = contact && contact.presence && contact.presence !== 'offline';
            const name = contact?.name || c.name || t('talk.missed.someone-no-longer');
            return (
              <div key={c.id} className={`tk-member${c.seen ? '' : ' new'}`} data-missed={c.id}>
                <span className="tk-who"><strong>{name}</strong><small>{[missedText(c), formatWhen(c.at), missedWhy(c.why)].filter(Boolean).join(' · ')}</small></span>
                {contact ? <Button ghost className="tk-missed-open" label={t('mail.archive.write')} onClick={() => { onOpen(contact); onClose(); }} /> : null}
                {contact ? <Button icon={c.video ? 'camera' : 'phone'} className="tk-missed-back" label={t('talk.missed.call-back')} disabled={!reachable} title={reachable ? '' : t('talk.calls-reach')} onClick={() => { onCallBack(contact, c.video); onClose(); }} /> : null}
              </div>
            );
          })}
        </div>
      )}
    </Dialog>
  );
}
