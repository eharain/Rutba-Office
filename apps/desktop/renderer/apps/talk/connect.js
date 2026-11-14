// Connect (Office 2.25): someone who already uses Talk is asked at once, in
// Talk, with no email; someone who asked first becomes a contact at once.
// Where a person stands comes from what this copy has seen (talk.reach): the
// introducer never says whether an address has an account.

import React from 'react';
import { Button } from '@rutba/office-ui';
import { t } from '../talk-words.js';

/** How a request made with Connect was told, as the list of people asked shows it. */
export const askedInTalkText = () => t('talk.connect.asked-in-talk');

/** A short mark of where a person stands, or null. */
export function reachText(reach) {
  switch (reach) {
    case 'contact': return t('talk.connect.talk-contacts');
    case 'asked-you': return t('talk.connect.asked-contact');
    case 'asked': return t('talk.connect.asked-already');
    case 'uses-talk': return t('talk.connect.uses-talk');
    default: return null;
  }
}

/** Whether Connect is offered: they use Talk, or asked first. */
export const canConnect = (reach) => reach === 'uses-talk' || reach === 'asked-you';

/** Connect, and say how it went. */
export async function connectWith({ shell, person, toast }) {
  const name = person.name || person.email;
  const r = await shell.talk.connect({ email: person.email }).catch((err) => ({ error: err?.message || 'server' }));
  if (r.contact) toast(t('talk.invite.name-contacts-now', { name: r.contact.name || name }), { tone: 'good', ms: 4000 });
  else if (r.request) toast(t('talk.connect.asked-name-talk', { name }), { tone: 'good', ms: 4000 });
  return r;
}

export function ConnectButton({ shell, person, toast, onDone }) {
  return (
    <Button
      primary
      label={t('common.connect')}
      title={t('talk.connect.ask-name-talk', { name: person.name || person.email })}
      className="tk-connect-now"
      onClick={async (e) => {
        e.preventDefault();
        e.stopPropagation();
        const r = await connectWith({ shell, person, toast });
        onDone?.(r);
      }}
    />
  );
}
