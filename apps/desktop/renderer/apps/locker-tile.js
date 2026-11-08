// The Locker's card on the launcher, in a build made with its switch only.
import React from 'react';
import { t } from './locker-words.js';

export default function LockerTile({ Card, onOpen }) {
  return <Card app={{ key: 'locker', short: t('Locker'), icon: 'lock', tagline: t('Lock folders so only you can open them') }} onOpen={onOpen} />;
}
