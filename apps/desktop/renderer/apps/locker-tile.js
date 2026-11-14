// The Locker's card on the launcher.
import React from 'react';
import { t } from './locker-words.js';

export default function LockerTile({ Card, onOpen }) {
  return <Card app={{ key: 'locker', short: t('locker.locker'), icon: 'lock', tagline: t('locker.tile.lock-folders-only') }} onOpen={onOpen} />;
}
