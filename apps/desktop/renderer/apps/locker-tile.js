// The Locker's card on the launcher.
import React from 'react';
import { t } from '@rutba/office-ui';

export default function LockerTile({ Card, onOpen }) {
  return <Card app={{ key: 'locker', short: t('Locker'), icon: 'lock', tagline: t('Lock folders so only you can open them') }} onOpen={onOpen} />;
}
