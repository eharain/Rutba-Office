// What a source is called and drawn as, for the Set up columns and what reads them out.

import { t } from '@rutba/office-ui';

export const KIND_ICON = { screen: 'screen', window: 'window', camera: 'camera', title: 'textbox', picture: 'picture', web: 'globe', mic: 'mic', system: 'speaker' };

/** A kind's name, for what reads it out and for the source bar. */
export const kindName = (kind) => ({
  screen: t('Screen'), window: t('Window'), camera: t('Camera'), title: t('Title'), picture: t('Picture'), web: t('Web page'), mic: t('Microphone'), system: t('Computer sound'),
}[kind] || kind);

/** An item's name as shown: its own, or its kind's. */
export const itemName = (it) => (it.kind === 'title' ? it.text || t('Title') : it.kind === 'picture' ? (it.path ? it.path.split(/[\\/]/).pop() : t('Picture')) : it.name || kindName(it.kind));
