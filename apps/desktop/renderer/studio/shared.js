// Two small pieces both the Record and the Live panel show.

import React from 'react';
import { Button, Icon, basename, t } from '@rutba/office-ui';

/** A length as a clock: 1:02:03 or 2:03; with `long`, always hours, as a recording's Stop shows it. */
export const clock = (ms, long = false) => {
  const s = Math.floor(ms / 1000);
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const ss = String(s % 60).padStart(2, '0');
  if (long) return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${ss}`;
  return h ? `${h}:${String(m).padStart(2, '0')}:${ss}` : `${m}:${ss}`;
};

/** Take Picture, as today: the button, and the key. */
export function PictureButton({ s }) {
  return (
    <button type="button" className="st-bb st-shoot" data-tip={`${t('Take Picture')}${s.pictureKey ? ` (${s.pictureKey})` : ''}`} disabled={s.shotCount != null || !s.settings.folder} onClick={s.takePicture}>
      <Icon name="picture" size={16} />{s.shotCount != null ? t('In {seconds} s', { seconds: s.shotCount }) : t('Picture')}
    </button>
  );
}

/** The picture just taken, with what to do with it. */
export function PictureStrip({ s }) {
  const { picture, shell } = s;
  if (!picture) return null;
  return (
    <div className="st-shot" role="status">
      <img src={picture.url} alt="" />
      <span className="st-shot-name">{basename(picture.path)}</span>
      <span className="st-shot-meta">{picture.width} × {picture.height}</span>
      <Button icon="picture" label={t('Open')} className="st-shot-open" onClick={() => shell.win.create({ app: 'pictures', file: picture.path })} />
      <Button icon="image" label={t('Edit in Image')} className="st-shot-edit" onClick={() => shell.win.create({ app: 'image', file: picture.path })} />
      <Button icon="copy" label={t('Copy')} className="st-shot-copy" onClick={s.copyPicture} />
      <Button icon="folderOpen" label={t('Show in folder')} onClick={() => shell.shell.showInFolder({ path: picture.path })} />
    </div>
  );
}
