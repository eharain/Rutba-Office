// The brand kit (docs/STUDIO.md X6): a logo, three colours, a font and the
// lower third's name and role, kept for this person in the app's settings
// and drawn into every scene of an introduction.
//
// The box shows the kit as it will look: a small frame in the background
// colour with the logo, the title in the words' colour and the line under it
// in the accent, and the lower third as it comes in over a recording.

import React, { useState } from 'react';
import { Button, Dialog, Field, Input, Select, t } from '@rutba/office-ui';
import { normaliseBrand, BRAND_FONTS } from '@rutba/studio/intro';
import { fileUrl } from '../media.js';

/** The key the brand kit is kept under, per person. */
export const BRAND_KEY = 'video.brand';

/** A colour picked, with its value written beside it. */
function Colour({ label, value, onChange, name }) {
  return (
    <label className="vi-colour">
      <input type="color" className={`vi-colour-${name}`} value={value} onChange={(e) => onChange(e.target.value)} aria-label={label} />
      <span className="vi-colour-text">
        <b>{label}</b>
        <small>{value.toUpperCase()}</small>
      </span>
    </label>
  );
}

export function BrandDialog({ shell, brand, onClose, onSave }) {
  const [kit, setKit] = useState(() => normaliseBrand(brand));
  const set = (patch) => setKit((k) => normaliseBrand({ ...k, ...patch, colors: { ...k.colors, ...(patch.colors || {}) } }));
  const chooseLogo = async () => {
    const files = await shell.dialog.open({
      title: t('common.choose-a-logo'),
      filters: [{ name: t('common.pictures'), extensions: ['png', 'jpg', 'jpeg', 'webp', 'svg', 'gif'] }],
      multiple: false,
    });
    const file = files?.[0];
    if (file) set({ logo: { url: file, name: file.split(/[\\/]/).pop() } });
  };
  const rtl = /[֐-ࣿﭐ-﷿ﹰ-ﻼ]/.test(`${kit.name}${kit.role}`);
  const font = `"${kit.font}", "Segoe UI", system-ui, sans-serif`;
  return (
    <Dialog
      title={t('common.brand-kit')}
      onClose={onClose}
      width={620}
      className="vi-brand-dialog"
      actions={
        <>
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary className="vi-brand-save" label={t('common.use-every')} onClick={() => onSave(kit)} />
        </>
      }
    >
      <div className="vi-brand">
        <div className="vi-brand-form">
          <Field label={t('common.name-lower-third')}>
            <Input className="rw-input vi-brand-name" dir="auto" value={kit.name} placeholder={t('common.your-name')} onChange={(e) => set({ name: e.target.value })} />
          </Field>
          <Field label={t('common.role-under-name')}>
            <Input className="rw-input vi-brand-role" dir="auto" value={kit.role} placeholder={t('common.do-line')} onChange={(e) => set({ role: e.target.value })} />
          </Field>
          <Field label={t('common.logo')}>
            <div className="vi-logo-row">
              <span className="vi-logo-name">{kit.logo ? kit.logo.name : t('common.no-logo')}</span>
              <Button className="vi-logo-pick" label={kit.logo ? t('common.change-2') : t('common.choose')} onClick={chooseLogo} />
              {kit.logo ? <Button ghost icon="close" title={t('common.remove-the-logo')} onClick={() => set({ logo: null })} /> : null}
            </div>
          </Field>
          <Field label={t('common.colours')}>
            <div className="vi-colours">
              <Colour name="background" label={t('common.background')} value={kit.colors.background} onChange={(v) => set({ colors: { background: v } })} />
              <Colour name="text" label={t('common.words')} value={kit.colors.text} onChange={(v) => set({ colors: { text: v } })} />
              <Colour name="accent" label={t('common.accent')} value={kit.colors.accent} onChange={(v) => set({ colors: { accent: v } })} />
            </div>
          </Field>
          <Field label={t('common.font')}>
            <Select className="rw-select vi-brand-font" value={kit.font} onChange={(e) => set({ font: e.target.value })}>
              {[...new Set([...BRAND_FONTS, kit.font])].map((f) => <option key={f} value={f} style={{ fontFamily: `"${f}"` }}>{f}</option>)}
            </Select>
          </Field>
          <label className="vi-check"><input type="checkbox" className="vi-brand-third-on" checked={kit.lowerThird} onChange={(e) => set({ lowerThird: e.target.checked })} /> {t('common.show-name-role')}</label>
          <label className="vi-check"><input type="checkbox" className="vi-brand-mark" checked={kit.watermark} onChange={(e) => set({ watermark: e.target.checked })} /> {t('common.keep-logo-corner')}</label>
        </div>
        <div className="vi-brand-look" aria-label={t('common.how-it-looks')}>
          <div className="vi-brand-card" style={{ background: kit.colors.background, fontFamily: font }}>
            {kit.logo ? <img src={fileUrl(kit.logo.url)} alt="" className="vi-brand-logo" /> : null}
            <b style={{ color: kit.colors.text }} dir="auto">{kit.name || t('common.your-name')}</b>
            <span style={{ color: kit.colors.accent }} dir="auto">{kit.role || t('common.do-line')}</span>
          </div>
          <div className="vi-brand-scene" style={{ fontFamily: font }}>
            <span className="vi-brand-face" aria-hidden="true" />
            {kit.lowerThird ? (
              <span className={`vi-brand-third${rtl ? ' rtl' : ''}`} style={{ background: kit.colors.accent, color: kit.colors.background }}>
                <b dir="auto">{kit.name || t('common.your-name')}</b>
                <small dir="auto">{kit.role || t('common.do-line')}</small>
              </span>
            ) : null}
            {kit.logo && kit.watermark ? <img src={fileUrl(kit.logo.url)} alt="" className="vi-brand-mark-img" /> : null}
          </div>
          <small className="vi-note">{t('common.kept-used-every')}</small>
        </div>
      </div>
    </Dialog>
  );
}
