// Insert → Signature Line: Word's Signature Setup — who is asked to sign,
// their title and e-mail, what they are told, whether the date shows — and
// the picture of the line itself, drawn here: an X, the line, and the
// signer's name and title under it, as Word draws one before it is signed.

import React, { useState } from 'react';
import { Button, Dialog, Field, t } from '@rutba/office-ui';

const WIDTH = 256;
const HEIGHT = 128;

/** The unsigned line as a PNG, at twice its size for a sharp print. */
export async function signatureLinePng({ signer = '', title = '' } = {}) {
  const scale = 2;
  const canvas = document.createElement('canvas');
  canvas.width = WIDTH * scale;
  canvas.height = HEIGHT * scale;
  const g = canvas.getContext('2d');
  g.scale(scale, scale);
  g.fillStyle = '#ffffff';
  g.fillRect(0, 0, WIDTH, HEIGHT);
  g.strokeStyle = '#000000';
  g.fillStyle = '#000000';
  g.lineWidth = 1.5;
  // The X, then the line it sits on.
  g.beginPath();
  g.moveTo(12, 46); g.lineTo(30, 64);
  g.moveTo(30, 46); g.lineTo(12, 64);
  g.stroke();
  g.lineWidth = 1;
  g.beginPath();
  g.moveTo(8, 72); g.lineTo(WIDTH - 8, 72);
  g.stroke();
  g.font = '13px "Segoe UI", Calibri, sans-serif';
  g.textBaseline = 'top';
  const fit = (text, y) => {
    let t = String(text || '');
    while (t && g.measureText(t).width > WIDTH - 16) t = t.slice(0, -1);
    if (t) g.fillText(t, 8, y);
  };
  fit(signer, 78);
  g.font = '12px "Segoe UI", Calibri, sans-serif';
  fit(title, 96);
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
  return { png: new Uint8Array(await blob.arrayBuffer()), widthPx: WIDTH, heightPx: HEIGHT };
}

export function SignatureSetupDialog({ onInsert, onClose }) {
  const [s, setS] = useState({ signer: '', title: '', email: '', instructions: '', allowComments: false, showDate: true });
  const set = (patch) => setS((v) => ({ ...v, ...patch }));
  const field = (key, label, extra = {}) => (
    <Field label={label}>
      <input className={`rw-input wd-sig-${key}`} value={s[key]} onChange={(e) => set({ [key]: e.target.value })} {...extra} />
    </Field>
  );
  return (
    <Dialog
      title={t('Signature Setup')}
      width={440}
      onClose={onClose}
      actions={<><Button label={t('Cancel')} onClick={onClose} /><Button primary label={t('OK')} className="wd-sig-ok" onClick={() => onInsert(s)} /></>}
    >
      <div className="wd-sig">
        {field('signer', t('Suggested signer (for example, Jo Bloggs)'), { autoFocus: true })}
        {field('title', t('Suggested signer\'s title (for example, Manager)'))}
        {field('email', t('Suggested signer\'s e-mail address'))}
        <Field label={t('Instructions to the signer')}>
          <textarea className="rw-input wd-sig-instructions" rows={2} value={s.instructions} onChange={(e) => set({ instructions: e.target.value })} />
        </Field>
        <label className="wd-sig-check"><input type="checkbox" checked={s.allowComments} onChange={(e) => set({ allowComments: e.target.checked })} /> {t('Allow the signer to add comments in the Sign dialog')}</label>
        <label className="wd-sig-check"><input type="checkbox" className="wd-sig-date" checked={s.showDate} onChange={(e) => set({ showDate: e.target.checked })} /> {t('Show sign date in signature line')}</label>
        <p className="wd-sig-note">{t('Word offers to sign the line; this suite keeps the signer\'s details and the line as Word writes them.')}</p>
      </div>
    </Dialog>
  );
}

export const SIGNATURE_CSS = `
.wd-sig { display: flex; flex-direction: column; gap: 8px; font-size: 12.5px; }
.wd-sig .rw-input { width: 100%; box-sizing: border-box; }
.wd-sig textarea { resize: vertical; font: inherit; }
.wd-sig-check { display: flex; align-items: center; gap: 6px; }
.wd-sig-note { margin: 2px 0 0; color: var(--ink-3); font-size: 11.5px; line-height: 1.4; }
`;
