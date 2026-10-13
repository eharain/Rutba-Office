// Insert > QR Code, shared by Documents, Worksheets and Presentations: a
// code a phone can scan, for a link or any words, encoded by the suite's own
// encoder (@rutba/barcode) and put in as a picture. The picture's alt text
// is the code's text and its title says it is a QR Code, at which level and
// in which colour, so a double-click on it opens this dialog again with all
// three, and OK draws it anew in its place.

import React, { useEffect, useMemo, useState } from 'react';
import { Button, Dialog, t, msg } from '@rutba/office-ui';
import { encodeQr } from '@rutba/barcode';

/** The title a QR Code picture carries: "QR Code; M; #1F3864". */
export const qrTitle = ({ level, colour }) => `QR Code; ${level}; ${String(colour).toUpperCase()}`;

/** A picture's QR Code, read from its alt text and title: { text, level, colour }, or null for any other picture. */
export function qrOf(picture) {
  const m = /^QR Code; ([LMQH]); (#[0-9A-F]{6})$/i.exec(String(picture?.title ?? ''));
  if (!m || picture?.descr == null) return null;
  return { text: String(picture.descr), level: m[1].toUpperCase(), colour: m[2].toUpperCase() };
}

const LEVELS = [['L', msg('Low (7%)')], ['M', msg('Medium (15%)')], ['Q', msg('Quartile (25%)')], ['H', msg('High (30%)')]];
const SIZES = [[128, msg('Small')], [192, msg('Medium')], [256, msg('Large')], [384, msg('Extra large')]];
const COLOURS = [['#000000', msg('Black')], ['#1F3864', msg('Navy')], ['#1F4E79', msg('Dark blue')], ['#00695C', msg('Dark teal')], ['#1B5E20', msg('Dark green')], ['#7B1FA2', msg('Purple')], ['#B71C1C', msg('Dark red')], ['#3E2723', msg('Dark brown')]];

/** The symbol's modules, or the reason there are none. */
function symbolOf(text, level) {
  if (!text) return { error: t('Type the words or the address the code should hold.') };
  try {
    return { qr: encodeQr(text, { level }) };
  } catch {
    return { error: t('That is too much for a QR Code at this level of error correction. Choose a lower level, or shorten the text.') };
  }
}

/** The symbol as one SVG path, a module a unit, the quiet zone four modules round it. */
function pathOf(qr) {
  let d = '';
  qr.modules.forEach((row, y) => row.forEach((on, x) => { if (on) d += `M${x + 4} ${y + 4}h1v1h-1z`; }));
  return d;
}

/**
 * The code as a PNG: whole pixels a module, so a phone and the checks read
 * it cleanly, with four modules of white round it; about `size` pixels
 * across. Answers { bytes, widthPx, heightPx }.
 */
export async function qrPng(qr, { size = 192, colour = '#000000' } = {}) {
  const n = qr.size + 8;
  const scale = Math.max(2, Math.round(size / n));
  const px = n * scale;
  const canvas = document.createElement('canvas');
  canvas.width = px;
  canvas.height = px;
  const g = canvas.getContext('2d');
  g.fillStyle = '#FFFFFF';
  g.fillRect(0, 0, px, px);
  g.fillStyle = colour;
  qr.modules.forEach((row, y) => row.forEach((on, x) => { if (on) g.fillRect((x + 4) * scale, (y + 4) * scale, scale, scale); }));
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
  return { bytes: new Uint8Array(await blob.arrayBuffer()), widthPx: px, heightPx: px };
}

/**
 * @param {{ initial?: { text, level, colour, size }, onInsert, onClose }} props
 *   onInsert({ bytes, widthPx, heightPx, contentType, name, descr, title }): the picture, drawn
 */
export function QrDialog({ initial = null, onInsert, onClose }) {
  const [text, setText] = useState(initial?.text ?? '');
  const [level, setLevel] = useState(initial?.level ?? 'M');
  const [size, setSize] = useState(() => SIZES.reduce((best, [s]) => (Math.abs(s - (initial?.size ?? 192)) < Math.abs(best - (initial?.size ?? 192)) ? s : best), 192));
  const [colour, setColour] = useState(initial?.colour ?? '#000000');
  const [busy, setBusy] = useState(false);
  const [shown, setShown] = useState(() => symbolOf(initial?.text ?? '', initial?.level ?? 'M'));
  // The preview follows the typing a beat behind, so a long paste does not encode every keystroke.
  useEffect(() => {
    const timer = setTimeout(() => setShown(symbolOf(text, level)), 120);
    return () => clearTimeout(timer);
  }, [text, level]);
  const path = useMemo(() => (shown.qr ? pathOf(shown.qr) : ''), [shown]);
  const editing = Boolean(initial?.text);

  const insert = async () => {
    const now = symbolOf(text, level);
    if (!now.qr || busy) return;
    setBusy(true);
    try {
      const picture = await qrPng(now.qr, { size, colour });
      await onInsert({ ...picture, contentType: 'image/png', name: 'QR Code', descr: text, title: qrTitle({ level, colour }) });
    } finally {
      setBusy(false);
    }
  };

  return (
    <Dialog
      title={editing ? t('Edit QR Code') : t('Insert QR Code')}
      width={600}
      onClose={onClose}
      actions={<><Button label={t('Cancel')} onClick={onClose} /><Button primary label={editing ? t('OK') : t('Insert')} className="qr-ok" disabled={!shown.qr || busy || !text} onClick={insert} /></>}
    >
      <style>{QR_CSS}</style>
      <div className="qr">
        <div className="qr-form">
          <label className="qr-field">
            <span>{t('Text or web address')}</span>
            <textarea className="rw-input qr-text" rows={4} value={text} spellCheck={false} autoFocus placeholder="https://" onChange={(e) => setText(e.target.value)} />
          </label>
          <label className="qr-field">
            <span>{t('Error correction')}</span>
            <select className="rw-input qr-level" value={level} onChange={(e) => setLevel(e.target.value)}>
              {LEVELS.map(([v, label]) => <option key={v} value={v}>{t(label)}</option>)}
            </select>
            <small>{t('More correction still scans when part of the code is covered or worn, at the cost of a denser code.')}</small>
          </label>
          <label className="qr-field">
            <span>{t('Size')}</span>
            <select className="rw-input qr-size" value={size} onChange={(e) => setSize(Number(e.target.value))}>
              {SIZES.map(([v, label]) => <option key={v} value={v}>{t('{name} ({px} px)', { name: t(label), px: v })}</option>)}
            </select>
          </label>
          <div className="qr-field">
            <span>{t('Colour')}</span>
            <div className="qr-colours" role="radiogroup" aria-label={t('Colour')}>
              {COLOURS.map(([c, label]) => (
                <button key={c} type="button" role="radio" aria-checked={colour === c} className={`qr-swatch${colour === c ? ' on' : ''}`} data-colour={c} data-tip={t(label)} style={{ background: c }} onClick={() => setColour(c)} />
              ))}
            </div>
          </div>
        </div>
        <div className="qr-preview">
          <div className="qr-art" aria-label={t('Preview')}>
            {shown.qr ? (
              <svg viewBox={`0 0 ${shown.qr.size + 8} ${shown.qr.size + 8}`} shapeRendering="crispEdges">
                <rect width="100%" height="100%" fill="#fff" />
                <path d={path} fill={colour} />
              </svg>
            ) : <span className="qr-none">{shown.error}</span>}
          </div>
          {shown.qr ? <small className="qr-about">{t('Version {version}, {n} by {n} modules', { version: shown.qr.version, n: shown.qr.size })}</small> : null}
        </div>
      </div>
    </Dialog>
  );
}

const QR_CSS = `
.qr { display: flex; gap: 18px; align-items: stretch; }
.qr-form { flex: 1; min-width: 0; display: flex; flex-direction: column; gap: 12px; }
.qr-field { display: flex; flex-direction: column; gap: 5px; font-size: 12.5px; color: var(--ink-2); }
.qr-field > span { font-weight: 600; color: var(--ink); }
.qr-field small { color: var(--ink-3); font-size: 11.5px; line-height: 1.35; }
.qr-text { resize: vertical; min-height: 76px; font-family: var(--font-mono, ui-monospace, monospace); font-size: 12.5px; box-sizing: border-box; width: 100%; }
.qr-colours { display: flex; flex-wrap: wrap; gap: 7px; }
.qr-swatch { width: 22px; height: 22px; border-radius: 50%; border: 2px solid var(--surface); box-shadow: 0 0 0 1px var(--line); cursor: pointer; padding: 0; }
.qr-swatch.on { box-shadow: 0 0 0 2px var(--accent); }
.qr-preview { width: 212px; flex: none; display: flex; flex-direction: column; align-items: center; gap: 8px; }
.qr-art {
  width: 212px; height: 212px; border-radius: var(--r-3); border: 1px solid var(--line); background: #fff;
  display: grid; place-items: center; overflow: hidden;
}
.qr-art svg { width: 196px; height: 196px; display: block; }
.qr-none { padding: 18px; text-align: center; color: var(--ink-3); font-size: 12px; line-height: 1.4; }
.qr-about { color: var(--ink-3); font-size: 11.5px; }
`;
