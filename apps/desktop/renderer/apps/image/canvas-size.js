// Canvas size (I15): the canvas made larger or smaller without scaling
// anything on it, grown or cut from one of nine anchors, as Paint.NET and
// Photoshop ask it: a width and a height in pixels and a grid of arrows that
// says where the picture stays.

import React, { useState } from 'react';
import { Dialog, Button, Input, t } from '@rutba/office-ui';
import { ANCHORS, MAX_SIDE } from '@rutba/studio/layers';
import { ANCHOR_NAMES } from './words.js';

/** The arrow each anchor cell shows, pointing away from the anchor. */
function arrow(anchor, cell) {
  if (anchor === cell) return '●';
  const [ar, ac] = [Math.floor(ANCHORS.indexOf(anchor) / 3), ANCHORS.indexOf(anchor) % 3];
  const [cr, cc] = [Math.floor(ANCHORS.indexOf(cell) / 3), ANCHORS.indexOf(cell) % 3];
  const dr = cr - ar;
  const dc = cc - ac;
  if (Math.abs(dr) > 1 || Math.abs(dc) > 1) return '';
  return { '-1,-1': '↖', '-1,0': '↑', '-1,1': '↗', '0,-1': '←', '0,1': '→', '1,-1': '↙', '1,0': '↓', '1,1': '↘' }[`${dr},${dc}`] || '';
}

export function CanvasSizeDialog({ size, onApply, onClose }) {
  const [width, setWidth] = useState(size.width);
  const [height, setHeight] = useState(size.height);
  const [anchor, setAnchor] = useState('c');
  const w = Math.max(1, Math.min(MAX_SIDE, Math.round(Number(width) || 0)));
  const h = Math.max(1, Math.min(MAX_SIDE, Math.round(Number(height) || 0)));
  return (
    <Dialog
      title={t('Canvas Size')}
      width={380}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('OK')} onClick={() => onApply({ width: w, height: h, anchor })} data-role="canvas-ok" />
        </>
      }
    >
      <div className="im-canvas-size">
        <div className="im-resize-grid">
          <span>{t('Width')}</span>
          <Input type="number" min={1} max={MAX_SIDE} value={width} data-role="canvas-width" onChange={(e) => setWidth(e.target.value)} style={{ width: 96 }} />
          <span>{t('px')}</span>
          <span>{t('Height')}</span>
          <Input type="number" min={1} max={MAX_SIDE} value={height} data-role="canvas-height" onChange={(e) => setHeight(e.target.value)} style={{ width: 96 }} />
          <span>{t('px')}</span>
        </div>
        <div className="im-anchor-row">
          <span className="im-anchor-label">{t('Anchor')}</span>
          {/* The grid is the picture's place, so it reads the same way in every language. */}
          <div className="im-anchor" role="radiogroup" aria-label={t('Anchor')} dir="ltr">
            {ANCHORS.map((a) => (
              <button
                key={a}
                type="button"
                role="radio"
                aria-checked={anchor === a}
                aria-label={ANCHOR_NAMES[a]}
                data-tip={ANCHOR_NAMES[a]}
                data-anchor={a}
                className={anchor === a ? 'on' : ''}
                onClick={() => setAnchor(a)}
              >
                {arrow(anchor, a)}
              </button>
            ))}
          </div>
          <div className="rw-hint im-anchor-hint">{t('The picture stays at the anchor; the canvas grows or is cut on the other sides. Nothing is scaled.')}</div>
        </div>
        <div className="rw-hint">{t('Now {width} × {height}; after, {newWidth} × {newHeight}.', { width: size.width, height: size.height, newWidth: w, newHeight: h })}</div>
      </div>
    </Dialog>
  );
}

export const CANVAS_CSS = `
.im-canvas-size { display: flex; flex-direction: column; gap: 14px; }
.im-anchor-row { display: grid; grid-template-columns: 64px auto 1fr; gap: 10px; align-items: start; font-size: 12.5px; color: var(--ink-2); }
.im-anchor-label { padding-top: 4px; }
.im-anchor { display: grid; grid-template-columns: repeat(3, 26px); grid-template-rows: repeat(3, 26px); gap: 3px; }
.im-anchor button { border: 1px solid var(--line); background: var(--surface); border-radius: 4px; font-size: 13px; line-height: 1; color: var(--ink-2); cursor: pointer; padding: 0; }
.im-anchor button:hover { border-color: var(--n-50); color: var(--ink); }
.im-anchor button.on { background: var(--accent); border-color: var(--accent); color: #fff; }
.im-anchor-hint { font-size: 11.5px; line-height: 1.45; }
`;
