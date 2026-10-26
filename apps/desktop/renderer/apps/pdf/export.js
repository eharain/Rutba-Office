// Pages out as pictures, and a smaller PDF.
//
// Export as Pictures draws each page chosen with pdf.js at the resolution
// chosen and writes it as a PNG or a JPEG into a folder, named as Acrobat
// names them (Report_Page_1.png). Reduce File Size has the main process find
// the pictures shown at more than the resolution chosen (@rutba/pdf/compress);
// they are drawn smaller here, on a canvas, and encoded as JPEG, and the main
// process writes the file whole with them in place of the old, as a new file.

import React, { useState } from 'react';
import { Button, Dialog, Select, t } from '@rutba/office-ui';
import { getPage } from '../../pdfjs.js';

export const DPI_CHOICES = [72, 96, 150, 200, 300, 600];

/** A page drawn as a picture's bytes: PNG, or JPEG at `quality`. */
export async function pageBytes(doc, index, { rotate = 0, dpi = 150, type = 'png', quality = 0.9 }) {
  const page = await getPage(doc, index);
  const viewport = page.getViewport({ scale: dpi / 72, rotation: rotate });
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.floor(viewport.width));
  canvas.height = Math.max(1, Math.floor(viewport.height));
  const ctx = canvas.getContext('2d', { alpha: false });
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvasContext: ctx, canvas, viewport, background: '#ffffff' }).promise;
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, type === 'jpeg' ? 'image/jpeg' : 'image/png', quality));
  const bytes = new Uint8Array(await blob.arrayBuffer());
  canvas.width = 0;
  canvas.height = 0;
  return { bytes, width: Math.floor(viewport.width), height: Math.floor(viewport.height) };
}

/** Export as Pictures: which pages, in what format, at what resolution. */
export function ExportPicturesDialog({ count, picked, onExport, onClose }) {
  const [scope, setScope] = useState(picked > 1 ? 'picked' : 'all');
  const [format, setFormat] = useState('png');
  const [dpi, setDpi] = useState(150);
  const pages = scope === 'all' ? count : scope === 'picked' ? picked : 1;
  return (
    <Dialog
      title={t('Export as Pictures')}
      onClose={onClose}
      actions={(
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('Export…')} data-role="export-go" onClick={() => onExport({ scope, format, dpi })} />
        </>
      )}
    >
      <div className="pdf-export-form" data-role="export-dialog">
        <fieldset className="pdf-export-scope">
          <legend>{t('Pages')}</legend>
          <label><input type="radio" name="pdf-export-scope" checked={scope === 'all'} onChange={() => setScope('all')} />{t('All pages')}</label>
          <label><input type="radio" name="pdf-export-scope" checked={scope === 'picked'} disabled={!picked} onChange={() => setScope('picked')} />{t('The selected pages')}</label>
          <label><input type="radio" name="pdf-export-scope" checked={scope === 'current'} onChange={() => setScope('current')} />{t('This page')}</label>
        </fieldset>
        <div className="pdf-export-row">
          <label htmlFor="pdf-export-format">{t('Format')}</label>
          <Select id="pdf-export-format" value={format} data-role="export-format" onChange={(e) => setFormat(e.target.value)}>
            <option value="png">{t('PNG: sharp, for text and drawings')}</option>
            <option value="jpeg">{t('JPEG: smaller, for photographs')}</option>
          </Select>
        </div>
        <div className="pdf-export-row">
          <label htmlFor="pdf-export-dpi">{t('Resolution')}</label>
          <Select id="pdf-export-dpi" value={dpi} data-role="export-dpi" onChange={(e) => setDpi(Number(e.target.value))}>
            {DPI_CHOICES.map((d) => <option key={d} value={d}>{t('{dpi} dots per inch', { dpi: d })}</option>)}
          </Select>
        </div>
        <div className="pdf-export-note">{t('Each page is written as a picture into a folder you choose: {count} in all. This document is not changed.', { count: pages })}</div>
      </div>
    </Dialog>
  );
}

export const QUALITIES = [
  { key: 'high', dpi: 225, label: () => t('High quality (225 dots per inch)') },
  { key: 'standard', dpi: 150, label: () => t('Standard (150 dots per inch)') },
  { key: 'small', dpi: 96, label: () => t('Smallest (96 dots per inch)') },
];

/** Reduce File Size: how small the pictures may be drawn. */
export function CompressDialog({ size, onCompress, onClose }) {
  const [quality, setQuality] = useState('standard');
  return (
    <Dialog
      title={t('Reduce File Size')}
      onClose={onClose}
      actions={(
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('Reduce and Save As…')} data-role="compress-go" onClick={() => onCompress(QUALITIES.find((q) => q.key === quality))} />
        </>
      )}
    >
      <div className="pdf-export-form" data-role="compress-dialog">
        <p className="pdf-export-lead">{t('Pictures shown at more detail than the page needs are drawn again at the resolution chosen. Words, drawings and pictures already small enough are left as they are.')}</p>
        <fieldset className="pdf-export-scope">
          <legend>{t('Pictures')}</legend>
          {QUALITIES.map((q) => (
            <label key={q.key}><input type="radio" name="pdf-compress" checked={quality === q.key} data-quality={q.key} onChange={() => setQuality(q.key)} />{q.label()}</label>
          ))}
        </fieldset>
        <div className="pdf-export-note">{t('The smaller PDF is saved as a new file; this one is not changed. It is now {size}.', { size })}</div>
      </div>
    </Dialog>
  );
}

/**
 * A compression plan's pictures drawn at their new size and encoded as
 * JPEG: { [ref]: { data, width, height, filter, comps } }. A picture the
 * browser cannot decode is left out, and stays as it was.
 */
export async function encodePlan(plan, { quality = 0.8 } = {}) {
  const out = {};
  for (const item of plan) {
    try {
      let source;
      if (/DCT/.test(item.filter)) {
        source = await createImageBitmap(new Blob([item.data], { type: 'image/jpeg' }));
      } else {
        const rgba = new Uint8ClampedArray(item.width * item.height * 4);
        for (let i = 0, o = 0; o < rgba.length; o += 4, i += item.comps) {
          rgba[o] = item.data[i];
          rgba[o + 1] = item.data[item.comps === 1 ? i : i + 1];
          rgba[o + 2] = item.data[item.comps === 1 ? i : i + 2];
          rgba[o + 3] = 255;
        }
        source = await createImageBitmap(new ImageData(rgba, item.width, item.height));
      }
      const canvas = document.createElement('canvas');
      canvas.width = item.toWidth;
      canvas.height = item.toHeight;
      const ctx = canvas.getContext('2d', { alpha: false });
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high';
      ctx.drawImage(source, 0, 0, item.toWidth, item.toHeight);
      source.close?.();
      const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/jpeg', quality));
      out[item.ref] = { data: new Uint8Array(await blob.arrayBuffer()), width: item.toWidth, height: item.toHeight, filter: 'DCTDecode', comps: 3 };
    } catch {
      /* left as it was */
    }
  }
  return out;
}
