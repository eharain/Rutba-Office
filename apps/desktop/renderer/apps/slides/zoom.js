// Insert → Zoom: Summary Zoom, Section Zoom and Slide Zoom — a picture of
// each slide (or section's first slide) picked, that takes the show there
// with a click and, at the end of that slide (or section), back to the
// zoom, as PowerPoint's zooms return. Kept in the deck as pictures linked
// to their slides, named "Slide Zoom", "Section Zoom" or "Summary Zoom",
// which any PowerPoint opens; the return is this suite's show's own.

import React, { useState } from 'react';
import { Button, Dialog, t } from '@rutba/office-ui';

/** Whether a shape is one of the zooms this inserts, and of which kind. */
export const zoomKind = (shape) => /^(Slide|Section|Summary) Zoom\b/.exec(shape?.name || '')?.[1]?.toLowerCase() || null;

/** A slide's picture, as PNG bytes, from its drawing — pictures in it carried along. */
export async function slidePicture(svg, width, height) {
  const urls = [...new Set([...svg.matchAll(/(?:href|xlink:href)="(rutba:[^"]+|blob:[^"]+)"/g)].map((m) => m[1]))];
  let own = svg.replace(/<foreignObject\b[\s\S]*?<\/foreignObject>/g, '');
  for (const url of urls) {
    try {
      const blob = await (await fetch(url)).blob();
      const data = await new Promise((resolve) => { const r = new FileReader(); r.onload = () => resolve(r.result); r.readAsDataURL(blob); });
      own = own.split(url).join(String(data));
    } catch { /* left out */ }
  }
  const img = new Image();
  img.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(own)}`;
  await img.decode();
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  canvas.getContext('2d').drawImage(img, 0, 0, width, height);
  return new Uint8Array(await (await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'))).arrayBuffer());
}

/**
 * The Zoom box: the slides (or, for a Section Zoom, the sections) with a
 * box each; `onInsert(list)` with the slide indexes picked — a section by
 * its first slide.
 */
export function ZoomDialog({ kind, outline = [], sections = [], onInsert, onClose }) {
  const items = kind === 'section'
    ? sections.map((s) => ({ key: s.start ?? s.first ?? s.slides?.[0] ?? 0, label: s.name, thumb: outline[s.start ?? s.first ?? s.slides?.[0] ?? 0]?.thumbnail }))
    : outline.map((o) => ({ key: o.index, label: `${o.index + 1}. ${o.title || t('Slide {number}', { number: o.index + 1 })}`, thumb: o.thumbnail }));
  const [picked, setPicked] = useState([]);
  const title = { summary: t('Insert Summary Zoom'), section: t('Insert Section Zoom'), slide: t('Insert Slide Zoom') }[kind];
  return (
    <Dialog
      title={title}
      width={620}
      onClose={onClose}
      actions={<><Button label={t('Cancel')} onClick={onClose} /><Button primary label={t('Insert')} className="sl-zoom-ok" disabled={!picked.length} onClick={() => onInsert([...picked].sort((a, b) => a - b))} /></>}
    >
      <div className="sl-zoom-grid">
        {items.map((it) => (
          <label key={it.key} className={`sl-zoom-item${picked.includes(it.key) ? ' on' : ''}`} data-slide={it.key}>
            <input type="checkbox" checked={picked.includes(it.key)} onChange={(e) => setPicked((p) => (e.target.checked ? [...p, it.key] : p.filter((x) => x !== it.key)))} />
            <span className="sl-zoom-thumb" dangerouslySetInnerHTML={it.thumb ? { __html: it.thumb } : undefined} />
            <span className="sl-zoom-label">{it.label}</span>
          </label>
        ))}
        {!items.length ? <p className="sl-zoom-empty">{kind === 'section' ? t('This deck has no sections. Add one from the slides strip first.') : t('There are no slides to zoom to.')}</p> : null}
      </div>
    </Dialog>
  );
}

export const ZOOM_CSS = `
.sl-zoom-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(160px, 1fr)); gap: 10px; max-height: 52vh; overflow: auto; }
.sl-zoom-item { display: flex; flex-direction: column; gap: 4px; border: 1px solid var(--line-soft); border-radius: var(--r-2); padding: 6px; cursor: pointer; font-size: 12px; position: relative; }
.sl-zoom-item.on { border-color: var(--accent); background: var(--selected); }
.sl-zoom-item input { position: absolute; top: 8px; left: 8px; }
.sl-zoom-thumb { display: block; aspect-ratio: 16 / 9; background: #fff; border: 1px solid var(--line-soft); overflow: hidden; }
.sl-zoom-thumb svg { width: 100%; height: auto; display: block; }
.sl-zoom-label { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sl-zoom-empty { color: var(--ink-3); }
`;
