// Review → Compare in Presentations: the Revisions pane and the markers on
// the slide, as PowerPoint shows a reviewer's changes. The pane lists the
// changes on this slide (each shape's, a box to tick for each) and the
// presentation's (slides added, deleted, moved); a marker sits by every
// changed shape, and a press on it lists that shape's changes to tick. A
// tick is Accept, an untick Reject; Accept and Reject on the ribbon do one,
// this slide's or all of them.

import React, { useState } from 'react';
import { Button, Icon, t, tn } from '@rutba/office-ui';

const EMU_PX = 9525;

/** A change, in words. */
export function changeWords(c) {
  const quote = (list) => list.map((w) => '"' + (w.length > 40 ? w.slice(0, 40) + '…' : w) + '"').join(', ');
  switch (c.kind) {
    case 'text': {
      if (c.detail?.table) return t('Table changed');
      if (c.detail?.formatted) return t('Text formatting changed');
      const parts = [];
      if (c.detail?.inserted?.length) parts.push(t('Inserted {words}', { words: quote(c.detail.inserted) }));
      if (c.detail?.deleted?.length) parts.push(t('Deleted {words}', { words: quote(c.detail.deleted) }));
      return parts.join('; ') || t('Text changed');
    }
    case 'moved': return t('Moved');
    case 'resized': return t('Resized');
    case 'restyled': return t('Formatting changed');
    case 'picture': return t('Picture changed');
    case 'shapeAdded': return c.detail?.words ? t('Inserted {name}: {words}', { name: c.name || t('a shape'), words: c.detail.words }) : t('Inserted {name}', { name: c.name || t('a shape') });
    case 'shapeDeleted': return t('Deleted {name}', { name: c.name || t('a shape') });
    case 'slideAdded': return c.name ? t('Slide inserted: {title}', { title: c.name }) : t('Slide inserted');
    case 'slideDeleted': return c.name ? t('Slide deleted: {title}', { title: c.name }) : t('Slide deleted');
    case 'slideMoved': return t('Slide moved from {from} to {to}', { from: c.detail.from + 1, to: c.detail.to + 1 });
    default: return t('Changed');
  }
}

/** The changes on one slide, by the shape they are about, in the order of the shapes. */
function byShape(changes) {
  const groups = new Map();
  for (const c of changes) {
    const key = c.kind === 'shapeAdded' ? 'added:' + c.reviewedShapeId : 'shape:' + c.shapeId;
    if (!groups.has(key)) groups.set(key, { key, name: c.name, list: [] });
    groups.get(key).list.push(c);
  }
  return [...groups.values()];
}

function ChangeBox({ c, onDecide }) {
  return (
    <label className={`sl-rev-item${c.accepted ? ' on' : ''}`} data-change={c.id} data-kind={c.kind}>
      <input type="checkbox" className="sl-rev-tick" checked={Boolean(c.accepted)} onChange={(e) => onDecide([c.id], e.target.checked)} />
      <span>{changeWords(c)}</span>
    </label>
  );
}

/**
 * The Revisions pane: Slide Changes and Presentation Changes. `review` is the
 * model's: `{ reviewer, name, changes }` with each change's `accepted` and
 * `slide` (where it is now).
 */
export function RevisionsPane({ review, slideIndex, slideId, onDecide, onGo, onEnd }) {
  const changes = review?.changes || [];
  const here = changes.filter((c) => c.scope === 'shape' && String(c.slideId) === String(slideId));
  const deck = changes.filter((c) => c.scope === 'slide');
  const slidesWith = [...new Set(changes.filter((c) => c.scope === 'shape').map((c) => c.slide))].filter((i) => i >= 0).sort((a, b) => a - b);
  return (
    <div className="sl-rev">
      <div className="sl-rev-by"><Icon name="comment" size={13} />{review?.reviewer ? t('Changes by {person}', { person: review.reviewer }) : t('Changes from {file}', { file: review?.name || '' })}</div>
      <h4>{t('Slide Changes')}</h4>
      {here.length ? byShape(here).map((g) => (
        <div key={g.key} className="sl-rev-group">
          <div className="sl-rev-shape">{g.name || t('Shape')}</div>
          {g.list.map((c) => <ChangeBox key={c.id} c={c} onDecide={onDecide} />)}
        </div>
      )) : <div className="sl-rev-none">{slidesWith.length ? t('No changes on this slide.') : t('No changes to the slides\' contents.')}</div>}
      {slidesWith.length && !here.length ? (
        <div className="sl-rev-jump">{t('Slides with changes:')} {slidesWith.map((i) => <button key={i} type="button" className="sl-rev-goto" onClick={() => onGo(i)}>{i + 1}</button>)}</div>
      ) : null}
      <h4>{t('Presentation Changes')}</h4>
      {deck.length ? deck.map((c) => (
        <div key={c.id} className="sl-rev-group">
          <ChangeBox c={c} onDecide={onDecide} />
          {c.slide >= 0 ? <button type="button" className="sl-rev-goto wide" onClick={() => onGo(c.slide)}>{t('Go to slide {number}', { number: c.slide + 1 })}</button> : null}
        </div>
      )) : <div className="sl-rev-none">{t('No slides were added, deleted or moved.')}</div>}
      <div className="sl-rev-foot">
        <span>{tn(changes.filter((c) => c.accepted).length, '{count} accepted', '{count} accepted')} · {tn(changes.length, '{count} change', '{count} changes')}</span>
        <Button className="sl-rev-end" icon="close" label={t('End Review')} onClick={onEnd} />
      </div>
    </div>
  );
}

/**
 * The markers on the slide: one by each changed shape (or, for a shape the
 * reviewer added, where it would go, outlined), its changes listed when
 * pressed, each with a box to tick.
 */
export function ReviewMarkers({ review, slideId, shapes, scale = 1, onDecide }) {
  const [open, setOpen] = useState(null);
  const here = (review?.changes || []).filter((c) => c.scope === 'shape' && String(c.slideId) === String(slideId));
  if (!here.length) return null;
  const groups = byShape(here);
  const boxOf = (g) => {
    const c = g.list[0];
    if (c.kind === 'shapeAdded') {
      const b = c.detail?.box;
      return b ? { x: b.x / EMU_PX, y: b.y / EMU_PX, w: b.w / EMU_PX, h: b.h / EMU_PX, ghost: !c.accepted, words: c.detail?.words || '' } : null;
    }
    const s = shapes.find((x) => String(x.id) === String(c.shapeId));
    return s?.geometry ? { x: s.geometry.x, y: s.geometry.y, w: s.geometry.w, h: s.geometry.h } : (c.detail?.box ? { x: c.detail.box.x / EMU_PX, y: c.detail.box.y / EMU_PX, w: c.detail.box.w / EMU_PX, h: c.detail.box.h / EMU_PX } : null);
  };
  const inv = 1 / (scale || 1);
  return (
    <div className="sl-rev-layer">
      {groups.map((g) => {
        const b = boxOf(g);
        if (!b) return null;
        const pending = g.list.some((c) => !c.accepted);
        return (
          <React.Fragment key={g.key}>
            {b.ghost ? <div className="sl-rev-ghost" style={{ left: b.x, top: b.y, width: b.w, height: b.h, borderWidth: 1.5 * inv }}>{b.words ? <span className="sl-rev-ghost-words">{b.words}</span> : null}</div> : null}
            <button
              type="button"
              className={`sl-rev-marker${pending ? '' : ' done'}`}
              data-marker={g.key}
              style={{ left: b.x - 26 * inv, top: b.y, transform: `scale(${inv})`, transformOrigin: 'top left' }}
              data-tip={g.list.map(changeWords).join('; ')}
              onMouseDown={(e) => e.stopPropagation()}
              onClick={(e) => { e.stopPropagation(); setOpen(open === g.key ? null : g.key); }}
            >
              <Icon name="pen" size={12} />
            </button>
            {open === g.key ? (
              <div className="sl-rev-pop" style={{ left: b.x, top: b.y + 4 * inv, transform: `scale(${inv})`, transformOrigin: 'top left' }} onMouseDown={(e) => e.stopPropagation()}>
                <div className="sl-rev-pop-head">{g.name || t('Shape')}</div>
                <label className="sl-rev-item all">
                  <input type="checkbox" className="sl-rev-all" checked={g.list.every((c) => c.accepted)} onChange={(e) => onDecide(g.list.map((c) => c.id), e.target.checked)} />
                  <span>{t('All changes to {name}', { name: g.name || t('this shape') })}</span>
                </label>
                {g.list.map((c) => <ChangeBox key={c.id} c={c} onDecide={onDecide} />)}
              </div>
            ) : null}
          </React.Fragment>
        );
      })}
    </div>
  );
}

export const REVISIONS_CSS = `
.sl-rev { display: flex; flex-direction: column; gap: 6px; padding: 8px 10px 12px; font-size: 12.5px; }
.sl-rev h4 { margin: 10px 0 2px; font-size: 11.5px; font-weight: 600; letter-spacing: 0.02em; text-transform: uppercase; color: var(--ink-3); }
.sl-rev-by { display: flex; align-items: center; gap: 6px; padding: 6px 8px; border-radius: 6px; background: var(--surface-2); color: var(--ink-2); }
.sl-rev-group { border: 1px solid var(--line); border-radius: 7px; padding: 6px 8px; display: flex; flex-direction: column; gap: 2px; }
.sl-rev-shape { font-weight: 600; font-size: 12px; margin-bottom: 2px; }
.sl-rev-item { display: flex; align-items: flex-start; gap: 7px; padding: 3px 0; cursor: pointer; line-height: 1.4; }
.sl-rev-item input { margin-top: 2px; flex: none; }
.sl-rev-item.on span { color: var(--ink-2); }
.sl-rev-none { color: var(--ink-3); padding: 4px 2px; }
.sl-rev-jump { display: flex; flex-wrap: wrap; gap: 4px; align-items: center; color: var(--ink-2); }
.sl-rev-goto { border: 1px solid var(--line); background: var(--surface); border-radius: 5px; padding: 1px 7px; font: inherit; color: var(--accent); cursor: pointer; }
.sl-rev-goto.wide { align-self: flex-start; margin-top: 2px; font-size: 11.5px; }
.sl-rev-foot { display: flex; align-items: center; justify-content: space-between; gap: 8px; margin-top: 10px; padding-top: 8px; border-top: 1px solid var(--line-soft); color: var(--ink-2); font-size: 11.5px; }
.sl-rev-layer { position: absolute; inset: 0; pointer-events: none; z-index: 40; }
.sl-rev-marker { pointer-events: auto; position: absolute; width: 22px; height: 22px; border-radius: 4px 4px 4px 0; border: 1px solid #b4009e; background: #fdf0fb; color: #b4009e; display: grid; place-items: center; cursor: pointer; box-shadow: 0 1px 3px rgba(0,0,0,.18); padding: 0; }
.sl-rev-marker.done { border-color: #8a8a8a; background: #f2f2f2; color: #6a6a6a; }
.sl-rev-ghost { position: absolute; border: 1.5px dashed #b4009e; border-radius: 2px; background: rgba(180, 0, 158, 0.05); overflow: hidden; }
.sl-rev-ghost-words { display: block; padding: 6px 10px; font: 18px/1.3 var(--font, "Segoe UI", system-ui, sans-serif); color: #b4009e; opacity: .7; }
.sl-rev-pop { pointer-events: auto; position: absolute; min-width: 240px; max-width: 340px; background: var(--surface, #fff); border: 1px solid var(--line, #c8ccd2); border-radius: 8px; box-shadow: 0 8px 24px rgba(15,23,42,.2); padding: 8px 10px; font: 12.5px/1.35 var(--font, "Segoe UI", system-ui, sans-serif); color: var(--ink, #1b1f24); margin-top: 22px; }
.sl-rev-pop-head { font-weight: 600; margin-bottom: 4px; }
.sl-rev-item.all { border-bottom: 1px solid var(--line-soft, #e3e6ea); padding-bottom: 5px; margin-bottom: 3px; font-weight: 600; }
.sl-thumb-rev { display: inline-block; width: 8px; height: 8px; border-radius: 2px; background: #b4009e; margin-top: 3px; }
`;
