// The scenes of an introduction (docs/STUDIO.md section 9): the panel down
// the left of the Video window when the project is an introduction.
//
// Each scene is a card: its number, a drawing of its layout, its name and
// how long it runs. The scene picked opens out: its layout, switched with a
// press even after recording (X2); for a card, its words and its length;
// for a recorded scene, its script (the teleprompter's words, X3), Record
// or Retake, and its takes, any of which can be put back (X4); and moving
// it earlier or later, or taking it out. At the foot: Add Scene, and the
// brand kit every scene is drawn in (X6).

import React, { useState } from 'react';
import { Button, Icon, useMenu, t, tn } from '@rutba/office-ui';
import { timecode } from '@rutba/media/timeline';
import { LAYOUTS, RECORDED, sceneSpans, activeTake, scriptSeconds } from '@rutba/studio/intro';

/** Each layout by name, as the window shows it. */
export const LAYOUT_NAMES = () => ({
  title: t('Title card'),
  talking: t('Talking head'),
  screenCamera: t('Screen with camera'),
  sideBySide: t('Side by side'),
  end: t('End card'),
});

/** What each layout is for, after its name in a tip. */
export const LAYOUT_TIPS = () => ({
  title: t('your name and a line, in the brand\'s colours'),
  talking: t('the camera, the whole frame'),
  screenCamera: t('the screen, the camera in a corner'),
  sideBySide: t('the camera and the screen next to each other'),
  end: t('thanks and where to find you'),
});

/** A layout drawn small: the frame, the screen and the camera where they sit. */
export function LayoutGlyph({ layout, size = 'small' }) {
  return (
    <span className={`vi-glyph l-${layout} g-${size}`} aria-hidden="true">
      <i className="scr" />
      <i className="cam" />
      <i className="ln1" />
      <i className="ln2" />
    </span>
  );
}

export default function ScenesPanel({ doc, current, onPick, onRecord, onLayout, onMove, onRemove, onAdd, onChange, onUseTake, onBrand, speed }) {
  const names = LAYOUT_NAMES();
  const tips = LAYOUT_TIPS();
  const spans = sceneSpans(doc);
  const scenes = doc.intro.scenes;
  const addMenu = useMenu();
  const [draft, setDraft] = useState(null);

  return (
    <div className="vi-scenes" data-count={scenes.length}>
      <ol className="vi-list">
        {scenes.map((s, i) => {
          const span = spans[i];
          const take = activeTake(s);
          const open = s.id === current;
          const recorded = RECORDED.has(s.layout);
          return (
            <li
              key={s.id}
              className={`vi-scene${open ? ' open' : ''}${recorded && !take ? ' waiting' : ''}`}
              data-scene={s.id}
              data-layout={s.layout}
              data-takes={s.takes.length}
              data-take={s.take || ''}
            >
              <button type="button" className="vi-scene-head" onClick={() => onPick(s.id, span.start)} aria-expanded={open}>
                <span className="vi-num">{i + 1}</span>
                <LayoutGlyph layout={s.layout} />
                <span className="vi-scene-name">
                  <b>{names[s.layout]}</b>
                  <small>
                    {recorded ? (take ? tn(s.takes.length, '{count} take', '{count} takes') : t('Not recorded yet')) : (s.title || t('No words yet'))}
                  </small>
                </span>
                <span className="vi-len">{timecode(span.end - span.start)}</span>
              </button>
              {open ? (
                <div className="vi-scene-body">
                  <div className="vi-label">{t('Layout')}</div>
                  <div className="vi-layouts" role="radiogroup" aria-label={t('Layout')}>
                    {LAYOUTS.map((l) => (
                      <button
                        key={l}
                        type="button"
                        role="radio"
                        aria-checked={s.layout === l}
                        className={`vi-layout${s.layout === l ? ' on' : ''}`}
                        data-layout={l}
                        data-tip={`${names[l]}: ${tips[l]}`}
                        onClick={() => s.layout !== l && onLayout(s.id, l)}
                      >
                        <LayoutGlyph layout={l} />
                      </button>
                    ))}
                  </div>

                  {recorded ? (
                    <>
                      <label className="vi-label" htmlFor={`vi-script-${s.id}`}>{t('Script')}</label>
                      <textarea
                        id={`vi-script-${s.id}`}
                        className="rw-input vi-script"
                        dir="auto"
                        rows={4}
                        placeholder={t('What you will say. It scrolls above the camera while you record.')}
                        value={draft?.id === s.id ? draft.script : s.script || ''}
                        onChange={(e) => setDraft({ id: s.id, script: e.target.value })}
                        onBlur={() => { if (draft?.id === s.id && draft.script !== (s.script || '')) onChange(s.id, { script: draft.script }); setDraft(null); }}
                      />
                      {(draft?.id === s.id ? draft.script : s.script) ? (
                        <div className="vi-note">{t('About {seconds} s to read at the teleprompter\'s speed', { seconds: Math.max(1, Math.round(scriptSeconds(draft?.id === s.id ? draft.script : s.script, speed))) })}</div>
                      ) : null}
                      <Button primary icon="record" className="vi-record" label={take ? t('Retake') : t('Record')} onClick={() => { if (draft?.id === s.id && draft.script !== (s.script || '')) onChange(s.id, { script: draft.script }); setDraft(null); onRecord(s.id); }} />
                      {s.takes.length ? (
                        <div className="vi-takes">
                          {s.takes.map((k, n) => (
                            <div key={k.id} className={`vi-take${k.id === s.take ? ' on' : ''}`} data-take={k.id}>
                              <Icon name={k.id === s.take ? 'check' : 'video'} size={13} />
                              <span>{t('Take {n}', { n: n + 1 })}</span>
                              <span className="vi-take-len">{timecode(k.duration || 0)}</span>
                              {k.id === s.take ? <span className="vi-take-used">{t('On the timeline')}</span> : <Button ghost className="vi-use" label={t('Use')} onClick={() => onUseTake(s.id, k.id)} />}
                            </div>
                          ))}
                        </div>
                      ) : null}
                    </>
                  ) : (
                    <>
                      <label className="vi-label" htmlFor={`vi-title-${s.id}`}>{t('Words')}</label>
                      <input id={`vi-title-${s.id}`} className="rw-input vi-title" dir="auto" value={s.title || ''} placeholder={s.layout === 'title' ? t('Your name') : t('Thanks for watching')} onChange={(e) => onChange(s.id, { title: e.target.value }, `title-${s.id}`)} />
                      <input className="rw-input vi-subtitle" dir="auto" value={s.subtitle || ''} placeholder={s.layout === 'title' ? t('What you do, in a line') : t('Where to find you')} onChange={(e) => onChange(s.id, { subtitle: e.target.value }, `subtitle-${s.id}`)} />
                      <div className="vi-row">
                        <span className="vi-label">{t('Length')}</span>
                        <input className="rw-input vi-seconds" type="number" min={1} max={30} step={0.5} value={s.seconds} onChange={(e) => { const v = Number(e.target.value); if (v >= 1 && v <= 30) onChange(s.id, { seconds: v }, `seconds-${s.id}`); }} />
                        <span className="vi-unit">{t('seconds')}</span>
                      </div>
                    </>
                  )}
                  <div className="vi-scene-tools">
                    <Button icon="chevronUp" title={t('Move Earlier')} className="vi-up" disabled={i === 0} onClick={() => onMove(s.id, -1)} />
                    <Button icon="chevronDown" title={t('Move Later')} className="vi-down" disabled={i === scenes.length - 1} onClick={() => onMove(s.id, 1)} />
                    <span className="vi-grow" />
                    <Button icon="trash" title={t('Remove Scene')} className="vi-remove" disabled={scenes.length < 2} onClick={() => onRemove(s.id)} />
                  </div>
                </div>
              ) : null}
            </li>
          );
        })}
      </ol>
      <div className="vi-foot">
        <Button icon="plus" label={t('Add Scene')} className="vi-add" onClick={(e) => addMenu.open(e, LAYOUTS.map((l) => ({ label: names[l], icon: 'scenes', run: () => onAdd(l) })))} />
        <Button icon="brand" label={t('Brand Kit')} className="vi-brand-open" onClick={onBrand} />
      </div>
      {addMenu.node}
    </div>
  );
}
