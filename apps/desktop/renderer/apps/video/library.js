// Video's library (the 2.7 plan: "library and panels: media bin, titles,
// word art, shapes, icons, sound effects"): a panel beside the preview with
// the project's own media and everything Video can add, a click putting it on
// the timeline at the playhead. Media is added with a double-click or its
// buttons, as one click there picks it to look at.

import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Button, Icon, t } from '@rutba/office-ui';
import { timecode } from '@rutba/media/timeline';
import { TITLE_PRESETS } from '@rutba/studio/project';
import { paintShapeTile, ICONS, WORDART_PRESETS, buildPlan, paintFrame } from '@rutba/studio/host';
import { insertCatalogFor } from '@rutba/studio/insert-catalog';
import { SOUND_EFFECTS, effectSamples } from '@rutba/studio/effects';
import { SHAPE_GROUP_NAMES, SHAPE_NAMES, ICON_GROUP_NAMES, ICON_NAMES, WORDART_NAMES } from './library-words.js';
import { TitleTile, TITLE_NAMES } from './ribbon.js';

export const LIBRARY_TABS = ['media', 'titles', 'shapes', 'icons', 'sounds'];

const TAB_NAMES = () => ({ media: t('Media'), titles: t('Titles'), shapes: t('Shapes'), icons: t('Icons'), sounds: t('Sounds') });
const TAB_ICONS = { media: 'video', titles: 'textbox', shapes: 'shape', icons: 'star', sounds: 'volume' };

export const EFFECT_NAMES = () => ({
  whoosh: t('Whoosh'), swoosh: t('Swoosh'), pop: t('Pop'), click: t('Click'), chime: t('Chime'),
  ding: t('Ding'), rise: t('Rise'), drop: t('Drop'), beep: t('Beep'), heartbeat: t('Heartbeat'),
});

/** A shape drawn by the studio's own painter, so the tile is what it adds. */
function ShapeTile({ geometry }) {
  const ref = useRef(null);
  useEffect(() => {
    const c = ref.current;
    if (!c) return;
    const g = c.getContext('2d');
    g.clearRect(0, 0, c.width, c.height);
    g.save();
    g.translate(c.width / 2, c.height / 2);
    g.fillStyle = getComputedStyle(c).color || '#5b8def';
    try { paintShapeTile(g, geometry, c.width * 0.7, c.height * 0.7); } catch { /* a geometry this painter does not know */ }
    g.restore();
  }, [geometry]);
  return <canvas ref={ref} className="vd-lib-shape" width={56} height={40} aria-hidden="true" />;
}

/** A word art style drawn by the studio's own painter on a small dark frame. */
function WordArtTile({ preset }) {
  const ref = useRef(null);
  useEffect(() => {
    const c = ref.current;
    if (!c) return;
    try {
      const plan = buildPlan({ canvas: c, images: [], options: { aspect: 'landscape', frame: { width: 192, height: 108 } }, layerPatches: [{ id: 'w', type: 'wordart', text: 'Abc', preset, fx: 0.5, fy: 0.55, sizeFrac: 0.26 }], videos: {}, assets: {} });
      paintFrame(c.getContext('2d'), plan, 0);
    } catch { /* a style this painter does not know draws nothing */ }
  }, [preset]);
  return <canvas ref={ref} className="vd-lib-wordart" width={192} height={108} aria-hidden="true" />;
}

function IconTile({ name }) {
  return (
    <svg className="vd-lib-icon" viewBox="0 0 24 24" width="28" height="28" aria-hidden="true">
      <path d={ICONS[name]} fill="currentColor" />
    </svg>
  );
}

/** The first thumbnail of a source, drawn when the library has it. */
function SourceThumb({ media, url, kind }) {
  const ref = useRef(null);
  const [, redraw] = useState(0);
  useEffect(() => media.onChange((k) => { if (k === 'thumb') redraw((n) => n + 1); }), [media]);
  useEffect(() => {
    const c = ref.current;
    const frame = media.thumbs?.get(url)?.frames?.[0];
    if (!c || !frame) return;
    const g = c.getContext('2d');
    g.clearRect(0, 0, c.width, c.height);
    const s = Math.max(c.width / frame.width, c.height / frame.height);
    g.drawImage(frame, (c.width - frame.width * s) / 2, (c.height - frame.height * s) / 2, frame.width * s, frame.height * s);
  });
  if (kind === 'audio') return <span className="vd-lib-thumb sound"><Icon name="volume" size={18} /></span>;
  return <canvas ref={ref} className="vd-lib-thumb" width={64} height={36} aria-hidden="true" />;
}

let previewContext = null;
/** An effect heard once, from its own samples. */
function hear(key) {
  try {
    previewContext = previewContext || new AudioContext();
    const samples = effectSamples(key, previewContext.sampleRate);
    const buffer = previewContext.createBuffer(1, samples.length, previewContext.sampleRate);
    buffer.getChannelData(0).set(samples);
    const src = previewContext.createBufferSource();
    src.buffer = buffer;
    src.connect(previewContext.destination);
    src.start();
  } catch { /* no sound device: nothing to hear */ }
}

/**
 * @param {{ doc: object, media: object, tab: string, setTab: (k: string) => void,
 *   onAddSource: (url: string, where: string) => void, onAddFiles: () => void,
 *   onTitle: (preset: string) => void, onGraphic: (patch: object) => void,
 *   onEffect: (key: string) => void }} p
 */
export default function Library({ doc, media, tab, setTab, onAddSource, onAddFiles, onTitle, onGraphic, onEffect }) {
  const [query, setQuery] = useState('');
  const [picked, setPicked] = useState(null);
  const q = query.trim().toLowerCase();
  const has = (name) => !q || String(name).toLowerCase().includes(q);
  const catalogue = useMemo(() => insertCatalogFor('video'), []);
  const tabNames = TAB_NAMES();
  const sources = Object.entries(doc.sources || {}).filter(([url, s]) => url && s && s.kind !== 'project');

  let body = null;
  if (tab === 'media') {
    const shown = sources.filter(([url, s]) => has(s.name || url));
    body = (
      <>
        <div className="vd-lib-row-actions">
          <Button icon="plus" label={t('Add Media')} className="vd-lib-add-files" onClick={onAddFiles} />
        </div>
        {shown.length ? (
          <ul className="vd-lib-media" role="listbox" aria-label={t('Media in this project')}>
            {shown.map(([url, s]) => (
              <li
                key={url}
                role="option"
                aria-selected={picked === url}
                className={`vd-lib-source${picked === url ? ' on' : ''}`}
                data-url={url}
                data-tip={t('{name}: double-click to add it after the clips on the picked track', { name: s.name || url })}
                onClick={() => setPicked(url)}
                onDoubleClick={() => onAddSource(url, 'end')}
              >
                <SourceThumb media={media} url={url} kind={s.kind} />
                <span className="vd-lib-source-text">
                  <span className="vd-lib-name">{s.name || url}</span>
                  <span className="vd-lib-meta">{s.kind === 'image' ? t('Picture') : s.duration ? timecode(s.duration) : ''}</span>
                </span>
                <span className="vd-lib-source-actions">
                  <Button icon="plus" label="" title={t('Add after the clips on the picked track')} className="vd-lib-add" onClick={(e) => { e.stopPropagation(); onAddSource(url, 'end'); }} />
                  <Button icon="move" label="" title={t('Insert at the playhead')} className="vd-lib-insert" onClick={(e) => { e.stopPropagation(); onAddSource(url, 'insert'); }} />
                </span>
              </li>
            ))}
          </ul>
        ) : (
          <p className="vd-lib-empty">{q ? t('Nothing here by that name.') : t('The videos, sound and pictures in this project are kept here, to add again.')}</p>
        )}
      </>
    );
  } else if (tab === 'titles') {
    const names = TITLE_NAMES();
    const styles = WORDART_NAMES();
    const arts = Object.keys(WORDART_PRESETS).filter((k) => has(styles[k] || k));
    body = (
      <>
        <div className="vd-lib-grid wide">
          {Object.keys(TITLE_PRESETS).filter((p) => has(names[p] || p)).map((p) => (
            <button key={p} type="button" className="vd-lib-tile vd-lib-title" data-key={p} data-tip={t('{name}: added at the playhead', { name: names[p] || p })} onClick={() => onTitle(p)}>
              <TitleTile preset={p} />
              <span className="vd-lib-name">{names[p] || p}</span>
            </button>
          ))}
        </div>
        {arts.length ? (
          <div className="vd-lib-group">
            <h4>{t('Word art')}</h4>
            <div className="vd-lib-grid wide">
              {arts.map((k) => (
                <button key={k} type="button" className="vd-lib-tile vd-lib-wordart-tile" data-key={k} aria-label={styles[k] || k} data-tip={t('{name}: added at the playhead', { name: styles[k] || k })} onClick={() => onGraphic({ type: 'wordart', text: t('Your words'), preset: k, fx: 0.5, fy: 0.3, sizeFrac: 0.1 })}>
                  <WordArtTile preset={k} />
                  <span className="vd-lib-name">{styles[k] || k}</span>
                </button>
              ))}
            </div>
          </div>
        ) : null}
      </>
    );
  } else if (tab === 'shapes' || tab === 'icons') {
    const section = catalogue.find((s) => s.key === tab);
    const groupNames = tab === 'shapes' ? SHAPE_GROUP_NAMES() : ICON_GROUP_NAMES();
    const itemNames = tab === 'shapes' ? SHAPE_NAMES() : ICON_NAMES();
    const keyOf = (it) => (tab === 'shapes' ? it.patch.geometry : it.patch.name);
    body = (section?.groups || []).map((g) => {
      const items = g.items.filter((it) => has(itemNames[keyOf(it)] || it.label));
      if (!items.length) return null;
      return (
        <div key={g.label || 'all'} className="vd-lib-group">
          {g.label ? <h4>{groupNames[g.label] || g.label}</h4> : null}
          <div className="vd-lib-grid">
            {items.map((it) => {
              const name = itemNames[keyOf(it)] || it.label;
              return (
                <button key={it.id} type="button" className={`vd-lib-tile vd-lib-${tab === 'shapes' ? 'shape-tile' : 'icon-tile'}`} data-key={keyOf(it)} aria-label={name} data-tip={t('{name}: added at the playhead', { name })} onClick={() => onGraphic(tab === 'icons' ? { ...it.patch, label: name, color: '#ffffff' } : { ...it.patch, name, color: '#ffffff' })}>
                  {tab === 'shapes' ? <ShapeTile geometry={it.patch.geometry} /> : <IconTile name={it.patch.name} />}
                </button>
              );
            })}
          </div>
        </div>
      );
    });
  } else if (tab === 'sounds') {
    const names = EFFECT_NAMES();
    body = (
      <ul className="vd-lib-sounds">
        {SOUND_EFFECTS.filter((e) => has(names[e.key] || e.key)).map((e) => (
          <li key={e.key} className="vd-lib-sound" data-key={e.key}>
            <Button icon="play" label="" title={t('Listen')} className="vd-lib-listen" onClick={() => hear(e.key)} />
            <span className="vd-lib-name">{names[e.key] || e.key}</span>
            <span className="vd-lib-meta">{t('{seconds} s', { seconds: e.seconds })}</span>
            <Button icon="plus" label="" title={t('Add at the playhead')} className="vd-lib-add-sound" onClick={() => onEffect(e.key)} />
          </li>
        ))}
      </ul>
    );
  }

  return (
    <div className="vd-library" data-tab={tab}>
      <div className="vd-lib-tabs" role="tablist" aria-label={t('Library')}>
        {LIBRARY_TABS.map((k) => (
          <button key={k} type="button" role="tab" aria-selected={tab === k} className={tab === k ? 'on' : ''} data-tab={k} data-tip={tabNames[k]} aria-label={tabNames[k]} onClick={() => setTab(k)}>
            <Icon name={TAB_ICONS[k]} size={16} />
          </button>
        ))}
      </div>
      <h3 className="vd-lib-title-bar">{tabNames[tab]}</h3>
      <input type="search" className="vd-lib-search" placeholder={t('Search')} aria-label={t('Search the library')} value={query} onChange={(e) => setQuery(e.target.value)} />
      <div className="vd-lib-body">{body}</div>
    </div>
  );
}

export const LIBRARY_CSS = `
.vd-library { display: flex; flex-direction: column; min-height: 0; height: 100%; font-size: 12.5px; }
.vd-lib-tabs { display: flex; gap: 2px; padding: 6px 8px 4px; border-bottom: 1px solid var(--line-soft); }
.vd-lib-tabs button { flex: 1; display: grid; place-items: center; height: 30px; border: 0; border-radius: 6px; background: transparent; color: var(--ink-2); cursor: pointer; }
.vd-lib-tabs button:hover { background: var(--hover); }
.vd-lib-tabs button.on { background: var(--accent-soft); color: var(--accent); box-shadow: inset 0 0 0 1px var(--accent-line); }
.vd-lib-title-bar { margin: 8px 12px 4px; font-size: 12px; font-weight: 600; color: var(--ink-2); }
.vd-lib-search { margin: 0 10px 6px; padding: 5px 8px; border: 1px solid var(--line-strong); border-radius: 6px; background: var(--surface); color: inherit; font: inherit; }
.vd-lib-body { flex: 1; min-height: 0; overflow: auto; padding: 2px 10px 12px; }
.vd-lib-row-actions { display: flex; margin-bottom: 6px; }
.vd-lib-media, .vd-lib-sounds { list-style: none; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 2px; }
.vd-lib-source { display: flex; align-items: center; gap: 8px; padding: 4px; border-radius: 6px; cursor: default; }
.vd-lib-source:hover { background: var(--hover); }
.vd-lib-source.on { background: var(--accent-soft); box-shadow: inset 0 0 0 1px var(--accent-line); }
.vd-lib-thumb { width: 64px; height: 36px; border-radius: 4px; background: #1c1f26; flex: none; display: grid; place-items: center; color: #aab0bb; }
.vd-lib-source-text { flex: 1; min-width: 0; display: flex; flex-direction: column; }
.vd-lib-name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.vd-lib-meta { color: var(--ink-3); font-size: 11.5px; font-variant-numeric: tabular-nums; }
.vd-lib-source-actions { display: flex; opacity: 0; }
.vd-lib-source:hover .vd-lib-source-actions, .vd-lib-source.on .vd-lib-source-actions, .vd-lib-source:focus-within .vd-lib-source-actions { opacity: 1; }
.vd-lib-empty { color: var(--ink-3); line-height: 1.45; margin: 8px 2px; }
.vd-lib-group h4 { margin: 10px 2px 4px; font-size: 11.5px; font-weight: 600; color: var(--ink-3); }
.vd-lib-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(56px, 1fr)); gap: 4px; }
.vd-lib-grid.wide { grid-template-columns: repeat(auto-fill, minmax(96px, 1fr)); gap: 6px; }
.vd-lib-tile { display: grid; place-items: center; gap: 4px; min-height: 48px; padding: 4px; border: 1px solid var(--line-soft); border-radius: 6px; background: var(--surface); color: var(--ink-2); cursor: pointer; font: inherit; }
.vd-lib-tile:hover { border-color: var(--accent-line); background: var(--accent-soft); color: var(--accent); }
.vd-lib-title { padding: 8px 4px; }
.vd-lib-title .vd-title-tile { width: 72px; height: 40px; }
.vd-lib-shape { width: 56px; height: 40px; color: var(--accent); }
.vd-lib-wordart { width: 96px; height: 54px; border-radius: 4px; background: #111; }
.vd-lib-sound { display: flex; align-items: center; gap: 6px; padding: 2px 4px; border-radius: 6px; }
.vd-lib-sound:hover { background: var(--hover); }
.vd-lib-sound .vd-lib-name { flex: 1; }
.vd-panel-rail { width: 30px; flex: none; display: flex; flex-direction: column; align-items: center; padding-top: 6px; gap: 4px; background: var(--sunken); border-inline: 1px solid var(--line-soft); }
.vd-panel-rail button { writing-mode: vertical-rl; display: flex; align-items: center; gap: 6px; padding: 8px 4px; border: 0; border-radius: 6px; background: transparent; color: var(--ink-2); cursor: pointer; font: inherit; font-size: 12px; }
.vd-panel-rail button:hover { background: var(--hover); color: var(--accent); }
`;
