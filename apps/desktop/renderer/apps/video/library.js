// Video's library (the 2.7 plan: "library and panels: media bin, titles,
// word art, shapes, icons, sound effects"): a panel beside the preview with
// the project's own media and everything Video can add, a click putting it on
// the timeline at the playhead. Media is added with a double-click or its
// buttons, as one click there picks it to look at.
//
// Beside media, titles, shapes, icons and word art it adds the other kinds of
// layer the drawing engine has (D-2026-11-20-19): a paragraph or a list, a
// chart, a colour look over the tracks under it, a QR code, the shade along
// the bottom, a progress bar, a fade at the two ends and an outro card.
//
// The Sounds tab holds the audio library (D-2026-11-22-10): a list of sound
// and video files added from this computer and of addresses that play from
// the web, and nothing more. An entry goes onto the timeline as sound. The
// ten built-in effects stay under it.

import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Button, Icon, t } from '@rutba/office-ui';
import { timecode } from '@rutba/media/timeline';
import { TITLE_PRESETS, OVERLAYS } from '@rutba/studio/project';
import { paintShapeTile, ICONS, WORDART_PRESETS, CHART_KINDS, ADJUST_PRESETS, buildPlan, paintFrame } from '@rutba/studio/host';
import { insertCatalogFor } from '@rutba/studio/insert-catalog';
import { SOUND_EFFECTS, effectSamples } from '@rutba/studio/effects';
import { SHAPE_GROUP_NAMES, SHAPE_NAMES, ICON_GROUP_NAMES, ICON_NAMES, WORDART_NAMES } from './library-words.js';
import { TitleTile, TITLE_NAMES } from './ribbon.js';
import { CHART_NAMES, ADJUST_NAMES, LIST_NAMES, OVERLAY_NAMES, KIND_ICONS } from './layer-names.js';

export const LIBRARY_TABS = ['media', 'titles', 'shapes', 'icons', 'charts', 'overlays', 'sounds'];

const TAB_NAMES = () => ({ media: t('common.media'), titles: t('video.titles'), shapes: t('common.shapes'), icons: t('common.icons'), charts: t('common.charts'), overlays: t('video.library.overlays'), sounds: t('video.library.sounds') });
const TAB_ICONS = { media: 'video', titles: 'textbox', shapes: 'shape', icons: 'star', charts: 'chart', overlays: 'layers', sounds: 'volume' };

/** What each new kind of layer is when the library adds it; a function, as its sample words are in the reader's language. */
export const NEW_LAYERS = {
  textblock: (list) => ({
    type: 'textblock', list, fx: 0.5, fy: 0.55, fw: 0.66, sizeFrac: 0.035, color: '#ffffff',
    text: list === 'none' ? t('common.your-words-here') : [t('video.library.first-point'), t('video.library.second-point'), t('video.library.third-point')].join('\n'),
  }),
  chart: (chart) => ({
    type: 'chart', chart, fx: 0.5, fy: 0.5, fw: 0.6, fh: 0.5,
    ...(chart === 'pie' || chart === 'donut'
      ? { labels: [t('image.insert.ours'), t('image.insert.theirs'), t('common.other')], values: [58, 27, 15] }
      : { labels: [t('image.insert.mon'), t('image.insert.tue'), t('image.insert.wed'), t('image.insert.thu')], values: [12, 19, 9, 24] }),
  }),
  adjust: (preset) => ({ type: 'adjust', preset }),
  overlay: (key) => ({ ...OVERLAYS[key] }),
};

/** The overlays that are added over the whole project, and the one added after it; the rest go at the playhead. */
const WHOLE_PROJECT = ['progress', 'edges'];
const AFTER_PROJECT = ['outro'];

/** A band of colour a look is seen on, in its tile. */
const LOOK_SAMPLE = { id: 'sample', type: 'shape', geometry: 'rect', fx: 0.5, fy: 0.5, fw: 1, fh: 1, fill: { kind: 'linear', angle: 0, stops: [[0, '#e0393e'], [0.35, '#f5d90a'], [0.65, '#2e9e4f'], [1, '#2f7de1']] } };

export const EFFECT_NAMES = () => ({
  whoosh: t('video.library.whoosh'), swoosh: t('video.library.swoosh'), pop: t('video.pop'), click: t('video.library.click'), chime: t('video.library.chime'),
  ding: t('video.library.ding'), rise: t('video.rise'), drop: t('video.library.drop'), beep: t('video.library.beep'), heartbeat: t('video.library.heartbeat'),
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

/** Any layers drawn by the studio's own painter on a small dark frame, so the tile is what it adds. */
function LayerTile({ patches }) {
  const ref = useRef(null);
  const key = JSON.stringify(patches);
  useEffect(() => {
    const c = ref.current;
    if (!c) return;
    try {
      const plan = buildPlan({ canvas: c, images: [], options: { aspect: 'landscape', frame: { width: 192, height: 108 } }, layerPatches: JSON.parse(key), videos: {}, assets: {} });
      paintFrame(c.getContext('2d'), plan, 0);
    } catch { /* a layer this painter does not know draws nothing */ }
  }, [key]);
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

/** A decoded sound played from its start; answers what stops it. `onEnd` is told when it ends by itself. */
function play(buffer, onEnd) {
  try {
    previewContext = previewContext || new AudioContext();
    const src = previewContext.createBufferSource();
    src.buffer = buffer;
    src.connect(previewContext.destination);
    src.onended = () => onEnd?.();
    src.start();
    return () => { src.onended = null; try { src.stop(); } catch { /* ended already */ } };
  } catch {
    onEnd?.();
    return () => {};
  }
}

const ADDRESS_ERRORS = () => ({
  'not-an-address': t('video.library.not-web-address'),
  'not-https': t('video.library.address-must-start'),
  listed: t('video.library.address-list-already'),
});

/**
 * The audio library: the list, the two ways to add to it, and each entry
 * listened to, put on the timeline as sound, or taken off the list.
 */
function AudioLibrary({ audio, has }) {
  const [address, setAddress] = useState('');
  const [problem, setProblem] = useState('');
  const [playing, setPlaying] = useState(null);
  const [working, setWorking] = useState(null);
  const stop = useRef(null);
  const hush = () => { stop.current?.(); stop.current = null; setPlaying(null); };
  useEffect(() => () => { stop.current?.(); }, []);
  const entries = audio.list.filter((e) => has(e.name));
  const fetching = audio.fetching || [];

  const addAddress = () => {
    if (!address.trim()) return;
    const error = audio.onAddAddress(address);
    setProblem(error ? ADDRESS_ERRORS()[error] || error : '');
    if (!error) setAddress('');
  };
  const listen = async (entry) => {
    const was = playing;
    hush();
    if (was === entry.id) return;
    setWorking(entry.id);
    try {
      const buffer = await audio.onSound(entry);
      if (buffer) {
        stop.current = play(buffer, () => { stop.current = null; setPlaying(null); });
        setPlaying(entry.id);
      }
    } finally {
      setWorking(null);
    }
  };
  const add = async (entry) => {
    setWorking(entry.id);
    try { await audio.onAdd(entry); } finally { setWorking(null); }
  };

  return (
    <div className="vd-lib-group vd-audio-library">
      <h4>{t('video.library.audio-library')}</h4>
      <div className="vd-lib-row-actions">
        <Button icon="folder" label={t('video.library.add-files-computer')} className="vd-audio-add-files" onClick={audio.onAddFiles} />
      </div>
      <form className="vd-audio-address" onSubmit={(e) => { e.preventDefault(); addAddress(); }}>
        <input type="text" inputMode="url" className="vd-audio-address-box" placeholder="https://" aria-label={t('video.library.web-address-sound')} value={address} onChange={(e) => { setAddress(e.target.value); setProblem(''); }} />
        <Button icon="link" label="" title={t('video.library.add-address-list')} className="vd-audio-add-address" disabled={!address.trim()} onClick={addAddress} />
      </form>
      {problem ? <p className="vd-lib-problem" role="alert">{problem}</p> : null}
      <p className="vd-audio-rights">{t('video.library.use-only-music')}</p>
      {entries.length ? (
        <ul className="vd-lib-sounds vd-audio-list" aria-label={t('video.library.audio-library')}>
          {entries.map((e) => (fetching.includes(e.id) ? (
            // An address being fetched says so, and can be stopped: a long file is a long wait.
            <li key={e.id} className="vd-lib-sound vd-audio-entry fetching" data-kind={e.kind} title={e.url}>
              <Icon name="globe" size={13} />
              <span className="vd-lib-name">{e.name}</span>
              <span className="vd-audio-fetching" role="status">{t('video.library.fetching')}</span>
              <Button label={t('common.cancel')} className="vd-audio-cancel" onClick={() => audio.onCancel(e.id)} />
            </li>
          ) : (
            <li key={e.id} className={`vd-lib-sound vd-audio-entry${working === e.id ? ' working' : ''}`} data-kind={e.kind} title={e.kind === 'remote' ? e.url : e.path}>
              <Button icon={playing === e.id ? 'stop' : 'play'} label="" title={playing === e.id ? t('common.stop') : t('video.library.listen')} className="vd-lib-listen" disabled={working === e.id} onClick={() => listen(e)} />
              <Icon name={e.kind === 'remote' ? 'globe' : 'volume'} size={13} />
              <span className="vd-lib-name">{e.name}</span>
              <Button icon="plus" label="" title={t('video.library.add-playhead')} className="vd-lib-add-sound" disabled={working === e.id} onClick={() => add(e)} />
              <Button icon="trash" label="" title={t('common.remove-list')} className="vd-audio-remove" onClick={() => { if (playing === e.id) hush(); audio.onRemove(e.id); }} />
            </li>
          )))}
        </ul>
      ) : (
        <p className="vd-lib-empty">{audio.list.length ? t('video.library.nothing-name') : t('video.library.sound-video-files')}</p>
      )}
    </div>
  );
}

/**
 * @param {{ doc: object, media: object, tab: string, setTab: (k: string) => void,
 *   onAddSource: (url: string, where: string) => void, onAddFiles: () => void,
 *   onTitle: (preset: string) => void, onGraphic: (patch: object) => void,
 *   onEffect: (key: string) => void,
 *   audio: { list: object[], fetching: string[], onAddFiles: () => void, onAddAddress: (address: string) => string | null,
 *     onSound: (entry: object) => Promise<AudioBuffer | null>, onAdd: (entry: object) => Promise<void>, onRemove: (id: string) => void,
 *     onCancel: (id: string) => void } }} p   `fetching` holds the ids of the entries whose address is being fetched; onCancel stops one
 */
export default function Library({ doc, media, tab, setTab, onAddSource, onAddFiles, onTitle, onGraphic, onEffect, onQuick = null, audio = null }) {
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
          <Button icon="plus" label={t('video.add-media')} className="vd-lib-add-files" onClick={onAddFiles} />
        </div>
        {shown.length ? (
          <ul className="vd-lib-media" aria-label={t('video.library.media-project')}>
            {shown.map(([url, s]) => (
              // The row is a button to pick (a double-click adds), with its two actions beside it rather than inside it.
              <li key={url} className={`vd-lib-source${picked === url ? ' on' : ''}`} data-url={url}>
                <button
                  type="button"
                  className="vd-lib-source-pick"
                  aria-pressed={picked === url}
                  data-tip={t('video.library.name-double-click', { name: s.name || url })}
                  onClick={() => setPicked(url)}
                  onDoubleClick={() => onAddSource(url, 'end')}
                >
                  <SourceThumb media={media} url={url} kind={s.kind} />
                  <span className="vd-lib-source-text">
                    <span className="vd-lib-name">{s.name || url}</span>
                    <span className="vd-lib-meta">{s.kind === 'image' ? t('common.picture') : s.duration ? timecode(s.duration) : ''}</span>
                  </span>
                </button>
                <span className="vd-lib-source-actions">
                  <Button icon="plus" label="" title={t('video.library.add-after-clips')} className="vd-lib-add" onClick={() => onAddSource(url, 'end')} />
                  <Button icon="move" label="" title={t('video.library.insert-playhead')} className="vd-lib-insert" onClick={() => onAddSource(url, 'insert')} />
                </span>
              </li>
            ))}
          </ul>
        ) : (
          <p className="vd-lib-empty">{q ? t('video.library.nothing-name') : t('video.library.videos-sound-pictures')}</p>
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
            <button key={p} type="button" className="vd-lib-tile vd-lib-title" data-key={p} data-tip={t('video.library.name-added-playhead', { name: names[p] || p })} onClick={() => onTitle(p)}>
              <TitleTile preset={p} />
              <span className="vd-lib-name">{names[p] || p}</span>
            </button>
          ))}
        </div>
        {onQuick ? (() => {
          // The caption layer and the quick adds of the social app's video tool (S17).
          const quick = [
            ['caption', 'captions', t('common.caption'), t('video.library.words-typed-over')],
            ['pill', 'textbox', t('video.library.words-pill'), t('video.library.short-word-coloured')],
            ['qr', 'qrcode', t('common.qr-code'), t('video.library.qr-code-type')],
            ['logo', 'picture', t('common.logo'), t('video.library.logo-brand-kit')],
            ['footer', 'minus', t('common.footer'), t('video.library.small-line-along')],
          ].filter(([, , name]) => has(name));
          if (!quick.length) return null;
          return (
            <div className="vd-lib-group">
              <h4>{t('video.library.captions-quick-adds')}</h4>
              <ul className="vd-lib-sounds vd-lib-overlays vd-lib-quick">
                {quick.map(([k, icon, name, tip]) => (
                  <li key={k} className="vd-lib-sound">
                    <button type="button" className="vd-lib-overlay" data-key={`quick-${k}`} data-tip={tip} onClick={() => onQuick(k)}>
                      <Icon name={icon} size={16} />
                      <span className="vd-lib-name">{name}</span>
                      <Icon name="plus" size={13} />
                    </button>
                  </li>
                ))}
              </ul>
            </div>
          );
        })() : null}
        {arts.length ? (
          <div className="vd-lib-group">
            <h4>{t('video.word-art')}</h4>
            <div className="vd-lib-grid wide">
              {arts.map((k) => (
                <button key={k} type="button" className="vd-lib-tile vd-lib-wordart-tile" data-key={k} aria-label={styles[k] || k} data-tip={t('video.library.name-added-playhead', { name: styles[k] || k })} onClick={() => onGraphic({ type: 'wordart', text: t('video.library.your-words'), preset: k, fx: 0.5, fy: 0.3, sizeFrac: 0.1 })}>
                  <WordArtTile preset={k} />
                  <span className="vd-lib-name">{styles[k] || k}</span>
                </button>
              ))}
            </div>
          </div>
        ) : null}
        {(() => {
          const lists = LIST_NAMES();
          const kinds = Object.keys(lists).filter((k) => has(lists[k]));
          if (!kinds.length) return null;
          return (
            <div className="vd-lib-group">
              <h4>{t('video.library.paragraphs-and-lists')}</h4>
              <div className="vd-lib-grid wide">
                {kinds.map((k) => (
                  <button key={k} type="button" className="vd-lib-tile vd-lib-layer-tile" data-key={`textblock-${k}`} aria-label={lists[k]} data-tip={t('video.library.name-added-playhead', { name: lists[k] })} onClick={() => onGraphic(NEW_LAYERS.textblock(k))}>
                    <LayerTile patches={[{ ...NEW_LAYERS.textblock(k), id: 'tile', fy: 0.5, fw: 0.86, sizeFrac: 0.13 }]} />
                    <span className="vd-lib-name">{lists[k]}</span>
                  </button>
                ))}
              </div>
            </div>
          );
        })()}
      </>
    );
  } else if (tab === 'charts') {
    const names = CHART_NAMES();
    body = (
      <div className="vd-lib-grid wide">
        {CHART_KINDS.filter((k) => has(names[k] || k)).map((k) => (
          <button key={k} type="button" className="vd-lib-tile vd-lib-layer-tile" data-key={`chart-${k}`} aria-label={names[k] || k} data-tip={t('video.library.name-added-playhead', { name: names[k] || k })} onClick={() => onGraphic(NEW_LAYERS.chart(k))}>
            <LayerTile patches={[{ ...NEW_LAYERS.chart(k), id: 'tile', fw: 0.92, fh: 0.86, showLegend: false, showLabels: false }]} />
            <span className="vd-lib-name">{names[k] || k}</span>
          </button>
        ))}
      </div>
    );
  } else if (tab === 'overlays') {
    const looks = ADJUST_NAMES();
    const overlays = OVERLAY_NAMES();
    const lookKeys = Object.keys(ADJUST_PRESETS).filter((k) => k !== 'none' && has(looks[k] || k));
    const overlayKeys = Object.keys(OVERLAYS).filter((k) => has(overlays[k] || k));
    const tipOf = (k, name) => (WHOLE_PROJECT.includes(k) ? t('video.library.name-added-over', { name }) : AFTER_PROJECT.includes(k) ? t('video.library.name-added-after', { name }) : t('video.library.name-added-playhead', { name }));
    body = (
      <>
        {overlayKeys.length ? (
          <div className="vd-lib-group">
            <h4>{t('video.library.on-the-frame')}</h4>
            <ul className="vd-lib-sounds vd-lib-overlays">
              {overlayKeys.map((k) => (
                <li key={k} className="vd-lib-sound">
                  <button type="button" className="vd-lib-overlay" data-key={`overlay-${k}`} data-tip={tipOf(k, overlays[k])} onClick={() => onGraphic(NEW_LAYERS.overlay(k))}>
                    <Icon name={KIND_ICONS[k]} size={16} />
                    <span className="vd-lib-name">{overlays[k]}</span>
                    <Icon name="plus" size={13} />
                  </button>
                </li>
              ))}
            </ul>
          </div>
        ) : null}
        {lookKeys.length ? (
          <div className="vd-lib-group">
            <h4>{t('video.library.colour-looks')}</h4>
            <p className="vd-lib-empty">{t('video.look-changes-colour')}</p>
            <div className="vd-lib-grid wide">
              {lookKeys.map((k) => (
                <button key={k} type="button" className="vd-lib-tile vd-lib-layer-tile" data-key={`adjust-${k}`} aria-label={looks[k] || k} data-tip={t('video.library.name-added-playhead', { name: looks[k] || k })} onClick={() => onGraphic(NEW_LAYERS.adjust(k))}>
                  <LayerTile patches={[LOOK_SAMPLE, { ...NEW_LAYERS.adjust(k), id: 'tile' }]} />
                  <span className="vd-lib-name">{looks[k] || k}</span>
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
                <button key={it.id} type="button" className={`vd-lib-tile vd-lib-${tab === 'shapes' ? 'shape-tile' : 'icon-tile'}`} data-key={keyOf(it)} aria-label={name} data-tip={t('video.library.name-added-playhead', { name })} onClick={() => onGraphic(tab === 'icons' ? { ...it.patch, label: name, color: '#ffffff' } : { ...it.patch, name, color: '#ffffff' })}>
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
      <>
        {audio ? <AudioLibrary audio={audio} has={has} /> : null}
        <div className="vd-lib-group">
          <h4>{t('video.library.sound-effects')}</h4>
          <ul className="vd-lib-sounds vd-lib-effects">
            {SOUND_EFFECTS.filter((e) => has(names[e.key] || e.key)).map((e) => (
              <li key={e.key} className="vd-lib-sound" data-key={e.key}>
                <Button icon="play" label="" title={t('video.library.listen')} className="vd-lib-listen" onClick={() => hear(e.key)} />
                <span className="vd-lib-name">{names[e.key] || e.key}</span>
                <span className="vd-lib-meta">{t('common.seconds-s', { seconds: e.seconds })}</span>
                <Button icon="plus" label="" title={t('video.library.add-playhead')} className="vd-lib-add-sound" onClick={() => onEffect(e.key)} />
              </li>
            ))}
          </ul>
        </div>
      </>
    );
  }

  return (
    <div className="vd-library" data-tab={tab}>
      <div className="vd-lib-tabs" role="tablist" aria-label={t('video.library')}>
        {LIBRARY_TABS.map((k) => (
          <button key={k} type="button" role="tab" aria-selected={tab === k} className={tab === k ? 'on' : ''} data-tab={k} data-tip={tabNames[k]} aria-label={tabNames[k]} onClick={() => setTab(k)}>
            <Icon name={TAB_ICONS[k]} size={16} />
          </button>
        ))}
      </div>
      <h3 className="vd-lib-title-bar">{tabNames[tab]}</h3>
      <input type="search" className="vd-lib-search" placeholder={t('common.search')} aria-label={t('video.library.search-the-library')} value={query} onChange={(e) => setQuery(e.target.value)} />
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
.vd-lib-source-pick { flex: 1; min-width: 0; display: flex; align-items: center; gap: 8px; padding: 0; border: 0; background: none; color: inherit; font: inherit; text-align: start; cursor: default; }
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
.vd-lib-layer-tile { padding: 6px 4px; }
.vd-lib-overlay { flex: 1; min-width: 0; display: flex; align-items: center; gap: 8px; padding: 5px 4px; border: 0; border-radius: 6px; background: none; color: inherit; font: inherit; text-align: start; cursor: pointer; }
.vd-lib-overlay:hover { color: var(--accent); }
.vd-audio-address { display: flex; gap: 4px; margin-bottom: 6px; }
.vd-audio-address-box { flex: 1; min-width: 0; padding: 5px 8px; border: 1px solid var(--line-strong); border-radius: 6px; background: var(--surface); color: inherit; font: inherit; }
.vd-lib-problem { margin: 0 2px 6px; color: var(--bad, #c62828); line-height: 1.4; }
.vd-audio-entry.working { opacity: 0.6; }
.vd-audio-entry .vd-lib-name { min-width: 0; }
.vd-audio-entry.fetching { padding-block: 4px; }
.vd-audio-fetching { flex: none; color: var(--ink-3); }
.vd-audio-rights { margin: 0 2px 8px; color: var(--ink-3); font-size: 11.5px; line-height: 1.4; }
.vd-audio-entry > svg { flex: none; color: var(--ink-3); }
.vd-panel-rail { width: 30px; flex: none; display: flex; flex-direction: column; align-items: center; padding-top: 6px; gap: 4px; background: var(--sunken); border-inline: 1px solid var(--line-soft); }
.vd-panel-rail button { writing-mode: vertical-rl; display: flex; align-items: center; gap: 6px; padding: 8px 4px; border: 0; border-radius: 6px; background: transparent; color: var(--ink-2); cursor: pointer; font: inherit; font-size: 12px; }
.vd-panel-rail button:hover { background: var(--hover); color: var(--accent); }
`;
