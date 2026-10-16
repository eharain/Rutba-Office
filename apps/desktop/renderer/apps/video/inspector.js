// The properties pane: what is picked, and every setting it has.
//
// With nothing picked it is the project's own settings (the frame's shape,
// how long a still lasts, ripple and snapping), its markers with their
// names, and its sources. With a clip picked: where it sits and how long it
// runs; for a picture its fit, turn, flips and crop; for a clip with sound
// its volume and fades; the transition into it; for a title its words,
// size, colour and the animations in and out.

import React from 'react';
import { Button, Icon, Select, labelFirstControl, t, tn } from '@rutba/office-ui';
import { timecode } from '@rutba/media/timeline';
import { clipById, partnersOf, previousOnLane, lengthOf, animationName, TEXT_ANIMATIONS, TRANSITIONS, setTransition, trimEnd, moveClip, updateMarker, removeMarker } from '@rutba/studio/project';
import { insetRect, cornerOf } from '@rutba/studio/recorder';
import { pointerOf, applyAutoZoom, updateZoom, setZooms } from '@rutba/studio/pointer';
import { composeIntro } from '@rutba/studio/intro';
import { transitionName } from './timeline.js';

const COLOURS = ['#ffffff', '#111111', '#ffc107', '#e2445c', '#2bb673', '#3b82f6', '#c2408f'];

const round = (n, k = 100) => Math.round(n * k) / k;

/** A number box in seconds. */
function Seconds({ value, onChange, min = 0, step = 0.1, className }) {
  return (
    <input
      className={`rw-input vd-num${className ? ` ${className}` : ''}`}
      type="number"
      min={min}
      step={step}
      value={round(value)}
      onChange={(e) => {
        const v = Number(e.target.value);
        if (Number.isFinite(v)) onChange(v);
      }}
    />
  );
}

/** A fraction of the frame shown and typed as a percentage. */
function Percent({ value, onChange, min = -100, max = 200, className }) {
  return (
    <input
      className={`rw-input vd-num${className ? ` ${className}` : ''}`}
      type="number"
      min={min}
      max={max}
      step={0.5}
      value={Math.round((value ?? 0) * 1000) / 10}
      onChange={(e) => {
        const v = Number(e.target.value);
        if (Number.isFinite(v)) onChange(Math.max(min, Math.min(max, v)) / 100);
      }}
    />
  );
}

const BLEND_MODES = () => [
  ['', t('Normal')],
  ['multiply', t('Multiply')],
  ['screen', t('Screen')],
  ['overlay', t('Overlay')],
  ['darken', t('Darken')],
  ['lighten', t('Lighten')],
  ['difference', t('Difference')],
];

/**
 * Where a clip stands on the frame and how: its place, size and angle as
 * numbers (what the grips on the preview change), its opacity and blend, a
 * clip or still that fills the frame made a picture in picture, and back.
 */
function PlacementProps({ clip, set, aspect }) {
  const fills = (clip.type === 'video' && !clip.fw) || (clip.type === 'image' && Boolean(clip.fit));
  const boxed = clip.type === 'image' || clip.type === 'video' || clip.type === 'shape';
  const toInset = () => {
    if (clip.type === 'video') {
      const r = insetRect({ corner: 'br', size: 0.35, shape: 'rounded', W: 1920, H: 1080, aspect });
      set(t('Picture in picture'), { fx: r.fx, fy: r.fy, fw: r.fw, fh: r.fh, insetShape: 'rounded' });
    } else set(t('Picture in picture'), { fit: undefined, fx: 0.3, fy: 0.25, fw: 0.4, rot: 0 });
  };
  return (
    <section className="vd-insp-section vd-place-props">
      <h3>{t('Position and size')}</h3>
      {fills ? (
        <>
          <div className="vd-prop-note">{t('This clip fills the frame. Drag a corner on the preview, or make it a picture in picture, to move, size and turn it.')}</div>
          <Button icon="window" label={t('Picture in picture')} className="vd-to-inset" onClick={toInset} />
        </>
      ) : (
        <>
          <Row label={t('Across')}>
            <Percent className="vd-fx" value={clip.fx ?? 0.5} onChange={(v) => set(t('Move'), { fx: v }, `fx-${clip.id}`)} />
          </Row>
          <Row label={t('Down')}>
            <Percent className="vd-fy" value={clip.fy ?? 0.5} onChange={(v) => set(t('Move'), { fy: v }, `fy-${clip.id}`)} />
          </Row>
          {boxed ? (
            <Row label={t('Width')}>
              <Percent className="vd-fw" min={2} max={200} value={clip.fw ?? 0.16} onChange={(v) => set(t('Resize'), clip.fh && clip.fw ? { fw: v, fh: +(clip.fh * (v / clip.fw)).toFixed(4) } : { fw: v }, `fw-${clip.id}`)} />
            </Row>
          ) : null}
          {boxed && clip.fh ? (
            <Row label={t('Height')}>
              <Percent className="vd-fh" min={2} max={200} value={clip.fh} onChange={(v) => set(t('Resize'), { fh: v }, `fh-${clip.id}`)} />
            </Row>
          ) : null}
          <Row label={t('Angle')}>
            <input
              className="rw-input vd-num vd-rot"
              type="number"
              min={-180}
              max={180}
              step={1}
              value={Math.round((clip.rot || 0) * 10) / 10}
              onChange={(e) => {
                const v = Number(e.target.value);
                if (Number.isFinite(v)) set(t('Rotate'), { rot: (((v % 360) + 540) % 360) - 180 }, `rot-${clip.id}`);
              }}
            />
          </Row>
          {clip.type === 'image' ? (
            <Button icon="maximize" label={t('Fill the whole frame')} className="vd-still-full" onClick={() => set(t('Fill the whole frame'), { fit: 'cover', fx: undefined, fy: undefined, fw: undefined, rot: 0 })} />
          ) : null}
        </>
      )}
      <Row label={t('Opacity')}>
        <Slider value={clip.opacity ?? 1} min={0} max={1} step={0.05} onChange={(x) => set(t('Opacity'), { opacity: x }, `op-${clip.id}`)} format={(x) => `${Math.round(x * 100)}%`} />
      </Row>
      <Row label={t('Blend')}>
        <Select className="vd-blend" value={clip.blend || ''} onChange={(e) => set(t('Blend'), { blend: e.target.value || undefined })}>
          {BLEND_MODES().map(([k, l]) => <option key={k} value={k}>{l}</option>)}
        </Select>
      </Row>
    </section>
  );
}

function Row({ label, children }) {
  const ref = React.useRef(null);
  const id = React.useId();
  // The row's name is its field's name to a screen reader.
  React.useLayoutEffect(() => labelFirstControl(ref.current, `${id}label`));
  return (
    <div className="vd-prop" ref={ref}>
      <span className="vd-prop-label" id={`${id}label`}>{label}</span>
      <span className="vd-prop-value">{children}</span>
    </div>
  );
}

function Slider({ value, min, max, step, onChange, format }) {
  return (
    <span className="vd-slider">
      <input type="range" min={min} max={max} step={step} value={value} onChange={(e) => onChange(Number(e.target.value))} />
      <span className="vd-slider-value">{format ? format(value) : value}</span>
    </span>
  );
}

const ANIMATIONS = () => [
  ['none', t('None')],
  ['fade', t('Fade')],
  ['rise', t('Rise')],
  ['type', t('Type on')],
  ['pop', t('Pop')],
];

/**
 * A clip drawn as an inset over the others (the camera of a recording): how
 * big, which corner, what shape, mirrored or not, or back to the whole frame.
 */
function InsetProps({ clip, set, aspect = 9 / 16 }) {
  const shape = clip.insetShape || 'rounded';
  const corner = cornerOf(clip);
  const place = (patch) => {
    const next = { corner, size: clip.fw, shape, ...patch };
    const r = insetRect({ corner: next.corner, size: next.size, shape: next.shape, W: 1920, H: 1080, aspect });
    return { fx: r.fx, fy: r.fy, fw: r.fw, fh: r.fh };
  };
  const corners = [['tl', t('Top left')], ['tr', t('Top right')], ['bl', t('Bottom left')], ['br', t('Bottom right')]];
  return (
    <section className="vd-insp-section vd-inset-props">
      <h3>{t('Picture in picture')}</h3>
      <Row label={t('Size')}>
        <Slider value={round(clip.fw, 1000)} min={0.1} max={0.6} step={0.01} onChange={(x) => set(t('Inset size'), place({ size: x }), `inset-size-${clip.id}`)} format={(x) => `${Math.round(x * 100)}%`} />
      </Row>
      <Row label={t('Corner')}>
        <span className="vd-corners">
          {corners.map(([k, label]) => (
            <button key={k} type="button" className={`vd-corner c-${k}${corner === k ? ' on' : ''}`} data-corner={k} data-tip={label} aria-label={label} onClick={() => set(t('Inset corner'), place({ corner: k }))}><i /></button>
          ))}
        </span>
      </Row>
      <Row label={t('Shape')}>
        <Select value={shape} onChange={(e) => set(t('Inset shape'), { insetShape: e.target.value, ...place({ shape: e.target.value }) })}>
          <option value="circle">{t('Circle')}</option>
          <option value="rounded">{t('Rounded')}</option>
          <option value="rect">{t('Rectangle')}</option>
        </Select>
      </Row>
      <Row label={t('Mirror')}>
        <input type="checkbox" checked={!!clip.flipX} onChange={(e) => set(t('Mirror'), { flipX: e.target.checked })} />
      </Row>
      <Button icon="maximize" label={t('Fill the whole frame')} className="vd-inset-full" onClick={() => set(t('Fill the whole frame'), { fx: undefined, fy: undefined, fw: undefined, fh: undefined, insetShape: undefined, flipX: false })} />
    </section>
  );
}

const POINTER_COLOURS = ['#ffd54a', '#ff5c8a', '#4cc9f0', '#7ee081', '#ffffff'];

/**
 * A screen recording's pointer (R7) and its automatic zoom (X5): the look of
 * the highlight and the rings, and the zooms found from the clicks, each
 * made stronger, weaker or taken away.
 */
function PointerProps({ doc, screen, pointer, edit, onSeek }) {
  const setPointer = (label, change, group) => edit(label, (d) => ({ ...d, patches: d.patches.map((p) => (p.id === pointer.id ? { ...p, ...change } : p)) }), group);
  const zooms = screen?.autoZoom || [];
  const levels = [1.4, 1.8, 2.2, 2.6];
  return (
    <>
      <section className="vd-insp-section vd-pointer-props">
        <h3>{t('Pointer')}</h3>
        <Row label={t('Highlight')}>
          <input type="checkbox" className="vd-pointer-highlight" checked={pointer.highlight !== false} onChange={(e) => setPointer(t('Highlight'), { highlight: e.target.checked })} />
        </Row>
        <Row label={t('Click rings')}>
          <input type="checkbox" className="vd-pointer-rings" checked={pointer.rings !== false} onChange={(e) => setPointer(t('Click rings'), { rings: e.target.checked })} />
        </Row>
        <Row label={t('Larger pointer')}>
          <input type="checkbox" checked={Boolean(pointer.arrow)} onChange={(e) => setPointer(t('Larger pointer'), { arrow: e.target.checked })} />
        </Row>
        <Row label={t('Size')}>
          <Slider value={pointer.size || 1} min={0.5} max={2.5} step={0.1} onChange={(x) => setPointer(t('Size'), { size: x }, `ptr-size-${pointer.id}`)} format={(x) => `${Math.round(x * 100)}%`} />
        </Row>
        <Row label={t('Colour')}>
          <span className="vd-swatches">
            {POINTER_COLOURS.map((c) => (
              <button key={c} type="button" className={`vd-swatch${String(pointer.color).toLowerCase() === c ? ' on' : ''}`} style={{ background: c }} aria-label={c} onClick={() => setPointer(t('Colour'), { color: c })} />
            ))}
          </span>
        </Row>
        <div className="vd-prop-note">{tn((pointer.clicks || []).length, '{count} click was recorded.', '{count} clicks were recorded.')}</div>
      </section>
      {screen ? (
        <section className="vd-insp-section vd-zoom-props">
          <h3>{t('Automatic zoom')}</h3>
          {zooms.length ? (
            <>
              {zooms.map((z, i) => (
                <div key={`${z.start}-${i}`} className="vd-zoom-row" data-level={z.level}>
                  <button type="button" className="vd-marker-time" onClick={() => onSeek(screen.timing.start + z.start)}>{timecode(screen.timing.start + z.start)}</button>
                  <span className="vd-zoom-len">{t('for {seconds} s', { seconds: Math.round((z.end - z.start) * 10) / 10 })}</span>
                  <Select value={String(z.level)} aria-label={t('How close')} onChange={(e) => edit(t('Zoom'), (d) => updateZoom(d, screen.id, i, { level: Number(e.target.value) }))}>
                    {[...new Set([...levels, z.level])].sort((a, b) => a - b).map((l) => <option key={l} value={String(l)}>{t('{level} times', { level: l })}</option>)}
                  </Select>
                  <Button icon="trash" title={t('Remove this zoom')} onClick={() => edit(t('Remove zoom'), (d) => updateZoom(d, screen.id, i, null))} />
                </div>
              ))}
              <Button icon="close" label={t('Take the zoom off')} className="vd-zoom-off" onClick={() => edit(t('Take the zoom off'), (d) => setZooms(d, screen.id, []))} />
            </>
          ) : (
            <>
              <div className="vd-prop-note">{(pointer.clicks || []).length ? t('Zoom in where the clicks are, and follow them.') : t('No clicks were recorded to zoom on.')}</div>
              <Button icon="zoomIn" label={t('Zoom on Clicks')} className="vd-zoom-on" disabled={!(pointer.clicks || []).length} onClick={() => edit(t('Zoom on Clicks'), (d) => applyAutoZoom(d, screen.id))} />
            </>
          )}
        </section>
      ) : null}
    </>
  );
}

export default function Inspector({ doc, selected, edit, onSeek, missing, kindName }) {
  const clip = selected.length ? clipById(doc, selected[0]) : null;
  if (!clip) return <ProjectProps doc={doc} edit={edit} onSeek={onSeek} missing={missing} />;

  const set = (label, change, group) => edit(label, (d) => ({ ...d, patches: d.patches.map((p) => (p.id === clip.id ? { ...p, ...change } : p)) }), group);
  const sound = clip.type === 'sound' ? clip : partnersOf(doc, clip.id).find((p) => p.type === 'sound');
  const setSound = (label, change, group) => edit(label, (d) => ({ ...d, patches: d.patches.map((p) => (p.id === sound.id ? { ...p, ...change } : p)) }), group);
  const isPicture = clip.type === 'video' || clip.type === 'image';
  const before = isPicture ? previousOnLane(doc, clip.id) : null;
  const crop = clip.crop || { x: 0, y: 0, w: 1, h: 1 };
  const setCrop = (side, v) => {
    const c = { left: crop.x, top: crop.y, right: 1 - crop.x - crop.w, bottom: 1 - crop.y - crop.h, [side]: v };
    const next = { x: c.left, y: c.top, w: Math.max(0.05, 1 - c.left - c.right), h: Math.max(0.05, 1 - c.top - c.bottom) };
    set(t('Crop'), { crop: next.x === 0 && next.y === 0 && next.w === 1 && next.h === 1 ? null : next }, `crop-${clip.id}`);
  };

  return (
    <div className="vd-inspector">
      <div className="vd-insp-head">
        <span className={`vd-insp-kind k-${clip.type}`}><Icon name={clip.caption ? 'captions' : { video: 'video', image: 'picture', sound: 'volume', text: 'textbox', shape: 'shape', pointer: 'pointer' }[clip.type] || 'file'} size={14} /></span>
        <div className="vd-insp-title">
          <strong>{clip.type === 'text' ? clip.text || t('Title') : clip.name || kindName(clip.type)}</strong>
          <small>{kindName(clip.caption ? 'caption' : clip.type)}{selected.length > 1 ? `, ${t('{count} picked', { count: selected.length })}` : ''}</small>
        </div>
      </div>

      <section className="vd-insp-section">
        <h3>{t('Timing')}</h3>
        <Row label={t('Starts at')}>
          <Seconds value={clip.timing.start} onChange={(v) => edit(t('Move clip'), (d) => moveClip(d, clip.id, { start: v }), `start-${clip.id}`)} />
        </Row>
        <Row label={t('Length')}>
          <Seconds value={lengthOf(clip)} min={0.1} onChange={(v) => edit(t('Trim clip'), (d) => trimEnd(d, clip.id, clip.timing.start + v, { ripple: d.settings?.ripple !== false }), `len-${clip.id}`)} />
        </Row>
        {clip.type === 'video' || clip.type === 'sound' ? (
          <Row label={t('From the source')}>
            <span className="vd-prop-note">{t('{from} to {to}', { from: timecode(clip.offset || 0), to: timecode((clip.offset || 0) + lengthOf(clip)) })}</span>
          </Row>
        ) : null}
        {missing?.has?.(clip.url) ? <div className="vd-warn">{t('This clip’s file cannot be found: {path}', { path: clip.url })}</div> : null}
      </section>

      {isPicture ? (
        <section className="vd-insp-section">
          <h3>{t('Picture')}</h3>
          <Row label={t('Fit')}>
            <Select value={clip.fit || 'cover'} onChange={(e) => set(t('Fit'), { fit: e.target.value })}>
              <option value="cover">{t('Fill the frame')}</option>
              <option value="contain">{t('Fit inside (black bars)')}</option>
              <option value="blur">{t('Fit inside (blurred fill)')}</option>
            </Select>
          </Row>
          <Row label={t('Rotate')}>
            <Button icon="rotate" title={t('Rotate left')} className="flip-x" onClick={() => set(t('Rotate'), { turn: (((clip.turn || 0) - 90) % 360 + 360) % 360 })} />
            <Button icon="rotate" title={t('Rotate right')} onClick={() => set(t('Rotate'), { turn: ((clip.turn || 0) + 90) % 360 })} />
            <Button icon="flip" title={t('Flip horizontal')} pressed={!!clip.flipX} onClick={() => set(t('Flip'), { flipX: !clip.flipX })} />
            <Button icon="flip" title={t('Flip vertical')} className="turn-90" pressed={!!clip.flipY} onClick={() => set(t('Flip'), { flipY: !clip.flipY })} />
          </Row>
          {[['left', t('Crop left')], ['right', t('Crop right')], ['top', t('Crop top')], ['bottom', t('Crop bottom')]].map(([side, label]) => {
            const v = side === 'left' ? crop.x : side === 'top' ? crop.y : side === 'right' ? 1 - crop.x - crop.w : 1 - crop.y - crop.h;
            return (
              <Row key={side} label={label}>
                <Slider value={round(Math.max(0, v), 1000)} min={0} max={0.45} step={0.01} onChange={(x) => setCrop(side, x)} format={(x) => `${Math.round(x * 100)}%`} />
              </Row>
            );
          })}
        </section>
      ) : null}

      {['video', 'image', 'text', 'shape'].includes(clip.type) && !clip.caption ? <PlacementProps clip={clip} set={set} aspect={doc.sources?.[clip.url]?.width ? doc.sources[clip.url].height / doc.sources[clip.url].width : 9 / 16} /> : null}

      {(() => {
        // A screen recording with its pointer, or the pointer itself: its look and the zoom on its clicks.
        const screen = clip.type === 'video' ? clip : clip.type === 'pointer' ? doc.patches.find((p) => p.type === 'video' && p.url === clip.screenUrl && p.timing && p.timing.start < clip.timing.end && p.timing.end > clip.timing.start && (!clip.group || p.group === clip.group)) || null : null;
        const pointer = clip.type === 'pointer' ? clip : clip.type === 'video' ? pointerOf(doc, clip.id) : null;
        return pointer ? <PointerProps doc={doc} screen={screen} pointer={pointer} edit={edit} onSeek={onSeek} /> : null;
      })()}

      {clip.type === 'video' && clip.fw ? <InsetProps clip={clip} set={set} aspect={doc.sources?.[clip.url]?.width ? doc.sources[clip.url].height / doc.sources[clip.url].width : 9 / 16} /> : null}

      {before ? (
        <section className="vd-insp-section">
          <h3>{t('Transition in')}</h3>
          <Row label={t('Kind')}>
            <Select value={clip.transition?.kind || 'none'} onChange={(e) => edit(t('Transition'), (d) => setTransition(d, clip.id, { kind: e.target.value, seconds: clip.transition?.seconds || 1 }))}>
              <option value="none">{t('None')}</option>
              {TRANSITIONS.map((k) => <option key={k} value={k}>{transitionName(k)}</option>)}
            </Select>
          </Row>
          {clip.transition ? (
            <Row label={t('Duration')}>
              <Seconds value={clip.transition.seconds} min={0.1} onChange={(v) => edit(t('Transition'), (d) => setTransition(d, clip.id, { kind: clip.transition.kind, seconds: v }), `tr-${clip.id}`)} />
            </Row>
          ) : null}
        </section>
      ) : null}

      {sound ? (
        <section className="vd-insp-section">
          <h3>{t('Sound')}</h3>
          <Row label={t('Volume')}>
            <Slider value={sound.volume ?? 1} min={0} max={1} step={0.01} onChange={(x) => setSound(t('Volume'), { volume: x }, `vol-${sound.id}`)} format={(x) => `${Math.round(x * 100)}%`} />
          </Row>
          <Row label={t('Mute')}>
            <input type="checkbox" checked={!!sound.muted} onChange={(e) => setSound(t('Mute'), { muted: e.target.checked })} />
          </Row>
          <Row label={t('Track volume')}>
            <Slider
              value={doc.lanes?.a?.[sound.lane || 0]?.volume ?? 1}
              min={0}
              max={1}
              step={0.01}
              onChange={(x) => edit(t('Track volume'), (d) => ({ ...d, lanes: { ...d.lanes, a: { ...(d.lanes?.a || {}), [sound.lane || 0]: { ...((d.lanes?.a || {})[sound.lane || 0] || {}), volume: x } } } }), `lane-vol-${sound.lane || 0}`)}
              format={(x) => `${Math.round(x * 100)}%`}
            />
          </Row>
          <Row label={t('Fade In')}>
            <Seconds value={sound.enter?.seconds || 0} onChange={(v) => setSound(t('Fade In'), { enter: { kind: 'fade', seconds: Math.max(0, v) } }, `fi-${sound.id}`)} />
          </Row>
          <Row label={t('Fade Out')}>
            <Seconds value={sound.exit?.seconds || 0} onChange={(v) => setSound(t('Fade Out'), { exit: { kind: 'fade', seconds: Math.max(0, v) } }, `fo-${sound.id}`)} />
          </Row>
        </section>
      ) : null}

      {clip.type === 'text' ? (
        <section className="vd-insp-section">
          <h3>{t('Words')}</h3>
          <textarea className="rw-input vd-text" dir="auto" rows={3} value={clip.text || ''} onChange={(e) => set(t('Typing'), { text: e.target.value }, `text-${clip.id}`)} />
          <Row label={t('Size')}>
            <Slider value={clip.sizeFrac || 0.04} min={0.015} max={0.14} step={0.002} onChange={(x) => set(t('Size'), { sizeFrac: x }, `size-${clip.id}`)} format={(x) => `${Math.round(x * 1000) / 10}`} />
          </Row>
          <Row label={t('Bold')}>
            <input type="checkbox" checked={(clip.weight || 600) >= 700} onChange={(e) => set(t('Bold'), { weight: e.target.checked ? 800 : 500 })} />
          </Row>
          <Row label={t('Colour')}>
            <Swatches value={clip.color} onPick={(c) => set(t('Colour'), { color: c })} />
          </Row>
          <Row label={t('Animation in')}>
            <Select value={animationName(clip.enter)} onChange={(e) => set(t('Animation'), { enter: TEXT_ANIMATIONS[e.target.value] })}>
              {ANIMATIONS().map(([k, l]) => <option key={k} value={k}>{l}</option>)}
            </Select>
          </Row>
          <Row label={t('Animation out')}>
            <Select value={animationName(clip.exit)} onChange={(e) => set(t('Animation'), { exit: e.target.value === 'type' ? TEXT_ANIMATIONS.fade : TEXT_ANIMATIONS[e.target.value] })}>
              {ANIMATIONS().filter(([k]) => k !== 'type').map(([k, l]) => <option key={k} value={k}>{l}</option>)}
            </Select>
          </Row>
        </section>
      ) : null}

      {clip.type === 'shape' ? (
        <section className="vd-insp-section">
          <h3>{t('Background')}</h3>
          <Row label={t('Colour')}>
            <Swatches value={clip.color} onPick={(c) => set(t('Colour'), { color: c })} />
          </Row>
        </section>
      ) : null}
    </div>
  );
}

function Swatches({ value, onPick }) {
  return (
    <span className="vd-swatches">
      {COLOURS.map((c) => (
        <button key={c} type="button" className={`vd-swatch${String(value).toLowerCase() === c ? ' on' : ''}`} style={{ background: c }} aria-label={c} onClick={() => onPick(c)} />
      ))}
    </span>
  );
}

/** The project's own settings, markers and sources. */
function ProjectProps({ doc, edit, onSeek, missing }) {
  const sources = Object.entries(doc.sources || {}).filter(([url]) => doc.patches.some((p) => p.url === url));
  return (
    <div className="vd-inspector">
      <div className="vd-insp-head">
        <span className="vd-insp-kind k-project"><Icon name="video" size={14} /></span>
        <div className="vd-insp-title">
          <strong>{t('Project')}</strong>
          <small>{t('Pick a clip to see its settings.')}</small>
        </div>
      </div>
      <section className="vd-insp-section">
        <h3>{t('Settings')}</h3>
        <Row label={t('Frame')}>
          <Select value={doc.options.aspect} onChange={(e) => edit(t('Frame'), (d) => composeIntro({ ...d, options: { ...d.options, aspect: e.target.value } }))}>
            <option value="landscape">{t('Landscape 16:9')}</option>
            <option value="vertical">{t('Vertical 9:16')}</option>
            <option value="square">{t('Square 1:1')}</option>
          </Select>
        </Row>
        <Row label={t('A still lasts')}>
          <Seconds value={doc.settings?.stillSeconds ?? 5} min={0.5} step={0.5} onChange={(v) => edit(t('Settings'), (d) => ({ ...d, settings: { ...d.settings, stillSeconds: Math.max(0.5, v) } }))} />
        </Row>
        <Row label={t('Ripple')}>
          <input type="checkbox" checked={doc.settings?.ripple !== false} onChange={(e) => edit(t('Settings'), (d) => ({ ...d, settings: { ...d.settings, ripple: e.target.checked } }))} />
        </Row>
        <Row label={t('Snapping')}>
          <input type="checkbox" checked={doc.settings?.snap !== false} onChange={(e) => edit(t('Settings'), (d) => ({ ...d, settings: { ...d.settings, snap: e.target.checked } }))} />
        </Row>
      </section>
      <section className="vd-insp-section">
        <h3>{t('Markers')}</h3>
        {(doc.markers || []).length ? (
          (doc.markers || []).map((m) => (
            <div key={m.id} className="vd-marker-row">
              <button type="button" className="vd-marker-time" onClick={() => onSeek(m.t)}>{timecode(m.t)}</button>
              <input className="rw-input" value={m.name} placeholder={t('Marker name')} onChange={(e) => edit(t('Rename marker'), (d) => updateMarker(d, m.id, { name: e.target.value }), `marker-${m.id}`)} />
              <Button icon="trash" title={t('Remove marker')} onClick={() => edit(t('Remove marker'), (d) => removeMarker(d, m.id))} />
            </div>
          ))
        ) : (
          <div className="vd-prop-note">{t('Add a marker at the playhead with M.')}</div>
        )}
      </section>
      <section className="vd-insp-section">
        <h3>{t('Sources')}</h3>
        {sources.length ? (
          sources.map(([url, s]) => (
            <div key={url} className={`vd-source${missing?.has?.(url) ? ' missing' : ''}`} title={url}>
              <Icon name={s.kind === 'image' ? 'picture' : s.kind === 'audio' ? 'volume' : 'video'} size={13} />
              <span className="vd-source-name">{s.name || url}</span>
              <span className="vd-source-len">{s.duration ? timecode(s.duration) : ''}</span>
            </div>
          ))
        ) : (
          <div className="vd-prop-note">{t('Nothing on the timeline yet.')}</div>
        )}
      </section>
    </div>
  );
}
