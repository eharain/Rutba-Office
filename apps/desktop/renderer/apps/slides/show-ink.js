// The slide show's pen, as PowerPoint's: a small bar at the bottom left of
// the show (Pen, Highlighter, Eraser, Erase All Ink on Slide, the colours),
// Ctrl+P, Ctrl+I and Ctrl+E for the pen, the highlighter and the eraser,
// E to erase the slide's ink, Ctrl+A or Escape for the arrow again. While a
// pen is in hand a press draws rather than moving the show on. When the
// show ends with ink drawn, the editor asks "Do you want to keep your ink
// annotations?": Keep puts each slide's strokes on it as ink, saved as
// PowerPoint saves ink; Discard lets them go.

import React, { useCallback, useEffect, useRef, useState } from 'react';
import { Button, Dialog, Icon, t, tn } from '@rutba/office-ui';

const SHOW_COLOURS = ['#E71224', '#FFC114', '#0000FF', '#00A0FF', '#008C3A', '#000000', '#FFFFFF'];
const PEN = { tool: 'pen', width: 2.25, alpha: 1, cap: 'rnd' };
const HIGHLIGHTER = { tool: 'highlighter', width: 12, alpha: 0.5, cap: 'sq' };

/** The stage's box on screen, and how a screen point maps to the slide's own pixels. */
function stageFrame(size) {
  const el = document.querySelector('.sl-present .sl-show-stage');
  if (!el || !size) return null;
  const r = el.getBoundingClientRect();
  return { left: r.left, top: r.top, k: r.width / size.width };
}

/**
 * The show's pen for one window. `present` is whether the show is on,
 * `index` the slide on screen, `size` the slide's in pixels; `onKeep`
 * receives `[{ slide, strokes }]` to put on the slides.
 */
export function useShowInk({ present, index, size, onKeep }) {
  const [tool, setTool] = useState(null);
  const [colour, setColour] = useState(SHOW_COLOURS[0]);
  const [strokes, setStrokes] = useState({});
  const [live, setLive] = useState(null);
  const [ask, setAsk] = useState(null);
  const [, redraw] = useState(0);
  const drag = useRef(null);
  const wasPresent = useRef(false);

  // When the show ends with ink drawn, the question is asked in the editor, as PowerPoint asks it.
  useEffect(() => {
    if (present) { wasPresent.current = true; return; }
    if (!wasPresent.current) return;
    wasPresent.current = false;
    const drawn = Object.entries(strokes).filter(([, list]) => list.length).map(([slide, list]) => ({ slide: Number(slide), strokes: list }));
    if (drawn.length) setAsk(drawn);
    setStrokes({});
    setTool(null);
    setLive(null);
  }, [present]); // eslint-disable-line react-hooks/exhaustive-deps

  // The keys PowerPoint gives the pen in a show.
  useEffect(() => {
    if (!present) return undefined;
    const key = (e) => {
      const k = e.key.toLowerCase();
      const pick = (v) => { e.preventDefault(); e.stopImmediatePropagation(); setTool(v); };
      if (e.ctrlKey && k === 'p') pick('pen');
      else if (e.ctrlKey && k === 'i') pick('highlighter');
      else if (e.ctrlKey && k === 'e') pick('eraser');
      else if (e.ctrlKey && k === 'a') pick(null);
      else if (!e.ctrlKey && k === 'e' && !e.altKey) { e.preventDefault(); e.stopImmediatePropagation(); setStrokes((s) => ({ ...s, [index]: [] })); }
      else if (e.key === 'Escape' && tool) { e.preventDefault(); e.stopImmediatePropagation(); setTool(null); }
    };
    const resize = () => redraw((n) => n + 1);
    window.addEventListener('keydown', key, true);
    window.addEventListener('resize', resize);
    return () => { window.removeEventListener('keydown', key, true); window.removeEventListener('resize', resize); };
  }, [present, index, tool]);

  const toSlide = (e) => {
    const f = stageFrame(size);
    return f ? [(e.clientX - f.left) / f.k, (e.clientY - f.top) / f.k] : null;
  };
  const eraseAt = useCallback((p) => {
    setStrokes((s) => {
      const list = s[index] || [];
      const near = (st) => st.points.some(([x, y]) => Math.hypot(x - p[0], y - p[1]) < Math.max(8, st.width));
      const kept = list.filter((st) => !near(st));
      return kept.length === list.length ? s : { ...s, [index]: kept };
    });
  }, [index]);

  const down = (e) => {
    if (!tool || e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    e.currentTarget.setPointerCapture?.(e.pointerId);
    const p = toSlide(e);
    if (!p) return;
    if (tool === 'eraser') { drag.current = { erase: true }; eraseAt(p); return; }
    const look = tool === 'highlighter' ? HIGHLIGHTER : PEN;
    drag.current = { points: [p], color: tool === 'highlighter' && colour === SHOW_COLOURS[0] ? '#FFFF00' : colour, width: look.width, alpha: look.alpha, cap: look.cap };
    setLive({ ...drag.current });
  };
  const move = (e) => {
    const d = drag.current;
    if (!d) return;
    e.preventDefault();
    const p = toSlide(e);
    if (!p) return;
    if (d.erase) { eraseAt(p); return; }
    d.points.push(p);
    setLive({ ...d, points: d.points.slice() });
  };
  const up = (e) => {
    const d = drag.current;
    drag.current = null;
    setLive(null);
    if (!d || d.erase) return;
    e.stopPropagation();
    if (d.points.length) setStrokes((s) => ({ ...s, [index]: [...(s[index] || []), d] }));
  };

  const f = present ? stageFrame(size) : null;
  const path = (st) => {
    if (!f || !st.points.length) return '';
    const pts = st.points.map(([x, y]) => [f.left + x * f.k, f.top + y * f.k]);
    return 'M' + pts.map(([x, y]) => `${x.toFixed(1)} ${y.toFixed(1)}`).join(' L');
  };
  const here = [...(strokes[index] || []), ...(live ? [live] : [])];
  const count = Object.values(strokes).reduce((n, list) => n + list.length, 0);

  const overlay = present ? (
    <>
      <style>{SHOW_INK_CSS}</style>
      <svg
        className={`sl-show-ink${tool ? ' on' : ''}${tool === 'eraser' ? ' erasing' : ''}`}
        onPointerDown={down}
        onPointerMove={move}
        onPointerUp={up}
        onClick={(e) => { if (tool) e.stopPropagation(); }}
      >
        {here.map((st, i) => (
          <path key={i} className="sl-show-stroke" d={path(st)} fill="none" stroke={st.color} strokeOpacity={st.alpha} strokeWidth={Math.max(1, st.width * (96 / 72) * (f?.k || 1))} strokeLinecap={st.cap === 'sq' ? 'square' : 'round'} strokeLinejoin="round" />
        ))}
      </svg>
      <div className="sl-show-pens" onClick={(e) => e.stopPropagation()} onPointerDown={(e) => e.stopPropagation()}>
        <button type="button" className={`sl-show-pen${tool === 'pen' ? ' on' : ''}`} data-tool="pen" data-tip={t('Pen (Ctrl+P)')} onClick={() => setTool(tool === 'pen' ? null : 'pen')}><Icon name="pen" size={16} /></button>
        <button type="button" className={`sl-show-pen${tool === 'highlighter' ? ' on' : ''}`} data-tool="highlighter" data-tip={t('Highlighter (Ctrl+I)')} onClick={() => setTool(tool === 'highlighter' ? null : 'highlighter')}><Icon name="highlighter" size={16} /></button>
        <button type="button" className={`sl-show-pen${tool === 'eraser' ? ' on' : ''}`} data-tool="eraser" data-tip={t('Eraser (Ctrl+E)')} onClick={() => setTool(tool === 'eraser' ? null : 'eraser')}><Icon name="eraser" size={16} /></button>
        <button type="button" className="sl-show-pen" data-tool="clear" data-tip={t('Erase All Ink on Slide (E)')} onClick={() => setStrokes((s) => ({ ...s, [index]: [] }))}><Icon name="trash" size={16} /></button>
        <span className="sl-show-sep" />
        {SHOW_COLOURS.map((c) => <button key={c} type="button" className={`sl-show-colour${colour === c ? ' on' : ''}`} data-colour={c} style={{ background: c }} data-tip={t('Ink Color')} onClick={() => { setColour(c); if (!tool || tool === 'eraser') setTool('pen'); }} />)}
      </div>
    </>
  ) : null;

  const dialog = ask ? (
    <Dialog
      title={t('Rutba Presentations')}
      width={440}
      onClose={() => setAsk(null)}
      className="sl-keep-ink"
      actions={<><Button className="sl-keep-ink-discard" label={t('Discard')} onClick={() => setAsk(null)} /><Button primary className="sl-keep-ink-keep" label={t('Keep')} onClick={() => { const list = ask; setAsk(null); onKeep(list); }} /></>}
    >
      <p style={{ margin: 0 }}>{t('Do you want to keep your ink annotations?')}</p>
      <p style={{ margin: '8px 0 0', color: 'var(--ink-2)', fontSize: 12.5 }}>{tn(ask.reduce((n, a) => n + a.strokes.length, 0), '{count} stroke drawn in the show.', '{count} strokes drawn in the show.')}</p>
    </Dialog>
  ) : null;

  return { overlay, dialog, tool, setTool, count, strokes };
}

const SHOW_INK_CSS = `
.sl-show-ink { position: fixed; inset: 0; width: 100vw; height: 100vh; z-index: 210; pointer-events: none; }
.sl-show-ink.on { pointer-events: auto; cursor: crosshair; }
.sl-show-ink.erasing { cursor: cell; }
.sl-show-pens { position: fixed; left: 18px; bottom: 18px; z-index: 220; display: flex; align-items: center; gap: 4px; padding: 5px 6px; border-radius: 10px; background: rgba(32, 32, 32, 0.72); opacity: 0.35; transition: opacity 150ms ease; }
.sl-show-pens:hover, .sl-present:hover .sl-show-pens { opacity: 1; }
.sl-show-pen { width: 30px; height: 30px; border: 0; border-radius: 7px; background: transparent; color: #fff; display: grid; place-items: center; cursor: pointer; }
.sl-show-pen:hover { background: rgba(255, 255, 255, 0.16); }
.sl-show-pen.on { background: rgba(255, 255, 255, 0.28); box-shadow: inset 0 0 0 1px rgba(255,255,255,.6); }
.sl-show-sep { width: 1px; height: 20px; background: rgba(255,255,255,.3); margin: 0 3px; }
.sl-show-colour { width: 16px; height: 16px; border-radius: 50%; border: 1.5px solid rgba(255,255,255,.75); cursor: pointer; padding: 0; }
.sl-show-colour.on { box-shadow: 0 0 0 2px #fff; }
`;
