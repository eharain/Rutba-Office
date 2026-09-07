// Presenter view.
//
// The audience window shows the slide. This one shows what the speaker needs
// and the audience must never see: the notes, the slide that is coming, how
// long they have been talking and what the time is.
//
// It is a window of its own so it can be dragged to the other screen, which is
// the entire point — a presenter view that shares a display with the slide is a
// notes panel with ambitions. The two windows agree through the `present`
// namespace in the main process, because two renderers have no way to speak to
// each other and duplicating the position in both is how they drift apart.

import React, { useCallback, useEffect, useRef, useState } from 'react';
import { Button, Icon, Spinner, Empty } from '@rutba/office-ui';

/** hh:mm:ss, or mm:ss for a talk that has not gone on too long. */
function elapsed(from) {
  if (!from) return '0:00';
  const total = Math.max(0, Math.floor((Date.now() - from) / 1000));
  const s = String(total % 60).padStart(2, '0');
  const m = Math.floor(total / 60) % 60;
  const h = Math.floor(total / 3600);
  return h ? `${h}:${String(m).padStart(2, '0')}:${s}` : `${m}:${s}`;
}

export default function Presenter({ shell, docId }) {
  const [model, setModel] = useState(null);
  const [next, setNext] = useState(null);
  const [state, setState] = useState({ index: 0, running: false, startedAt: null, blank: false });
  const [now, setNow] = useState(Date.now());
  const [error, setError] = useState(null);
  const index = state.index || 0;

  // A clock that ticks. One second is the right resolution for a talk, and a
  // faster one would re-render the notes for no reason.
  useEffect(() => {
    const timer = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(timer);
  }, []);

  useEffect(() => {
    let live = true;
    shell.present.state().then((s) => live && setState((current) => ({ ...current, ...s }))).catch(() => {});
    const off = shell.on('present:state', (s) => setState((current) => ({ ...current, ...s })));
    return () => {
      live = false;
      off?.();
    };
  }, [shell]);

  // This slide and the one after it. Two calls rather than one, because the
  // deck model is per-slide and asking for both is cheaper than a model that
  // carries every slide's scene.
  useEffect(() => {
    let live = true;
    Promise.all([
      shell.doc.model({ id: docId, slide: index, width: 1280 }),
      shell.doc.model({ id: docId, slide: index + 1, width: 640 }),
    ])
      .then(([current, upcoming]) => {
        if (!live) return;
        setModel(current);
        setNext(upcoming.index === index ? null : upcoming);
      })
      .catch((err) => live && setError(err.message));
    return () => {
      live = false;
    };
  }, [shell, docId, index]);

  const move = useCallback(
    (delta) => {
      const count = model?.count ?? 1;
      shell.present.set({ index: Math.max(0, Math.min(count - 1, index + delta)) }).catch(() => {});
    },
    [shell, model, index]
  );

  useEffect(() => {
    const onKey = (e) => {
      if (e.key === 'ArrowRight' || e.key === 'PageDown' || e.key === ' ' || e.key === 'Enter') { e.preventDefault(); move(1); }
      else if (e.key === 'ArrowLeft' || e.key === 'PageUp' || e.key === 'Backspace') { e.preventDefault(); move(-1); }
      else if (e.key === 'b' || e.key === '.') shell.present.set({ blank: !state.blank }).catch(() => {});
      else if (e.key === 'Home') shell.present.set({ index: 0 }).catch(() => {});
      else if (e.key === 'End') shell.present.set({ index: (model?.count ?? 1) - 1 }).catch(() => {});
      else if (e.key === 'Escape') shell.present.set({ running: false }).catch(() => {});
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [move, shell, state.blank, model]);

  if (error) return <Empty icon="info" title="The presentation is not open">{error}</Empty>;
  if (!model) {
    return (
      <div className="pv" style={{ placeItems: 'center', display: 'grid' }}>
        <Spinner />
      </div>
    );
  }

  const notes = model.slide?.notes || '';

  return (
    <div className="pv">
      <style>{CSS}</style>

      <header className="pv-bar">
        <span className="pv-clock" title="How long you have been talking">
          <Icon name="clock" size={14} />
          {elapsed(state.startedAt)}
        </span>
        <Button
          icon={state.running ? 'pause' : 'play'}
          label={state.running ? 'Pause' : 'Start'}
          onClick={() => shell.present.set({ running: !state.running })}
        />
        <Button icon="refresh" label="Reset" onClick={() => shell.present.set({ restart: true, running: true })} />
        <span className="pv-spacer" />
        <span className="pv-position">
          Slide {index + 1} of {model.count}
        </span>
        <span className="pv-spacer" />
        <Button icon={state.blank ? 'eye' : 'stop'} label={state.blank ? 'Show' : 'Black'} onClick={() => shell.present.set({ blank: !state.blank })} />
        <span className="pv-time">{new Date(now).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</span>
      </header>

      <div className="pv-body">
        <section className="pv-current">
          <div className="pv-stage" dangerouslySetInnerHTML={{ __html: model.slide?.svg || '' }} />
          <div className="pv-controls">
            <Button icon="skipBack" label="Previous" disabled={index === 0} onClick={() => move(-1)} />
            <Button primary icon="skipForward" label="Next" disabled={index >= model.count - 1} onClick={() => move(1)} />
          </div>
        </section>

        <aside className="pv-side">
          <div className="pv-next">
            <div className="pv-label">{next ? 'Next' : 'End of the deck'}</div>
            {next ? <div className="pv-thumb" dangerouslySetInnerHTML={{ __html: next.slide?.svg || '' }} /> : <div className="pv-thumb pv-empty">Nothing after this one</div>}
          </div>

          <div className="pv-notes">
            <div className="pv-label">Notes</div>
            {notes ? (
              <div className="pv-notes-text">
                {notes.split('\n').map((line, i) => (
                  <p key={i}>{line}</p>
                ))}
              </div>
            ) : (
              <p className="pv-none">No notes on this slide. Add them from the Home tab.</p>
            )}
          </div>
        </aside>
      </div>
    </div>
  );
}

const CSS = `
.pv { flex: 1; display: flex; flex-direction: column; min-height: 0; background: var(--sunken); }
.pv-bar {
  display: flex; align-items: center; gap: 10px; padding: 8px 14px;
  border-bottom: 1px solid var(--line); background: var(--chrome);
}
.pv-spacer { flex: 1; }
.pv-clock {
  display: inline-flex; align-items: center; gap: 7px; font-variant-numeric: tabular-nums;
  font-size: 19px; font-weight: 600; min-width: 96px;
}
.pv-position { font-size: 13px; color: var(--ink-2); font-variant-numeric: tabular-nums; }
.pv-time { font-size: 15px; color: var(--ink-2); font-variant-numeric: tabular-nums; min-width: 62px; text-align: right; }

.pv-body { flex: 1; display: grid; grid-template-columns: 1.55fr 1fr; gap: 14px; padding: 14px; min-height: 0; }
.pv-current { display: flex; flex-direction: column; gap: 10px; min-width: 0; min-height: 0; }
.pv-stage {
  flex: 1; min-height: 0; display: grid; place-items: center; overflow: hidden;
  background: #000; border-radius: var(--r-2); box-shadow: var(--shadow-1);
}
.pv-stage svg, .pv-thumb svg { width: 100%; height: 100%; display: block; }
.pv-controls { display: flex; gap: 8px; justify-content: center; }

.pv-side { display: flex; flex-direction: column; gap: 14px; min-height: 0; }
.pv-label {
  font-size: 11px; text-transform: uppercase; letter-spacing: .06em;
  color: var(--ink-3); margin-bottom: 6px;
}
.pv-next { flex: none; }
.pv-thumb {
  aspect-ratio: 16 / 9; background: #000; border-radius: var(--r-2);
  overflow: hidden; box-shadow: var(--shadow-1);
}
.pv-thumb.pv-empty {
  display: grid; place-items: center; background: var(--chrome); color: var(--ink-3);
  font-size: 12px; box-shadow: none; border: 1px dashed var(--line);
}
.pv-notes {
  flex: 1; min-height: 0; display: flex; flex-direction: column;
  background: var(--surface); border: 1px solid var(--line); border-radius: var(--r-2); padding: 12px 14px;
}
.pv-notes-text { overflow: auto; font-size: 17px; line-height: 1.55; }
.pv-notes-text p { margin: 0 0 .7em; }
.pv-none { color: var(--ink-3); font-size: 13px; margin: 0; }
`;
