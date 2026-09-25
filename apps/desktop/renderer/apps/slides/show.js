// The slideshow's stage, and the editor's Preview.
//
// The show keeps the slide it is leaving on screen while the next one comes
// in: two layers in one box, the transition played between them (see
// motion.js), then the old layer dropped. Moving forward plays the
// incoming slide's own transition, as PowerPoint does; stepping back cuts.

import React, { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { playTransition } from './motion.js';

const plays = (t) => Boolean(t && t.type && t.type !== 'none' && Number(t.duration) > 0);

/**
 * @param {{ slide: { index: number, svg: string, transition?: object }, hidden?: boolean, onShown?: (index: number) => void }} props
 * `onShown` is told each slide's index once it is fully on screen — after
 * its transition — which is when "advance after" starts counting.
 */
export function ShowStage({ slide, size = null, hidden = false, onShown }) {
  const [layers, setLayers] = useState([]);
  const w = size?.width || 16;
  const h = size?.height || 9;
  const stageRef = useRef(null);
  const runRef = useRef(null);
  const shownRef = useRef(onShown);
  shownRef.current = onShown;

  useEffect(() => {
    if (!slide) return;
    setLayers((prev) => {
      const current = prev[prev.length - 1];
      const next = { key: `slide-${slide.index}`, index: slide.index, svg: slide.svg };
      if (current && current.index === slide.index) {
        return current.svg === slide.svg ? prev : [...prev.slice(0, -1), { ...current, svg: slide.svg }];
      }
      // The show opens on its first slide's transition from black; after
      // that, only a step forward plays one.
      const forward = !current || slide.index > current.index;
      if (!forward || !plays(slide.transition)) return [next];
      return current ? [{ ...current, role: 'from', transition: null }, { ...next, role: 'to', transition: slide.transition }] : [{ ...next, role: 'to', transition: slide.transition }];
    });
  }, [slide?.index, slide?.svg, slide?.transition]);

  useLayoutEffect(() => {
    const top = layers[layers.length - 1];
    if (!top) return undefined;
    if (!top.transition) {
      shownRef.current?.(top.index);
      return undefined;
    }
    const stage = stageRef.current;
    runRef.current?.cancel();
    const run = playTransition({
      stage,
      fromEl: stage?.querySelector('[data-layer="from"]') || null,
      toEl: stage?.querySelector('[data-layer="to"]') || null,
      transition: top.transition,
    });
    runRef.current = run;
    let live = true;
    run.finished.then(() => {
      if (!live) return;
      setLayers((prev) => (prev[prev.length - 1]?.key === top.key ? [{ key: top.key, index: top.index, svg: prev[prev.length - 1].svg }] : prev));
    });
    return () => {
      live = false;
    };
  }, [layers]);

  // Leaving the show mid-effect leaves nothing running.
  useEffect(() => () => runRef.current?.cancel(), []);

  return (
    <div
      className="sl-show-stage"
      ref={stageRef}
      style={{
        // As large as the screen allows at the deck's own shape.
        aspectRatio: `${w} / ${h}`,
        width: `min(100vw, ${((100 * w) / h).toFixed(3)}vh)`,
        visibility: hidden ? 'hidden' : undefined,
      }}
    >
      {layers.map((l) => (
        <div
          key={l.key}
          className={`sl-show-layer${l.role === 'to' ? ' in' : l.role === 'from' ? ' out' : ''}`}
          data-layer={l.role || 'shown'}
          data-slide={l.index}
          data-transition={l.role === 'to' ? l.transition?.type : undefined}
          dangerouslySetInnerHTML={{ __html: l.svg }}
        />
      ))}
    </div>
  );
}

/**
 * Transitions → Preview on the editing stage: the slide before this one
 * (its picture from the strip, or black for the first) and this slide,
 * played through this slide's transition over the stage, then gone.
 */
export function TransitionPreview({ fromSvg, toSvg, transition, onDone }) {
  const ref = useRef(null);
  const doneRef = useRef(onDone);
  doneRef.current = onDone;
  // The transition as it was when Preview was pressed: a change made while
  // it plays (a new duration typed in) waits for the next Preview.
  const playing = useRef(transition);
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return undefined;
    const run = playTransition({ stage: el, fromEl: el.querySelector('[data-layer="from"]'), toEl: el.querySelector('[data-layer="to"]'), transition: playing.current });
    let live = true;
    run.finished.then(() => live && doneRef.current?.());
    return () => {
      live = false;
      run.cancel();
    };
  }, []);
  return (
    <div className="sl-preview" ref={ref} data-transition={transition?.type}>
      <div className="sl-preview-layer" data-layer="from" dangerouslySetInnerHTML={{ __html: fromSvg || '' }} />
      <div className="sl-preview-layer" data-layer="to" dangerouslySetInnerHTML={{ __html: toSvg || '' }} />
    </div>
  );
}
