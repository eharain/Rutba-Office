// The slideshow's stage, and the editor's Previews.
//
// The show keeps the slide it is leaving on screen while the next one comes
// in: two layers in one box, the transition played between them (see
// motion.js), then the old layer dropped. Moving forward plays the
// incoming slide's own transition, as PowerPoint does; stepping back cuts.
//
// On each slide the animations play a click group at a time (animate.js):
// shapes that enter start hidden — already hidden while the slide's
// transition plays — the group that starts with the slide plays once the
// transition is over, and each step the show is told about plays the next
// group, or puts the shapes straight where that step has them.

import React, { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { playTransition } from './motion.js';
import { sequence, applyState, playGroup } from './animate.js';
import { Markup, FILL } from './markup.js';

const plays = (t) => Boolean(t && t.type && t.type !== 'none' && Number(t.duration) > 0);

/**
 * @param {{
 *   slide: { index: number, svg: string, transition?: object, animations?: object[], shapes?: object[] },
 *   size?: { width: number, height: number },
 *   step?: { index: number, value: number },
 *   hidden?: boolean,
 *   onSettled?: (state: { index: number, step: number, clicks: number }) => void,
 *   control?: { current: any },
 * }} props
 * `step` is how many clicks of the slide at `step.index` have played; it
 * is only acted on once that slide is the one on the stage. `onSettled` is
 * told whenever nothing is moving any more — the transition over and the
 * group that was playing done — which is when Advance Slide → After counts.
 * `control.current.finish()` jumps whatever is moving to its end and says
 * whether anything was, so a click during an animation completes it.
 */
export function ShowStage({ slide, size = null, step = null, hidden = false, onSettled, control }) {
  const [layers, setLayers] = useState([]);
  const w = size?.width || 16;
  const h = size?.height || 9;
  const stageRef = useRef(null);
  const runRef = useRef(null);
  const groupRef = useRef(null);
  const phase = useRef({ key: null, index: null, step: 0, entering: false, autoDone: true, svg: null });
  const settledRef = useRef(onSettled);
  settledRef.current = onSettled;
  const slideRef = useRef(slide);
  slideRef.current = slide;

  const topEl = () => stageRef.current?.querySelector('.sl-show-layer:last-child') || null;
  const seqOf = () => sequence(slideRef.current?.animations || []);
  const settle = () => {
    const p = phase.current;
    settledRef.current?.({ index: p.index, step: p.step, clicks: seqOf().clicks.length });
  };
  const runGroup = (group, then) => {
    groupRef.current?.finish();
    const player = playGroup(topEl(), group, { shapes: slideRef.current?.shapes || [], size });
    groupRef.current = player;
    player.finished.then(() => {
      if (groupRef.current !== player) return;
      groupRef.current = null;
      then?.();
      settle();
    });
  };
  // The slide is on screen: the group that starts with it plays now.
  const entered = () => {
    const p = phase.current;
    p.entering = false;
    const seq = seqOf();
    if (!p.autoDone && seq.auto) {
      runGroup(seq.auto, () => { p.autoDone = true; });
    } else {
      p.autoDone = true;
      settle();
    }
  };

  if (control) {
    control.current = {
      finish() {
        let moving = false;
        if (runRef.current && phase.current.entering) { runRef.current.cancel(); moving = true; }
        if (groupRef.current?.running()) { groupRef.current.finish(); moving = true; }
        return moving;
      },
    };
  }

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

  // A new slide on top: its shapes put where the sequence starts (entering
  // shapes hidden) before the first frame, then its transition, then the
  // group that starts with it.
  useLayoutEffect(() => {
    const top = layers[layers.length - 1];
    if (!top) return undefined;
    const p = phase.current;
    const el = topEl();
    const isNew = p.key !== top.key;
    if (isNew) {
      groupRef.current?.finish();
      groupRef.current = null;
      const seq = seqOf();
      const at = step && step.index === top.index ? Math.min(Math.max(0, step.value), seq.clicks.length) : 0;
      phase.current = { key: top.key, index: top.index, step: at, entering: true, autoDone: at > 0 || !seq.auto, svg: top.svg };
      applyState(el, slideRef.current?.animations || [], at, at > 0);
    } else if (p.svg !== top.svg) {
      // The same slide drawn again: the shapes put back where the show has them.
      p.svg = top.svg;
      applyState(el, slideRef.current?.animations || [], p.step, p.autoDone);
    }
    if (!top.transition) {
      if (phase.current.entering) entered();
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
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [layers]);

  // A step on the slide that is showing: one more click plays its group;
  // any other change puts the shapes straight where that step has them.
  useLayoutEffect(() => {
    const p = phase.current;
    if (!step || step.index !== p.index || p.key == null) return;
    const seq = seqOf();
    const target = Math.min(Math.max(0, step.value), seq.clicks.length);
    if (target === p.step) return;
    if (p.entering) {
      runRef.current?.cancel();
      p.entering = false;
    }
    if (target === p.step + 1) {
      // Whatever was still moving — the slide's own opening group, or the
      // click before — ends where it was going, then the next click's plays.
      groupRef.current?.finish();
      if (!p.autoDone) {
        applyState(topEl(), slideRef.current?.animations || [], p.step, true);
        p.autoDone = true;
      }
      p.step = target;
      runGroup(seq.clicks[target - 1]);
    } else {
      groupRef.current?.finish();
      groupRef.current = null;
      p.step = target;
      p.autoDone = true;
      applyState(topEl(), slideRef.current?.animations || [], target, true);
      settle();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [step?.index, step?.value]);

  // Leaving the show mid-effect leaves nothing running.
  useEffect(() => () => {
    runRef.current?.cancel();
    groupRef.current?.finish();
  }, []);

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
        >
          <Markup html={l.svg} style={FILL} />
        </div>
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
      <div className="sl-preview-layer" data-layer="from">{fromSvg ? <Markup html={fromSvg} style={FILL} /> : null}</div>
      <div className="sl-preview-layer" data-layer="to"><Markup html={toSvg} style={FILL} /></div>
    </div>
  );
}

/**
 * Animations → Preview on the editing stage: the slide with its entering
 * shapes hidden, then every group in turn — the one that starts with the
 * slide, then each click's, a short beat apart — then gone. `only` plays
 * just one effect (by its index in the list): a pick from the gallery
 * previews the effect it made, the way PowerPoint's does.
 */
export function AnimationPreview({ slide, size, only = null, onDone }) {
  const ref = useRef(null);
  const doneRef = useRef(onDone);
  doneRef.current = onDone;
  const playing = useRef({ animations: slide?.animations || [], shapes: slide?.shapes || [], svg: slide?.svg || '' });
  useLayoutEffect(() => {
    const el = ref.current?.firstElementChild;
    if (!el) return undefined;
    const { animations, shapes } = playing.current;
    const list = only == null ? animations : animations.filter((e) => e.index === only).map((e) => ({ ...e, trigger: 'onClick', delay: 0 }));
    const seq = sequence(list);
    const groups = [...(seq.auto ? [seq.auto] : []), ...seq.clicks];
    applyState(el, list, 0, false);
    let live = true;
    let player = null;
    let timer = null;
    const next = (i) => {
      if (!live) return;
      if (i >= groups.length) { timer = setTimeout(() => live && doneRef.current?.(), 400); return; }
      player = playGroup(el, groups[i], { shapes, size });
      player.finished.then(() => { if (live) timer = setTimeout(() => next(i + 1), 250); });
    };
    timer = setTimeout(() => next(0), 150);
    return () => {
      live = false;
      clearTimeout(timer);
      player?.finish();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  return (
    <div className="sl-preview sl-preview-anim" ref={ref} data-animating="true">
      <div className="sl-preview-layer"><Markup html={playing.current.svg} style={FILL} /></div>
    </div>
  );
}
