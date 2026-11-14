// The small pieces the studio's panels share: the picture's slot, the two-way
// switch, a level meter, and the two little dialogs.

import React, { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { Button, Dialog, Field, Input, t } from '@rutba/office-ui';

/** Holds a canvas the program draws on and records from, so it goes on working when the panel around it changes. */
export function CanvasSlot({ canvas }) {
  const ref = useRef(null);
  useLayoutEffect(() => {
    const el = ref.current;
    if (el && canvas.parentNode !== el) el.appendChild(canvas);
  }, [canvas]);
  return <div ref={ref} className="st-canvas-slot" />;
}

/** The picture at its real shape, drawn over by what the panel needs. */
export function Frame({ canvas, frameRef, label, className = '', children, ...rest }) {
  return (
    <div ref={frameRef} aria-label={label} className={`st-frame ${className}`.trim()} dir="ltr" {...rest}>
      <CanvasSlot canvas={canvas} />
      {children}
    </div>
  );
}

/**
 * The panels, as Office's tabs: one Tab stop, the arrow keys move along it.
 * `value` is the open one's key; options are [key, label, tip].
 */
export function PanelSwitch({ value, options, onChange, label }) {
  const refs = useRef({});
  const onKey = (e, i) => {
    let next = null;
    if (e.key === 'ArrowRight' || e.key === 'ArrowDown') next = (i + 1) % options.length;
    else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') next = (i - 1 + options.length) % options.length;
    else if (e.key === 'Home') next = 0;
    else if (e.key === 'End') next = options.length - 1;
    if (next == null) return;
    e.preventDefault();
    const [key] = options[next];
    onChange(key);
    refs.current[key]?.focus();
  };
  return (
    <div className="st-seg st-panels" role="tablist" aria-label={label}>
      {options.map(([key, text, tip], i) => (
        <button key={key} ref={(n) => { refs.current[key] = n; }} type="button" role="tab" aria-selected={value === key} tabIndex={value === key ? 0 : -1}
          className={value === key ? 'on' : ''} data-panel={key} data-tip={tip || text} onClick={() => onChange(key)} onKeyDown={(e) => onKey(e, i)}>{text}</button>
      ))}
    </div>
  );
}

/** A small row of choices, one on. Options are [value, label]. */
export function Seg({ value, options, onChange, label, className = '', disabled }) {
  return (
    <div className={`st-seg small ${className}`.trim()} role="radiogroup" aria-label={label}>
      {options.map(([v, text]) => (
        <button key={v} type="button" role="radio" aria-checked={value === v} data-value={v} className={value === v ? 'on' : ''} disabled={disabled} onClick={() => onChange(v)}>{text}</button>
      ))}
    </div>
  );
}

/** A small copy of a canvas, drawn a few times a second: On air now, over Set up. */
export function Mirror({ program, label }) {
  const ref = useRef(null);
  useEffect(() => {
    if (!program) return undefined;
    let n = 0;
    return program.onFrame((src) => {
      n += 1;
      if (n % 6 !== 0) return;
      const c = ref.current;
      if (!c) return;
      c.width = 240;
      c.height = 135;
      c.getContext('2d').drawImage(src, 0, 0, 240, 135);
    });
  }, [program]);
  return <canvas ref={ref} className="st-mirror-canvas" width={240} height={135} aria-label={label} />;
}

/** A sound's meter, read ten times a second. */
export function Meter({ media, id, live, className = '' }) {
  const bar = useRef(null);
  useEffect(() => {
    if (!live) { if (bar.current) bar.current.style.width = '0%'; return undefined; }
    const timer = setInterval(() => {
      const v = media.level(id);
      const el = bar.current;
      if (!el) return;
      el.style.width = `${Math.round(v * 100)}%`;
      el.dataset.level = v > 0.9 ? 'loud' : v > 0.02 ? 'on' : 'quiet';
    }, 100);
    return () => clearInterval(timer);
  }, [media, id, live]);
  return <span className={`rec-meter st-meter ${className}`.trim()} role="meter" aria-label={t('studio.level')} aria-valuemin={0} aria-valuemax={100}><i ref={bar} /></span>;
}

/** Asks for a name. */
export function NameDialog({ title, label, value = '', okLabel, onOk, onClose }) {
  const [text, setText] = useState(value);
  const ok = () => { const v = text.trim(); if (v) onOk(v); };
  return (
    <Dialog title={title} width={380} className="st-name-dialog" onClose={onClose} actions={<>
      <Button label={t('common.cancel')} onClick={onClose} />
      <Button primary label={okLabel} className="st-name-ok" disabled={!text.trim()} onClick={ok} />
    </>}>
      <Field label={label}><Input className="rw-input st-name-input" autoFocus value={text} onChange={(e) => setText(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); ok(); } }} /></Field>
    </Dialog>
  );
}

/** Asks before something that cannot be undone. */
export function ConfirmDialog({ title, body, okLabel, onOk, onClose, tone = 'bad' }) {
  return (
    <Dialog title={title} width={420} className="st-confirm-dialog" onClose={onClose} actions={<>
      <Button label={t('common.cancel')} onClick={onClose} />
      <Button primary label={okLabel} className={`st-confirm-ok tone-${tone}`} onClick={() => { onOk(); onClose(); }} />
    </>}>
      <p style={{ margin: 0, lineHeight: 1.5 }}>{body}</p>
    </Dialog>
  );
}
