// Typing on the canvas (docs/STUDIO.md 7.2, I8): a text layer's words edited
// in place on the picture, as photo editors and slide programs edit them.
//
// A double-click on a text layer (or a new Text from Insert) lays a text box
// exactly over the layer, in its own font, size, weight, slant, colour,
// alignment, line spacing and turn, at the window's scale, while the engine
// leaves that layer out of the picture; so the words are typed where they
// will be, with the caret, a selection by mouse or Shift+arrows, the
// clipboard and the language's own input, all of it the platform's. Each
// change goes into the layer at once (one undo step for the whole spell of
// typing), and the box follows the engine's own layout of the layer, so it
// grows and wraps as the picture will. The type bar above the picture sets
// the layer's font, size, bold, italic, alignment and colour while typing.
// Escape, Ctrl+Enter, Done, or a click away from the words finishes.

import React, { useEffect, useLayoutEffect, useRef } from 'react';
import { Button, Icon, t } from '@rutba/office-ui';
import { faceOf, stackOf, inkOf, sizeOf, sizeFracOf, isBold } from '@rutba/studio/type';
import { FONT_CHOICES } from '../slides/design.js';

/**
 * The text box over the layer. `layer` is the compiled layer from the
 * window's plan (its box and look in the plan's pixels), `patch` the stored
 * one; `cssPerPx` turns the plan's pixels into the window's.
 */
export function TextEditor({ layer, bounds, patch, cssPerPx, selectAll = false, onText, onStyle = null, onDone }) {
  const ref = useRef(null);
  const chosen = useRef(selectAll);
  const k = cssPerPx;
  const sizePx = layer.sizePx || 24;

  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    el.focus();
    if (chosen.current) {
      el.select();
      chosen.current = false;
    } else if (document.activeElement === el && el.selectionStart === 0 && el.selectionEnd === 0) {
      // Opened on words already there: the caret at their end, as a double-click on a text box puts it.
      el.setSelectionRange(el.value.length, el.value.length);
    }
  }, []);

  // A click anywhere but the words or the type bar finishes typing.
  useEffect(() => {
    const away = (e) => {
      const el = ref.current;
      if (!el || el.contains(e.target) || e.target.closest?.('.im-type-bar, .rw-menu, .rw-popover, .rw-dialog')) return;
      onDone();
    };
    window.addEventListener('mousedown', away, true);
    return () => window.removeEventListener('mousedown', away, true);
  }, [onDone]);

  const align = layer.align || 'left';
  const style = {
    left: bounds.x * k,
    top: bounds.y * k,
    width: bounds.w * k,
    height: Math.max(bounds.h, sizePx * (layer.lineHeight || 1.35)) * k,
    font: `${layer.italic ? 'italic ' : ''}${layer.weight || 500} ${sizePx * k}px ${layer.family || stackOf('Segoe UI')}`,
    lineHeight: String(layer.lineHeight || 1.35),
    letterSpacing: layer.tracking ? `${layer.tracking * sizePx * k}px` : undefined,
    textAlign: align,
    color: inkOf(layer),
    caretColor: inkOf(layer),
    transform: layer.rot ? `rotate(${layer.rot}deg)` : undefined,
    paddingLeft: layer.list && layer.list !== 'none' ? (layer.indent || 1.35) * sizePx * k : 0,
  };
  return (
    <textarea
      ref={ref}
      className="im-type-box"
      style={style}
      value={patch.text || ''}
      dir="auto"
      spellCheck
      data-role="type-box"
      data-layer={patch.id}
      aria-label={t('The layer\'s words')}
      onChange={(e) => onText(e.target.value)}
      onKeyDown={(e) => {
        if (e.key === 'Escape' || (e.key === 'Enter' && (e.ctrlKey || e.metaKey))) {
          e.preventDefault();
          e.stopPropagation();
          onDone();
        } else if ((e.ctrlKey || e.metaKey) && !e.altKey && onStyle && /^[bi]$/i.test(e.key)) {
          // Bold and italic from the keyboard, as in any text box.
          e.preventDefault();
          e.stopPropagation();
          if (e.key.toLowerCase() === 'b') onStyle({ weight: isBold(patch) ? 500 : 700 });
          else onStyle({ italic: !patch.italic });
        }
      }}
    />
  );
}

const SIZES = [12, 16, 20, 24, 32, 40, 48, 64, 72, 96, 128, 160, 200];

/** The type bar: the layer's font, size, bold, italic, alignment and colour, while its words are typed. */
export function TypeBar({ patch, canvasWidth, onChange, onDone, refocus }) {
  const face = faceOf(patch.family);
  const faces = FONT_CHOICES.includes(face) ? FONT_CHOICES : [face, ...FONT_CHOICES];
  const size = sizeOf(patch, canvasWidth);
  const bold = isBold(patch);
  const align = patch.align || 'left';
  // The text box keeps its caret and selection: every control here hands the focus back.
  const keep = (e) => e.preventDefault();
  const set = (change, label) => {
    onChange(change, label);
    refocus();
  };
  return (
    <div className="im-options im-type-bar" data-role="type-bar">
      <span className="im-opt-tool"><Icon name="mkText" size={14} /><span>{t('Type')}</span></span>
      <label className="im-opt" data-opt="font">
        <span className="im-opt-label">{t('Font')}</span>
        <select className="rw-select im-type-font" value={face} data-role="type-font" style={{ fontFamily: stackOf(face) }} onChange={(e) => set({ family: stackOf(e.target.value) }, t('Font'))}>
          {faces.map((f) => <option key={f} value={f} style={{ fontFamily: stackOf(f) }}>{f}</option>)}
        </select>
      </label>
      <label className="im-opt" data-opt="size">
        <span className="im-opt-label">{t('Size')}</span>
        <input
          className="rw-input im-type-size"
          type="number"
          min={8}
          max={2000}
          list="im-type-sizes"
          value={size}
          data-role="type-size"
          onChange={(e) => { const v = Number(e.target.value); if (v >= 8) onChange({ sizeFrac: sizeFracOf(v, canvasWidth) }, t('Size')); }}
          onBlur={refocus}
          onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); refocus(); } }}
        />
        <datalist id="im-type-sizes">{SIZES.map((s) => <option key={s} value={s} />)}</datalist>
        <span className="im-opt-unit">{t('px')}</span>
      </label>
      <span className="im-opt-sep" />
      <Button icon="bold" title={t('Bold (Ctrl+B)')} pressed={bold} onMouseDown={keep} onClick={() => set({ weight: bold ? 500 : 700 }, t('Bold'))} data-role="type-bold" />
      <Button icon="italic" title={t('Italic (Ctrl+I)')} pressed={Boolean(patch.italic)} onMouseDown={keep} onClick={() => set({ italic: !patch.italic }, t('Italic'))} data-role="type-italic" />
      <span className="im-opt-sep" />
      {[['left', 'alignLeft', t('Align left')], ['center', 'alignCenter', t('Centre')], ['right', 'alignRight', t('Align right')]].map(([v, icon, tip]) => (
        <Button key={v} icon={icon} title={tip} pressed={align === v} onMouseDown={keep} onClick={() => set({ align: v }, t('Align'))} data-role={`type-align-${v}`} />
      ))}
      <span className="im-opt-sep" />
      <label className="im-opt im-type-colour" data-opt="colour" data-tip={t('Colour of the words')}>
        <span className="im-type-swatch" style={{ background: inkOf(patch) }} />
        <input type="color" value={/^#[0-9a-f]{6}$/i.test(inkOf(patch)) ? inkOf(patch) : '#ffffff'} data-role="type-colour" onChange={(e) => onChange({ fill: { kind: 'solid', color: e.target.value } }, t('Colour'))} onBlur={refocus} aria-label={t('Colour of the words')} />
      </label>
      <span className="im-opt-grow" />
      <Button icon="check" label={t('Done')} title={t('Done: stop typing (Esc)')} onMouseDown={keep} onClick={onDone} data-role="type-done" />
    </div>
  );
}

export const TYPE_CSS = `
.im-type-box { position: absolute; box-sizing: content-box; margin: 0; padding: 0; border: 0; outline: 1px dashed var(--accent); outline-offset: 3px; background: transparent; resize: none; overflow: hidden; white-space: pre-wrap; overflow-wrap: break-word; transform-origin: 50% 50%; z-index: 4; }
.im-type-box::selection { background: rgba(38, 110, 220, 0.55); color: #ffffff; }
.im-type-bar .im-type-font { width: 168px; }
.im-type-bar .im-type-size { width: 78px; }
.im-type-bar .im-opt-unit { color: var(--ink-3); font-size: 12px; }
.im-type-bar .im-opt-tool { display: inline-flex; align-items: center; gap: 6px; font-weight: 600; color: var(--ink); padding-inline-end: 4px; }
.im-type-bar .im-opt-sep { width: 1px; align-self: stretch; margin: 4px 2px; background: var(--line-soft); }
.im-type-bar .im-opt-grow { flex: 1; }
.im-type-colour { position: relative; display: inline-flex; align-items: center; cursor: pointer; }
.im-type-swatch { width: 22px; height: 22px; border-radius: 4px; border: 1px solid var(--line); box-shadow: inset 0 0 0 1px rgba(255,255,255,0.6); }
.im-type-colour input[type="color"] { position: absolute; inset: 0; opacity: 0; width: 100%; height: 100%; cursor: pointer; }
`;
