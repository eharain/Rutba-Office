// Design → Themes, Variants, Colours, Fonts and Effects.
//
// The galleries ask the deck for its choices and, for themes and variants,
// for the slide on screen drawn under each one — by the same renderer that
// draws the stage — so a theme is picked by looking at this deck in it, not
// at a stock picture. A pick is one operation on the deck (one undo step);
// the theme part is rewritten and every slide follows.

import React, { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { Button, Dialog, Icon } from '@rutba/office-ui';
import { Markup } from './markup.js';

/** The twelve slots, in Customise Colours' order, with PowerPoint's own words for them. */
export const COLOUR_ROWS = [
  ['dk1', 'Text/Background – Dark 1'],
  ['lt1', 'Text/Background – Light 1'],
  ['dk2', 'Text/Background – Dark 2'],
  ['lt2', 'Text/Background – Light 2'],
  ['accent1', 'Accent 1'],
  ['accent2', 'Accent 2'],
  ['accent3', 'Accent 3'],
  ['accent4', 'Accent 4'],
  ['accent5', 'Accent 5'],
  ['accent6', 'Accent 6'],
  ['hlink', 'Hyperlink'],
  ['folHlink', 'Followed Hyperlink'],
];

/** Faces offered by Customise Fonts: ones Windows ships, most of them on macOS too. Any other name can be typed. */
export const FONT_CHOICES = [
  'Arial', 'Bahnschrift', 'Calibri', 'Cambria', 'Candara', 'Century Gothic', 'Consolas', 'Constantia', 'Corbel',
  'Courier New', 'Franklin Gothic Book', 'Franklin Gothic Medium', 'Garamond', 'Georgia', 'Gill Sans MT',
  'Palatino Linotype', 'Segoe UI', 'Segoe UI Light', 'Segoe UI Semibold', 'Tahoma', 'Times New Roman',
  'Trebuchet MS', 'Verdana',
];

const hex = (v) => `#${String(v || '000000').replace('#', '')}`;
/** A face and the kind of face to fall back to, for the window's own previews. */
const stack = (font) => `"${font}", ${/georgia|times|palatino|cambria|constantia|garamond|book antiqua/i.test(font) ? 'Georgia, serif' : '"Segoe UI", system-ui, sans-serif'}`;

/** A theme's six accents as a thin bar, under its picture. */
function Accents({ colors }) {
  if (!colors) return null;
  return (
    <span className="sl-accents" aria-hidden="true">
      {[1, 2, 3, 4, 5, 6].map((n) => <span key={n} style={{ background: hex(colors[`accent${n}`]) }} />)}
    </span>
  );
}

/** The swatches a palette is known by: the second dark and light, then the six accents. */
function Swatches({ colors, size = 16 }) {
  const keys = ['dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
  return (
    <span className="sl-dg-swatches">
      {keys.map((k) => <span key={k} className="sl-dg-swatch" style={{ width: size, height: size, background: hex(colors?.[k]) }} />)}
    </span>
  );
}

/**
 * A popover gallery under the button that opened it: the deck's choices
 * for one kind, the current one ringed. Escape or a click elsewhere closes
 * it; a pick closes it and hands the choice to `onPick`.
 */
export function DesignGallery({ kind, anchor, shell, docId, slide, onPick, onCustomise, onClose }) {
  const [data, setData] = useState(null);
  const [error, setError] = useState(null);
  const ref = useRef(null);
  const [pos, setPos] = useState({ left: anchor?.left ?? 100, top: anchor?.bottom ?? 100 });

  useEffect(() => {
    let alive = true;
    shell.doc.deckDesign({ id: docId, slide, kind, width: kind === 'themes' ? 184 : 196 })
      .then((d) => { if (alive) setData(d); })
      .catch((err) => { if (alive) setError(err.message); });
    return () => { alive = false; };
  }, [shell, docId, slide, kind]);

  useEffect(() => {
    const onKey = (e) => { if (e.key === 'Escape') onClose(); };
    const onDown = (e) => { if (ref.current && !ref.current.contains(e.target)) onClose(); };
    window.addEventListener('keydown', onKey);
    window.addEventListener('mousedown', onDown, true);
    return () => {
      window.removeEventListener('keydown', onKey);
      window.removeEventListener('mousedown', onDown, true);
    };
  }, [onClose]);

  // Kept on screen: a gallery opened near the right edge moves left.
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    const left = Math.max(8, Math.min(anchor?.left ?? 100, window.innerWidth - r.width - 8));
    const top = Math.min(anchor?.bottom ?? 100, Math.max(8, window.innerHeight - r.height - 8));
    if (left !== pos.left || top !== pos.top) setPos({ left, top });
  });

  const title = { themes: 'Themes', variants: 'Variants', colours: 'Colours', fonts: 'Fonts', effects: 'Effects' }[kind];
  const items = data?.items || [];
  return (
    <div className={`sl-dg sl-dg-${kind}`} ref={ref} style={{ left: pos.left, top: pos.top }} role="dialog" aria-label={title}>
      <div className="sl-dg-head">
        <span>{title}</span>
        {data?.info?.name && kind === 'themes' ? <span className="sl-dg-sub">This presentation: {data.info.name}</span> : null}
        {kind === 'colours' && data?.info?.colorName ? <span className="sl-dg-sub">Now: {data.info.colorName}</span> : null}
        {kind === 'fonts' && data?.info?.fonts ? <span className="sl-dg-sub">Now: {data.info.fonts.major} / {data.info.fonts.minor}</span> : null}
      </div>
      {error ? <div className="sl-dg-empty">{error}</div> : !data ? <div className="sl-dg-empty">Drawing this slide in each one…</div> : null}
      {kind === 'themes' || kind === 'variants' ? (
        <div className="sl-dg-grid">
          {items.map((it) => (
            <button
              key={it.id}
              type="button"
              className={`sl-dg-card${it.current ? ' current' : ''}`}
              data-theme={kind === 'themes' ? it.id : undefined}
              data-variant={kind === 'variants' ? it.id : undefined}
              data-tip={kind === 'themes' ? `${it.name} — ${it.fonts.major} and ${it.fonts.minor}${it.dark ? ', dark' : ''}` : `${it.name}${it.dark ? ' (dark)' : ''}`}
              onClick={() => onPick(it)}
            >
              {it.svg ? <Markup as="span" className="sl-dg-pic" html={it.svg} /> : <span className="sl-dg-pic sl-dg-nopic" />}
              <span className="sl-dg-name">{it.current ? <Icon name="check" size={12} /> : null}<span className="sl-dg-label">{it.name}</span><Accents colors={it.colors} /></span>
            </button>
          ))}
        </div>
      ) : null}
      {kind === 'colours' ? (
        <div className="sl-dg-list">
          {items.map((it) => (
            <button key={it.id} type="button" className={`sl-dg-row${it.current ? ' current' : ''}`} data-palette={it.id} onClick={() => onPick(it)}>
              <Swatches colors={it.colors} />
              <span className="sl-dg-rowname">{it.name}</span>
              {it.current ? <Icon name="check" size={13} /> : null}
            </button>
          ))}
        </div>
      ) : null}
      {kind === 'fonts' ? (
        <div className="sl-dg-list">
          {items.map((it) => (
            <button key={it.id} type="button" className={`sl-dg-row sl-dg-fontrow${it.current ? ' current' : ''}`} data-pair={it.id} onClick={() => onPick(it)}>
              <span className="sl-dg-aa" style={{ fontFamily: stack(it.major) }}>Aa</span>
              <span className="sl-dg-faces">
                <span className="sl-dg-rowname">{it.name}</span>
                <span className="sl-dg-major" style={{ fontFamily: stack(it.major) }}>{it.major}</span>
                <span className="sl-dg-minor" style={{ fontFamily: stack(it.minor) }}>{it.minor}</span>
              </span>
              {it.current ? <Icon name="check" size={13} /> : null}
            </button>
          ))}
        </div>
      ) : null}
      {kind === 'effects' ? (
        <div className="sl-dg-grid sl-dg-effects">
          {items.map((it) => (
            <button key={it.id} type="button" className={`sl-dg-card${it.current ? ' current' : ''}`} data-effects={it.id} data-tip={`${it.name} — ${it.description}`} onClick={() => onPick(it)}>
              {it.svg ? <Markup as="span" className="sl-dg-pic" html={it.svg} /> : <span className="sl-dg-pic sl-dg-nopic" />}
              <span className="sl-dg-name">{it.current ? <Icon name="check" size={12} /> : null}{it.name}</span>
            </button>
          ))}
        </div>
      ) : null}
      {kind === 'colours' || kind === 'fonts' ? (
        <div className="sl-dg-foot">
          <button type="button" className="sl-dg-custom" onClick={() => onCustomise(data?.info || null)}>
            <Icon name="settings" size={14} /> {kind === 'colours' ? 'Customise Colours…' : 'Customise Fonts…'}
          </button>
        </div>
      ) : null}
      {kind === 'effects' ? <div className="sl-dg-foot sl-dg-note">Effects reach shapes styled from the theme — PowerPoint's shape styles are.</div> : null}
    </div>
  );
}

/** Customise Colours: the twelve slots, a sample of the palette at work, and a name to save it under. */
export function CustomColoursDialog({ info, onClose, onSave }) {
  const start = () => Object.fromEntries(COLOUR_ROWS.map(([k]) => [k, hex(info?.colors?.[k] || '000000')]));
  const [colors, setColors] = useState(start);
  const [name, setName] = useState(info?.colorName && !/^Custom/.test(info.colorName) ? 'Custom 1' : info?.colorName || 'Custom 1');
  const set = (k, v) => setColors((c) => ({ ...c, [k]: v }));
  const valid = (v) => /^#[0-9a-f]{6}$/i.test(v);
  const sample = (dark) => {
    const bg = dark ? colors.dk1 : colors.lt1;
    const ink = dark ? colors.lt1 : colors.dk1;
    const bg2 = dark ? colors.dk2 : colors.lt2;
    return (
      <svg viewBox="0 0 160 96" className="sl-cc-sample" aria-hidden="true">
        <rect x="0" y="0" width="160" height="96" fill={bg} />
        <rect x="0" y="70" width="160" height="26" fill={bg2} />
        <text x="10" y="24" fontSize="15" fontWeight="600" fill={ink}>Text</text>
        <text x="10" y="42" fontSize="9" fill={ink}>Body words</text>
        <text x="62" y="42" fontSize="9" fill={colors.hlink} textDecoration="underline">Link</text>
        {[1, 2, 3, 4, 5, 6].map((n, i) => <rect key={n} x={10 + i * 23} y={50 + (i % 2) * 0} width="18" height={30 - (i % 3) * 5} fill={colors[`accent${n}`]} />)}
      </svg>
    );
  };
  const ok = COLOUR_ROWS.every(([k]) => valid(colors[k])) && name.trim();
  return (
    <Dialog
      title="Create new theme colours"
      width={560}
      onClose={onClose}
      actions={
        <>
          <Button label="Reset" className="sl-cc-reset" onClick={() => setColors(start())} />
          <span style={{ flex: 1 }} />
          <Button label="Cancel" onClick={onClose} />
          <Button primary label="Save" className="sl-cc-save" disabled={!ok} onClick={() => onSave(Object.fromEntries(Object.entries(colors).map(([k, v]) => [k, v.replace('#', '').toUpperCase()])), name.trim())} />
        </>
      }
    >
      <div className="sl-cc">
        <div className="sl-cc-slots">
          <div className="sl-cc-caption">Theme colours</div>
          {COLOUR_ROWS.map(([k, label]) => (
            <label key={k} className="sl-cc-slot" data-slot={k}>
              <span className="sl-cc-label">{label}</span>
              <input type="color" className={`sl-cc-pick sl-cc-pick-${k}`} value={valid(colors[k]) ? colors[k] : '#000000'} onChange={(e) => set(k, e.target.value)} />
              <input className={`rw-input sl-cc-hex sl-cc-hex-${k}`} value={colors[k]} maxLength={7} spellCheck={false} onChange={(e) => set(k, e.target.value.startsWith('#') ? e.target.value : `#${e.target.value}`)} />
            </label>
          ))}
        </div>
        <div className="sl-cc-side">
          <div className="sl-cc-caption">Sample</div>
          <div className="sl-cc-samples">{sample(false)}{sample(true)}</div>
          <label className="sl-cc-name">
            <span>Name</span>
            <input className="rw-input sl-cc-namefield" value={name} onChange={(e) => setName(e.target.value)} />
          </label>
          <p className="rw-hint" style={{ margin: 0 }}>Saved into this deck's theme: every slide that takes its colours from the theme follows, and Undo puts the old ones back.</p>
        </div>
      </div>
    </Dialog>
  );
}

/** Customise Fonts: a heading face and a body face, a sample of the two, and a name. */
export function CustomFontsDialog({ info, onClose, onSave }) {
  const [major, setMajor] = useState(info?.fonts?.major || 'Segoe UI Semibold');
  const [minor, setMinor] = useState(info?.fonts?.minor || 'Segoe UI');
  const [name, setName] = useState(info?.fontName && !/^Custom/.test(info.fontName) ? 'Custom 1' : info?.fontName || 'Custom 1');
  const ok = major.trim() && minor.trim() && name.trim();
  return (
    <Dialog
      title="Create new theme fonts"
      width={520}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button primary label="Save" className="sl-cf-save" disabled={!ok} onClick={() => onSave({ major: major.trim(), minor: minor.trim() }, name.trim())} />
        </>
      }
    >
      <div className="sl-cf">
        <datalist id="sl-cf-fonts">{FONT_CHOICES.map((f) => <option key={f} value={f} />)}</datalist>
        <div className="sl-cf-fields">
          <label className="sl-cf-field"><span>Heading font</span><input className="rw-input sl-cf-major" list="sl-cf-fonts" value={major} onChange={(e) => setMajor(e.target.value)} /></label>
          <label className="sl-cf-field"><span>Body font</span><input className="rw-input sl-cf-minor" list="sl-cf-fonts" value={minor} onChange={(e) => setMinor(e.target.value)} /></label>
          <label className="sl-cf-field"><span>Name</span><input className="rw-input sl-cf-name" value={name} onChange={(e) => setName(e.target.value)} /></label>
        </div>
        <div className="sl-cf-sample">
          <div className="sl-cc-caption">Sample</div>
          <div className="sl-cf-card">
            <div className="sl-cf-heading" style={{ fontFamily: stack(major) }}>Heading</div>
            <div className="sl-cf-body" style={{ fontFamily: stack(minor) }}>Body text body text body text. Body text body text.</div>
          </div>
        </div>
      </div>
    </Dialog>
  );
}

/**
 * The ribbon's own strip of a gallery: a few live thumbnails in a row and
 * a button that opens the whole gallery — PowerPoint's in-ribbon gallery.
 */
export function RibbonStrip({ items = [], kind, count = 4, onPick, onMore, label }) {
  // The deck's own theme leads the strip, the way PowerPoint puts it first.
  const ordered = kind === 'themes' ? [...(items || [])].sort((a, b) => (b.current ? 1 : 0) - (a.current ? 1 : 0)) : items || [];
  const shown = ordered.slice(0, count);
  return (
    <div className={`sl-rs sl-rs-${kind}`}>
      <div className="sl-rs-items">
        {shown.length ? shown.map((it) => (
          <button
            key={it.id}
            type="button"
            className={`sl-rs-item${it.current ? ' current' : ''}`}
            data-tip={kind === 'themes' ? `${it.name} — apply this theme to every slide` : `${it.name} — this theme in these colours`}
            data-theme={kind === 'themes' ? it.id : undefined}
            data-variant={kind === 'variants' ? it.id : undefined}
            onClick={() => onPick(it)}
          >
            {it.svg ? <Markup as="span" className="sl-rs-pic" html={it.svg} /> : <span className="sl-rs-pic sl-dg-nopic" />}
            <Accents colors={it.colors} />
          </button>
        )) : <span className="sl-rs-wait">{label}</span>}
      </div>
      <button type="button" className="sl-rs-more" data-tip={kind === 'themes' ? 'More Themes — every theme, drawn on this slide' : 'More Variants — the four colourings of this theme'} onClick={onMore}>
        <Icon name="chevronDown" size={12} />
      </button>
    </div>
  );
}

export const DESIGN_CSS = `
/* Design → the galleries: a popover under the button, cards of this slide drawn in each choice. */
.sl-dg { position: fixed; z-index: 130; background: var(--surface); border: 1px solid var(--line); border-radius: 12px; box-shadow: var(--shadow-3); padding: 10px 12px 12px; animation: rw-pop 120ms var(--ease); max-height: calc(100vh - 140px); overflow: auto; }
.sl-dg-themes, .sl-dg-variants { width: 836px; max-width: calc(100vw - 24px); }
.sl-dg-colours, .sl-dg-fonts { width: 330px; }
.sl-dg-effects { width: 408px; }
.sl-dg-head { display: flex; align-items: baseline; gap: 10px; padding: 0 2px 8px; font-size: 12px; font-weight: 600; color: var(--ink-2); text-transform: uppercase; letter-spacing: .05em; }
.sl-dg-sub { font-weight: 400; text-transform: none; letter-spacing: 0; color: var(--ink-3); font-size: 12px; margin-left: auto; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sl-dg-empty { padding: 18px 6px; color: var(--ink-3); font-size: 12.5px; }
.sl-dg-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; }
.sl-dg-grid.sl-dg-effects { grid-template-columns: repeat(2, 1fr); width: auto; }
.sl-dg-card { display: flex; flex-direction: column; gap: 6px; padding: 5px; border: 1px solid transparent; border-radius: 9px; background: transparent; cursor: pointer; text-align: left; transition: background var(--fast), border-color var(--fast), box-shadow var(--fast); }
.sl-dg-card:hover { background: var(--hover); border-color: var(--line); }
.sl-dg-card.current { border-color: var(--accent); box-shadow: 0 0 0 2px var(--accent-soft); background: var(--selected); }
.sl-dg-pic { display: block; line-height: 0; border-radius: 5px; overflow: hidden; box-shadow: 0 0 0 1px rgba(0,0,0,.10), 0 1px 3px rgba(15,20,30,.12); background: #fff; }
.sl-dg-pic svg { display: block; width: 100%; height: auto; }
.sl-dg-nopic { aspect-ratio: 16 / 9; background: var(--surface-2); }
.sl-dg-name { display: flex; align-items: center; gap: 4px; font-size: 12px; color: var(--ink); padding: 0 2px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sl-dg-card.current .sl-dg-name { color: var(--accent); font-weight: 600; }
.sl-dg-list { display: flex; flex-direction: column; gap: 1px; }
.sl-dg-row { display: flex; align-items: center; gap: 10px; padding: 6px 8px; border: 0; border-radius: 7px; background: transparent; cursor: pointer; text-align: left; color: var(--ink); font-size: 12.5px; }
.sl-dg-row:hover { background: var(--hover); }
.sl-dg-row.current { background: var(--selected); color: var(--accent); }
.sl-dg-rowname { flex: 1; min-width: 0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sl-dg-swatches { display: inline-flex; gap: 2px; flex: none; padding: 2px; border-radius: 4px; background: var(--surface-2); }
.sl-dg-swatch { display: inline-block; border-radius: 2px; box-shadow: inset 0 0 0 1px rgba(0,0,0,.12); }
.sl-dg-fontrow { align-items: center; }
.sl-dg-aa { width: 44px; flex: none; font-size: 24px; line-height: 1; text-align: center; color: var(--ink); }
.sl-dg-faces { display: flex; flex-direction: column; gap: 1px; flex: 1; min-width: 0; }
.sl-dg-faces .sl-dg-rowname { font-size: 10.5px; text-transform: uppercase; letter-spacing: .05em; color: var(--ink-3); }
.sl-dg-major { font-size: 14px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sl-dg-minor { font-size: 12px; color: var(--ink-2); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sl-dg-row.current .sl-dg-major, .sl-dg-row.current .sl-dg-aa { color: var(--accent); }
.sl-dg-foot { border-top: 1px solid var(--line-soft); margin-top: 8px; padding-top: 6px; }
.sl-dg-note { font-size: 11.5px; color: var(--ink-3); padding: 8px 4px 0; }
.sl-dg-custom { display: flex; align-items: center; gap: 8px; width: 100%; border: 0; background: transparent; padding: 7px 8px; border-radius: 7px; font-size: 12.5px; color: var(--ink); cursor: pointer; text-align: left; }
.sl-dg-custom:hover { background: var(--hover); }
/* The ribbon's own strip: live thumbnails in a row, the rest behind a button. */
.sl-rs { display: flex; align-items: stretch; gap: 0; height: 60px; border: 1px solid var(--line-soft); border-radius: 8px; background: var(--surface-2); padding: 3px; }
.sl-rs-items { display: flex; gap: 4px; align-items: center; }
.sl-rs-item { width: 88px; height: 52px; padding: 0; border: 1.5px solid transparent; border-radius: 5px; background: transparent; cursor: pointer; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 2px; transition: border-color var(--fast), transform var(--fast); }
.sl-rs-item:hover { border-color: var(--line-strong); }
.sl-rs-item.current { border-color: var(--accent); box-shadow: 0 0 0 2px var(--accent-soft); }
.sl-rs-pic { display: block; width: 80px; line-height: 0; border-radius: 3px; overflow: hidden; box-shadow: 0 0 0 1px rgba(0,0,0,.10); background: #fff; }
.sl-rs-pic svg { display: block; width: 100%; height: auto; }
.sl-accents { display: flex; width: 80px; height: 3px; border-radius: 2px; overflow: hidden; flex: none; }
.sl-accents > span { flex: 1; }
.sl-dg-name .sl-accents { width: 54px; margin-left: auto; height: 6px; border-radius: 3px; }
.sl-dg-label { overflow: hidden; text-overflow: ellipsis; }
.sl-rs-wait { font-size: 11.5px; color: var(--ink-3); padding: 0 10px; white-space: nowrap; }
.sl-rs-more { width: 18px; margin-left: 3px; border: 0; border-left: 1px solid var(--line-soft); background: transparent; color: var(--ink-2); cursor: pointer; display: grid; place-items: center; border-radius: 0 5px 5px 0; }
.sl-rs-more:hover { background: var(--hover); color: var(--accent); }
/* Customise Colours and Customise Fonts. */
.sl-cc { display: grid; grid-template-columns: 1fr 190px; gap: 18px; }
.sl-cc-caption { font-size: 11px; text-transform: uppercase; letter-spacing: .05em; color: var(--ink-3); font-weight: 600; margin-bottom: 6px; }
.sl-cc-slots { display: flex; flex-direction: column; gap: 3px; }
.sl-cc-slot { display: grid; grid-template-columns: 1fr 34px 82px; gap: 8px; align-items: center; font-size: 12.5px; color: var(--ink); }
.sl-cc-label { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sl-cc-pick { width: 34px; height: 24px; padding: 0 2px; border: 1px solid var(--line); border-radius: 5px; background: var(--surface); cursor: pointer; }
.sl-cc-hex { height: 24px; font-size: 12px; font-family: ui-monospace, Consolas, monospace; text-transform: uppercase; }
.sl-cc-side { display: flex; flex-direction: column; gap: 10px; }
.sl-cc-samples { display: grid; gap: 8px; }
.sl-cc-sample { width: 100%; height: auto; border-radius: 6px; box-shadow: 0 0 0 1px var(--line); display: block; }
.sl-cc-name { display: flex; flex-direction: column; gap: 4px; font-size: 12px; color: var(--ink-2); }
.sl-cf { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; }
.sl-cf-fields { display: flex; flex-direction: column; gap: 10px; }
.sl-cf-field { display: flex; flex-direction: column; gap: 4px; font-size: 12px; color: var(--ink-2); }
.sl-cf-card { border: 1px solid var(--line); border-radius: 8px; padding: 14px 16px; background: #fff; color: #1b1b1b; min-height: 112px; }
.sl-cf-heading { font-size: 26px; line-height: 1.2; margin-bottom: 8px; }
.sl-cf-body { font-size: 13px; line-height: 1.45; color: #333; }
`;
