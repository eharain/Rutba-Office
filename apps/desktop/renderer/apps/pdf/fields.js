// Filling a PDF's form in the window, and Fill & Sign.
//
// Each field's widgets get a control laid over the page where the file puts
// them: a text box, a check box, a radio button, a drop-down or a list. The
// page itself, drawn by pdf.js, shows the appearance the file holds; a
// control is see-through over it, tinted as Acrobat tints fields, and shows
// its own words only while they are being typed or are on their way into
// the file. A value is written when it is committed (a text box left or
// Enter pressed, a box ticked, an option chosen): the main process fills the
// field as an incremental update with appearance streams any reader shows
// (@rutba/pdf/forms), and the file with it becomes the pages' new source, so
// Undo takes the value away again.
//
// A signature or initials, typed, drawn on the pad or taken from a picture
// of a signature on paper, is placed as a stamp: a picture
// of a signature, drawn into the file as an annotation. It is not a digital
// signature and the window says so wherever one is made.

import React, { useEffect, useRef, useState } from 'react';
import { Button, Dialog, Input, t } from '@rutba/office-ui';
import { getPage } from '../../pdfjs.js';

/** A page's pdf.js viewport at a scale and turn, once the page is read. */
export function usePageViewport(source, entry, scale) {
  const [vp, setVp] = useState(null);
  useEffect(() => {
    let live = true;
    if (!source?.doc) return undefined;
    getPage(source.doc, entry.page).then((page) => { if (live) setVp(page.getViewport({ scale, rotation: entry.rotate })); }).catch(() => {});
    return () => { live = false; };
  }, [source, entry.page, entry.rotate, scale]);
  return vp;
}

/** A rectangle in the page's coordinates as a box on the screen page. */
export function boxOf(vp, r) {
  const a = vp.convertToViewportPoint(r[0], r[1]);
  const b = vp.convertToViewportPoint(r[2], r[3]);
  return { left: Math.min(a[0], b[0]), top: Math.min(a[1], b[1]), width: Math.abs(b[0] - a[0]), height: Math.abs(b[1] - a[1]) };
}

/** The controls for the fields on one page. */
export function FieldLayer({ entry, source, scale, fields, highlight, onFill }) {
  const vp = usePageViewport(source, entry, scale);
  const mine = (fields || []).flatMap((f) => f.widgets.filter((w) => w.page === entry.page && w.rect).map((w) => ({ field: f, widget: w })));
  if (!vp || !mine.length) return null;
  return (
    <div className={`pdf-fields${highlight ? ' tinted' : ''}`} data-ready="yes">
      {mine.map(({ field, widget }) => (
        <FieldControl key={`${field.id}:${widget.id}`} field={field} widget={widget} box={boxOf(vp, widget.rect)} scale={scale} onFill={(value) => onFill(entry.source, field, value)} />
      ))}
    </div>
  );
}

function FieldControl({ field, widget, box, scale, onFill }) {
  const [draft, setDraft] = useState(field.value);
  const [focused, setFocused] = useState(false);
  // A value the file comes back with (Undo, or the write finishing) replaces the draft.
  useEffect(() => { setDraft(field.value); }, [field.value]);
  const style = { left: box.left, top: box.top, width: box.width, height: box.height };
  const common = {
    className: 'pdf-field',
    'data-field': field.name,
    'data-type': field.type,
    'aria-label': field.name,
    disabled: field.readOnly,
  };
  const pending = focused || (typeof draft === 'string' && draft !== field.value);
  if (field.type === 'text' || field.type === 'signature') {
    if (field.type === 'signature') return <div {...common} className="pdf-field pdf-field-sig" style={style} title={t('A signature field: Fill & Sign places a picture of a signature anywhere on the page')} />;
    const size = field.fontSize ? field.fontSize * scale : Math.max(8, Math.min(16 * scale, box.height * 0.62));
    const commit = () => { if (draft !== field.value) onFill(draft); };
    const props = {
      ...common,
      className: `pdf-field pdf-field-text${pending ? ' shown' : ''}${field.multiline ? ' multi' : ''}`,
      style: { ...style, fontSize: size, textAlign: field.align === 1 ? 'center' : field.align === 2 ? 'right' : 'left' },
      value: draft || '',
      maxLength: field.maxLength || undefined,
      onFocus: () => setFocused(true),
      onBlur: () => { setFocused(false); commit(); },
      onChange: (e) => setDraft(e.target.value),
      onKeyDown: (e) => {
        if (e.key === 'Enter' && !field.multiline) { e.preventDefault(); e.currentTarget.blur(); }
        if (e.key === 'Escape') { setDraft(field.value); e.currentTarget.blur(); }
      },
      spellCheck: false,
    };
    return field.multiline ? <textarea {...props} /> : <input type={field.password ? 'password' : 'text'} {...props} />;
  }
  if (field.type === 'checkbox' || field.type === 'radio') {
    const on = field.type === 'checkbox' ? field.value !== 'Off' && field.value !== '' : field.value === widget.onState;
    return (
      <button
        type="button"
        {...common}
        role={field.type}
        aria-checked={on}
        className={`pdf-field pdf-field-${field.type}`}
        style={style}
        data-state={widget.onState || ''}
        onClick={() => onFill(field.type === 'checkbox' ? !on : (on && field.widgets.length > 1 ? widget.onState : widget.onState))}
      />
    );
  }
  if (field.type === 'combo') {
    return (
      <select {...common} className="pdf-field pdf-field-choice" style={style} value={field.value || ''} onChange={(e) => onFill(e.target.value)}>
        {!field.options.some((o) => o.value === field.value) ? <option value={field.value || ''}>{field.value || ''}</option> : null}
        {field.options.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
      </select>
    );
  }
  if (field.type === 'list') {
    const chosen = Array.isArray(field.value) ? field.value : field.value ? [field.value] : [];
    return (
      <select
        {...common}
        multiple
        className="pdf-field pdf-field-list"
        style={style}
        value={chosen}
        onChange={(e) => {
          const picked = [...e.target.selectedOptions].map((o) => o.value);
          onFill(field.multiSelect ? picked : picked.slice(-1));
        }}
      >
        {field.options.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
      </select>
    );
  }
  return null;
}

// ── signatures ────────────────────────────────────────────────────────────

const STORE_KEY = (kind) => `rutba-pdf-${kind}`;
export function savedSignature(kind) {
  try { return JSON.parse(localStorage.getItem(STORE_KEY(kind)) || 'null'); } catch { return null; }
}
export function forgetSignature(kind) {
  try { localStorage.removeItem(STORE_KEY(kind)); } catch { /* nothing kept */ }
}
const SCRIPT_FONT = '"Segoe Script", "Lucida Handwriting", "Brush Script MT", "Ink Free", cursive';

/** A typed name drawn in a handwriting face, cropped to its ink: { png: Uint8Array, width, height } in pixels. */
export async function typedPicture(text, colour = '#10235c') {
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d');
  const px = 96;
  ctx.font = `${px}px ${SCRIPT_FONT}`;
  const w = Math.ceil(ctx.measureText(text).width) + px;
  canvas.width = w;
  canvas.height = Math.ceil(px * 1.8);
  ctx.font = `${px}px ${SCRIPT_FONT}`;
  ctx.fillStyle = colour;
  ctx.textBaseline = 'middle';
  ctx.fillText(text, px / 2, canvas.height / 2);
  // Cropped to the ink, with a little room.
  const { data } = ctx.getImageData(0, 0, canvas.width, canvas.height);
  let x0 = canvas.width; let y0 = canvas.height; let x1 = 0; let y1 = 0;
  for (let y = 0; y < canvas.height; y++) {
    for (let x = 0; x < canvas.width; x++) {
      if (data[(y * canvas.width + x) * 4 + 3] > 8) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
    }
  }
  if (x1 <= x0 || y1 <= y0) return null;
  const pad = 6;
  const out = document.createElement('canvas');
  out.width = x1 - x0 + pad * 2;
  out.height = y1 - y0 + pad * 2;
  out.getContext('2d').drawImage(canvas, x0 - pad, y0 - pad, out.width, out.height, 0, 0, out.width, out.height);
  const blob = await new Promise((resolve) => out.toBlob(resolve, 'image/png'));
  return { png: new Uint8Array(await blob.arrayBuffer()), width: out.width, height: out.height };
}

/**
 * A picture of a signature (a photo or scan of one signed on paper) made
 * ready to place: no larger than it needs to be, its paper made see-through
 * when `clear` (every pixel near white, faded at the edge of the ink), and
 * cropped to the ink. { url (a PNG data URL), width, height } in pixels.
 */
export async function imagePicture(bytes, { clear = true } = {}) {
  const bitmap = await createImageBitmap(new Blob([bytes]));
  const scale = Math.min(1, 900 / bitmap.width, 400 / bitmap.height);
  const w = Math.max(1, Math.round(bitmap.width * scale));
  const h = Math.max(1, Math.round(bitmap.height * scale));
  const canvas = document.createElement('canvas');
  canvas.width = w;
  canvas.height = h;
  const ctx = canvas.getContext('2d');
  ctx.drawImage(bitmap, 0, 0, w, h);
  bitmap.close?.();
  const img = ctx.getImageData(0, 0, w, h);
  const d = img.data;
  let x0 = w; let y0 = h; let x1 = 0; let y1 = 0;
  for (let i = 0; i < d.length; i += 4) {
    if (clear) {
      const luma = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
      // Paper goes; ink stays; between the two, the ink's edge fades.
      const alpha = luma >= 225 ? 0 : luma <= 160 ? 255 : Math.round(((225 - luma) / 65) * 255);
      d[i + 3] = Math.min(d[i + 3], alpha);
    }
    if (d[i + 3] > 16) {
      const p = i / 4;
      const x = p % w;
      const y = (p - x) / w;
      if (x < x0) x0 = x;
      if (x > x1) x1 = x;
      if (y < y0) y0 = y;
      if (y > y1) y1 = y;
    }
  }
  if (x1 <= x0 || y1 <= y0) return null;
  ctx.putImageData(img, 0, 0);
  const pad = 4;
  const out = document.createElement('canvas');
  out.width = Math.min(w, x1 - x0 + pad * 2);
  out.height = Math.min(h, y1 - y0 + pad * 2);
  out.getContext('2d').drawImage(canvas, Math.max(0, x0 - pad), Math.max(0, y0 - pad), out.width, out.height, 0, 0, out.width, out.height);
  return { url: out.toDataURL('image/png'), width: out.width, height: out.height };
}

const IMAGE_TYPES = ['png', 'jpg', 'jpeg', 'bmp', 'webp', 'gif'];

/**
 * A signature or initials, typed, drawn on the pad or taken from a picture,
 * made for placing on the page. `kind`: 'signature' | 'initials'.
 * `onDone({ kind, mode, text?, strokes?, url?, width, height })`.
 */
export function SignatureDialog({ kind, shell, onDone, onClose }) {
  const kept = savedSignature(kind);
  const [mode, setMode] = useState(kept?.mode || 'type');
  const [text, setText] = useState(kept?.mode === 'type' ? kept.text : '');
  const [strokes, setStrokes] = useState(kept?.mode === 'draw' ? kept.strokes : []);
  const [picture, setPicture] = useState(kept?.mode === 'image' ? kept : null);
  const [source, setSource] = useState(null);
  const [clear, setClear] = useState(true);
  const [failed, setFailed] = useState(null);
  // The picture made again whenever its file or the choice to clear the paper changes.
  useEffect(() => {
    if (!source) return undefined;
    let live = true;
    imagePicture(source, { clear }).then((p) => {
      if (!live) return;
      if (p) { setPicture({ mode: 'image', ...p }); setFailed(null); } else setFailed(t('No ink was found in that picture.'));
    }).catch(() => { if (live) setFailed(t('That file could not be read as a picture.')); });
    return () => { live = false; };
  }, [source, clear]);
  const choosePicture = async () => {
    const [file] = await shell.dialog.open({ title: t('Choose a picture of your signature'), filters: [{ name: t('Pictures'), extensions: IMAGE_TYPES }] }) || [];
    if (!file) return;
    try {
      const { bytes } = await shell.fs.read({ path: file });
      setSource(bytes);
    } catch (err) {
      setFailed(err.message);
    }
  };
  const pad = useRef(null);
  const PAD_W = 460;
  const PAD_H = 150;

  useEffect(() => {
    const c = pad.current;
    if (!c || mode !== 'draw') return;
    const ctx = c.getContext('2d');
    const ratio = window.devicePixelRatio || 1;
    c.width = PAD_W * ratio;
    c.height = PAD_H * ratio;
    ctx.setTransform(ratio, 0, 0, ratio, 0, 0);
    ctx.clearRect(0, 0, PAD_W, PAD_H);
    ctx.strokeStyle = '#10235c';
    ctx.lineWidth = 2.4;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    for (const s of strokes) {
      ctx.beginPath();
      s.forEach(([x, y], i) => (i ? ctx.lineTo(x, y) : ctx.moveTo(x, y)));
      if (s.length === 1) ctx.lineTo(s[0][0] + 0.1, s[0][1]);
      ctx.stroke();
    }
  }, [strokes, mode]);

  const down = (e) => {
    const box = pad.current.getBoundingClientRect();
    const at = (ev) => [Math.round((ev.clientX - box.left) * 10) / 10, Math.round((ev.clientY - box.top) * 10) / 10];
    let stroke = [at(e)];
    setStrokes((s) => [...s, stroke]);
    const move = (ev) => {
      stroke = [...stroke, at(ev)];
      setStrokes((s) => [...s.slice(0, -1), stroke]);
    };
    const up = () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  const ready = mode === 'type' ? Boolean(text.trim()) : mode === 'image' ? Boolean(picture) : strokes.some((s) => s.length > 1);
  const title = kind === 'initials' ? t('Add Initials') : t('Add Signature');
  const place = () => {
    const made = mode === 'type' ? { kind, mode, text: text.trim() }
      : mode === 'image' ? { kind, mode, url: picture.url, width: picture.width, height: picture.height }
        : { kind, mode, strokes, width: PAD_W, height: PAD_H };
    try { localStorage.setItem(STORE_KEY(kind), JSON.stringify(made)); } catch { /* kept for this window only */ }
    onDone(made);
  };
  return (
    <Dialog
      title={title}
      onClose={onClose}
      width={520}
      actions={(
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('Place on the Page')} data-role="sign-place" disabled={!ready} onClick={place} />
        </>
      )}
    >
      <div className="pdf-sign" data-role="sign-dialog" data-kind={kind}>
        <div className="pdf-sign-tabs" role="tablist">
          <button type="button" role="tab" aria-selected={mode === 'type'} className={mode === 'type' ? 'on' : ''} data-role="sign-type" onClick={() => setMode('type')}>{t('Typed')}</button>
          <button type="button" role="tab" aria-selected={mode === 'draw'} className={mode === 'draw' ? 'on' : ''} data-role="sign-draw" onClick={() => setMode('draw')}>{t('Drawn')}</button>
          <button type="button" role="tab" aria-selected={mode === 'image'} className={mode === 'image' ? 'on' : ''} data-role="sign-image" onClick={() => setMode('image')}>{t('Picture')}</button>
        </div>
        {mode === 'type' ? (
          <div className="pdf-sign-type">
            <Input value={text} autoFocus placeholder={kind === 'initials' ? t('Your initials') : t('Your name')} aria-label={kind === 'initials' ? t('Your initials') : t('Your name')} data-role="sign-text" onChange={(e) => setText(e.target.value)} />
            <div className="pdf-sign-preview" style={{ fontFamily: SCRIPT_FONT }}>{text || '\u00a0'}</div>
          </div>
        ) : mode === 'image' ? (
          <div className="pdf-sign-image">
            <div className="pdf-sign-preview pdf-sign-picture" data-role="sign-picture">
              {picture ? <img src={picture.url} alt={t('Your signature')} /> : <span className="pdf-sign-hint">{failed || t('A photo or scan of your signature on white paper')}</span>}
            </div>
            <div className="pdf-sign-image-row">
              <Button icon="picture" label={picture ? t('Choose Another…') : t('Choose a Picture…')} onClick={choosePicture} data-role="sign-choose" />
              <label className="pdf-sign-clear">
                <input type="checkbox" checked={clear} disabled={!source} onChange={(e) => setClear(e.target.checked)} />
                <span>{t('Make the paper see-through')}</span>
              </label>
            </div>
          </div>
        ) : (
          <div className="pdf-sign-draw">
            <canvas ref={pad} className="pdf-sign-pad" style={{ width: PAD_W, height: PAD_H }} data-role="sign-pad" onMouseDown={down} />
            <div className="pdf-sign-line" />
            <Button ghost label={t('Clear')} disabled={!strokes.length} onClick={() => setStrokes([])} />
          </div>
        )}
        <p className="pdf-sign-note" data-role="sign-note">
          {t('This places a picture of your signature on the page. It is not a digital signature: it does not prove who signed the PDF, or that the PDF has not been changed since.')}
        </p>
      </div>
    </Dialog>
  );
}

/**
 * The stamp a made signature becomes at a place on the page: its rect in
 * the page's coordinates, and its drawing (strokes as paths, or a typed
 * name as a picture).
 */
export async function signatureStamp(made, { point, rect }) {
  const wide = made.kind === 'initials' ? 64 : 150;
  let aspect;
  let picture = null;
  if (made.mode === 'type') {
    picture = await typedPicture(made.text);
    if (!picture) return null;
    aspect = picture.height / picture.width;
  } else if (made.mode === 'image') {
    const res = await fetch(made.url);
    picture = { png: new Uint8Array(await res.arrayBuffer()), width: made.width, height: made.height };
    aspect = made.height / made.width;
  } else {
    const xs = made.strokes.flat().map((p) => p[0]);
    const ys = made.strokes.flat().map((p) => p[1]);
    const bx = [Math.min(...xs), Math.max(...xs)];
    const by = [Math.min(...ys), Math.max(...ys)];
    made = { ...made, bounds: { x: bx[0], y: by[0], w: Math.max(1, bx[1] - bx[0]), h: Math.max(1, by[1] - by[0]) } };
    aspect = made.bounds.h / made.bounds.w;
  }
  let box;
  if (rect && rect[2] - rect[0] > 8) {
    const w = rect[2] - rect[0];
    const h = Math.min(rect[3] - rect[1], w * aspect) || w * aspect;
    box = [rect[0], rect[3] - h, rect[0] + h / aspect, rect[3]];
  } else {
    const h = wide * aspect;
    box = [point[0], point[1] - h, point[0] + wide, point[1]];
  }
  const spec = { type: 'Stamp', rect: box.map((v) => Math.round(v * 100) / 100), icon: made.kind === 'initials' ? 'Initials' : 'Signature', subject: made.kind === 'initials' ? 'Initials' : 'Signature', color: [0.06, 0.14, 0.36] };
  if (picture) spec.image = picture.png;
  else {
    const { x, y, w, h } = made.bounds;
    const sx = (box[2] - box[0]) / w;
    const sy = (box[3] - box[1]) / h;
    spec.paths = made.strokes.filter((s) => s.length).map((s) => s.flatMap(([px, py]) => [Math.round((box[0] + (px - x) * sx) * 100) / 100, Math.round((box[3] - (py - y) * sy) * 100) / 100]));
    spec.width = Math.max(0.8, 2.4 * Math.min(sx, sy));
  }
  return spec;
}
