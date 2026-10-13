// Text recognition in the windows, shared by PDF (Recognise Text), Pictures
// and Image (Copy Text from Picture) and Worksheets (Data from Picture).
//
// The reading itself is done by the main process's pool of threads
// (window.rutbaOffice.ocr, apps/desktop/main/ocr.js); a window only hands
// over pixels and shows what came back. Here: which languages to read, as
// the person last chose them (or the window's own language and English),
// the picker that chooses them with Urdu marked beta, pixels from a picture
// or a canvas, and the words laid over a picture so they can be selected.

import React, { useLayoutEffect, useRef, useState } from 'react';
import { t, language } from '@rutba/office-ui';
import { languagesFor, CODES } from '@rutba/ocr/languages';

const KEY = 'rutba-ocr-languages';

/** Every language there is a model for, its name in the window's language. */
export const ocrLanguageNames = () => ({
  eng: t('English'),
  urd: t('Urdu'),
  ara: t('Arabic'),
  hin: t('Hindi'),
  chi_sim: t('Chinese (Simplified)'),
  jpn: t('Japanese'),
});

/** The languages to read: the last chosen, or the window's own and English. */
export function chosenOcrLanguages() {
  try {
    const saved = JSON.parse(localStorage.getItem(KEY) || 'null');
    if (Array.isArray(saved) && saved.length && saved.every((c) => CODES.includes(c))) return saved;
  } catch { /* nothing kept, or nothing readable */ }
  return languagesFor(language());
}

export function rememberOcrLanguages(list) {
  try { localStorage.setItem(KEY, JSON.stringify(list)); } catch { /* the next window asks again */ }
}

/** Checkboxes for the six languages; Urdu is marked beta, with a word on why when it is ticked. */
export function OcrLanguagePicker({ value, onChange }) {
  const names = ocrLanguageNames();
  const toggle = (code) => {
    const on = value.includes(code);
    const next = on ? value.filter((c) => c !== code) : [...value, code];
    if (next.length) onChange(CODES.filter((c) => next.includes(c)));
  };
  return (
    <div className="ocr-langs" role="group" aria-label={t('Languages in the picture')}>
      <div className="ocr-langs-grid">
        {CODES.map((code) => (
          <label key={code} className="ocr-lang" data-code={code}>
            <input type="checkbox" checked={value.includes(code)} onChange={() => toggle(code)} />
            <span>{names[code]}</span>
            {code === 'urd' ? <span className="ocr-beta" title={t('Urdu text recognition is in beta')}>{t('Beta')}</span> : null}
          </label>
        ))}
      </div>
      {value.includes('urd') ? (
        <div className="ocr-note" data-role="ocr-urdu-note">{t('Urdu text recognition is in beta: it reads Naskh type well, but not yet Nastaliq print reliably.')}</div>
      ) : null}
    </div>
  );
}

export const OCR_CSS = `
.ocr-langs { display: flex; flex-direction: column; gap: 8px; }
.ocr-langs-grid { display: grid; grid-template-columns: repeat(2, minmax(150px, 1fr)); gap: 6px 18px; }
.ocr-lang { display: flex; align-items: center; gap: 8px; font-size: 13px; min-height: 26px; cursor: default; }
.ocr-lang input { margin: 0; }
.ocr-beta { font-size: 10.5px; font-weight: 600; letter-spacing: 0.02em; text-transform: uppercase; padding: 1px 6px; border-radius: 999px; color: var(--accent); background: color-mix(in srgb, var(--accent) 12%, transparent); border: 1px solid color-mix(in srgb, var(--accent) 30%, transparent); }
.ocr-note { font-size: 12px; line-height: 1.45; color: var(--ink-2); padding: 7px 10px; border-radius: var(--r-2); background: color-mix(in srgb, var(--warn) 10%, var(--surface)); border: 1px solid color-mix(in srgb, var(--warn) 28%, var(--line)); max-width: 420px; }
.ocr-form { display: flex; flex-direction: column; gap: 14px; min-width: 400px; }
.ocr-form-label { font-size: 12px; font-weight: 600; color: var(--ink-2); margin-bottom: 2px; }
.ocr-choice { display: flex; align-items: center; gap: 8px; font-size: 13px; min-height: 24px; }
.ocr-foot { font-size: 12px; color: var(--ink-3); line-height: 1.45; max-width: 420px; }
.ocr-progress { display: flex; flex-direction: column; gap: 10px; min-width: 400px; }
.ocr-progress-line { font-size: 13px; color: var(--ink); font-variant-numeric: tabular-nums; }

/* Recognised words over a picture: invisible, but there to select and copy. */
.ocr-layer { position: absolute; inset: 0; z-index: 3; user-select: text; -webkit-user-select: text; cursor: text; overflow: hidden; }
.ocr-layer br { position: absolute; }
.ocr-layer .ocr-line { position: absolute; white-space: pre; color: transparent; line-height: 1; transform-origin: 0 0; font-family: Arial, 'Segoe UI', 'Noto Naskh Arabic', 'Nirmala UI', 'Microsoft YaHei', 'Yu Gothic', sans-serif; }
.ocr-layer .ocr-line::selection { background: rgba(0, 92, 230, 0.32); color: transparent; }
.ocr-layer.show-lines .ocr-line { outline: 1px solid color-mix(in srgb, var(--accent) 55%, transparent); background: color-mix(in srgb, var(--accent) 7%, transparent); border-radius: 2px; }
`;

/**
 * A picture's pixels for recognition, from an <img>, a canvas or an
 * ImageBitmap; at most `max` pixels on its long side (the engine needs no
 * more, and the copy to the main process stays small).
 */
export function pixelsOf(source, { max = 5000 } = {}) {
  const w = source.naturalWidth || source.width;
  const h = source.naturalHeight || source.height;
  const k = Math.min(1, max / Math.max(w, h));
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(w * k));
  canvas.height = Math.max(1, Math.round(h * k));
  const ctx = canvas.getContext('2d', { willReadFrequently: true });
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(source, 0, 0, canvas.width, canvas.height);
  const data = ctx.getImageData(0, 0, canvas.width, canvas.height);
  return { width: canvas.width, height: canvas.height, data: data.data, scale: k };
}

/** A picture by its URL (rutba://file reads with CORS, so its pixels can be read back). */
export function loadPicture(url) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error(t('The picture could not be read.')));
    img.src = url;
  });
}

/** Read the text in pixels on the main process's pool; `job` names it so it can be cancelled. */
export async function recognisePixels(shell, pixels, { languages, job, dpi } = {}) {
  const { width, height, data } = pixels;
  return shell.ocr.recognize({ image: { width, height, data }, languages, job, dpi });
}

/**
 * The lines recognised in a picture, laid over it where they are, as
 * invisible text that can be selected and copied. `result` is what
 * recognisePixels answered for pixels at `pixelScale` of the picture's
 * natural size; `scale` is how many screen pixels a natural pixel is now.
 */
export function TextOverLayer({ result, pixelScale = 1, scale = 1, showLines = false, className = '' }) {
  const ref = useRef(null);
  const [, setTick] = useState(0);
  const k = scale / (pixelScale || 1);
  // Each line is stretched to its box's width once drawn, as a reader's text layer does.
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    for (const span of el.querySelectorAll('.ocr-line')) {
      span.style.transform = 'none';
      const natural = span.getBoundingClientRect().width;
      const want = Number(span.dataset.width);
      if (natural > 0 && want > 0) span.style.transform = `scaleX(${want / natural})`;
    }
  }, [result, k]);
  useLayoutEffect(() => { setTick((n) => n + 1); }, [scale]);
  if (!result) return null;
  return (
    <div ref={ref} className={`ocr-layer${showLines ? ' show-lines' : ''} ${className}`.trim()} data-lines={result.lines.length} data-role="ocr-layer">
      {result.lines.map((line, i) => (
        <React.Fragment key={i}>
          <span
            className="ocr-line"
            dir={line.direction}
            data-width={line.box.width * k}
            style={{ left: line.box.x * k, top: line.box.y * k, fontSize: Math.max(4, line.box.height * k * 0.86), height: line.box.height * k }}
          >
            {line.text}
          </span>
          <br />
        </React.Fragment>
      ))}
    </div>
  );
}
