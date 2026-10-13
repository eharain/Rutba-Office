// Text recognition in the windows, shared by PDF (Recognise Text), Pictures
// and Image (Copy Text from Picture) and Worksheets (Data from Picture).
//
// The reading itself is done by the main process's pool of threads
// (window.rutbaOffice.ocr, apps/desktop/main/ocr.js); a window only hands
// over pixels and shows what came back. Here: which languages to read, as
// the person last chose them (or the window's own language and English),
// the picker that chooses them with Urdu marked beta, pixels from a picture
// or a canvas, and the words laid over a picture so they can be selected.

import React, { useCallback, useLayoutEffect, useRef, useState } from 'react';
import { Button, Dialog, Spinner, t, tn, language } from '@rutba/office-ui';
import { languagesFor, CODES, LANGUAGES } from '@rutba/ocr/languages';

const KEY = 'rutba-ocr-languages';

/**
 * Every language there is a model for, its name in the window's language.
 * A language added to @rutba/ocr's table shows its English name until its
 * name is added here (and so to the catalogues).
 */
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
            <span>{names[code] || LANGUAGES.find((l) => l.code === code)?.name}</span>
            {code === 'urd' ? <span className="ocr-beta" title={t('Urdu text recognition is in beta')}>{t('Beta')}</span> : LANGUAGES.find((l) => l.code === code)?.beta ? <span className="ocr-beta">{t('Beta')}</span> : null}
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
.ocr-bar { position: sticky; top: 10px; grid-area: 1 / 1; align-self: start; justify-self: center; z-index: 6; display: flex; align-items: center; gap: 8px; padding: 5px 6px 5px 12px; max-width: calc(100% - 24px); background: var(--surface); color: var(--ink); border: 1px solid var(--line); border-radius: var(--r-3); box-shadow: var(--shadow-2); font-size: 12.5px; line-height: 1.3; cursor: default; }
.ocr-bar .rw-spinner { width: 16px; height: 16px; }
.ocr-bar-text { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; padding-inline-end: 4px; }
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

/**
 * Copy Text from Picture, as Pictures and Image use it: the picture read
 * once (by its `key`), all its text put on the clipboard, and the result
 * kept to lay its lines over the picture. A second press with the same
 * picture copies again without reading again.
 */
export function usePictureText({ shell, toast }) {
  const [state, setState] = useState(null); // { key, result, pixelScale }
  const [busy, setBusy] = useState(false);
  const [asking, setAsking] = useState(false);
  const seq = useRef(0);
  const live = useRef(null);

  const copyAll = useCallback(async (result) => {
    await shell.clipboard.writeText({ text: result.text });
  }, [shell]);

  const read = useCallback(async (key, getPixels) => {
    if (state?.key === key && state.result) {
      await copyAll(state.result);
      toast(t('Text copied'), { tone: 'good' });
      return;
    }
    const n = ++seq.current;
    live.current = `picture-${n}`;
    setBusy(true);
    try {
      const pixels = await getPixels();
      const result = await recognisePixels(shell, pixels, { languages: chosenOcrLanguages(), job: live.current });
      if (n !== seq.current) return;
      if (!result.words.length) {
        setState(null);
        toast(t('No text was found in this picture.'), { ms: 4000 });
        return;
      }
      await copyAll(result);
      setState({ key, result, pixelScale: pixels.scale || 1 });
      toast(tn(result.lines.length, 'Text copied: {count} line. Select words on the picture to copy only those.', 'Text copied: {count} lines. Select words on the picture to copy only those.'), { tone: 'good', ms: 5000 });
    } catch (err) {
      if (n === seq.current && err?.code !== 'OCR_CANCELLED') toast(t('Text could not be recognised: {reason}', { reason: err.message }), { tone: 'bad', ms: 7000 });
    } finally {
      if (n === seq.current) setBusy(false);
    }
  }, [state, shell, toast, copyAll]);

  /** Stop a reading under way, and take the words off the picture. */
  const clear = useCallback(() => {
    seq.current += 1;
    if (live.current) shell.ocr.cancel({ job: live.current }).catch(() => {});
    live.current = null;
    setBusy(false);
    setState(null);
  }, [shell]);

  return {
    state,
    busy,
    read,
    clear,
    copyAll: () => state && copyAll(state.result).then(() => toast(t('Text copied'), { tone: 'good' })),
    asking,
    setAsking,
  };
}

/** The bar over a picture while its words are read, and once they are there. */
export function PictureTextBar({ text }) {
  if (text.busy) {
    return (
      <div className="ocr-bar" data-role="ocr-bar" data-state="reading" onMouseDown={(e) => e.stopPropagation()}>
        <Spinner />
        <span className="ocr-bar-text">{t('Reading the words in the picture…')}</span>
        <Button ghost label={t('Cancel')} onClick={text.clear} />
      </div>
    );
  }
  if (!text.state) return null;
  const lines = text.state.result.lines.length;
  return (
    <div className="ocr-bar" data-role="ocr-bar" data-state="ready" data-lines={lines} onMouseDown={(e) => e.stopPropagation()}>
      <span className="ocr-bar-text">{tn(lines, '{count} line of text. Select words to copy them.', '{count} lines of text. Select words to copy them.')}</span>
      <Button icon="copy" label={t('Copy All')} data-role="ocr-copy-all" onClick={text.copyAll} />
      <Button ghost icon="close" title={t('Hide the text')} onClick={text.clear} data-role="ocr-hide" />
    </div>
  );
}

/** Text Languages: which languages Copy Text reads, remembered for every window. */
export function OcrLanguagesDialog({ onClose }) {
  const [value, setValue] = useState(chosenOcrLanguages);
  return (
    <Dialog
      title={t('Text Languages')}
      onClose={onClose}
      actions={(
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('OK')} data-role="ocr-languages-ok" onClick={() => { rememberOcrLanguages(value); onClose(); }} />
        </>
      )}
    >
      <div className="ocr-form" data-role="ocr-languages">
        <div className="ocr-form-label">{t('Languages in the picture')}</div>
        <OcrLanguagePicker value={value} onChange={setValue} />
        <div className="ocr-foot">{t('The words are read on this computer. Nothing is sent anywhere.')}</div>
      </div>
    </Dialog>
  );
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
