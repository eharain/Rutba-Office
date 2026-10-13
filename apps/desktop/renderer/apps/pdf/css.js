// The PDF window's look: the pages white on a grey desk, the strip of
// thumbnails beside them, as a PDF reader's pages are shown.

export const CSS = `
.pdf-app { flex: 1; min-width: 0; min-height: 0; display: flex; overflow: hidden; background: var(--sunken); }
.pdf-side { flex: none; display: flex; flex-direction: column; background: var(--surface); border-inline-end: 1px solid var(--line); min-height: 0; }
.pdf-side-head { flex: none; display: flex; align-items: center; gap: 6px; height: 34px; padding: 0 10px; font-size: 12px; font-weight: 600; color: var(--ink-2); border-bottom: 1px solid var(--line-soft); }
.pdf-side-head .spacer { flex: 1; }
.pdf-side-count { font-weight: 500; color: var(--ink-3); font-variant-numeric: tabular-nums; }
.pdf-strip { flex: 1; min-height: 0; overflow-y: auto; overflow-x: hidden; padding: 10px 0 24px; outline: none; }
.pdf-thumb { display: flex; flex-direction: column; align-items: center; gap: 5px; margin: 0 8px 6px; padding: 7px 6px 5px; border-radius: var(--r-2); cursor: default; user-select: none; position: relative; }
.pdf-thumb:hover { background: var(--hover); }
.pdf-thumb.selected { background: var(--selected); }
.pdf-thumb.selected .pdf-thumb-paper { outline: 2.5px solid var(--accent); outline-offset: 2px; }
.pdf-thumb.current:not(.selected) .pdf-thumb-paper { outline: 1.5px solid var(--accent-line); outline-offset: 2px; }
.pdf-thumb-paper { background: #fff; box-shadow: 0 0 0 1px rgba(0,0,0,0.12), 0 1px 3px rgba(0,0,0,0.18); position: relative; overflow: hidden; }
.pdf-thumb-canvas { width: 100%; height: 100%; display: block; }
.pdf-thumb-label { font-size: 11.5px; color: var(--ink-2); font-variant-numeric: tabular-nums; line-height: 1; }
.pdf-thumb.selected .pdf-thumb-label { color: var(--ink); font-weight: 600; }
.pdf-drop-line { height: 3px; margin: -2px 14px 1px; border-radius: 2px; background: var(--accent); box-shadow: 0 0 0 2px color-mix(in srgb, var(--accent) 25%, transparent); }
body.pdf-dragging, body.pdf-dragging * { cursor: grabbing !important; }
.pdf-split { flex: none; width: 5px; cursor: col-resize; margin-inline-start: -3px; z-index: 2; }
.pdf-split:hover { background: var(--accent-line); }

.pdf-main { flex: 1; min-width: 0; min-height: 0; display: flex; flex-direction: column; position: relative; overflow: hidden; }
.pdf-view { flex: 1; min-height: 0; overflow: auto; background: color-mix(in srgb, var(--sunken) 82%, #8a8f99 18%); }
:root[data-theme='dark'] .pdf-view { background: #2a2c31; }
.pdf-view-inner { display: flex; flex-direction: column; align-items: center; gap: 14px; padding: 18px 24px 40px; min-width: min-content; }
.pdf-row { display: flex; gap: 14px; align-items: flex-start; }
.pdf-page { position: relative; background: #fff; box-shadow: 0 0 0 1px rgba(0,0,0,0.10), 0 2px 8px rgba(0,0,0,0.22); flex: none; }
.pdf-page-canvas { position: absolute; inset: 0; width: 100%; height: 100%; display: block; }
.pdf-view.layout-single .pdf-view-inner { min-height: 100%; justify-content: center; }

/* pdf.js's text layer: the words, invisible, over the drawing, to select and find. */
.textLayer { user-select: text; -webkit-user-select: text; position: absolute; text-align: initial; inset: 0; overflow: clip; opacity: 1; line-height: 1; text-size-adjust: none; forced-color-adjust: none; transform-origin: 0 0; caret-color: CanvasText; z-index: 0;
  --scale-round-x: 1px; --scale-round-y: 1px;
  --min-font-size: 1; --text-scale-factor: calc(var(--total-scale-factor) * var(--min-font-size)); --min-font-size-inv: calc(1 / var(--min-font-size)); }
.textLayer :is(span, br) { color: transparent; position: absolute; white-space: pre; cursor: text; transform-origin: 0% 0%; }
.textLayer > :not(.markedContent), .textLayer .markedContent span:not(.markedContent) { z-index: 1; --font-height: 0; font-size: calc(var(--text-scale-factor) * var(--font-height)); --scale-x: 1; --rotate: 0deg; transform: rotate(var(--rotate)) scaleX(var(--scale-x)) scale(var(--min-font-size-inv)); }
.textLayer .markedContent { display: contents; }
.textLayer span[role='img'] { user-select: none; cursor: default; }
.textLayer ::selection { background: rgba(0, 92, 230, 0.28); }
.textLayer br::selection { background: transparent; }
.textLayer .endOfContent { display: block; position: absolute; inset: 100% 0 0; z-index: 0; cursor: default; user-select: none; }
.textLayer.selecting .endOfContent { top: 0; }
.textLayer[data-main-rotation='90'] { transform: rotate(90deg) translateY(-100%); }
.textLayer[data-main-rotation='180'] { transform: rotate(180deg) translate(-100%, -100%); }
.textLayer[data-main-rotation='270'] { transform: rotate(270deg) translateX(-100%); }
::highlight(pdf-find) { background-color: rgba(255, 196, 0, 0.42); }
::highlight(pdf-find-now) { background-color: rgba(255, 120, 0, 0.62); }

.pdf-find { position: absolute; top: 10px; inset-inline-end: 22px; z-index: 5; display: flex; align-items: center; gap: 4px; padding: 5px 6px 5px 10px; background: var(--surface); border: 1px solid var(--line); border-radius: var(--r-3); box-shadow: var(--shadow-2); }
.pdf-find input { width: 220px; height: 28px; }
.pdf-find-count { font-size: 12px; color: var(--ink-2); min-width: 72px; text-align: center; font-variant-numeric: tabular-nums; white-space: nowrap; }
.pdf-find-count[data-count='0'][data-query='yes'] { color: var(--bad); }

.pdf-notice { flex: none; display: flex; align-items: center; gap: 8px; padding: 7px 14px; font-size: 12.5px; background: color-mix(in srgb, var(--warn) 14%, var(--surface)); border-bottom: 1px solid color-mix(in srgb, var(--warn) 35%, var(--line)); color: var(--ink); }
.pdf-empty-actions { display: flex; gap: 8px; justify-content: center; flex-wrap: wrap; }
.pdf-busy { position: absolute; inset: 0; display: grid; place-items: center; background: color-mix(in srgb, var(--sunken) 60%, transparent); z-index: 6; }

.pdf-split-form { display: flex; flex-direction: column; gap: 12px; min-width: 380px; }
.pdf-split-choice { display: flex; align-items: center; gap: 8px; font-size: 13px; }
.pdf-split-choice input[type='number'] { width: 70px; }
.pdf-split-note { font-size: 12px; color: var(--ink-3); }

.pdf-print-root { display: none; }
@media print {
  #root { display: none !important; }
  .pdf-print-root { display: block !important; }
  .pdf-print-root img { display: block; break-after: page; max-width: 100vw; max-height: 100vh; margin: 0 auto; object-fit: contain; }
  .pdf-print-root img:last-child { break-after: auto; }
  @page { margin: 0; }
}
`;
