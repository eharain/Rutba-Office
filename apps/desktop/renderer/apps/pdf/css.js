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

/* Comments: the layer over a page, the pane beside the pages, the box a note is typed into. */
.pdf-annots { position: absolute; inset: 0; width: 100%; height: 100%; z-index: 2; pointer-events: none; overflow: visible; }
.pdf-annots.drawing { pointer-events: all; cursor: crosshair; }
.pdf-annots[data-tool='note'] { cursor: copy; }
.pdf-annots[data-tool='text'] { cursor: text; }
.pdf-annot polygon { fill: transparent; stroke: none; pointer-events: all; cursor: pointer; vector-effect: non-scaling-stroke; }
.pdf-annots.drawing .pdf-annot polygon { pointer-events: none; }
.pdf-annot:hover polygon { stroke: var(--accent-line); stroke-width: 1px; }
.pdf-annot.picked polygon { stroke: var(--accent); stroke-width: 1.5px; stroke-dasharray: 4 3; fill: color-mix(in srgb, var(--accent) 7%, transparent); }
.pdf-draft { vector-effect: non-scaling-stroke; }
.pdf-link { fill: transparent; stroke: none; pointer-events: all; cursor: pointer; }
.pdf-link:hover { fill: color-mix(in srgb, var(--accent) 10%, transparent); }
.pdf-annots.drawing .pdf-link { pointer-events: none; }
.pdf-side-tabs { display: flex; gap: 2px; }
.pdf-side-tabs button { height: 24px; padding: 0 8px; border: 0; border-radius: var(--r-1); background: none; font: inherit; font-size: 12px; font-weight: 600; color: var(--ink-3); cursor: pointer; }
.pdf-side-tabs button:hover { background: var(--hover); }
.pdf-side-tabs button.on { color: var(--ink); background: var(--selected); }
.pdf-bookmarks { flex: 1; min-height: 0; overflow: auto; padding: 8px 6px 24px; }
.pdf-bookmarks-name { font-size: 11px; font-weight: 600; color: var(--ink-3); text-transform: none; padding: 8px 6px 4px; }
.pdf-bookmark { display: flex; align-items: center; gap: 5px; min-height: 26px; padding-inline-end: 6px; border-radius: var(--r-1); font-size: 12.5px; color: var(--ink); cursor: pointer; user-select: none; }
.pdf-bookmark:hover { background: var(--hover); }
.pdf-bookmark.current { background: var(--selected); }
.pdf-bookmark.gone { color: var(--ink-3); cursor: default; }
.pdf-bookmark.gone .pdf-bookmark-title { text-decoration: line-through; text-decoration-color: color-mix(in srgb, var(--ink-3) 60%, transparent); }
.pdf-bookmark svg { flex: none; color: var(--ink-3); }
.pdf-bookmark-twist { flex: none; width: 16px; height: 16px; display: grid; place-items: center; border: 0; padding: 0; background: none; color: var(--ink-2); cursor: pointer; border-radius: 3px; }
.pdf-bookmark-twist:hover { background: var(--active); }
.pdf-bookmark-title { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.pdf-bookmarks-empty { padding: 18px 12px; font-size: 12.5px; color: var(--ink-3); text-align: center; }
.pdf-swatches { display: grid; grid-template-columns: repeat(3, 20px); gap: 6px; padding: 6px 4px; align-content: center; }
.pdf-swatch { width: 20px; height: 20px; border-radius: 50%; border: 1px solid rgba(0,0,0,0.18); background: var(--swatch); cursor: pointer; padding: 0; }
.pdf-swatch:hover { transform: scale(1.08); }
.pdf-swatch.on { box-shadow: 0 0 0 2px var(--surface), 0 0 0 4px var(--accent); }
.pdf-comments { flex: none; width: 300px; display: flex; flex-direction: column; min-height: 0; background: var(--surface); border-inline-start: 1px solid var(--line); }
.pdf-comments-head { flex: none; display: flex; align-items: center; gap: 7px; height: 38px; padding: 0 6px 0 12px; border-bottom: 1px solid var(--line-soft); color: var(--ink-2); }
.pdf-comments-head .spacer { flex: 1; }
.pdf-comments-title { font-size: 13px; font-weight: 600; color: var(--ink); }
.pdf-comments-count { font-size: 11.5px; min-width: 20px; padding: 1px 6px; border-radius: 10px; background: var(--surface-2); color: var(--ink-2); text-align: center; font-variant-numeric: tabular-nums; }
.pdf-comments-filter { flex: none; display: flex; gap: 2px; padding: 6px 10px; border-bottom: 1px solid var(--line-soft); }
.pdf-comments-filter button { flex: 1; height: 26px; border: 0; border-radius: var(--r-1); background: transparent; color: var(--ink-2); font: inherit; font-size: 12px; cursor: pointer; display: inline-flex; align-items: center; justify-content: center; gap: 5px; }
.pdf-comments-filter button:hover { background: var(--hover); }
.pdf-comments-filter button.on { background: var(--selected); color: var(--ink); font-weight: 600; }
.pdf-comments-n { font-size: 10.5px; padding: 0 5px; border-radius: 8px; background: color-mix(in srgb, var(--good) 18%, transparent); color: var(--ink); }
.pdf-comments-list { flex: 1; min-height: 0; overflow-y: auto; padding: 8px 10px 24px; display: flex; flex-direction: column; gap: 8px; }
.pdf-comments-empty { padding: 18px 8px; font-size: 12.5px; color: var(--ink-3); text-align: center; line-height: 1.5; }
.pdf-comment { border: 1px solid var(--line); border-radius: var(--r-2); padding: 8px 8px 9px 10px; background: var(--surface); cursor: default; display: flex; flex-direction: column; gap: 5px; }
.pdf-comment:hover { border-color: var(--line-strong); }
.pdf-comment.picked { border-color: var(--accent); box-shadow: 0 0 0 1px var(--accent); }
.pdf-comment.resolved { background: var(--surface-2); }
.pdf-comment.resolved .pdf-comment-text { color: var(--ink-2); }
.pdf-comment-head { display: flex; align-items: center; gap: 6px; min-width: 0; }
.pdf-comment-head .spacer { flex: 1; }
.pdf-comment-head .rw-btn.icon { width: 24px; height: 24px; min-width: 24px; }
.pdf-comment-resolve[aria-pressed='true'] { color: var(--good); }
.pdf-comment-kind { flex: none; width: 22px; height: 22px; border-radius: 6px; display: grid; place-items: center; color: var(--ink); background: color-mix(in srgb, var(--kind) 30%, var(--surface)); box-shadow: inset 0 0 0 1px color-mix(in srgb, var(--kind) 55%, transparent); }
.pdf-comment-who { font-size: 12.5px; font-weight: 600; color: var(--ink); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; min-width: 0; }
.pdf-comment-when { font-size: 11.5px; color: var(--ink-3); white-space: nowrap; }
.pdf-comment-meta { display: flex; align-items: center; gap: 5px; font-size: 11.5px; color: var(--ink-3); padding-inline-start: 28px; flex-wrap: wrap; }
.pdf-comment-done { display: inline-flex; align-items: center; gap: 3px; color: var(--good); font-weight: 600; margin-inline-start: 4px; }
.pdf-comment-text { font-size: 12.5px; line-height: 1.45; color: var(--ink); white-space: pre-wrap; overflow-wrap: anywhere; padding-inline-start: 28px; }
.pdf-comment-add { align-self: flex-start; margin-inline-start: 28px; border: 0; background: none; padding: 0; font: inherit; font-size: 12px; color: var(--accent); cursor: pointer; }
.pdf-comment-edit, .pdf-comment-reply { display: flex; gap: 6px; padding-inline-start: 28px; }
.pdf-comment-edit { flex-direction: column; }
.pdf-comment-edit textarea, .pdf-note-editor textarea { width: 100%; box-sizing: border-box; resize: vertical; font: inherit; font-size: 12.5px; padding: 6px 8px; border: 1px solid var(--line-strong); border-radius: var(--r-1); background: var(--surface); color: var(--ink); outline: none; }
.pdf-comment-edit textarea:focus, .pdf-note-editor textarea:focus { border-color: var(--accent); box-shadow: 0 0 0 2px var(--accent-soft); }
.pdf-comment-actions, .pdf-note-editor-actions { display: flex; gap: 6px; justify-content: flex-end; }
.pdf-comment-reply .rw-input { flex: 1; min-width: 0; height: 28px; }
.pdf-comment-replies { display: flex; flex-direction: column; gap: 6px; margin-inline-start: 28px; padding-inline-start: 9px; border-inline-start: 2px solid var(--line); }
.pdf-reply-head { display: flex; gap: 6px; align-items: baseline; }
.pdf-reply .pdf-comment-text { padding-inline-start: 0; }
.pdf-note-editor { position: fixed; z-index: 60; width: 280px; display: flex; flex-direction: column; gap: 8px; padding: 10px; background: var(--surface); border: 1px solid var(--line); border-radius: var(--r-3); box-shadow: var(--shadow-3); }
.pdf-note-editor-head { display: flex; align-items: center; gap: 6px; font-size: 12.5px; font-weight: 600; color: var(--ink); }

/* A form's fields: see-through over the page's own drawing of them, tinted as Acrobat tints them. */
/* No z-index: a stacking context of its own would blend the tint with nothing instead of the page beneath. */
.pdf-fields { position: absolute; inset: 0; pointer-events: none; }
.pdf-field { position: absolute; box-sizing: border-box; margin: 0; padding: 0 3px; border: 1px solid transparent; border-radius: 2px; background: transparent; pointer-events: all; font-family: Helvetica, Arial, sans-serif; color: transparent; caret-color: #111; outline: none; }
.pdf-fields.tinted .pdf-field { background: rgb(222, 229, 255); mix-blend-mode: multiply; }
.pdf-field:hover { border-color: rgba(43, 95, 217, 0.55); }
.pdf-field:focus, .pdf-field-text.shown { background: #fff !important; mix-blend-mode: normal !important; color: #111; border-color: var(--accent); box-shadow: 0 0 0 2px var(--accent-soft); z-index: 1; }
.pdf-field-text.multi { resize: none; padding-top: 2px; line-height: 1.15; overflow: hidden; }
.pdf-field-checkbox, .pdf-field-radio { cursor: pointer; }
.pdf-field-radio { border-radius: 50%; }
.pdf-field-choice, .pdf-field-list { opacity: 1; color: transparent; cursor: pointer; appearance: none; }
.pdf-field-list { overflow: hidden; scrollbar-width: none; }
.pdf-field-list:focus { overflow-y: auto; }
.pdf-field-choice:focus, .pdf-field-list:focus { color: #111; }
.pdf-field-choice option, .pdf-field-list option { color: #111; background: #fff; }
.pdf-field-list:not(:focus) option { background: transparent; color: transparent; }
.pdf-field:disabled { cursor: not-allowed; background: rgba(128, 128, 128, 0.12) !important; }
.pdf-field-sig { background: rgba(255, 214, 102, 0.28); border: 1px dashed rgba(201, 121, 27, 0.6); }
.pdf-notice.info { background: color-mix(in srgb, var(--info) 10%, var(--surface)); border-bottom-color: color-mix(in srgb, var(--info) 28%, var(--line)); }
.pdf-sign { display: flex; flex-direction: column; gap: 12px; }
.pdf-sign-tabs { display: flex; gap: 2px; border-bottom: 1px solid var(--line); }
.pdf-sign-tabs button { height: 30px; padding: 0 16px; border: 0; border-bottom: 2px solid transparent; margin-bottom: -1px; background: none; font: inherit; font-size: 13px; color: var(--ink-2); cursor: pointer; }
.pdf-sign-tabs button.on { color: var(--ink); font-weight: 600; border-bottom-color: var(--accent); }
.pdf-sign-type { display: flex; flex-direction: column; gap: 10px; }
.pdf-sign-preview { height: 96px; display: flex; align-items: center; justify-content: center; font-size: 44px; color: #10235c; background: #fff; border: 1px solid var(--line); border-radius: var(--r-2); overflow: hidden; white-space: nowrap; }
.pdf-sign-draw { position: relative; display: flex; flex-direction: column; align-items: flex-start; gap: 6px; }
.pdf-sign-draw > .rw-btn { align-self: flex-end; }
.pdf-sign-pad { display: block; background: #fff; border: 1px solid var(--line); border-radius: var(--r-2); cursor: crosshair; touch-action: none; }
.pdf-sign-line { position: absolute; left: 24px; width: 412px; top: 112px; border-top: 1px solid #c9cdd6; pointer-events: none; }
.pdf-sign-image { display: flex; flex-direction: column; gap: 10px; }
.pdf-sign-picture { background-color: #fff; background-image: linear-gradient(45deg, #f1f2f5 25%, transparent 25%, transparent 75%, #f1f2f5 75%), linear-gradient(45deg, #f1f2f5 25%, transparent 25%, transparent 75%, #f1f2f5 75%); background-size: 16px 16px; background-position: 0 0, 8px 8px; height: 130px; }
.pdf-sign-picture img { max-width: 92%; max-height: 112px; object-fit: contain; }
.pdf-sign-hint { font-family: var(--font); font-size: 12.5px; color: var(--ink-3); }
.pdf-sign-image-row { display: flex; align-items: center; justify-content: space-between; gap: 10px; }
.pdf-sign-clear { display: inline-flex; align-items: center; gap: 6px; font-size: 12.5px; color: var(--ink-2); }
.pdf-sign-note { margin: 0; font-size: 12px; line-height: 1.5; color: var(--ink-2); padding: 8px 10px; border-radius: var(--r-2); background: color-mix(in srgb, var(--warn) 9%, var(--surface)); border: 1px solid color-mix(in srgb, var(--warn) 25%, var(--line)); }

/* Export as Pictures and Reduce File Size. */
.pdf-export-form { display: flex; flex-direction: column; gap: 12px; min-width: 400px; font-size: 13px; }
.pdf-export-scope { display: flex; flex-direction: column; gap: 6px; margin: 0; padding: 0; border: 0; }
.pdf-export-scope legend { font-size: 12px; font-weight: 600; color: var(--ink-2); margin-bottom: 4px; }
.pdf-export-scope label { display: flex; align-items: center; gap: 8px; }
.pdf-export-row { display: grid; grid-template-columns: 90px 1fr; align-items: center; gap: 8px; }
.pdf-export-row label { font-size: 12px; font-weight: 600; color: var(--ink-2); }
.pdf-export-note { font-size: 12px; color: var(--ink-3); }
.pdf-export-lead { margin: 0; line-height: 1.5; color: var(--ink); }

/* Redaction: the Apply dialog. */
.pdf-redact-form { display: flex; flex-direction: column; gap: 10px; font-size: 13px; line-height: 1.5; color: var(--ink); max-width: 460px; }
.pdf-redact-form p { margin: 0; }
.pdf-redact-lead { display: flex; align-items: center; gap: 10px; font-weight: 600; }
.pdf-redact-icon { flex: none; width: 36px; height: 36px; border-radius: var(--r-2); display: grid; place-items: center; background: #1b1d22; color: #fff; }
.pdf-redact-warn { padding: 8px 10px; border-radius: var(--r-2); background: color-mix(in srgb, var(--warn) 10%, var(--surface)); border: 1px solid color-mix(in srgb, var(--warn) 28%, var(--line)); }
.pdf-redact-small { font-size: 12px; color: var(--ink-2); }

.pdf-print-root { display: none; }
@media print {
  #root { display: none !important; }
  .pdf-print-root { display: block !important; }
  .pdf-print-root img { display: block; break-after: page; max-width: 100vw; max-height: 100vh; margin: 0 auto; object-fit: contain; }
  .pdf-print-root img:last-child { break-after: auto; }
  @page { margin: 0; }
}
`;
