// What the network camera dialogs look like: the one by address and the one that finds cameras.

export const NET_CAMERA_CSS = `
.st-cam-dialog { width: min(520px, 92vw); }
.st-cam-pair { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
.st-cam-known { margin: 0 0 12px; display: flex; flex-direction: column; gap: 6px; }
.st-cam-known h4 { margin: 0; font-size: 12px; font-weight: 600; color: var(--ink-2); }
.st-cam-row { display: flex; align-items: center; gap: 8px; }
.st-cam-pick { flex: 1; min-width: 0; display: flex; flex-direction: column; align-items: flex-start; gap: 1px; padding: 7px 10px; border: 1px solid var(--line); border-radius: 7px; background: var(--surface); color: var(--ink); font: inherit; text-align: start; cursor: pointer; }
.st-cam-pick:hover { border-color: var(--accent); }
.st-cam-pick b { font-size: 13px; max-width: 100%; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.st-cam-pick small { font-size: 11.5px; color: var(--ink-3); max-width: 100%; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; direction: ltr; }
.st-cam-url { direction: ltr; }
.st-cam-note { margin: 10px 0 0; font-size: 12px; line-height: 1.45; color: var(--ink-3); }
.st-cam-err { margin: 10px 0 0; font-size: 12.5px; color: var(--danger, #b3261e); }
`;
