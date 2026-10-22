// Databases' look: a list of tables, a sheet of records, nothing more than is needed.

export const BASE_CSS = `
/* Scrolls itself: the window's body does not, and twenty-five samples are taller than most screens. Centred by
   auto margins rather than justify-content, which pushes a list taller than the window out of reach at both ends. */
.bs-start { height: 100%; overflow-y: auto; display: flex; flex-direction: column; align-items: center; gap: 8px; padding: 24px 16px; box-sizing: border-box; }
.bs-start > :first-child { margin-top: auto; }
.bs-start > :last-child { margin-bottom: auto; }
.bs-start > .rw-empty { flex: none; padding-bottom: 20px; }
.bs-templates { width: min(760px, 100%); }
.bs-templates h2 { margin: 0 0 10px; font-size: 12px; font-weight: 600; letter-spacing: .04em; text-transform: uppercase; color: var(--ink-3); text-align: center; }
.bs-template-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(230px, 1fr)); gap: 10px; }
.bs-template { display: flex; align-items: flex-start; gap: 10px; padding: 12px; text-align: start; font: inherit; color: var(--ink); background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--r-3, 8px); cursor: pointer; }
.bs-template:hover { border-color: var(--accent); box-shadow: var(--shadow-1); }
.bs-template:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
.bs-template:disabled { opacity: .6; cursor: default; }
.bs-template-icon { flex: none; display: grid; place-items: center; width: 34px; height: 34px; border-radius: var(--r-2); color: var(--accent); background: color-mix(in srgb, var(--accent) 12%, var(--surface)); }
.bs-template-words { display: flex; flex-direction: column; gap: 2px; min-width: 0; }
.bs-template-words strong { font-size: 13.5px; font-weight: 600; }
.bs-template-words small { font-size: 12px; color: var(--ink-2); line-height: 1.35; }
.bs-samples { margin-top: 18px; }
.bs-sample-group { margin-top: 14px; }
.bs-sample-group h3 { margin: 0 0 8px; font-size: 13px; font-weight: 600; color: var(--ink-2); }
.bs-samples-note { margin: -4px 0 10px; font-size: 12.5px; color: var(--ink-2); text-align: center; }
.bs-sample { display: flex; flex-wrap: wrap; align-items: flex-start; gap: 10px; padding: 12px; color: var(--ink); background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--r-3, 8px); }
.bs-sample > .bs-template-words { flex: 1 1 0; }
.bs-sample-actions { flex-basis: 100%; display: flex; flex-wrap: wrap; gap: 6px; }
.bs-template.needs .bs-template-icon, .bs-sample.needs .bs-template-icon { opacity: .45; filter: grayscale(.6); }
.bs-sample-update { font-size: 12px; color: var(--accent-text); }
.bs-sample-actions button { font: inherit; font-size: 12px; min-height: 26px; padding: 3px 10px; border-radius: var(--r-2); border: 1px solid var(--line); background: var(--surface); color: var(--ink); cursor: pointer; }
.bs-sample-actions .bs-sample-data { border-color: var(--accent); color: var(--accent); }
.bs-sample-actions button:hover { border-color: var(--accent); background: color-mix(in srgb, var(--accent) 8%, var(--surface)); }
.bs-sample-actions button:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
.bs-sample-actions button:disabled { opacity: .6; cursor: default; }
.bs-start-actions { display: flex; gap: 8px; align-items: center; justify-content: center; margin-top: 14px; }

.bs-app { flex: 1; width: 100%; height: 100%; display: flex; min-height: 0; min-width: 0; }
.bs-app input[type='checkbox'], .bs-start input[type='checkbox'] { accent-color: var(--accent); }
/* Tables, queries and forms one after another, the pane scrolling as a whole; the foot stays at the foot. */
.bs-side { width: 236px; flex: none; display: flex; flex-direction: column; border-inline-end: 1px solid var(--line-soft); background: var(--sunken); min-height: 0; overflow: auto; }
.bs-side-head { display: flex; align-items: center; justify-content: space-between; padding: 10px 10px 4px 14px; }
.bs-side h2 { margin: 0; font-size: 11.5px; font-weight: 650; text-transform: uppercase; letter-spacing: 0.04em; color: var(--ink-3); }
.bs-tables { list-style: none; margin: 0; padding: 4px 8px; flex: none; display: grid; align-content: start; gap: 2px; }
.bs-table { display: flex; align-items: center; gap: 8px; width: 100%; padding: 7px 8px; border: 0; border-radius: var(--r-2); background: transparent; color: var(--ink); font: inherit; text-align: start; cursor: pointer; }
.bs-table span { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.bs-table small { color: var(--ink-3); font-size: 12px; font-variant-numeric: tabular-nums; }
.bs-table:hover { background: var(--wash); }
.bs-table.on { background: var(--wash-strong); box-shadow: inset 3px 0 0 var(--accent); }
.bs-side-foot { display: grid; gap: 2px; padding: 8px; border-top: 1px solid var(--line-soft); margin-top: auto; position: sticky; bottom: 0; background: var(--sunken); }
.bs-side-foot .bs-import { justify-content: flex-start; margin-bottom: 4px; }

.bs-main { flex: 1 1 0; min-width: 0; display: flex; flex-direction: column; position: relative; min-height: 0; }
.bs-tabs { display: flex; align-items: center; gap: 4px; padding: 10px 16px 0; border-bottom: 1px solid var(--line-soft); }
.bs-tabs h1 { margin: 0 14px 8px 0; font-size: 16px; font-weight: 650; }
.bs-tab { padding: 7px 12px 9px; border: 0; background: transparent; color: var(--ink-2); font: inherit; cursor: pointer; border-bottom: 2px solid transparent; }
.bs-tab:hover { color: var(--ink); }
.bs-tab.on { color: var(--accent-text, var(--accent)); border-bottom-color: var(--accent); font-weight: 600; }
.bs-busy { position: absolute; inset: 0; display: grid; place-items: center; background: color-mix(in srgb, var(--window) 60%, transparent); }

.bs-sheet { flex: 1; min-height: 0; display: flex; flex-direction: column; }
.bs-tools { display: flex; align-items: center; gap: 10px; padding: 8px 16px; }
.bs-count { color: var(--ink-3); font-size: 12.5px; font-variant-numeric: tabular-nums; }
.bs-grid { flex: 1; min-height: 0; overflow: auto; position: relative; border-top: 1px solid var(--line-soft); }
.bs-row { display: grid; height: 30px; align-items: stretch; }
.bs-head { position: sticky; top: 0; z-index: 2; background: var(--window); border-bottom: 1px solid var(--line); width: max-content; min-width: 100%; }
.bs-body { position: relative; width: max-content; min-width: 100%; }
.bs-body .bs-row { position: absolute; left: 0; right: 0; }
.bs-th { display: flex; align-items: center; gap: 4px; padding: 0 8px; border: 0; border-inline-end: 1px solid var(--line-soft); background: transparent; color: var(--ink-2); font: inherit; font-weight: 600; font-size: 12.5px; text-align: start; cursor: pointer; overflow: hidden; }
.bs-th span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.bs-th:hover { background: var(--wash); }
.bs-th-cell { display: flex; align-items: center; min-width: 0; border-inline-end: 1px solid var(--line-soft); }
.bs-th-cell .bs-th { flex: 1; min-width: 0; height: 100%; border-inline-end: 0; }
.bs-td { display: flex; align-items: center; padding: 0 8px; border-inline-end: 1px solid var(--line-soft); border-bottom: 1px solid var(--line-soft); overflow: hidden; min-width: 0; }
.bs-td.t-number, .bs-td.t-integer, .bs-td.t-currency, .bs-td.t-autonumber { justify-content: flex-end; font-variant-numeric: tabular-nums; }
.bs-td.t-yesno { justify-content: center; }
.bs-text { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.bs-pick { display: grid; place-items: center; border-inline-end: 1px solid var(--line-soft); border-bottom: 1px solid var(--line-soft); color: var(--ink-3); }
.bs-row.picked .bs-td, .bs-row.picked .bs-pick { background: var(--wash-strong); }
.bs-body .bs-row:hover .bs-td { background: var(--wash); }
.bs-edit, .bs-new-input { width: 100%; height: 24px; border: 1px solid var(--accent); border-radius: 3px; padding: 0 6px; background: var(--window); color: var(--ink); font: inherit; }
.bs-new-input { border-color: transparent; background: transparent; }
.bs-new-input:focus { border-color: var(--accent); background: var(--window); outline: none; }
.bs-new { position: sticky; bottom: 0; z-index: 2; background: var(--sunken); width: max-content; min-width: 100%; border-top: 1px solid var(--line); }
.bs-auto { color: var(--ink-3); font-size: 12px; }
.bs-hint { color: var(--ink-3); font-size: 12.5px; line-height: 1.5; margin: 8px 16px; }
.bs-error { color: var(--bad, #c62828); margin: 4px 16px; white-space: pre-wrap; }

.bs-design { flex: 1; min-height: 0; overflow: auto; padding: 14px 16px; }
.bs-fields { border-collapse: collapse; width: 100%; max-width: 980px; }
.bs-fields th { text-align: start; font-size: 12px; font-weight: 600; color: var(--ink-3); padding: 4px 6px; border-bottom: 1px solid var(--line); }
.bs-fields td { padding: 5px 6px; border-bottom: 1px solid var(--line-soft); vertical-align: middle; }
.bs-fields .rw-input, .bs-fields select { width: 100%; }
.bs-f-size { max-width: 90px; }
.bs-fixed { color: var(--ink-3); }
.bs-f-new td { border-bottom: 0; padding-top: 10px; }
.bs-design-foot { display: flex; align-items: center; gap: 12px; margin-top: 18px; }
.bs-design-foot .bs-hint { margin: 0; }

.bs-relations { flex: 1; min-height: 0; overflow: auto; padding: 16px 20px; max-width: 860px; }
.bs-relations h2 { margin: 0 0 4px; font-size: 16px; font-weight: 650; }
.bs-relations .bs-hint { margin: 0 0 14px; }
.bs-rel-list { list-style: none; margin: 0 0 18px; padding: 0; display: grid; gap: 6px; }
.bs-rel { display: flex; align-items: center; gap: 10px; padding: 8px 10px; border: 1px solid var(--line-soft); border-radius: var(--r-2); }
.bs-rel .rw-btn { margin-inline-start: auto; }
.bs-chip { font-size: 11.5px; color: var(--ink-2); background: var(--sunken); border-radius: 10px; padding: 2px 8px; }
.bs-rel-form { display: grid; gap: 10px; padding: 14px; border-radius: var(--r-3, 8px); background: var(--sunken); }
.bs-rel-form label { display: grid; grid-template-columns: 120px 1fr 1fr; gap: 8px; align-items: center; }
.bs-rel-form label.bs-check { display: flex; gap: 8px; }
.bs-rel-form .bs-relate { justify-self: start; }

.bs-sql { flex: 1; min-height: 0; display: flex; flex-direction: column; padding-top: 12px; }
.bs-sql-text { margin: 0 16px; min-height: 110px; height: 140px; resize: vertical; font-family: var(--mono, Consolas, monospace); font-size: 13px; line-height: 1.5; padding: 8px 10px; }
.bs-result { flex: 1; min-height: 0; overflow: auto; border-top: 1px solid var(--line-soft); }
.bs-result table { border-collapse: collapse; font-size: 13px; }
.bs-result th { position: sticky; top: 0; background: var(--window); text-align: start; padding: 6px 10px; border-bottom: 1px solid var(--line); border-inline-end: 1px solid var(--line-soft); font-weight: 600; }
.bs-result td { padding: 5px 10px; border-bottom: 1px solid var(--line-soft); border-inline-end: 1px solid var(--line-soft); white-space: nowrap; max-width: 360px; overflow: hidden; text-overflow: ellipsis; }

.bs-field { display: grid; gap: 6px; }
.bs-field span { font-size: 12.5px; color: var(--ink-2); }

/* the side panel: pinned items, a pin on each, sections that fold (D-2026-10-22-05) */
.bs-item { position: relative; display: flex; align-items: center; }
.bs-item .bs-table { flex: 1; min-width: 0; }
.bs-pin { position: absolute; inset-inline-end: 4px; display: grid; place-items: center; width: 22px; height: 22px; border: 0; border-radius: var(--r-1, 4px); background: var(--sunken); color: var(--ink-3); cursor: pointer; opacity: 0; }
.bs-item:hover .bs-pin, .bs-pin:focus-visible { opacity: 1; }
.bs-pin.on { color: var(--accent-text, var(--accent)); }
.bs-pinned .bs-pin { opacity: 0.8; }
.bs-pin:hover { background: var(--wash-strong); color: var(--ink); }
.bs-fold { display: flex; align-items: center; gap: 4px; padding: 0; border: 0; background: none; color: inherit; font: inherit; cursor: pointer; }
.bs-fold h2 { margin: 0; }
`;
