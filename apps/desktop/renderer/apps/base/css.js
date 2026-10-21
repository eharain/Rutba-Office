// Databases' look: a list of tables, a sheet of records, nothing more than is needed.

export const BASE_CSS = `
.bs-start { height: 100%; display: grid; place-items: center; }
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
`;
