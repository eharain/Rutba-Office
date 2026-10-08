// Worksheets: Automate → scripts, in the window.
//
// A script is JavaScript in the shape of Excel's Office Scripts —
// `function main(workbook) { … }` — kept on this computer (not in the
// workbook, as Office Scripts are kept apart from it). New Script and All
// Scripts open the Code Editor pane: the scripts by name, each to run, edit
// or delete, and the editor with Run and Save. Record Actions writes a
// script from what is typed and formatted while it is on. A script runs in a
// worker that can reach nothing but the snapshot it is handed (no file, no
// network, no window); its edits come back and are applied as one undo step.

import React, { useEffect, useRef, useState } from 'react';
import { Button, Input } from '@rutba/office-ui';

const KEY = 'sheets.scripts';
export const SAMPLE = `function main(workbook) {
  // The sheet you are on, and a range of it.
  const sheet = workbook.getActiveWorksheet();
  const cell = sheet.getRange("A1");
  cell.setValue("Hello from a script");
  cell.getFormat().getFont().setBold(true);
  console.log("Wrote to " + cell.getAddress());
}
`;

/** The scripts kept on this computer. */
export async function loadScripts(shell) {
  const list = await shell.store.get({ key: KEY, fallback: [] }).catch(() => []);
  return Array.isArray(list) ? list : [];
}
export async function saveScripts(shell, list) {
  await shell.store.set({ key: KEY, value: list });
}

/**
 * A script run in its worker on `snapshot`: `{ edits, logs, error }`. A
 * script that runs on past `ms` is stopped, the worker with it.
 */
export function runInWorker(code, snapshot, ms = 15000) {
  return new Promise((resolve) => {
    let worker;
    try { worker = new Worker('scripts-worker.js'); } catch (err) { resolve({ edits: [], logs: [], error: `Scripts cannot run here: ${err.message}` }); return; }
    const timer = setTimeout(() => { worker.terminate(); resolve({ edits: [], logs: [], error: `The script ran for more than ${Math.round(ms / 1000)} seconds and was stopped` }); }, ms);
    worker.onmessage = (e) => { clearTimeout(timer); worker.terminate(); resolve(e.data); };
    worker.onerror = (e) => { clearTimeout(timer); worker.terminate(); resolve({ edits: [], logs: [], error: e.message || 'The script could not run' }); e.preventDefault?.(); };
    worker.postMessage({ code, snapshot });
  });
}

const q = (s) => JSON.stringify(String(s));

/**
 * Record Actions: a script line for each operation the window sends, where
 * one is known — a cell typed into, a format set, another sheet chosen —
 * given the selection it applied to. Others are left out.
 */
export function recordLines(ops, model) {
  const out = [];
  const sel = model?.selection;
  const a1 = (row, col) => { let s = ''; for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s + (row + 1); };
  const rangeOf = () => (sel ? (sel.top === sel.bottom && sel.left === sel.right ? a1(sel.top, sel.left) : `${a1(sel.top, sel.left)}:${a1(sel.bottom, sel.right)}`) : 'A1');
  let draft = null;
  for (const op of ops) {
    if (op.op === 'updateDraft') draft = op.text;
    else if (op.op === 'commitEdit' && draft != null && sel?.active) {
      const at = a1(sel.active.row, sel.active.col);
      out.push(String(draft).startsWith('=') ? `  sheet.getRange(${q(at)}).setFormula(${q(draft)});` : `  sheet.getRange(${q(at)}).setValue(${/^-?\d+(\.\d+)?$/.test(String(draft).trim()) ? Number(draft) : q(draft)});`);
      draft = null;
    } else if (op.op === 'setFormat' && op.delta) {
      const d = op.delta;
      const r = `sheet.getRange(${q(rangeOf())}).getFormat()`;
      if (d.bold !== undefined) out.push(`  ${r}.getFont().setBold(${d.bold === 'toggle' ? 'true' : Boolean(d.bold)});`);
      if (d.italic !== undefined) out.push(`  ${r}.getFont().setItalic(${d.italic === 'toggle' ? 'true' : Boolean(d.italic)});`);
      if (d.fontColour) out.push(`  ${r}.getFont().setColor(${q(d.fontColour)});`);
      if (d.fill) out.push(`  ${r}.getFill().setColor(${q(d.fill)});`);
      if (d.fontSize) out.push(`  ${r}.getFont().setSize(${Number(d.fontSize)});`);
      if (d.numberFormat) out.push(`  sheet.getRange(${q(rangeOf())}).setNumberFormat(${q(d.numberFormat)});`);
      if (d.align) out.push(`  ${r}.setHorizontalAlignment(${q(d.align)});`);
    } else if (op.op === 'sheet' && op.name) {
      out.push(`  sheet = workbook.getWorksheet(${q(op.name)});`, '  sheet.activate();');
    }
  }
  return out;
}

/** The script Record Actions writes from its lines. */
export const recordedScript = (lines) => `function main(workbook) {\n  let sheet = workbook.getActiveWorksheet();\n${lines.join('\n')}\n}\n`;

/** The Code Editor pane: the scripts, one open in the editor, and what its run said. */
export function ScriptsPane({ shell, run, initial = null, recording = false, onRecord, toast }) {
  const [list, setList] = useState([]);
  const fresh = (x) => (x ? { ...x, code: x.code ?? SAMPLE } : null);
  const [open, setOpen] = useState(() => fresh(initial));
  const [output, setOutput] = useState(null);
  const [busy, setBusy] = useState(false);
  const area = useRef(null);
  useEffect(() => { loadScripts(shell).then(setList); }, [shell]);
  useEffect(() => { if (initial) setOpen(fresh(initial)); }, [initial]);

  const store = async (next) => { setList(next); await saveScripts(shell, next); };
  const save = async (script) => {
    const s = { ...script, name: String(script.name || '').trim() || 'Script', modified: Date.now() };
    const next = list.some((x) => x.id === s.id) ? list.map((x) => (x.id === s.id ? s : x)) : [...list, s];
    await store(next);
    setOpen(s);
    toast?.(`Script "${s.name}" saved on this computer`, { tone: 'good', ms: 2500 });
  };
  const go = async (code) => {
    setBusy(true);
    setOutput(null);
    const out = await run(code);
    setOutput(out);
    setBusy(false);
  };

  return (
    <div className="sh-scripts">
      <div className="sh-scripts-tools">
        <Button icon="plus" label="New Script" className="sh-script-new" onClick={() => setOpen({ id: `s${Date.now().toString(36)}`, name: `Script ${list.length + 1}`, code: SAMPLE })} />
        <Button icon={recording ? 'stop' : 'video'} label={recording ? 'Stop' : 'Record'} pressed={recording} className="sh-script-record" title={recording ? 'Stop recording — what was done becomes a script' : 'Record Actions — what you type and format becomes a script'} onClick={onRecord} />
      </div>
      {open ? (
        <div className="sh-script-editor">
          <Input className="sh-script-name" value={open.name} onChange={(e) => setOpen((o) => ({ ...o, name: e.target.value }))} />
          <textarea
            ref={area}
            className="sh-script-code"
            spellCheck={false}
            value={open.code}
            onChange={(e) => setOpen((o) => ({ ...o, code: e.target.value }))}
            onKeyDown={(e) => {
              // Tab indents, as an editor's does.
              if (e.key === 'Tab') {
                e.preventDefault();
                const el = e.currentTarget;
                const { selectionStart: a, selectionEnd: b, value } = el;
                const next = value.slice(0, a) + '  ' + value.slice(b);
                setOpen((o) => ({ ...o, code: next }));
                requestAnimationFrame(() => { el.selectionStart = el.selectionEnd = a + 2; });
              }
            }}
          />
          <div className="sh-script-buttons">
            <Button primary icon="play" label={busy ? 'Running…' : 'Run'} className="sh-script-run" disabled={busy} onClick={() => go(open.code)} />
            <Button icon="save" label="Save" className="sh-script-save" onClick={() => save(open)} />
            <Button label="Close" onClick={() => { setOpen(null); setOutput(null); }} />
          </div>
          {output ? (
            <div className={`sh-script-output${output.error ? ' bad' : ''}`}>
              {output.error ? <div className="sh-script-error">{output.error}</div> : <div>{`Done: ${output.cells || 0} cell${output.cells === 1 ? '' : 's'} set${output.edits?.length ? `, ${output.edits.length} change${output.edits.length === 1 ? '' : 's'} in one undo step` : ''}.`}</div>}
              {(output.logs || []).map((l, i) => <div key={i} className="sh-script-log">{l}</div>)}
            </div>
          ) : null}
        </div>
      ) : null}
      <div className="sh-scripts-head">All Scripts</div>
      {list.length ? list.map((s) => (
        <div key={s.id} className="sh-script-item" data-script={s.id}>
          <span className="sh-script-title">{s.name}</span>
          <Button icon="play" className="sh-script-item-run" title={`Run ${s.name}`} onClick={() => { setOpen(s); go(s.code); }} />
          <Button icon="textbox" title={`Edit ${s.name}`} onClick={() => { setOpen(s); setOutput(null); }} />
          <Button icon="trash" className="sh-script-delete" title={`Delete ${s.name}`} onClick={() => store(list.filter((x) => x.id !== s.id))} />
        </div>
      )) : <div className="sh-scripts-empty">No scripts yet. New Script writes one; Record turns what you do into one.</div>}
    </div>
  );
}

export const SCRIPTS_CSS = `
.sh-scripts { display: flex; flex-direction: column; gap: 6px; padding: 8px; height: 100%; box-sizing: border-box; }
.sh-scripts-tools, .sh-script-buttons { display: flex; gap: 4px; flex-wrap: wrap; }
.sh-script-editor { display: flex; flex-direction: column; gap: 6px; }
.sh-script-code { min-height: 220px; resize: vertical; font: 12px/1.45 Consolas, 'Cascadia Mono', monospace; padding: 8px; border: 1px solid var(--line); border-radius: 4px; background: var(--surface, #fff); color: inherit; tab-size: 2; white-space: pre; }
.sh-script-output { font: 12px/1.4 Consolas, monospace; border: 1px solid var(--line); border-radius: 4px; padding: 6px; max-height: 160px; overflow: auto; }
.sh-script-output.bad, .sh-script-error { color: var(--bad, #c62828); }
.sh-script-log { opacity: .85; }
.sh-scripts-head { font-weight: 600; margin-top: 6px; }
.sh-script-item { display: flex; align-items: center; gap: 4px; border: 1px solid var(--line); border-radius: 4px; padding: 2px 4px 2px 8px; }
.sh-script-title { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sh-scripts-empty { font-size: 12px; opacity: .75; }
`;
