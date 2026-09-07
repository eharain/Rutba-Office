// Rutba Worksheets.
//
// The grid is a window onto a model that lives in the backend. Scrolling asks
// for the viewport that is about to be painted and nothing else, so the size of
// the workbook has no bearing on the size of this window's heap — a hundred
// thousand rows scroll exactly as fast as a hundred.
//
// Every edit is an operation sent by name. The engine owns the undo history,
// the formula graph and the file, which is why the same operations work whether
// the sheet came from a .xlsx, a .csv or an .ods.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Ribbon, Group, Button, Separator, Icon, Spacer, Chip, Empty, Spinner, useToast, useMenu, useCommands, menuItems, Input } from '@rutba/office-ui';
import { AppFrame, useAppMenu, pickOpen, pickSave, confirmDiscard, useFileDrop, openInApp } from '../shell.js';

const colLabel = (n) => {
  let s = '';
  let x = n;
  while (x >= 0) {
    s = String.fromCharCode(65 + (x % 26)) + s;
    x = Math.floor(x / 26) - 1;
  }
  return s;
};

export default function Sheets({ app, shell, boot }) {
  const toast = useToast();
  const [doc, setDoc] = useState(null);
  const [model, setModel] = useState(null);
  const [busy, setBusy] = useState(true);
  const [tab, setTab] = useState('home');
  const [error, setError] = useState(null);
  const gridRef = useRef(null);
  const editorRef = useRef(null);
  const menu = useMenu();
  const appMenu = useAppMenu({ shell, appKey: 'sheets', onNew: () => shell.win.create({ app: 'sheets' }), onOpen: () => openFileRef.current?.() });
  const openFileRef = useRef(null);

  const dispatch = useCallback(
    async (...ops) => {
      if (!doc) return null;
      try {
        const next = await shell.doc.apply({ id: doc.id, ops });
        setDoc(next);
        setModel(next.model);
        return next;
      } catch (err) {
        toast(err.message, { tone: 'bad' });
        return null;
      }
    },
    [doc, shell, toast]
  );

  /* ── opening ─────────────────────────────────────────────────────────── */

  const load = useCallback(
    async (fn) => {
      setBusy(true);
      setError(null);
      try {
        const opened = await fn();
        setDoc(opened);
        setModel(opened.model);
        if (opened.path) shell.app.addRecent({ path: opened.path, app: 'sheets' }).catch(() => {});
        if (opened.converted?.from) {
          toast(`Opened from ${opened.converted.from.toUpperCase()}. Saving will write a .xlsx.`, { ms: 5200 });
        }
      } catch (err) {
        setError(err.message);
      } finally {
        setBusy(false);
      }
    },
    [shell, toast]
  );

  useEffect(() => {
    const params = new URLSearchParams(location.search);
    const template = params.get('template');
    if (boot.file) load(() => shell.doc.open({ path: boot.file }));
    else load(() => shell.doc.new({ kind: 'sheets', template: template && template !== 'blank' ? template : 'sheet' }));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const openFile = useCallback(async () => {
    const file = await pickOpen(shell, 'sheets');
    if (!file) return;
    if (doc?.dirty) {
      const choice = await confirmDiscard(shell, doc.name);
      if (choice === 'cancel') return;
      if (choice === 'save') await save();
    }
    openInApp(shell, file, 'sheets');
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shell, doc]);

  const save = useCallback(
    async (as = false) => {
      if (!doc) return;
      let target = doc.path;
      if (as || !target) {
        target = await pickSave(shell, 'sheets', doc.path || doc.name);
        if (!target) return;
      }
      try {
        const saved = await shell.doc.save({ id: doc.id, path: target });
        setDoc((d) => ({ ...d, ...saved, dirty: false }));
        shell.app.addRecent({ path: saved.path, app: 'sheets' }).catch(() => {});
        toast(`Saved ${saved.path.split(/[\\/]/).pop()}`, { tone: 'good' });
      } catch (err) {
        toast(err.message, { tone: 'bad' });
      }
    },
    [doc, shell, toast]
  );

  const exportAs = useCallback(
    async (format) => {
      if (!doc) return;
      const target = await shell.dialog.save({
        title: `Export as ${format.toUpperCase()}`,
        defaultPath: (doc.path || doc.name).replace(/\.[^.]+$/, `.${format}`),
        filters: [{ name: format.toUpperCase(), extensions: [format] }],
      });
      if (!target) return;
      try {
        await shell.doc.export({ id: doc.id, format, path: target });
        toast(`Exported ${target.split(/[\\/]/).pop()}`, { tone: 'good' });
      } catch (err) {
        toast(err.message, { tone: 'bad' });
      }
    },
    [doc, shell, toast]
  );

  openFileRef.current = openFile;

  useFileDrop(useCallback((files) => files.forEach((f) => openInApp(shell, f, 'sheets')), [shell]));

  /* ── viewport ────────────────────────────────────────────────────────── */

  const syncViewport = useCallback(async () => {
    const el = gridRef.current;
    if (!el || !doc) return;
    const next = await shell.doc.viewport({
      id: doc.id,
      width: el.clientWidth,
      height: el.clientHeight,
      x: el.scrollLeft,
      y: el.scrollTop,
    });
    setModel(next);
  }, [doc, shell]);

  useEffect(() => {
    const el = gridRef.current;
    if (!el) return undefined;
    let frame = 0;
    const onScroll = () => {
      cancelAnimationFrame(frame);
      frame = requestAnimationFrame(syncViewport);
    };
    el.addEventListener('scroll', onScroll, { passive: true });
    const ro = new ResizeObserver(onScroll);
    ro.observe(el);
    return () => {
      el.removeEventListener('scroll', onScroll);
      ro.disconnect();
      cancelAnimationFrame(frame);
    };
  }, [syncViewport]);

  /* ── commands ────────────────────────────────────────────────────────── */

  const commands = useMemo(
    () => ({
      'file.new': { label: 'New', icon: 'new', key: 'Mod+N', run: () => shell.win.create({ app: 'sheets' }) },
      'file.open': { label: 'Open…', icon: 'open', key: 'Mod+O', run: openFile },
      'file.save': { label: 'Save', icon: 'save', key: 'Mod+S', run: () => save(false) },
      'file.saveAs': { label: 'Save as…', icon: 'save', key: 'Mod+Shift+S', run: () => save(true) },
      'edit.undo': { label: 'Undo', icon: 'undo', key: 'Mod+Z', run: async () => { const n = await shell.doc.undo({ id: doc.id }); setDoc(n); setModel(n.model); } },
      'edit.redo': { label: 'Redo', icon: 'redo', key: 'Mod+Y', run: async () => { const n = await shell.doc.redo({ id: doc.id }); setDoc(n); setModel(n.model); } },
      'edit.copy': { label: 'Copy', icon: 'copy', key: 'Mod+C', run: () => dispatch({ op: 'copy' }) },
      'edit.clear': { label: 'Clear', icon: 'close', key: 'Delete', run: () => dispatch({ op: 'clear' }) },
      'sheet.autoSum': { label: 'AutoSum', icon: 'sum', run: () => dispatch({ op: 'autoSum', fn: 'SUM' }) },
      'sheet.merge': { label: 'Merge cells', icon: 'table', run: () => dispatch({ op: 'merge' }) },
      'sheet.insertRow': { label: 'Insert row', icon: 'plus', run: () => dispatch({ op: 'insertRows', at: model?.selection.top ?? 0, count: 1 }) },
      'sheet.deleteRow': { label: 'Delete row', icon: 'minus', run: () => dispatch({ op: 'deleteRows', at: model?.selection.top ?? 0, count: 1 }) },
      'sheet.insertCol': { label: 'Insert column', icon: 'plus', run: () => dispatch({ op: 'insertCols', at: model?.selection.left ?? 0, count: 1 }) },
      'sheet.deleteCol': { label: 'Delete column', icon: 'minus', run: () => dispatch({ op: 'deleteCols', at: model?.selection.left ?? 0, count: 1 }) },
      'sheet.sortAsc': { label: 'Sort A→Z', icon: 'sort', run: () => dispatch({ op: 'sort', ascending: true }) },
      'sheet.sortDesc': { label: 'Sort Z→A', icon: 'sort', run: () => dispatch({ op: 'sort', ascending: false }) },
      'sheet.chart': { label: 'Chart', icon: 'chart', run: () => dispatch({ op: 'insertChart', kind: 'column' }) },
    }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [doc, model, dispatch, openFile, save, shell]
  );

  useCommands(commands, [doc, model]);

  /* ── grid interaction ────────────────────────────────────────────────── */

  const editing = model?.editing;

  useEffect(() => {
    if (editing && editorRef.current) {
      editorRef.current.focus();
      const el = editorRef.current;
      el.setSelectionRange(el.value.length, el.value.length);
    }
  }, [editing?.row, editing?.col]);

  const onKeyDown = useCallback(
    async (e) => {
      if (!model) return;
      if (editing) {
        if (e.key === 'Escape') {
          e.preventDefault();
          await dispatch({ op: 'cancelEdit' });
        } else if (e.key === 'Enter') {
          e.preventDefault();
          await dispatch({ op: 'commitEdit', move: e.shiftKey ? 'up' : 'down' });
        } else if (e.key === 'Tab') {
          e.preventDefault();
          await dispatch({ op: 'commitEdit', move: e.shiftKey ? 'left' : 'right' });
        }
        return;
      }

      const arrows = { ArrowUp: 'up', ArrowDown: 'down', ArrowLeft: 'left', ArrowRight: 'right' };
      if (arrows[e.key]) {
        e.preventDefault();
        await dispatch({ op: 'move', direction: arrows[e.key], extend: e.shiftKey, jump: e.ctrlKey || e.metaKey });
        return;
      }
      if (e.key === 'Tab') {
        e.preventDefault();
        await dispatch({ op: 'tab', back: e.shiftKey });
        return;
      }
      if (e.key === 'Enter') {
        e.preventDefault();
        await dispatch({ op: 'enter', back: e.shiftKey });
        return;
      }
      if (e.key === 'F2') {
        e.preventDefault();
        await dispatch({ op: 'beginEdit' });
        return;
      }
      if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        await dispatch({ op: 'clear' });
        return;
      }
      // Typing a printable character starts an edit that replaces the cell.
      if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        e.preventDefault();
        await dispatch({ op: 'beginEdit', replace: true, initial: e.key });
      }
    },
    [model, editing, dispatch]
  );

  const cellStyle = (cell) => {
    const s = cell.style || {};
    return {
      left: cell.x,
      top: cell.y,
      width: cell.width,
      height: cell.height,
      textAlign: cell.align || (typeof cell.text === 'string' && cell.isFormula ? 'right' : undefined),
      fontWeight: s.bold ? 700 : undefined,
      fontStyle: s.italic ? 'italic' : undefined,
      textDecoration: s.underline ? 'underline' : undefined,
      color: s.colour || cell.colour || undefined,
      background: s.fill || undefined,
      fontSize: s.size ? `${s.size}px` : undefined,
      fontFamily: s.font || undefined,
      justifyContent: cell.align === 'center' ? 'center' : cell.align === 'right' ? 'flex-end' : undefined,
      alignItems: cell.valign === 'middle' ? 'center' : cell.valign === 'top' ? 'flex-start' : undefined,
    };
  };

  if (error) {
    return (
      <AppFrame app={app} shell={shell} title="Worksheets" menu={appMenu}>
        <Empty icon="sheets" title="This file could not be opened">
          {error}
        </Empty>
      </AppFrame>
    );
  }

  const sel = model?.selection;
  const status = model?.status;

  return (
    <AppFrame
      app={app}
      shell={shell}
      title={doc?.name || 'Worksheets'}
      subtitle={doc?.converted ? `from ${doc.converted.from.toUpperCase()}` : null}
      dirty={doc?.dirty}
      menu={appMenu}
      ribbon={
        <Ribbon
          tabs={[
            { id: 'home', label: 'Home' },
            { id: 'insert', label: 'Insert' },
            { id: 'data', label: 'Data' },
            { id: 'view', label: 'View' },
          ]}
          active={tab}
          onTab={setTab}
          quick={
            <>
              <Button icon="save" title="Save" onClick={() => save(false)} />
              <Button icon="undo" title="Undo" disabled={!doc?.canUndo} onClick={() => commands['edit.undo'].run()} />
              <Button icon="redo" title="Redo" disabled={!doc?.canRedo} onClick={() => commands['edit.redo'].run()} />
            </>
          }
        >
          {tab === 'home' ? (
            <>
              <Group label="File">
                <Button tall icon="new" label="New" onClick={() => shell.win.create({ app: 'sheets' })} />
                <Button tall icon="open" label="Open" onClick={openFile} />
                <Button tall icon="save" label="Save" onClick={() => save(false)} />
              </Group>
              <Group label="Clipboard">
                <Button tall icon="copy" label="Copy" onClick={() => commands['edit.copy'].run()} />
                <Button tall icon="paste" label="Paste" onClick={async () => dispatch({ op: 'paste', text: await shell.clipboard.readText() })} />
                <Button tall icon="wand" label="Format" title="Copy formatting" onClick={() => dispatch({ op: 'formatBrush' })} />
              </Group>
              <Group label="Cells">
                <Button icon="table" label="Merge" onClick={() => dispatch({ op: 'merge' })} />
                <Button icon="minus" label="Unmerge" onClick={() => dispatch({ op: 'unmerge' })} />
                <Button icon="close" label="Clear" onClick={() => dispatch({ op: 'clear' })} />
              </Group>
              <Group label="Rows and columns">
                <Button icon="plus" label="Row" onClick={() => commands['sheet.insertRow'].run()} />
                <Button icon="minus" label="Row" onClick={() => commands['sheet.deleteRow'].run()} />
                <Button icon="plus" label="Column" onClick={() => commands['sheet.insertCol'].run()} />
                <Button icon="minus" label="Column" onClick={() => commands['sheet.deleteCol'].run()} />
              </Group>
              <Group label="Editing">
                <Button tall icon="sum" label="AutoSum" onClick={() => commands['sheet.autoSum'].run()} />
                <Button tall icon="sort" label="Sort" onClick={() => commands['sheet.sortAsc'].run()} />
              </Group>
            </>
          ) : tab === 'insert' ? (
            <>
              <Group label="Charts">
                <Button tall icon="chart" label="Column" onClick={() => dispatch({ op: 'insertChart', kind: 'column' })} />
                <Button tall icon="chart" label="Line" onClick={() => dispatch({ op: 'insertChart', kind: 'line' })} />
                <Button tall icon="chart" label="Pie" onClick={() => dispatch({ op: 'insertChart', kind: 'pie' })} />
              </Group>
              <Group label="Illustrations">
                <Button tall icon="shape" label="Shape" onClick={() => dispatch({ op: 'insertShape', geometry: 'rect', text: '' })} />
              </Group>
            </>
          ) : tab === 'data' ? (
            <>
              <Group label="Sort and filter">
                <Button tall icon="sort" label="A → Z" onClick={() => commands['sheet.sortAsc'].run()} />
                <Button tall icon="sort" label="Z → A" onClick={() => commands['sheet.sortDesc'].run()} />
              </Group>
              <Group label="Export">
                <Button tall icon="export" label="CSV" onClick={() => exportAs('csv')} />
                <Button tall icon="export" label="TSV" onClick={() => exportAs('tsv')} />
              </Group>
            </>
          ) : (
            <>
              <Group label="Zoom">
                <Button icon="zoomOut" label="Out" onClick={() => shell.win.zoom({ delta: -0.1 })} />
                <Button icon="zoomIn" label="In" onClick={() => shell.win.zoom({ delta: 0.1 })} />
                <Button icon="check" label="100%" onClick={() => shell.win.zoom({ reset: true })} />
              </Group>
              <Group label="Window">
                <Button icon="maximize" label="Full screen" onClick={() => shell.win.fullscreen({})} />
              </Group>
            </>
          )}
        </Ribbon>
      }
      status={
        <>
          <span>{doc?.path || 'Not saved yet'}</span>
          <Spacer />
          {status ? (
            <>
              <Chip>Count {status.count}</Chip>
              {status.numeric ? <Chip>Sum {formatNumber(status.sum)}</Chip> : null}
              {status.numeric ? <Chip>Average {formatNumber(status.average)}</Chip> : null}
            </>
          ) : null}
          <Chip>{model?.activeSheet || ''}</Chip>
          <Chip>{sel?.ref || ''}</Chip>
        </>
      }
    >
      {busy || !model ? (
        <div style={{ flex: 1, display: 'grid', placeItems: 'center' }}>
          <Spinner style={{ width: 22, height: 22 }} />
        </div>
      ) : (
        <div className="sh" onKeyDown={onKeyDown} tabIndex={0} ref={(el) => el && !editing && document.activeElement === document.body && el.focus()}>
          <style>{CSS}</style>

          <div className="sh-formula">
            <div className="sh-namebox">{sel?.ref}</div>
            <Icon name="formula" size={14} style={{ color: 'var(--ink-3)' }} />
            <Input
              value={editing ? editing.draft : model.formulaBar ?? ''}
              onChange={(e) => (editing ? dispatch({ op: 'updateDraft', text: e.target.value }) : dispatch({ op: 'beginEdit', replace: true, initial: e.target.value }))}
              onKeyDown={(e) => {
                if (e.key === 'Enter') dispatch({ op: 'commitEdit', move: 'down' });
                if (e.key === 'Escape') dispatch({ op: 'cancelEdit' });
              }}
              style={{ flex: 1, border: 0, background: 'transparent' }}
            />
          </div>

          {/*
            Four panes in a CSS grid, not four absolutely positioned layers.
            `position: sticky` needs an element that is in normal flow, and in a
            grid each pane keeps its own track — so the corner sticks to both
            edges, the headers stick to one each, and none of them displaces the
            cells the way stacked block elements would.
          */}
          <div className="sh-grid" ref={gridRef}>
            <div
              className="sh-canvas"
              style={{
                gridTemplateColumns: `${model.headerWidth}px ${model.total.width}px`,
                gridTemplateRows: `${model.headerHeight}px ${model.total.height}px`,
              }}
            >
              <div className="sh-corner" />

              <div className="sh-colheads" style={{ height: model.headerHeight }}>
                {(model.columns || []).map((c) => (
                  <div
                    key={c.index}
                    className={`sh-head${c.index >= (sel?.left ?? -1) && c.index <= (sel?.right ?? -2) ? ' active' : ''}`}
                    style={{ left: c.x, width: c.width, height: model.headerHeight }}
                    onClick={(e) => dispatch({ op: 'selectColumn', col: c.index, extend: e.shiftKey })}
                    onContextMenu={(e) => menu.open(e, menuItems(commands, ['sheet.insertCol', 'sheet.deleteCol', '-', 'sheet.sortAsc', 'sheet.sortDesc']))}
                  >
                    {c.label || colLabel(c.index)}
                  </div>
                ))}
              </div>

              <div className="sh-rowheads" style={{ width: model.headerWidth }}>
                {(model.rows || []).map((r) => (
                  <div
                    key={r.index}
                    className={`sh-head${r.index >= (sel?.top ?? -1) && r.index <= (sel?.bottom ?? -2) ? ' active' : ''}`}
                    style={{ top: r.y, height: r.height, width: model.headerWidth }}
                    onClick={(e) => dispatch({ op: 'selectRow', row: r.index, extend: e.shiftKey })}
                    onContextMenu={(e) => menu.open(e, menuItems(commands, ['sheet.insertRow', 'sheet.deleteRow']))}
                  >
                    {r.label ?? r.index + 1}
                  </div>
                ))}
              </div>

              <div className="sh-cells">
                {model.cells.map((cell) => (
                  <div
                    key={cell.ref}
                    className={`sh-cell${cell.selected ? ' sel' : ''}${cell.active ? ' active' : ''}${cell.isError ? ' err' : ''}`}
                    style={cellStyle(cell)}
                    onMouseDown={(e) => dispatch({ op: 'select', row: cell.row, col: cell.col, extend: e.shiftKey })}
                    onDoubleClick={() => dispatch({ op: 'beginEdit' })}
                    onContextMenu={(e) => menu.open(e, menuItems(commands, ['edit.copy', 'edit.clear', '-', 'sheet.insertRow', 'sheet.insertCol', '-', 'sheet.merge']))}
                    title={cell.note || undefined}
                  >
                    {cell.text}
                  </div>
                ))}

                {editing ? (
                  <input
                    ref={editorRef}
                    className="sh-editor"
                    value={editing.draft}
                    onChange={(e) => dispatch({ op: 'updateDraft', text: e.target.value })}
                    style={{
                      left: editing.x ?? model.cells.find((c) => c.active)?.x ?? 0,
                      top: editing.y ?? model.cells.find((c) => c.active)?.y ?? 0,
                      width: Math.max(80, model.cells.find((c) => c.active)?.width ?? 80),
                      height: model.cells.find((c) => c.active)?.height ?? 20,
                    }}
                  />
                ) : null}
              </div>
            </div>
          </div>

          <div className="sh-tabs">
            {(model.sheets || []).map((name) => (
              <button
                key={name}
                type="button"
                className={`sh-tab${name === model.activeSheet ? ' active' : ''}`}
                onClick={() => dispatch({ op: 'sheet', name })}
              >
                {name}
              </button>
            ))}
          </div>
          {menu.node}
        </div>
      )}
    </AppFrame>
  );
}

function formatNumber(n) {
  if (!Number.isFinite(n)) return '—';
  return Math.abs(n) >= 1000 ? n.toLocaleString(undefined, { maximumFractionDigits: 2 }) : String(Math.round(n * 1000) / 1000);
}

const CSS = `
.sh { flex: 1; display: flex; flex-direction: column; min-height: 0; outline: none; }
.sh-formula {
  display: flex; align-items: center; gap: 8px; padding: 4px 10px;
  background: var(--surface); border-bottom: 1px solid var(--line);
}
.sh-namebox {
  min-width: 78px; padding: 3px 8px; font-size: 12px; font-variant-numeric: tabular-nums;
  border: 1px solid var(--line); border-radius: var(--r-2); background: var(--chrome); text-align: center;
}
.sh-formula .rw-input { font-family: var(--mono); font-size: 12.5px; }
.sh-formula .rw-input:focus { box-shadow: none; }

.sh-grid { flex: 1; overflow: auto; position: relative; background: var(--surface); }
.sh-canvas { display: grid; }
.sh-corner {
  position: sticky; left: 0; top: 0; z-index: 4; background: var(--chrome);
  border-right: 1px solid var(--line); border-bottom: 1px solid var(--line);
  grid-row: 1; grid-column: 1;
}
.sh-colheads { position: sticky; top: 0; z-index: 3; grid-row: 1; grid-column: 2; }
.sh-rowheads { position: sticky; left: 0; z-index: 3; grid-row: 2; grid-column: 1; background: var(--chrome); }
.sh-cells { grid-row: 2; grid-column: 2; }
.sh-colheads, .sh-rowheads, .sh-cells { position: relative; }
.sh-colheads { position: sticky; }
.sh-rowheads { position: sticky; }
.sh-head {
  position: absolute; display: grid; place-items: center; font-size: 11.5px; color: var(--ink-2);
  background: var(--chrome); border-right: 1px solid var(--line-soft); border-bottom: 1px solid var(--line);
  font-variant-numeric: tabular-nums; user-select: none;
}
.sh-head.active { background: var(--selected); color: var(--accent); font-weight: 600; }
.sh-cell {
  position: absolute; display: flex; align-items: center; padding: 0 5px;
  border-right: 1px solid var(--line-soft); border-bottom: 1px solid var(--line-soft);
  font-size: 12.5px; overflow: hidden; white-space: nowrap; background: var(--surface);
}
.sh-cell.sel { background: var(--selected); }
.sh-cell.active { outline: 2px solid var(--accent); outline-offset: -1px; z-index: 2; background: var(--surface); }
.sh-cell.err { color: var(--bad); }
.sh-editor {
  position: absolute; z-index: 6; border: 2px solid var(--accent); border-radius: 2px;
  padding: 0 4px; font: inherit; font-size: 12.5px; background: var(--surface); color: var(--ink); outline: none;
}
.sh-tabs {
  display: flex; align-items: stretch; gap: 2px; padding: 3px 8px;
  background: var(--chrome); border-top: 1px solid var(--line); overflow-x: auto;
}
.sh-tab {
  border: 0; background: transparent; color: var(--ink-2); padding: 3px 12px;
  font-size: 12px; border-radius: var(--r-2); white-space: nowrap;
}
.sh-tab:hover { background: var(--hover); }
.sh-tab.active { background: var(--surface); color: var(--accent); font-weight: 600; box-shadow: var(--shadow-1); }
`;
