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
import { Button, Icon, Spacer, Chip, Empty, Spinner, Dialog, useToast, useMenu, useCommands, menuItems, Input } from '@rutba/office-ui';
import { AppFrame, useAppMenu, pickOpen, pickSave, confirmDiscard, useFileDrop, openInApp , useDirtyGuard } from '../shell.js';
import { PrintDialog, defaultPrintOptions } from '../print.js';
import SheetsRibbon, { FUNCTIONS } from './sheets/ribbon.js';
import { SITE } from '@rutba/office-formats/registry';
import { SymbolDialog } from './word/dialogs.js';
import {
  GoToDialog, FunctionDialog, StatisticsDialog, SheetShortcutsDialog, SizeDialog, parseRef,
} from './sheets/dialogs.js';
import {
  ConditionalDialog, ValidationDialog, GoalSeekDialog, DataTableDialog, NameManager, FindDialog, PivotDialog,
} from './sheets/dialogs.js';

const colLabel = (n) => {
  let s = '';
  let x = n;
  while (x >= 0) {
    s = String.fromCharCode(65 + (x % 26)) + s;
    x = Math.floor(x / 26) - 1;
  }
  return s;
};

/**
 * A selection move that scrolled nothing answers with the selection and the
 * active cell's fields rather than a frame; the cells already on screen are
 * re-flagged here. Cells whose flags did not change keep their identity.
 */
function withSelection(model, patch) {
  const sel = patch.selection;
  const inside = (row, col) => {
    if (row >= sel.top && row <= sel.bottom && col >= sel.left && col <= sel.right) return true;
    return Boolean(sel.ranges && sel.ranges.some((q) => row >= q.top && row <= q.bottom && col >= q.left && col <= q.right));
  };
  const cells = (model.cells || []).map((c) => {
    const selected = inside(c.row, c.col);
    const active = c.row === sel.active.row && c.col === sel.active.col;
    return selected === Boolean(c.selected) && active === Boolean(c.active) ? c : { ...c, selected, active };
  });
  return { ...model, ...patch, cells };
}

export default function Sheets({ app, shell, boot }) {
  const toast = useToast();
  const [doc, setDoc] = useState(null);
  const [model, setModel] = useState(null);
  const [busy, setBusy] = useState(true);
  const [tab, setTab] = useState('home');
  const [error, setError] = useState(null);
  // Which of the ribbon's dialogs is open, by name. One piece of state rather
  // than seven booleans, because only one of them can be open at a time.
  const [dialog, setDialog] = useState(null);
  /**
   * How the grid is shown. None of it is in the workbook: gridlines,
   * headings and the formula bar are Excel's View toggles, "show formulas"
   * is Ctrl+`, and the page setup is what the PDF export will lay out to.
   */
  const [view, setView] = useState({
    gridlines: true, headings: true, formulaBar: true, formulas: false,
    page: { orientation: 'portrait', margins: 'normal', size: 'A4' },
  });
  const patchView = useCallback((patch) => setView((v) => ({ ...v, ...(typeof patch === 'function' ? patch(v) : patch) })), []);
  const gridRef = useRef(null);
  /** The element that takes the keys: the grid's own container. */
  const shRef = useRef(null);

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
        if (next.patch) setModel((m) => (m ? withSelection(m, next.patch) : m));
        else setModel(next.model);
        return next;
      } catch (err) {
        toast(err.message, { tone: 'bad' });
        return null;
      }
    },
    [doc, shell, toast]
  );

  /**
   * Navigation, coalesced. A held arrow key fires thirty times a second and
   * each press used to be its own round trip and its own full frame, so the
   * grid fell further behind the key the longer it was held and caught up
   * seconds after it was released. Moves that arrive while a frame is on its
   * way are kept and sent together when it lands: one frame per reply,
   * however fast the keys.
   */
  const navQueue = useRef([]);
  const navBusy = useRef(false);
  const navigate = useCallback(
    (op) => {
      navQueue.current.push(op);
      if (navBusy.current) return;
      navBusy.current = true;
      (async () => {
        try {
          while (navQueue.current.length) {
            const ops = navQueue.current.splice(0);
            await dispatch(...ops);
          }
        } finally {
          navBusy.current = false;
        }
      })();
    },
    [dispatch]
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
          // What saving will actually do: a .csv opened here saves as .csv,
          // an .ods cannot be saved at all until it is given a new name.
          const was = opened.converted.from.toUpperCase();
          toast(
            opened.converted.writesBack
              ? `Opened from ${was}. Saving writes the ${was} back.`
              : `Opened from ${was}. This build cannot write ${was} — Save as will write a .xlsx.`,
            { ms: 5200 }
          );
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
    const recover = params.get('recover');
    // A recovery copy the launcher offered: opened as the workbook it came
    // from, dirty, because what is on screen is not what is on disk.
    if (recover) load(() => shell.doc.recover({ file: recover }).then((r) => { toast('Recovered unsaved work. Save it to keep it.', { ms: 6000 }); return r; }));
    else if (boot.file) load(() => shell.doc.open({ path: boot.file, kind: 'sheet' }));
    else load(() => shell.doc.new({ kind: 'sheets', template: template && template !== 'blank' ? template : 'sheet' }));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Closing a window with unsaved work must ask, not discard.
  useDirtyGuard({ shell, dirty: doc?.dirty, name: doc?.name, onSave: () => save(false) });
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
      if (!doc) return false;
      let target = doc.path;
      if (as || !target) {
        target = await pickSave(shell, 'sheets', doc.path || doc.name);
        // A cancelled Save As is not a save; the caller must know.
        if (!target) return false;
      }
      try {
        const saved = await shell.doc.save({ id: doc.id, path: target });
        setDoc((d) => ({ ...d, ...saved, dirty: false }));
        shell.app.addRecent({ path: saved.path, app: 'sheets' }).catch(() => {});
        toast(`Saved ${saved.path.split(/[\\/]/).pop()}`, { tone: 'good' });
        return true;
      } catch (err) {
        toast(err.message, { tone: 'bad' });
        return false;
      }
    },
    [doc, shell, toast]
  );

  const exportAs = useCallback(
    async (format, options = null) => {
      if (!doc) return;
      const target = await shell.dialog.save({
        title: `Export as ${format.toUpperCase()}`,
        defaultPath: (doc.path || doc.name).replace(/\.[^.]+$/, `.${format}`),
        filters: [{ name: format.toUpperCase(), extensions: [format] }],
      });
      // A cancelled Save As is not a save; the caller must know.
      if (!target) return false;
      try {
        // A PDF is a print job: the page setup decides where the pages fall,
        // and a workbook or a deck must be laid out before it is one.
        if (format === 'pdf') await shell.print.pdf({ id: doc.id, path: target, options: options || defaultPrintOptions('sheet') });
        else await shell.doc.export({ id: doc.id, format, path: target });
        toast(`Exported ${target.split(/[\\/]/).pop()}`, { tone: 'good' });
      } catch (err) {
        toast(err.message, { tone: 'bad' });
        return false;
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
      'file.print': { label: 'Print…', icon: 'print', key: 'Mod+P', global: true, run: () => setDialog('print') },
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

  /**
   * What is being typed lives here while it is being typed.
   *
   * The draft used to be sent to the engine on every keystroke and the input's
   * value read back from the answer, so each character made a round trip and
   * the field lagged behind the keyboard — fast typing dropped characters and
   * reordered them. The engine does not need to see a half-typed formula; it
   * needs the finished one. So the cell holds its own text and hands it over
   * once, on commit.
   */
  const [draft, setDraft] = useState(null);

  /**
   * An edit is starting, decided here and now.
   *
   * Beginning an edit is a round trip, and the characters that follow arrive
   * before it returns — typing "=A1*2" briskly started five separate edits and
   * kept the last character. This flag is set synchronously on the first key,
   * so every later character is text rather than a fresh edit, whether or not
   * the cell editor has taken focus yet.
   */
  const startingRef = useRef(false);

  useEffect(() => {
    if (editing) putDraft((d) => (d == null ? editing.draft ?? '' : d));
    else {
      putDraft(null);
      startingRef.current = false;
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [Boolean(editing), editing?.row, editing?.col]);

  useEffect(() => {
    if (editing && editorRef.current) {
      const el = editorRef.current;
      el.focus();
      el.setSelectionRange(el.value.length, el.value.length);
    }
  }, [editing?.row, editing?.col]);

  /** Hand the finished text to the engine, then move as a spreadsheet does. */
  /**
   * The draft as it stands, not as the last render saw it.
   *
   * Enter is handled by a callback that closed over `draft`, and a keystroke
   * 60 ms before it can arrive between the state update and the render that
   * carries it: "hello" typed briskly into the formula bar committed "hell".
   * The ref is written in the same breath as the state, so what is committed
   * is what was typed however far behind the rendering is.
   */
  const draftRef = useRef(null);
  const putDraft = useCallback((value) => {
    draftRef.current = typeof value === 'function' ? value(draftRef.current) : value;
    setDraft(draftRef.current);
  }, []);

  const commitDraft = useCallback(
    (move = 'down') => {
      const text = draftRef.current ?? '';
      putDraft(null);
      startingRef.current = false;
      return dispatch({ op: 'updateDraft', text }, { op: 'commitEdit', move });
    },
    [dispatch, putDraft]
  );

  const onKeyDown = useCallback(
    async (e) => {
      if (!model) return;
      if (editing) {
        if (e.key === 'Escape') {
          e.preventDefault();
          putDraft(null);
          await dispatch({ op: 'cancelEdit' });
        } else if (e.key === 'Enter') {
          e.preventDefault();
          await commitDraft(e.shiftKey ? 'up' : 'down');
        } else if (e.key === 'Tab') {
          e.preventDefault();
          await commitDraft(e.shiftKey ? 'left' : 'right');
        }
        return;
      }

      const arrows = { ArrowUp: 'up', ArrowDown: 'down', ArrowLeft: 'left', ArrowRight: 'right' };
      if (arrows[e.key]) {
        e.preventDefault();
        navigate({ op: 'move', direction: arrows[e.key], extend: e.shiftKey, jump: e.ctrlKey || e.metaKey });
        return;
      }
      if (e.key === 'Tab') {
        e.preventDefault();
        navigate({ op: 'tab', back: e.shiftKey });
        return;
      }
      if (e.key === 'Enter') {
        e.preventDefault();
        navigate({ op: 'enter', back: e.shiftKey });
        return;
      }
      if (e.key === 'F2') {
        e.preventDefault();
        await dispatch({ op: 'beginEdit' });
        return;
      }
      // Excel's own: Ctrl+D/R fill, Ctrl+G go to, Ctrl+` show formulas.
      if ((e.ctrlKey || e.metaKey) && !e.altKey) {
        const key = e.key.toLowerCase();
        if (key === 'd' || key === 'r') { e.preventDefault(); await act('fill', key === 'd' ? 'down' : 'right'); return; }
        if (key === 'g') { e.preventDefault(); setDialog('goto'); return; }
        if (key === 'a') { e.preventDefault(); await dispatch({ op: 'selectAll' }); return; }
        if (key === '`') { e.preventDefault(); await act('toggleFormulas'); return; }
      }
      if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        await dispatch({ op: 'clear' });
        return;
      }
      // Typing a printable character starts an edit that replaces the cell.
      if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        e.preventDefault();
        if (startingRef.current) {
          // The edit is already on its way; this character is text, and it is
          // kept here because the editor may not hold focus yet.
          putDraft((d) => (d ?? '') + e.key);
          return;
        }
        startingRef.current = true;
        putDraft(e.key);
        await dispatch({ op: 'beginEdit', replace: true, initial: e.key });
      }
    },
    [model, editing, dispatch]
  );

  /**
   * A cell's formatting, as the engine reports it.
   *
   * The engine hands over `style.font` (`bold`, `italic`, `underline`,
   * `strike`, `sizePt`, `family`, `colour`), `style.fill` (`{ colour }`) and
   * `style.border` (an edge map). This painter used to read flat keys the
   * engine never sends — `s.bold`, `s.fill` as a string, `s.size` — so every
   * format applied correctly in the file and painted nothing in the window,
   * which reads as "formatting does not work".
   */
  // Text wider than its cell spills over empty neighbours, as in Excel; the
  // boxes are worked out once per frame and applied on top of the cell style.
  const spills = useMemo(() => spillBoxes(model), [model]);
  const spillStyle = (cell) => {
    const base = cellStyle(cell);
    const sp = spills.get(cell.ref);
    return sp ? { ...base, left: sp.x, width: sp.width, zIndex: 1 } : base;
  };

  const cellStyle = (cell) => {
    const s = cell.style || {};
    const font = s.font || {};
    const edge = (e) => (e && e.style && e.style !== 'none' ? `${e.widthPx || 1}px ${e.style === 'dashed' || e.style === 'dotted' || e.style === 'double' ? e.style : 'solid'} ${e.colour || '#000'}` : undefined);
    return {
      left: cell.x,
      top: cell.y,
      width: cell.width,
      height: cell.height,
      textAlign: cell.align || (typeof cell.text === 'string' && cell.isFormula ? 'right' : undefined),
      fontWeight: font.bold ? 700 : undefined,
      fontStyle: font.italic ? 'italic' : undefined,
      textDecoration: [font.underline ? 'underline' : '', font.strike ? 'line-through' : ''].filter(Boolean).join(' ') || undefined,
      color: font.colour || cell.colour || undefined,
      background: s.fill?.colour || undefined,
      fontSize: font.sizePt ? `${font.sizePt}pt` : undefined,
      fontFamily: font.family || undefined,
      borderTop: edge(s.border?.top),
      borderBottom: edge(s.border?.bottom),
      borderLeft: edge(s.border?.left),
      borderRight: edge(s.border?.right),
      whiteSpace: cell.wrap ? 'normal' : undefined,
      justifyContent: cell.align === 'center' ? 'center' : cell.align === 'right' ? 'flex-end' : undefined,
      alignItems: cell.valign === 'center' || cell.valign === 'middle' ? 'center' : cell.valign === 'top' ? 'flex-start' : undefined,
    };
  };

  /**
   * Which cell is under a point in the cells layer.
   *
   * The engine does not emit empty cells — a sheet is mostly empty and
   * drawing a million nothings is not a plan — so a click on blank grid lands
   * on the layer itself, not on a cell. The frame carries every visible
   * column's x and every visible row's y, which is enough to say which cell
   * that was. Without this, clicking an empty cell did nothing at all.
   */
  const cellAt = (event) => {
    const layer = event.currentTarget;
    const rect = layer.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    const col = (model?.columns || []).find((c) => x >= c.x && x < c.x + c.width);
    const row = (model?.rows || []).find((r) => y >= r.y && y < r.y + r.height);
    return col && row ? { row: row.index, col: col.index } : null;
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

  /**
   * The ribbon's verbs that are not one engine operation: the View toggles,
   * zoom, windows, the freeze shortcuts, the Fill and function helpers.
   * Named so a check can press the button and read what changed.
   */
  const act = async (name, arg, opts = {}) => {
    const at = sel?.active || { row: 0, col: 0 };
    const range = sel?.range || { top: at.row, left: at.col, bottom: at.row, right: at.col };
    switch (name) {
      case 'toggleGridlines': patchView((v) => ({ gridlines: v.gridlines === false })); return;
      case 'toggleHeadings': patchView((v) => ({ headings: v.headings === false })); return;
      case 'toggleFormulaBar': patchView((v) => ({ formulaBar: v.formulaBar === false })); return;
      case 'toggleFormulas': patchView((v) => ({ formulas: !v.formulas })); return;
      case 'page': patchView((v) => ({ page: { ...v.page, ...arg } })); return;
      case 'view': return;
      case 'zoom': {
        // The window's zoom is additive on a factor; reset first so a chosen
        // level is that level and not that level times the last one.
        await shell.win.zoom({ reset: true });
        if (arg && arg !== 1) await shell.win.zoom({ delta: arg - 1 });
        return;
      }
      case 'newWindow':
        if (!doc?.path) return toast('Save the workbook first, so a second window can open the same file.', { ms: 5000 });
        shell.win.create({ app: 'sheets', file: doc.path });
        return;
      case 'freeze': {
        const spec = arg === 'row' ? { rows: 1, cols: 0 } : arg === 'col' ? { rows: 0, cols: 1 } : arg === 'here' ? { rows: at.row, cols: at.col } : { rows: 0, cols: 0 };
        if (arg === 'here' && !spec.rows && !spec.cols) return toast('Select a cell below and right of what should stay in view, then freeze.', { ms: 5000 });
        await dispatch({ op: 'freeze', ...spec });
        return;
      }
      case 'mergeCentre':
        await dispatch({ op: 'merge' }, { op: 'setFormat', delta: { align: 'center' } });
        return;
      case 'fill': {
        // Excel's Ctrl+D/Ctrl+R: one cell fills from its neighbour above or
        // left; a range fills from its own first row or column.
        const single = range.top === range.bottom && range.left === range.right;
        if (single) {
          const source = arg === 'right' ? { row: at.row, col: at.col - 1 } : { row: at.row - 1, col: at.col };
          if (source.row < 0 || source.col < 0) return toast('Nothing above or left to fill from.', { ms: 4000 });
          await dispatch({ op: 'select', row: source.row, col: source.col }, { op: 'fill', target: range }, { op: 'select', row: at.row, col: at.col });
        } else {
          const source = arg === 'right'
            ? { top: range.top, bottom: range.bottom, left: range.left, right: range.left }
            : { top: range.top, bottom: range.top, left: range.left, right: range.right };
          await dispatch(
            { op: 'select', row: source.top, col: source.left },
            { op: 'select', row: source.bottom, col: source.right, extend: true },
            { op: 'fill', target: range },
            { op: 'select', row: range.top, col: range.left },
            { op: 'select', row: range.bottom, col: range.right, extend: true },
          );
        }
        return;
      }
      case 'insertFunction': {
        // Start an edit with `=NAME(` typed, the arguments the person's to
        // type; a bare name (Use in Formula) goes in as it is.
        const text = opts.bare ? `=${arg}` : `=${arg}(`;
        putDraft(text);
        await dispatch({ op: 'beginEdit', replace: true, initial: text });
        setTimeout(() => editorRef.current?.focus(), 0);
        return;
      }
      case 'recalculate':
        await dispatch({ op: 'select', row: at.row, col: at.col });
        toast('Recalculated.', { tone: 'good', ms: 2000 });
        return;
      case 'refreshAll':
        await dispatch({ op: 'select', row: at.row, col: at.col });
        try { await shell.doc.apply({ id: doc.id, ops: [{ op: 'refreshPivot' }] }).then((next) => { setDoc(next); setModel(next.model); }); } catch { /* no pivot to refresh */ }
        toast('Refreshed.', { tone: 'good', ms: 2000 });
        return;
      case 'autoFit': {
        // Each selected column takes the widest text on screen, plus padding.
        const ops = [];
        for (let col = range.left; col <= range.right; col++) {
          const widest = Math.max(0, ...(model?.cells || []).filter((c) => c.col === col && c.text).map((c) => String(c.text).length));
          ops.push({ op: 'colWidth', col, width: Math.max(40, Math.min(600, Math.round(widest * 7.2 + 14))) });
        }
        if (ops.length) await dispatch(...ops);
        return;
      }
      case 'textBox':
        await dispatch({ op: 'insertShape', geometry: 'rect', text: 'Text' });
        return;
      case 'goto': {
        const at2 = parseRef(arg);
        if (at2) await dispatch({ op: 'select', row: at2.row, col: at2.col });
        else await dispatch({ op: 'gotoName', name: arg });
        return;
      }
      case 'symbol':
        if (editing) putDraft((d) => (d ?? '') + arg);
        else await dispatch({ op: 'beginEdit', replace: true, initial: (model?.formulaBar ?? '') + arg });
        return;
      case 'help': shell.shell.openExternal({ url: SITE.help }); return;
      case 'feedback': shell.shell.openExternal({ url: SITE.contact }); return;
      case 'about': shell.win.create({ app: 'home', query: { about: 1 } }); return;
      default:
        toast(`${name} is not wired yet.`, { ms: 3000 });
    }
  };

  return (
    <AppFrame
      app={app}
      shell={shell}
      title={doc?.name || 'Worksheets'}
      subtitle={doc?.converted ? `from ${doc.converted.from.toUpperCase()}` : null}
      dirty={doc?.dirty}
      menu={appMenu}
      ribbon={
        <SheetsRibbon
          tab={tab}
          setTab={setTab}
          model={model}
          doc={doc}
          dispatch={dispatch}
          commands={commands}
          shell={shell}
          menu={menu}
          save={save}
          openFile={openFile}
          exportAs={exportAs}
          openDialog={setDialog}
          act={act}
          view={view}
          sel={sel}
        />
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
        <div className={`sh${view.gridlines === false ? ' no-grid' : ''}${view.headings === false ? ' no-heads' : ''}`} onKeyDown={onKeyDown} tabIndex={0} ref={(el) => { shRef.current = el; if (el && !editing && document.activeElement === document.body) el.focus(); }}>
          <style>{CSS}</style>

          <div className="sh-formula" hidden={view.formulaBar === false}>
            <div className="sh-namebox">{sel?.ref}</div>
            <Icon name="formula" size={14} style={{ color: 'var(--ink-3)' }} />
            <Input
              // The draft wins the moment there is one, not once the engine
              // has answered. Beginning an edit is a round trip; until it
              // returned, `editing` was false and this fell back to the
              // model's value, which put the cell's old contents back into
              // the box and threw away the letter just typed. "hello" typed
              // into the formula bar arrived as "ello" — the same race the
              // grid solved with startingRef below, in the one place that had
              // not been given it.
              value={editing || draft != null ? draft ?? '' : model.formulaBar ?? ''}
              onChange={(e) => {
                putDraft(e.target.value);
                if (!editing && !startingRef.current) {
                  startingRef.current = true;
                  // A refused edit — a protected sheet, say — must not leave
                  // the flag standing, or the bar takes no further keys.
                  Promise.resolve(dispatch({ op: 'beginEdit', replace: true, initial: e.target.value })).catch(() => {
                    startingRef.current = false;
                  });
                }
              }}
              onKeyDown={(e) => {
                // Enter, Tab and Escape finish with the bar and hand the keys
                // back to the grid, as Excel does. The bar kept focus after
                // Enter, so the next arrow key moved the caret in the bar
                // and the selection did not move — "the focus was not
                // moving with my keys".
                if (e.key === 'Enter' || e.key === 'Tab') {
                  e.preventDefault();
                  commitDraft(e.key === 'Tab' ? (e.shiftKey ? 'left' : 'right') : e.shiftKey ? 'up' : 'down');
                  shRef.current?.focus();
                }
                if (e.key === 'Escape') {
                  putDraft(null);
                  dispatch({ op: 'cancelEdit' });
                  shRef.current?.focus();
                }
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
                    onClick={(e) => dispatch({ op: 'selectColumn', col: c.index, extend: e.shiftKey, add: e.ctrlKey || e.metaKey })}
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
                    onClick={(e) => dispatch({ op: 'selectRow', row: r.index, extend: e.shiftKey, add: e.ctrlKey || e.metaKey })}
                    onContextMenu={(e) => menu.open(e, menuItems(commands, ['sheet.insertRow', 'sheet.deleteRow']))}
                  >
                    {r.label ?? r.index + 1}
                  </div>
                ))}
              </div>

              <div
                className="sh-cells"
                onMouseDown={(e) => {
                  // A press on a drawn cell is handled by the cell. Anywhere else
                  // is empty grid, and empty grid is still grid.
                  if (e.button !== 0 || e.target.closest('.sh-cell, .sh-editor, .sh-drawing')) return;
                  const at = cellAt(e);
                  if (at) dispatch({ op: 'select', row: at.row, col: at.col, extend: e.shiftKey, add: e.ctrlKey || e.metaKey });
                }}
                onDoubleClick={(e) => {
                  if (e.target.closest('.sh-cell, .sh-editor, .sh-drawing')) return;
                  const at = cellAt(e);
                  if (at) dispatch({ op: 'select', row: at.row, col: at.col }, { op: 'beginEdit' });
                }}
              >
                {model.cells.map((cell) => (
                  <div
                    key={cell.ref}
                    className={`sh-cell${cell.selected ? ' sel' : ''}${cell.active ? ' active' : ''}${cell.isError ? ' err' : ''}`}
                    data-ref={cell.ref}
                    style={spillStyle(cell)}
                    onMouseDown={(e) => dispatch({ op: 'select', row: cell.row, col: cell.col, extend: e.shiftKey, add: e.ctrlKey || e.metaKey })}
                    onDoubleClick={() => dispatch({ op: 'beginEdit' })}
                    onContextMenu={(e) => menu.open(e, menuItems(commands, ['edit.copy', 'edit.clear', '-', 'sheet.insertRow', 'sheet.insertCol', '-', 'sheet.merge']))}
                    title={cell.note || undefined}
                  >
                    {view.formulas && cell.formula ? cell.formula : cell.text}
                  </div>
                ))}

                {/*
                  What is drawn over the cells: the shapes, pictures and charts
                  the engine reads from the sheet's drawing part, each already
                  an SVG at its own size, anchored in the same coordinates as
                  the cells. The engine has produced these since charts were
                  built; the window never painted them, so a diagram drawn in
                  Excel opened as an empty grid with its captions.
                */}
                {(model.drawings || []).map((d) => (
                  <div
                    key={d.id}
                    className={`sh-drawing${d.svg ? '' : ' unsupported'}`}
                    style={{ left: d.x, top: d.y, width: d.width, height: d.height }}
                    title={d.unsupported ? `${d.name || d.kind}: ${d.unsupported}` : d.name || undefined}
                    {...(d.svg ? { dangerouslySetInnerHTML: { __html: d.svg } } : {})}
                  >
                    {d.svg ? null : <span>{d.name || d.kind}</span>}
                  </div>
                ))}

                {editing ? (

                  <input
                    ref={editorRef}
                    className="sh-editor"
                    value={draft ?? ''}
                    onChange={(e) => putDraft(e.target.value)}
                    onBlur={() => editing && commitDraft('none')}
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

      {dialog === 'print' && doc ? (
        <PrintDialog
          shell={shell}
          doc={doc}
          kind="sheet"
          sheets={model?.sheets || []}
          onClose={() => setDialog(null)}
          onSaveAs={(options) => exportAs('pdf', options)}
        />
      ) : null}

      {dialog === 'goto' ? (
        <GoToDialog names={model?.names || []} onClose={() => setDialog(null)} onGo={async (ref) => { setDialog(null); await act('goto', ref); }} />
      ) : null}
      {dialog === 'function' ? (
        <FunctionDialog catalogue={FUNCTIONS} onClose={() => setDialog(null)} onPick={async (name) => { setDialog(null); await act('insertFunction', name); }} />
      ) : null}
      {dialog === 'symbol' ? (
        <SymbolDialog onClose={() => setDialog(null)} onInsert={(ch) => act('symbol', ch)} />
      ) : null}
      {dialog === 'statistics' ? <StatisticsDialog model={model} onClose={() => setDialog(null)} /> : null}
      {dialog === 'shortcuts' ? <SheetShortcutsDialog onClose={() => setDialog(null)} /> : null}
      {dialog === 'rowHeight' || dialog === 'colWidth' ? (
        <SizeDialog
          kind={dialog === 'rowHeight' ? 'row' : 'col'}
          current={dialog === 'rowHeight' ? model?.cells?.find((c) => c.active)?.height : model?.cells?.find((c) => c.active)?.width}
          onClose={() => setDialog(null)}
          onApply={async (n) => {
            const r = sel?.range || { top: sel?.active?.row ?? 0, bottom: sel?.active?.row ?? 0, left: sel?.active?.col ?? 0, right: sel?.active?.col ?? 0 };
            const ops = [];
            if (dialog === 'rowHeight') for (let row = r.top; row <= r.bottom; row++) ops.push({ op: 'rowHeight', row, height: n });
            else for (let col = r.left; col <= r.right; col++) ops.push({ op: 'colWidth', col, width: n });
            await dispatch(...ops);
            setDialog(null);
          }}
        />
      ) : null}

      {dialog === 'conditional' ? (
        <ConditionalDialog
          selection={sel?.ref}
          onClose={() => setDialog(null)}
          onApply={async (spec) => {
            await dispatch({ op: 'conditional', spec });
            setDialog(null);
          }}
          onClear={async (all) => {
            await dispatch({ op: 'clearConditional', all });
            setDialog(null);
          }}
        />
      ) : null}

      {dialog === 'validation' ? (
        <ValidationDialog
          selection={sel?.ref}
          onClose={() => setDialog(null)}
          onApply={async (spec) => {
            await dispatch({ op: 'validation', spec });
            setDialog(null);
          }}
          onClear={async (all) => {
            await dispatch({ op: 'clearValidation', all });
            setDialog(null);
          }}
        />
      ) : null}

      {dialog === 'goalSeek' ? (
        <GoalSeekDialog
          active={sel?.ref?.split(':')[0]}
          onClose={() => setDialog(null)}
          onRun={async ({ set, to, by }) => {
            const next = await dispatch({ op: 'goalSeek', set, to, by });
            return next?.model?.goalSeek ?? { converged: Boolean(next), value: by, reached: to };
          }}
        />
      ) : null}

      {dialog === 'dataTable' ? (
        <DataTableDialog
          selection={sel?.ref}
          onClose={() => setDialog(null)}
          onRun={async (spec) => {
            await dispatch({ op: 'dataTable', ...spec });
            setDialog(null);
          }}
        />
      ) : null}

      {dialog === 'names' ? (
        <NameManager
          names={model?.names || []}
          selection={sel?.ref}
          onClose={() => setDialog(null)}
          onDefine={(name, ref) => dispatch({ op: 'defineName', name, ref })}
          onDelete={(name) => dispatch({ op: 'deleteName', name })}
          onGoto={async (name) => {
            await dispatch({ op: 'gotoName', name });
            setDialog(null);
          }}
        />
      ) : null}

      {dialog === 'find' ? (
        <FindDialog
          onClose={() => setDialog(null)}
          onFind={async (text) => {
            const next = await dispatch({ op: 'findNext', text });
            return next ? 'Found the next match.' : 'Nothing else matches.';
          }}
          onReplace={async (find, replace) => {
            const next = await dispatch({ op: 'replaceNext', find, replace });
            return next ? 'Replaced one.' : 'Nothing left to replace.';
          }}
          onReplaceAll={async (find, replace) => {
            const next = await dispatch({ op: 'replaceAll', find, replace });
            return next ? 'Replaced every match.' : 'Nothing matched.';
          }}
        />
      ) : null}

      {dialog === 'pivot' ? (
        <PivotDialog
          selection={sel?.ref}
          sheets={model?.sheets || []}
          onClose={() => setDialog(null)}
          onCreate={async (spec) => {
            await dispatch({ op: 'pivot', ...spec });
            setDialog(null);
          }}
        />
      ) : null}

      {dialog === 'freeze' ? <FreezeDialog model={model} sel={sel} dispatch={dispatch} onClose={() => setDialog(null)} /> : null}
    </AppFrame>
  );
}

/**
 * Freeze panes.
 *
 * Excel freezes at the selection and leaves you to work out where that is.
 * Naming the three things people actually want — the top row, the first column,
 * both — and offering "at the selection" underneath is the same feature with
 * the guessing taken out.
 */
function FreezeDialog({ model, sel, dispatch, onClose }) {
  const frozen = model?.frozen || { rows: 0, cols: 0 };
  const at = sel?.active || { row: 0, col: 0 };

  const choices = [
    { label: 'Freeze the top row', rows: 1, cols: 0 },
    { label: 'Freeze the first column', rows: 0, cols: 1 },
    { label: 'Freeze the top row and first column', rows: 1, cols: 1 },
    { label: `Freeze above and left of ${sel?.ref?.split(':')[0] || 'the selection'}`, rows: at.row, cols: at.col },
    { label: 'Unfreeze', rows: 0, cols: 0 },
  ];

  return (
    <Dialog
      title="Freeze panes"
      width={420}
      onClose={onClose}
      actions={<Button label="Close" onClick={onClose} />}
    >
      <p style={{ marginTop: 0, fontSize: 12.5 }}>
        {frozen.rows || frozen.cols
          ? `Currently frozen: ${frozen.rows} row${frozen.rows === 1 ? '' : 's'}, ${frozen.cols} column${frozen.cols === 1 ? '' : 's'}.`
          : 'Nothing is frozen.'}
      </p>
      <div className="ml-found">
        {choices.map((c) => (
          <button
            key={c.label}
            type="button"
            className="ml-found-item"
            onClick={async () => {
              await dispatch({ op: 'freeze', rows: c.rows, cols: c.cols });
              onClose();
            }}
          >
            <span className="ml-found-logo"><Icon name="freeze" size={15} /></span>
            <span className="grow"><div className="who">{c.label}</div></span>
            {frozen.rows === c.rows && frozen.cols === c.cols ? <Icon name="check" size={14} /> : null}
          </button>
        ))}
      </div>
    </Dialog>
  );
}

/**
 * Excel lets text wider than its cell spill into empty neighbours — to the
 * right for left-aligned text, to the left for right-aligned, both ways for
 * centred — and stops at the first cell with something in it. Numbers never
 * spill (Excel shows #### instead), nor does wrapped or merged text. Measured
 * with the cell's own font on a canvas, so what spills is what is too wide;
 * a neighbour off the edge of the frame counts as a wall, which errs towards
 * clipping rather than towards drawing over something unseen.
 *
 * @returns {Map<string, {x: number, width: number}>} by cell ref
 */
const measureCtx = typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null;
function spillBoxes(model) {
  const out = new Map();
  const cells = model?.cells || [];
  const columns = model?.columns || [];
  if (!measureCtx || !cells.length || !columns.length) return out;
  const byRow = new Map();
  for (const c of cells) {
    if (!byRow.has(c.row)) byRow.set(c.row, new Map());
    byRow.get(c.row).set(c.col, c);
  }
  const colAt = new Map(columns.map((c) => [c.index, c]));
  for (const c of cells) {
    if (!c.text || c.wrap || c.merged || c.numeric || c.isError) continue;
    const font = c.style?.font || {};
    measureCtx.font = `${font.bold ? 'bold ' : ''}${font.italic ? 'italic ' : ''}${font.sizePt ? Math.round(font.sizePt * (96 / 72)) : 12.5}px ${font.family || 'Calibri, "Segoe UI", sans-serif'}`;
    const needed = measureCtx.measureText(String(c.text)).width + 12;
    if (needed <= c.width) continue;
    const row = byRow.get(c.row);
    const free = (col) => colAt.has(col) && !row.get(col)?.text;
    let left = c.x;
    let right = c.x + c.width;
    const growRight = (until) => { let col = c.col + 1; while (right < until && free(col)) { const k = colAt.get(col); right = k.x + k.width; col += 1; } };
    const growLeft = (until) => { let col = c.col - 1; while (left > until && free(col)) { left = colAt.get(col).x; col -= 1; } };
    if (c.align === 'right') growLeft(right - needed);
    else if (c.align === 'center') {
      const extra = (needed - c.width) / 2;
      growLeft(c.x - extra);
      growRight(c.x + c.width + extra);
    } else growRight(left + needed);
    if (left !== c.x || right !== c.x + c.width) out.set(c.ref, { x: left, width: right - left });
  }
  return out;
}

function formatNumber(n) {
  if (!Number.isFinite(n)) return '—';
  return Math.abs(n) >= 1000 ? n.toLocaleString(undefined, { maximumFractionDigits: 2 }) : String(Math.round(n * 1000) / 1000);
}

const CSS = `
/*
  min-width: 0 on both is load-bearing. A flex item's minimum width is its
  widest descendant unless told otherwise, and the widest descendant here is
  the canvas — which is sized to cover the grid's own width plus a margin.
  Without the zero, the column stretched to the canvas, the canvas grew to
  cover the column, the size observer asked for a frame, and the frame grew
  the canvas again: an idle window asked the main process for a wider frame
  eight times a second for as long as it lived, 1,088 px wider each time.
*/
.sh { flex: 1; display: flex; flex-direction: column; min-height: 0; min-width: 0; outline: none; }
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

.sh-grid { flex: 1; overflow: auto; position: relative; min-width: 0; min-height: 0; background: var(--surface); }
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
.sh-drawing { position: absolute; overflow: visible; z-index: 2; }
.sh-drawing > svg { display: block; overflow: visible; }
.sh-drawing.unsupported { display: grid; place-items: center; border: 1px dashed var(--line); color: var(--ink-3); font-size: 11px; background: rgba(255, 255, 255, 0.6); }

/* Excel's View toggles: gridlines off leaves the cells' own borders; headings off drops the rails. */
.sh.no-grid .sh-cell { border-right-color: transparent; border-bottom-color: transparent; }
.sh.no-heads .sh-corner, .sh.no-heads .sh-colheads, .sh.no-heads .sh-rowheads { display: none; }
.sh.no-heads .sh-canvas { grid-template-columns: 0 auto !important; grid-template-rows: 0 auto !important; }
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
