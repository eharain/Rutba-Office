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
import { Button, Icon, Spacer, Chip, Empty, Spinner, Dialog, ZoomSlider, Panel, useToast, useMenu, useCommands, menuItems, Input } from '@rutba/office-ui';
import { AppFrame, useAppMenu, pickOpen, pickSave, confirmDiscard, useFileDrop, openInApp , useDirtyGuard } from '../shell.js';
import { PrintDialog, defaultPrintOptions } from '../print.js';
import { usePasswordGate, openProtected, LockedAction, useProtection } from '../protect.js';
import SheetsRibbon, { FUNCTIONS, MARGIN_PRESETS } from './sheets/ribbon.js';
import { SITE } from '@rutba/office-formats/registry';
import { SymbolDialog } from './word/dialogs.js';
import { useSheetsReview } from './sheets/review.js';
import {
  GoToDialog, FunctionDialog, StatisticsDialog, SheetShortcutsDialog, SizeDialog, SortDialog, LinkDialog, NoteDialog, HeaderFooterDialog, SheetNameDialog, SheetDeleteDialog, SparklineDialog, parseRef,
  OutlineAxisDialog, SubtotalDialog, AdvancedFilterDialog, EvaluateDialog,
  ProtectDialog, PasswordDialog, EditRangesDialog, CustomViewsDialog, ConsolidateDialog, ForecastDialog,
} from './sheets/dialogs.js';
import { WorkbookGallery, SHEET_DESIGN_CSS } from './sheets/design.js';
import { CustomColoursDialog, CustomFontsDialog, DESIGN_CSS } from './slides/design.js';
import { SlicerPanel, InsertSlicersDialog, ObjectHandles, RotateHandle, SelectionPane, angleAt, followPointer, OBJECTS_CSS } from './sheets/objects.js';
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

/** A 0-based row/column as Excel writes a cell reference, e.g. (1, 4) → "E2". */
const refText = (row, col) => colLabel(col) + (row + 1);

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

/** What the tip layer says over a cell: its note, or where its link goes. */
const tipFor = (cell) => {
  if (cell.thread) {
    const t = cell.thread;
    const more = t.replies ? ` (${t.replies} ${t.replies === 1 ? 'reply' : 'replies'})` : '';
    return `${t.author ? t.author + ': ' : ''}${t.text}${more}${t.done ? ' — resolved' : ''} — click the corner to open the thread`;
  }
  if (cell.note) return `${cell.note.author ? cell.note.author + ': ' : ''}${cell.note.text}`;
  if (cell.link) return `${cell.link.tooltip ? cell.link.tooltip + ' — ' : ''}${cell.link.href || cell.link.location} (Ctrl+click to open)`;
  return undefined;
};

export default function Sheets({ app, shell, boot }) {
  const toast = useToast();
  const [doc, setDoc] = useState(null);
  const [model, setModel] = useState(null);
  const [busy, setBusy] = useState(true);
  const [tab, setTab] = useState('home');
  const [error, setError] = useState(null);
  // A protected file whose Password dialog was cancelled: its name.
  const [lockedOut, setLockedOut] = useState(null);
  const gate = usePasswordGate();
  // Which of the ribbon's dialogs is open, by name. One piece of state rather
  // than seven booleans, because only one of them can be open at a time.
  const [dialog, setDialog] = useState(null);
  // A column or row being resized by its heading's edge: which one, and the
  // size the pointer has dragged it to — drawn as a guide line until the
  // button is released and the engine is told.
  const [resizing, setResizing] = useState(null);
  // A drag from the fill handle: the range it has reached, drawn as a dashed
  // box until the button is released and the engine fills it.
  const [filling, setFilling] = useState(null);
  /**
   * How the grid is shown. None of it is in the workbook: gridlines,
   * headings and the formula bar are Excel's View toggles, "show formulas"
   * is Ctrl+`, and the page setup is what the PDF export will lay out to.
   */
  const [view, setView] = useState({
    zoom: 1,
    gridlines: true, headings: true, formulaBar: true, formulas: false,
    // The file's own page setup, read once the workbook is open: what the
    // Page Layout tab says now, and what a press there changes.
    page: null,
  });
  const patchView = useCallback((patch) => setView((v) => ({ ...v, ...(typeof patch === 'function' ? patch(v) : patch) })), []);
  /** Tracing arrows on the grid: each a range and the cell it points at, or from. */
  const [arrows, setArrows] = useState([]);
  /** The sheet tab a rename or delete dialog is about. */
  const [sheetTarget, setSheetTarget] = useState(null);
  /** Formulas → Watch Window: which row in the pane is selected, for Delete Watch. */
  const [watchSel, setWatchSel] = useState(null);
  /** Data → Group / Ungroup on a selection that is neither whole rows nor whole columns: which of the two is being asked. */
  const [outlineAsk, setOutlineAsk] = useState(null);
  /** The list the Subtotal and Advanced Filter dialogs act on: its range and its columns by header. */
  const [listInfo, setListInfo] = useState(null);
  /**
   * Review → comments: the card open on a cell — its thread, a reply box,
   * or the box a new thread starts in — and which comment is being edited.
   */
  const [card, setCard] = useState(null);
  /** The Comments pane's filter: every thread, the open ones, or the resolved. */
  const [commentFilter, setCommentFilter] = useState('all');
  /** Page Break Preview: a break being dragged — its axis, where it was, where the pointer is. */
  const [breakDrag, setBreakDrag] = useState(null);
  /** View → Split: a split bar being dragged, and where the pointer has it. */
  const [splitDrag, setSplitDrag] = useState(null);
  /** Formulas → Evaluate Formula: the cell it opened on and the dialog's first state. */
  const [evaluating, setEvaluating] = useState(null);
  /**
   * Review → Protect: a password being asked for — to unprotect the sheet
   * or the workbook, or to unlock an edit range before typing into it —
   * with the ops to send again once it is right, and the last refusal.
   */
  const [passwordAsk, setPasswordAsk] = useState(null);
  /** An edit refused because its cell is in a password range: the window asks for the password (set below). */
  const rangeLockRef = useRef(null);
  /** Page Layout → Background: the picture's part and the address its bytes are drawn from. */
  const [backdrop, setBackdrop] = useState(null);
  /** Data → Consolidate and Forecast Sheet: what each dialog opens on. */
  const [analysis, setAnalysis] = useState(null);
  /**
   * The drawings picked on the sheet — a slicer, a chart, a picture, a
   * shape — by their frame ids, and one being dragged or resized: its box
   * as the pointer has it, drawn until the button comes up.
   */
  const [picked, setPicked] = useState([]);
  const [live, setLive] = useState(null);
  /** Each slicer's Multi-Select, as Excel keeps it: per panel, for the session. */
  const [slicerMulti, setSlicerMulti] = useState({});
  /** Insert → PivotTable / PivotChart and Insert → Slicer: what the dialog opens on. */
  const [pivotAsk, setPivotAsk] = useState(null);
  const [slicerAsk, setSlicerAsk] = useState(null);
  /** Page Layout → Themes, Colours, Fonts, Effects: the gallery open (its kind and where), and a Customise dialog. */
  const [gallery, setGallery] = useState(null);
  const [customise, setCustomise] = useState(null);
  /** Page Layout → Selection Pane, open or not. */
  const [selPane, setSelPane] = useState(false);
  const gridRef = useRef(null);
  /** The element that takes the keys: the grid's own container. */
  const shRef = useRef(null);

  const editorRef = useRef(null);
  const menu = useMenu();
  const openFileRef = useRef(null);
  /** The ribbon's verbs, for a command defined before them (Ctrl+Alt+M). */
  const actRef = useRef(null);
  // File → Info: the Protect Workbook card — Encrypt with Password, and the
  // sheet and structure protection Review already offers, as Excel lists them.
  const protection = useProtection({
    app: 'sheets',
    shell,
    doc,
    setDoc,
    toast,
    items: [
      { id: 'sheet', icon: 'table', label: model?.protection?.sheet ? 'Unprotect Current Sheet' : 'Protect Current Sheet', detail: 'Control what types of changes people can make to the current sheet.', run: () => actRef.current?.('protectSheet') },
      { id: 'structure', icon: 'grid', label: model?.workbookProtection?.structure ? 'Unprotect Workbook Structure' : 'Protect Workbook Structure', detail: 'Prevent unwanted changes to the structure of the workbook, such as adding sheets.', run: () => actRef.current?.('protectWorkbook') },
    ],
    notes: [
      ...(model?.workbookProtection?.structure ? ['The workbook’s structure is locked to prevent unwanted changes, such as moving, deleting or adding sheets.'] : []),
      ...(model?.protection?.sheet ? [`${model?.activeSheet || 'This sheet'} is protected: locked cells take no edits.`] : []),
    ],
  });
  const appMenu = useAppMenu({ shell, appKey: 'sheets', onNew: () => shell.win.create({ app: 'sheets' }), onOpen: () => openFileRef.current?.(), extra: doc ? [protection.menuItem] : [] });

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
        // A cell in an edit range with a password: Excel's Unlock Range
        // asks for it, then the edit goes ahead — not a refusal to read.
        const range = /^The range "(.+)" is protected by a password/.exec(err.message || '');
        if (range && rangeLockRef.current) {
          rangeLockRef.current(range[1], ops);
          return null;
        }
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

  /** The page setup as the file keeps it, for the Page Layout tab. */
  const refreshPage = useCallback((id) => {
    if (!id) return;
    shell.doc.pageSetup({ id }).then((page) => patchView({ page })).catch(() => {});
  }, [shell, patchView]);

  const load = useCallback(
    async (fn) => {
      setBusy(true);
      setError(null);
      try {
        const opened = await fn();
        setDoc(opened);
        setModel(opened.model);
        refreshPage(opened.id);
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
        if (err.locked) setLockedOut(err.locked);
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
    // A password-protected file (or the encrypted copy of one) asks for its
    // password in this window before anything opens.
    if (recover) load(() => openProtected((password) => shell.doc.recover({ file: recover, password }), gate).then((r) => { toast('Recovered unsaved work. Save it to keep it.', { ms: 6000 }); return r; }));
    else if (boot.file) load(() => openProtected((password) => shell.doc.open({ path: boot.file, kind: 'sheet', password }), gate));
    else load(() => shell.doc.new({ kind: 'sheets', template: template && template !== 'blank' ? template : 'sheet' }));
    // A window an earlier build zoomed stays zoomed across restarts — the
    // level is kept per origin — and the grid carries the zoom now, so the
    // window itself goes back to 100%.
    shell.win.zoom({ reset: true }).catch(() => {});
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

  /**
   * One frame on its way at a time.
   *
   * A scroll fires every animation frame and each one asked the engine for a
   * frame, so a flick of the wheel queued a dozen requests behind the first
   * and the grid caught up seconds later, drawing rows a person had already
   * left. A scroll that arrives while a frame is on its way is noted, and
   * when the frame lands one more is asked for from wherever the grid is
   * now — every reply is a frame somebody can still see.
   */
  const frameBusy = useRef(false);
  const frameAgain = useRef(false);
  const syncViewport = useCallback(async () => {
    const el = gridRef.current;
    if (!el || !doc) return;
    if (frameBusy.current) {
      frameAgain.current = true;
      return;
    }
    frameBusy.current = true;
    try {
      do {
        frameAgain.current = false;
        const next = await shell.doc.viewport({
          id: doc.id,
          width: el.clientWidth,
          height: el.clientHeight,
          x: el.scrollLeft,
          y: el.scrollTop,
        });
        setModel(next);
      } while (frameAgain.current && gridRef.current);
    } catch {
      // A frame that fails is a frame not drawn; the next scroll asks again.
    } finally {
      frameBusy.current = false;
    }
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

  /**
   * View → Split: the wheel over the top pane scrolls it up and down, over
   * the left pane across (Shift+wheel, or a sideways wheel), over the
   * corner both — three rows or columns a notch, as Excel's do. Anything
   * else the wheel does is the grid's own scrolling.
   */
  const hasModel = Boolean(model);
  useEffect(() => {
    const el = gridRef.current;
    if (!el) return undefined;
    const onWheel = (e) => {
      const pin = e.target?.closest?.('.split');
      if (!pin) return;
      const corner = pin.classList.contains('sh-pin-corner');
      const top = corner || pin.classList.contains('sh-pin-rows') || pin.classList.contains('sh-pin-rowheads');
      const left = corner || pin.classList.contains('sh-pin-cols') || pin.classList.contains('sh-pin-colheads');
      const dy = e.shiftKey ? 0 : e.deltaY;
      const dx = e.shiftKey ? (e.deltaY || e.deltaX) : e.deltaX;
      const step = (d) => Math.sign(d) * Math.max(1, Math.round((Math.abs(d) / 100) * 3));
      const rows = top && dy ? step(dy) : 0;
      const cols = left && dx ? step(dx) : 0;
      if (!rows && !cols) return;
      e.preventDefault();
      navigate({ op: 'scrollSplit', rows, cols });
    };
    el.addEventListener('wheel', onWheel, { passive: false });
    return () => el.removeEventListener('wheel', onWheel);
  }, [navigate, hasModel]);

  /**
   * Page Layout → Background: the frame names the picture's part; its bytes
   * are fetched once, as an address the grid tiles behind the cells.
   */
  const backgroundPart = model?.background || null;
  useEffect(() => {
    if (!doc || !backgroundPart) { setBackdrop(null); return undefined; }
    let live = true;
    shell.doc.asset({ id: doc.id, ref: backgroundPart })
      .then((a) => { if (live) setBackdrop(a?.url ? { part: backgroundPart, url: a.url } : null); })
      .catch(() => { if (live) setBackdrop(null); });
    return () => { live = false; };
  }, [doc?.id, backgroundPart, shell]);

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
      'insert.link': { label: 'Link…', icon: 'link', key: 'Mod+K', run: () => setDialog('link') },
      'insert.note': { label: 'Note…', icon: 'reply', key: 'Shift+F2', run: () => setDialog('note') },
      'insert.comment': { label: 'New comment', icon: 'reply', key: 'Mod+Alt+M', run: () => actRef.current?.('newComment') },
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
    // The editor has gone (a commit, a cancel): its draft goes with it —
    // unless the next edit is already on its way. Keys typed briskly after
    // Enter arrive between the render that drops the editor and this effect,
    // and wiping the draft here threw them away: "=A1*2" typed straight
    // after "42⏎" reached the cell as "2".
    else if (!startingRef.current) putDraft(null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [Boolean(editing), editing?.row, editing?.col]);

  useEffect(() => {
    if (editing && editorRef.current) {
      const el = editorRef.current;
      el.focus();
      el.setSelectionRange(el.value.length, el.value.length);
      // From here the editor holds the keys; the next printable key that
      // reaches the grid starts a fresh edit.
      startingRef.current = false;
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

  rangeLockRef.current = (title, ops) => {
    startingRef.current = false;
    putDraft(null);
    setPasswordAsk({
      kind: 'range',
      title: 'Unlock Range',
      range: title,
      message: `The cell you are trying to change is in the range "${title}", which is protected by a password. Enter the password to change this cell:`,
      retry: ops.filter((o) => o.op !== 'updateDraft' && o.op !== 'commitEdit'),
    });
  };

  const commitDraft = useCallback(
    (move = 'down') => {
      const text = draftRef.current ?? '';
      putDraft(null);
      startingRef.current = false;
      return dispatch({ op: 'updateDraft', text }, { op: 'commitEdit', move });
    },
    [dispatch, putDraft]
  );

  /**
   * Insert → Pictures: a file, put at the active cell at its own
   * proportions, its longer side no more than 320 px. The size comes from
   * the picture's own pixels, read here where an image can be decoded.
   */
  const insertPicture = useCallback(async () => {
    const [file] = await shell.dialog.open({
      title: 'Insert picture',
      filters: [{ name: 'Pictures', extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp'] }],
    });
    if (!file) return;
    const { bytes, stat } = await shell.fs.read({ path: file });
    const ext = String(stat?.ext || file.split('.').pop()).replace('.', '').toLowerCase();
    const contentType = { png: 'image/png', gif: 'image/gif', webp: 'image/webp', bmp: 'image/bmp' }[ext] || 'image/jpeg';
    const size = await new Promise((resolve) => {
      const url = URL.createObjectURL(new Blob([bytes], { type: contentType }));
      const img = new Image();
      img.onload = () => { URL.revokeObjectURL(url); resolve({ w: img.naturalWidth || 320, h: img.naturalHeight || 240 }); };
      img.onerror = () => { URL.revokeObjectURL(url); resolve({ w: 320, h: 240 }); };
      img.src = url;
    });
    const scale = Math.min(1, 320 / Math.max(1, size.w, size.h));
    await dispatch({
      op: 'insertPicture', name: stat?.name || file.split(/[\\/]/).pop(), contentType, data: bytes,
      widthPx: Math.max(16, Math.round(size.w * scale)), heightPx: Math.max(16, Math.round(size.h * scale)),
    });
  }, [shell, dispatch]);

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
        } else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey && e.target !== editorRef.current) {
          // The editor is drawn but has not taken focus yet — the frame
          // between its render and the effect that focuses it — and a brisk
          // typist's next character arrived at the grid instead. It is text:
          // "4321" typed at 60 ms a key landed as "421" without this.
          e.preventDefault();
          putDraft((d) => (d ?? '') + e.key);
        }
        return;
      }

      // A picked drawing takes Delete and Escape before the cells do.
      if (picked.length) {
        if (e.key === 'Delete' || e.key === 'Backspace') {
          e.preventDefault();
          const ids = picked;
          setPicked([]);
          await dispatch({ op: 'deleteDrawings', ids });
          return;
        }
        if (e.key === 'Escape') {
          e.preventDefault();
          setPicked([]);
          return;
        }
        const nudge = { ArrowUp: [0, -1], ArrowDown: [0, 1], ArrowLeft: [-1, 0], ArrowRight: [1, 0] }[e.key];
        if (nudge && !e.altKey) {
          e.preventDefault();
          const step = e.ctrlKey || e.metaKey ? 10 : 1;
          const boxes = (model.drawings || []).filter((d) => picked.includes(d.id))
            .map((d) => ({ id: d.id, x: Math.max(0, d.x + nudge[0] * step), y: Math.max(0, d.y + nudge[1] * step), width: d.width, height: d.height }));
          if (boxes.length) await dispatch({ op: 'drawingBoxes', boxes });
          return;
        }
      }
      // Excel's own: Shift+Alt+Right groups, Shift+Alt+Left ungroups.
      if (e.altKey && e.shiftKey && !e.ctrlKey && (e.key === 'ArrowRight' || e.key === 'ArrowLeft')) {
        e.preventDefault();
        await act(e.key === 'ArrowRight' ? 'group' : 'ungroup');
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
      // Excel's own: F9 calculates every sheet, Shift+F9 this one,
      // Ctrl+Alt+F9 every formula whatever changed.
      if (e.key === 'F9') {
        e.preventDefault();
        await act('calculate', e.shiftKey ? 'sheet' : e.ctrlKey && e.altKey ? 'full' : 'workbook');
        return;
      }
      // Excel's own: Ctrl+D/R fill, Ctrl+G go to, Ctrl+` show formulas.
      if ((e.ctrlKey || e.metaKey) && !e.altKey) {
        const key = e.key.toLowerCase();
        if (key === 'd' || key === 'r') { e.preventDefault(); await act('fill', key === 'd' ? 'down' : 'right'); return; }
        if (key === 'g') { e.preventDefault(); setDialog('goto'); return; }
        if (key === 'a') { e.preventDefault(); await dispatch({ op: 'selectAll' }); return; }
        if (key === 'e') { e.preventDefault(); await act('flashFill'); return; }
        if (key === '`') { e.preventDefault(); await act('toggleFormulas'); return; }
      }
      if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        // Delete on an empty sparkline cell takes the sparkline off — a
        // sparkline cell is normally empty, so this is the natural way to
        // clear one without hunting for the ribbon or a right-click.
        if (e.key === 'Delete') {
          const active = model.selection?.active;
          const activeCell = active && (model.cells || []).find((c) => c.row === active.row && c.col === active.col);
          const spark = active && (model.sparklines || []).find((s) => s.at.row === active.row && s.at.col === active.col);
          if (spark && !activeCell?.text) {
            await dispatch({ op: 'removeSparklines', at: active.ref });
            return;
          }
        }
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
    [model, editing, dispatch, picked]
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

  /**
   * Excel's textRotation as CSS on the cell's words: 1..90 turns them
   * anticlockwise (CSS turns clockwise, hence the sign), 91..180 clockwise by
   * the value less 90, and 255 stacks the letters upright.
   */
  const rotationStyle = (r) => (r === 255
    ? { writingMode: 'vertical-lr', textOrientation: 'upright', lineHeight: 1.1 }
    : { display: 'inline-block', transform: `rotate(${r > 90 ? r - 90 : -r}deg)`, transformOrigin: 'center', whiteSpace: 'nowrap' });

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
      // Excel's indent, from the side the alignment reads from: one unit is
      // about three characters, nine pixels at the grid's size.
      paddingLeft: cell.indent && cell.align !== 'right' ? `${5 + cell.indent * 9}px` : undefined,
      paddingRight: cell.indent && cell.align === 'right' ? `${5 + cell.indent * 9}px` : undefined,
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
  /**
   * Drag the edge of a column or row heading to resize it — Excel's way, and
   * the first thing anyone tries. The guide follows the pointer; the engine
   * is told once, on release. A double-click on a column's edge fits it to
   * its widest text; on a row's, puts the row back to the default height.
   */
  const startResize = useCallback(
    (e, kind, index, size, start) => {
      if (e.button !== 0) return;
      e.preventDefault();
      e.stopPropagation();
      const origin = kind === 'col' ? e.clientX : e.clientY;
      const min = kind === 'col' ? 12 : 10;
      const max = kind === 'col' ? 1200 : 500;
      let next = size;
      setResizing({ kind, index, size, start });
      const move = (ev) => {
        // Pointer travel is in screen pixels; the grid's sizes are its own, zoomed.
        next = Math.max(min, Math.min(max, Math.round(size + ((kind === 'col' ? ev.clientX : ev.clientY) - origin) / (view.zoom || 1))));
        setResizing({ kind, index, size: next, start });
      };
      const stop = () => {
        window.removeEventListener('mousemove', move);
        window.removeEventListener('mouseup', stop);
        setResizing(null);
        if (next !== size) dispatch(kind === 'col' ? { op: 'colWidth', col: index, width: next } : { op: 'rowHeight', row: index, height: next });
      };
      window.addEventListener('mousemove', move);
      window.addEventListener('mouseup', stop);
    },
    [dispatch, view.zoom]
  );

  const fitColumn = useCallback(
    (col) => {
      const widest = Math.max(0, ...(model?.cells || []).filter((c) => c.col === col && c.text).map((c) => String(c.text).length));
      dispatch({ op: 'colWidth', col, width: Math.max(40, Math.min(600, Math.round(widest * 7.2 + 14))) });
    },
    [dispatch, model]
  );

  /**
   * The fill handle: the square at the bottom-right of the selection. Drag
   * it down or across and the cells it passes over are filled from the
   * selection on release — a value copied, a formula's references moved, a
   * series continued, as the engine's fill does for Ctrl+D. The dashed box
   * follows the pointer; the direction is whichever the pointer has gone
   * further in, as in Excel.
   */
  const startFill = useCallback(
    (e, source) => {
      if (e.button !== 0 || !source) return;
      e.preventDefault();
      e.stopPropagation();
      const layer = gridRef.current?.querySelector('.sh-cells');
      if (!layer) return;
      const rect = layer.getBoundingClientRect();
      const z = view.zoom || 1;
      const columns = model?.columns || [];
      const rows = model?.rows || [];
      const corner = { x: (e.clientX - rect.left) / z, y: (e.clientY - rect.top) / z };
      let target = source;
      const move = (ev) => {
        const px = (ev.clientX - rect.left) / z;
        const py = (ev.clientY - rect.top) / z;
        const col = columns.find((c) => px >= c.x && px < c.x + c.width) || (px >= (columns[columns.length - 1]?.x ?? 0) ? columns[columns.length - 1] : columns[0]);
        const row = rows.find((r) => py >= r.y && py < r.y + r.height) || (py >= (rows[rows.length - 1]?.y ?? 0) ? rows[rows.length - 1] : rows[0]);
        if (!col || !row) return;
        const down = Math.abs(py - corner.y) >= Math.abs(px - corner.x);
        target = down
          ? { top: source.top, left: source.left, right: source.right, bottom: Math.max(source.bottom, row.index) }
          : { top: source.top, bottom: source.bottom, left: source.left, right: Math.max(source.right, col.index) };
        setFilling({ target });
      };
      const stop = () => {
        window.removeEventListener('mousemove', move);
        window.removeEventListener('mouseup', stop);
        setFilling(null);
        if (target.bottom > source.bottom || target.right > source.right) {
          dispatch(
            { op: 'select', row: source.top, col: source.left },
            { op: 'select', row: source.bottom, col: source.right, extend: true },
            { op: 'fill', target },
            { op: 'select', row: target.top, col: target.left },
            { op: 'select', row: target.bottom, col: target.right, extend: true },
          );
        }
      };
      window.addEventListener('mousemove', move);
      window.addEventListener('mouseup', stop);
    },
    [dispatch, model, view.zoom]
  );

  /** A range's box in the cells layer, from the visible columns and rows; null when it is scrolled away. */
  /**
   * Frozen panes, pinned.
   *
   * The frame carries the frozen rows and columns in every viewport; the
   * window drew them at their own place, so a frozen heading row scrolled
   * away with the rest. They are drawn in layers that stick — the rows
   * under the column headings, the columns beside the row headings, the
   * corner at both — `position: sticky` in normal flow inside the cells
   * layer, the way the headings already stick in the grid's own tracks.
   */
  const frozen = model?.frozen || { rows: 0, cols: 0 };
  const frozenH = frozen.rows ? (model?.rows || []).reduce((m, r) => (r.index < frozen.rows ? Math.max(m, r.y + r.height) : m), 0) : 0;
  const frozenW = frozen.cols ? (model?.columns || []).reduce((m, c) => (c.index < frozen.cols ? Math.max(m, c.x + c.width) : m), 0) : 0;
  /**
   * The outline gutter, Excel's: a lane per level beside the row headings
   * (above the column headings), the brackets and − / + boxes of the
   * groups in them, and the level buttons in the corner. It widens the
   * heading track rather than overlaying the cells, so the grid keeps its
   * coordinates and a frozen pane its pins. Hidden with the headings.
   */
  const OUTLINE_LANE = 16;
  const outline = view.headings ? model?.outline : null;
  const rowLevels = outline?.rows?.levels || 0;
  const colLevels = outline?.cols?.levels || 0;
  const gutW = rowLevels ? (rowLevels + 1) * OUTLINE_LANE + 4 : 0;
  const gutH = colLevels ? (colLevels + 1) * OUTLINE_LANE + 4 : 0;
  const lane = (level) => 2.5 + (level - 1) * OUTLINE_LANE + OUTLINE_LANE / 2;
  const headTop = view.headings ? (model?.headerHeight ?? 0) + gutH : 0;
  const headLeft = view.headings ? (model?.headerWidth ?? 0) + gutW : 0;
  const pane = (at) => (at.row < frozen.rows ? (at.col < frozen.cols ? 'corner' : 'rows') : at.col < frozen.cols ? 'cols' : 'main');
  /**
   * View → Split: the top and left panes are the frozen panes' pinned
   * layers, each showing the rows or columns it has scrolled to — the top
   * pane scrolls across with the main one and up and down on its own, the
   * left pane the other way round. The main layer draws the viewport's own
   * cells; a row can be in both, drawn in each.
   */
  const split = model?.split && !(frozen.rows || frozen.cols) ? model.split : null;
  const splitRows = new Set(split?.rows || []);
  const splitCols = new Set(split?.cols || []);
  const splitH = split?.rows?.length ? split.height : 0;
  const splitW = split?.cols?.length ? split.width : 0;
  const vpr = model?.viewport;
  const inMain = (at) => (split
    ? Boolean(vpr) && at.row >= vpr.firstRow && at.row <= vpr.lastRow && at.col >= vpr.firstCol && at.col <= vpr.lastCol
    : pane(at) === 'main');

  const boxOf = (range) => {
    const left = (model?.columns || []).find((c) => c.index === range.left);
    const right = (model?.columns || []).find((c) => c.index === range.right);
    const top = (model?.rows || []).find((r) => r.index === range.top);
    const bottom = (model?.rows || []).find((r) => r.index === range.bottom);
    if (!left || !right || !top || !bottom) return null;
    return { x: left.x, y: top.y, right: right.x + right.width, bottom: bottom.y + bottom.height };
  };

  const cellAt = (event) => {
    // A press inside a pinned layer is measured from the layer, which has
    // slid with the scroll; the columns layer starts under the frozen rows.
    const pin = event.target?.closest?.('.sh-pin');
    const layer = pin || event.currentTarget;
    // Screen pixels to the grid's own: the grid is zoomed, the pointer is not.
    const rect = layer.getBoundingClientRect();
    const z = view.zoom || 1;
    // Each pinned layer says how far its content is from where it sits:
    // the frozen columns start under the frozen rows, a split pane shows
    // the rows and columns it has scrolled to.
    const x = (event.clientX - rect.left) / z + Number(pin?.dataset.ox || 0);
    const y = (event.clientY - rect.top) / z + Number(pin?.dataset.oy || 0);
    const col = (model?.columns || []).find((c) => x >= c.x && x < c.x + c.width);
    const row = (model?.rows || []).find((r) => y >= r.y && y < r.y + r.height);
    return col && row ? { row: row.index, col: col.index } : null;
  };

  /**
   * What floats over the cells: charts, pictures and shapes as the engine
   * drew them, slicers as panels of buttons, and the handles of whatever is
   * picked. A drag moves or resizes the picked drawing on screen and tells
   * the engine once, when the button comes up.
   */
  const drawingsNode = () => {
    const z = view.zoom || 1;
    // A drag's boxes, by id, as the pointer has them; a turn as it goes.
    const boxFor = (d) => live?.boxes?.[d.id] || { x: d.x, y: d.y, width: d.width, height: d.height };
    const shown = (model.drawings || []).filter((d) => !d.hidden);
    const pickedShown = shown.filter((d) => picked.includes(d.id));
    const finish = (op) => Promise.resolve(dispatch(op)).finally(() => setLive(null));
    // Resize one, from a handle.
    const resize = (e, d, handle) => followPointer(e, {
      box: boxFor(d), mode: 'resize', handle, zoom: z,
      onMove: (box) => setLive({ boxes: { [d.id]: box } }),
      onDone: (box) => (box ? finish({ op: 'drawingBox', id: d.id, ...box }) : setLive(null)),
    });
    // Move every picked drawing by the same step, as one undo step.
    const move = (e, lead, group, { pickAlone = false } = {}) => {
      const start = boxFor(lead);
      followPointer(e, {
        box: start, mode: 'move', zoom: z,
        onMove: (box) => {
          const dx = box.x - start.x;
          const dy = box.y - start.y;
          setLive({ boxes: Object.fromEntries(group.map((g) => [g.id, { x: Math.max(0, g.x + dx), y: Math.max(0, g.y + dy), width: g.width, height: g.height }])) });
        },
        onDone: (box) => {
          // A click, not a drag, on one of several picked picks it alone.
          if (!box) { setLive(null); if (pickAlone) setPicked([lead.id]); return; }
          const dx = box.x - start.x;
          const dy = box.y - start.y;
          finish(group.length === 1
            ? { op: 'drawingBox', id: group[0].id, x: Math.max(0, group[0].x + dx), y: Math.max(0, group[0].y + dy), width: group[0].width, height: group[0].height }
            : { op: 'drawingBoxes', boxes: group.map((g) => ({ id: g.id, x: Math.max(0, g.x + dx), y: Math.max(0, g.y + dy), width: g.width, height: g.height })) });
        },
      });
    };
    // Turn one about its centre, from the rotation handle.
    const turn = (e, d) => {
      const el = e.currentTarget.closest('.sh-cells') || e.currentTarget.parentElement;
      const rect = el.getBoundingClientRect();
      const b = boxFor(d);
      const cx = rect.left + (b.x + b.width / 2) * z;
      const cy = rect.top + (b.y + b.height / 2) * z;
      let last = null;
      const onMove = (ev) => {
        last = angleAt(cx, cy, ev.clientX, ev.clientY, ev.shiftKey);
        setLive({ turn: { id: d.id, deg: last } });
      };
      const onUp = () => {
        window.removeEventListener('mousemove', onMove, true);
        window.removeEventListener('mouseup', onUp, true);
        if (last === null) { setLive(null); return; }
        finish({ op: 'rotateDrawings', ids: [d.id], to: last });
      };
      window.addEventListener('mousemove', onMove, true);
      window.addEventListener('mouseup', onUp, true);
    };
    // A press picks (Ctrl or Shift adds or takes away) and starts a move.
    const press = (e, d) => {
      if (e.button !== 0) return;
      e.preventDefault();
      e.stopPropagation();
      shRef.current?.focus({ preventScroll: true });
      const adding = e.ctrlKey || e.metaKey || e.shiftKey;
      let next;
      if (adding) next = picked.includes(d.id) ? picked.filter((id) => id !== d.id) : [...picked, d.id];
      else next = picked.includes(d.id) ? picked : [d.id];
      setPicked(next);
      if (!next.includes(d.id)) return;
      const group = shown.filter((x) => next.includes(x.id)).map((x) => ({ id: x.id, ...boxFor(x) }));
      move(e, d, group, { pickAlone: !adding && next.length > 1 });
    };
    const one = pickedShown.length === 1 ? pickedShown[0] : null;
    const turnOf = (d) => (live?.turn?.id === d.id ? live.turn.deg : d.rot || 0);
    return (
      <>
        {shown.map((d) => {
          const box = boxFor(d);
          if (d.kind === 'slicer') {
            const name = d.slicer?.name || d.name;
            return (
              <SlicerPanel
                key={d.id}
                d={d}
                box={box}
                picked={picked.includes(d.id)}
                multi={Boolean(slicerMulti[name])}
                onPick={(e) => {
                  shRef.current?.focus({ preventScroll: true });
                  setPicked((p) => ((e?.ctrlKey || e?.shiftKey) ? (p.includes(d.id) ? p : [...p, d.id]) : [d.id]));
                }}
                onStartMove={(e) => move(e, d, [{ id: d.id, ...box }])}
                onToggleMulti={() => setSlicerMulti((m) => ({ ...m, [name]: !m[name] }))}
                onChoose={(values) => dispatch({ op: 'slicerSelect', name, values })}
                onClear={() => dispatch({ op: 'slicerSelect', name, values: null })}
              />
            );
          }
          // A shape or a picture draws its own turn; a group is turned whole,
          // and so is anything while its rotation handle is being dragged.
          const dragging = live?.turn?.id === d.id;
          const extra = dragging && d.kind !== 'group' ? turnOf(d) - (d.rot || 0) : 0;
          const transform = [
            d.kind === 'group' && (turnOf(d) || 0) ? `rotate(${turnOf(d)}deg)` : '',
            extra ? `rotate(${extra}deg)` : '',
            d.kind === 'group' && (d.flipH || d.flipV) ? `scale(${d.flipH ? -1 : 1}, ${d.flipV ? -1 : 1})` : '',
          ].filter(Boolean).join(' ');
          return (
            <div
              key={d.id}
              className={`sh-drawing${d.svg || d.members ? '' : ' unsupported'}${picked.includes(d.id) ? ' picked' : ''}`}
              data-id={d.id}
              data-kind={d.kind}
              onContextMenu={(e) => menu.open(e, [{ label: 'Edit Alt Text…', icon: 'textbox', run: () => review.openAltText({ sheet: model.activeSheet, anchor: Number(String(d.id).replace('drawing-', '')) || 0 }) }])}
              data-name={d.name || ''}
              style={{ left: box.x, top: box.y, width: box.width, height: box.height, ...(transform ? { transform } : {}) }}
              title={d.unsupported ? `${d.name || d.kind}: ${d.unsupported}` : d.pivot ? `${d.name || 'PivotChart'} — a PivotChart of ${d.pivot}` : d.name || undefined}
              onMouseDown={(e) => press(e, d)}
              {...(d.svg ? { dangerouslySetInnerHTML: { __html: d.svg } } : {})}
            >
              {d.svg ? null : d.members ? d.members.map((m) => (
                <div key={m.key} className="sh-member" style={{ left: m.x * (box.width / d.width), top: m.y * (box.height / d.height), width: m.width * (box.width / d.width), height: m.height * (box.height / d.height) }} dangerouslySetInnerHTML={{ __html: m.svg || '' }} />
              )) : <span>{d.name || d.kind}</span>}
            </div>
          );
        })}
        {pickedShown.length > 1 ? pickedShown.map((d) => {
          const b = boxFor(d);
          return <div key={`ring-${d.id}`} className="sh-obj-ring several" style={{ left: b.x, top: b.y, width: b.width, height: b.height }} />;
        }) : null}
        {one ? (() => {
          // The ring and handles turn with the drawing, as Excel's do.
          const b = boxFor(one);
          const deg = turnOf(one);
          return (
            <div className="sh-obj-turn" style={deg ? { position: 'absolute', left: 0, top: 0, width: 0, height: 0, transformOrigin: `${b.x + b.width / 2}px ${b.y + b.height / 2}px`, transform: `rotate(${deg}deg)` } : { position: 'absolute', left: 0, top: 0, width: 0, height: 0 }}>
              <ObjectHandles box={b} onHandle={(e, handle) => resize(e, one, handle)} />
              {['shape', 'image', 'group'].includes(one.kind) ? <RotateHandle box={b} onStart={(e) => turn(e, one)} /> : null}
            </div>
          );
        })() : null}
      </>
    );
  };

  /** The tracing arrows, drawn over the cells: a dot at the source, a head at the target, a box round a range. */
  const arrowsNode = () => {
    if (!arrows.length) return null;
    const colAt = (c) => (model?.columns || model?.cols || []).find((x) => x.index === c);
    const rowAt = (r) => (model?.rows || []).find((x) => x.index === r);
    const centre = ({ row, col }) => {
      const c = colAt(col);
      const r = rowAt(row);
      return c && r ? { x: c.x + c.width / 2, y: r.y + r.height / 2 } : null;
    };
    return (
      <svg className="sh-arrows" style={{ position: 'absolute', left: 0, top: 0, width: '100%', height: '100%', overflow: 'visible', pointerEvents: 'none', zIndex: 4 }}>
        <defs>
          <marker id="sh-arrowhead" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto" markerUnits="userSpaceOnUse">
            <path d="M0,0 L8,4 L0,8 z" fill="#1f5f8b" />
          </marker>
        </defs>
        {arrows.map((a, i) => {
          const rangeCell = { row: a.top, col: a.left };
          const from = centre(a.kind === 'precedents' ? rangeCell : a.at);
          const to = centre(a.kind === 'precedents' ? a.at : rangeCell);
          if (!from || !to) return null;
          const c1 = colAt(a.left);
          const c2 = colAt(a.right);
          const r1 = rowAt(a.top);
          const r2 = rowAt(a.bottom);
          const box = (a.top !== a.bottom || a.left !== a.right) && c1 && c2 && r1 && r2
            ? <rect x={c1.x} y={r1.y} width={c2.x + c2.width - c1.x} height={r2.y + r2.height - r1.y} fill="rgba(31, 95, 139, 0.08)" stroke="#1f5f8b" strokeWidth="1" />
            : null;
          return (
            <g key={i} className="sh-arrow">
              {box}
              <circle cx={from.x} cy={from.y} r="3" fill="#1f5f8b" />
              <line x1={from.x} y1={from.y} x2={to.x} y2={to.y} stroke="#1f5f8b" strokeWidth="1.5" markerEnd="url(#sh-arrowhead)" />
            </g>
          );
        })}
      </svg>
    );
  };

  /** Sparklines in view, keyed by "row:col" so a cell paints its own without scanning the list. */
  const sparkAt = useMemo(() => {
    const m = new Map();
    for (const s of model?.sparklines || []) m.set(s.at.row + ':' + s.at.col, s);
    return m;
  }, [model?.sparklines]);

  /**
   * A sparkline drawn in its own cell: a line through the values (blanks as
   * gaps, the line breaking over them rather than crossing), or a bar per
   * value — both scaled to the cell's box with a pixel of padding, in the
   * group's own colour. The cell's own text still paints over it, as Excel
   * does; a sparkline cell is normally empty.
   */
  const sparkSvg = (spark, width, height) => {
    const w = Math.max(1, width - 2);
    const h = Math.max(1, height - 2);
    const n = spark.values.length;
    if (!n || !spark.values.some((v) => typeof v === 'number')) return null;
    const nums = spark.values.filter((v) => typeof v === 'number');
    const min = Math.min(0, ...nums);
    const max = Math.max(0, ...nums);
    const span = max - min || 1;
    const colour = '#' + (spark.colour || '376092');
    const x = (i) => 1 + (n > 1 ? (i * w) / (n - 1) : w / 2);
    const yOf = (v) => 1 + h - ((v - min) / span) * h;
    if (spark.type === 'column') {
      const gap = Math.min(2, w / (n * 4));
      const bw = Math.max(1, w / n - gap);
      return (
        <svg className="sh-spark" data-ref={refText(spark.at.row, spark.at.col)} width={width} height={height}>
          {spark.values.map((v, i) => (typeof v === 'number' ? (
            <rect key={i} x={1 + (i * w) / n + gap / 2} y={Math.min(yOf(v), yOf(0))} width={bw} height={Math.max(0.5, Math.abs(yOf(v) - yOf(0)))} fill={colour} />
          ) : null))}
        </svg>
      );
    }
    // One polyline per unbroken run of numbers — a blank cell in the data
    // range is a gap in the line, not a dip to zero.
    const segments = [];
    let run = [];
    spark.values.forEach((v, i) => {
      if (typeof v === 'number') run.push(x(i) + ',' + yOf(v));
      else if (run.length) { segments.push(run); run = []; }
    });
    if (run.length) segments.push(run);
    return (
      <svg className="sh-spark" data-ref={refText(spark.at.row, spark.at.col)} width={width} height={height}>
        {segments.map((seg, i) => <polyline key={i} points={seg.join(' ')} fill="none" stroke={colour} strokeWidth="1" />)}
      </svg>
    );
  };

  /** A cell, drawn where it sits — `dy` above it when its layer starts lower down. */
  const cellNode = (cell, dy = 0, dx = 0) => {
    const spark = sparkAt.get(cell.row + ':' + cell.col);
    return (
      <div
        key={cell.ref}
        className={`sh-cell${cell.selected ? ' sel' : ''}${cell.active ? ' active' : ''}${cell.isError ? ' err' : ''}${cell.link ? ' link' : ''}${cell.note ? ' noted' : ''}${cell.thread ? (cell.thread.done ? ' threaded resolved' : ' threaded') : ''}`}
        data-ref={cell.ref}
        style={dy || dx ? (() => { const s = spillStyle(cell); return { ...s, top: cell.y - dy, left: s.left - dx }; })() : spillStyle(cell)}
        onMouseDown={(e) => {
          // Ctrl+click on a link follows it, as in Word; a plain click selects, as in Excel.
          if (cell.link && (e.ctrlKey || e.metaKey)) { e.preventDefault(); act('follow', cell.link); return; }
          dispatch({ op: 'select', row: cell.row, col: cell.col, extend: e.shiftKey, add: e.ctrlKey || e.metaKey });
        }}
        onDoubleClick={() => dispatch({ op: 'beginEdit' })}
        onContextMenu={(e) => menu.open(e, spark
          ? menuItems(commands, ['edit.copy', 'edit.clear', '-']).concat([
            { label: 'Remove sparkline', icon: 'close', run: () => dispatch({ op: 'removeSparklines', at: cell.ref }) },
          ])
          : menuItems(commands, ['edit.copy', 'edit.clear', '-', 'insert.link', 'insert.comment', 'insert.note', '-', 'sheet.insertRow', 'sheet.insertCol', '-', 'sheet.merge']))}
        data-tip={tipFor(cell)}
      >
        {spark ? sparkSvg(spark, cell.width, cell.height) : null}
        {cell.thread ? (
          <span
            className="sh-thread-mark"
            onMouseDown={(e) => {
              // The corner opens the thread, as Excel's does; the cell is selected first.
              e.stopPropagation();
              dispatch({ op: 'select', row: cell.row, col: cell.col }).then(() => setCard({ row: cell.row, col: cell.col, mode: 'thread' }));
            }}
          />
        ) : null}
        {cell.rotation
          ? <span className="sh-rot" style={rotationStyle(cell.rotation)}>{view.formulas && cell.formula ? cell.formula : cell.text}</span>
          : (view.formulas && cell.formula ? cell.formula : cell.text)}
      </div>
    );
  };

  /** The cell editor, over the active cell, in whichever layer holds it. */
  const editorNode = (dy = 0) => {
    const active = model.cells.find((c) => c.active);
    return (
      <input
        ref={editorRef}
        className="sh-editor"
        value={draft ?? ''}
        onChange={(e) => putDraft(e.target.value)}
        onBlur={() => editing && commitDraft('none')}
        style={{
          left: editing.x ?? active?.x ?? 0,
          top: (editing.y ?? active?.y ?? 0) - dy,
          width: Math.max(80, active?.width ?? 80),
          height: active?.height ?? 20,
        }}
      />
    );
  };

  /**
   * One axis of the outline gutter: the bracket of each open group in its
   * level's lane — a hook at the far end, a line to the summary — and a
   * box at the summary, − to fold the group and + to open it. Drawn in the
   * sheet's own coordinates, like the headings beside it.
   */
  const gutterNode = (axis, pinned = false) => {
    const o = outline?.[axis === 'row' ? 'rows' : 'cols'];
    if (!o) return null;
    const row = axis === 'row';
    const BOX = 11;
    const lines = [];
    const boxes = [];
    // The copy in a frozen pane's pinned headings draws only what reaches
    // into the pane; the rest is under it in the sliding copy.
    const reach = row ? frozenH : frozenW;
    for (const g of o.groups) {
      const c = lane(g.level);
      const mid = g.at === null ? null : g.at + g.size / 2;
      if (pinned && !(g.from < reach || (mid !== null && mid - BOX / 2 < reach))) continue;
      const key = `${g.level}:${g.start}`;
      if (!g.collapsed && g.to > g.from) {
        // Along the group, from its first row to the box (or its last row).
        const a = g.from + 3;
        const b = g.to - 3;
        const near = mid === null ? null : o.below ? mid - BOX / 2 : mid + BOX / 2;
        const from = o.below ? a : near ?? a;
        const to = o.below ? near ?? b : b;
        const hookAt = o.below ? a : b;
        const pts = row
          ? `${c + 5},${hookAt} ${c},${hookAt} ${c},${o.below ? to : from}`
          : `${hookAt},${c + 5} ${hookAt},${c} ${o.below ? to : from},${c}`;
        lines.push(<polyline key={key} points={pts} fill="none" />);
        if (mid === null) {
          const end = o.below ? b : a;
          lines.push(<polyline key={key + 'e'} points={row ? `${c},${end} ${c + 5},${end}` : `${end},${c} ${end},${c + 5}`} fill="none" />);
        }
      }
      if (mid !== null) {
        const what = row ? 'rows' : 'columns';
        const span = row ? `${g.start + 1} to ${g.end + 1}` : `${colLabel(g.start)} to ${colLabel(g.end)}`;
        boxes.push(
          <div
            key={key + 'b'}
            className={`sh-ol-box${g.collapsed ? ' folded' : ''}`}
            data-axis={axis}
            data-level={g.level}
            data-start={g.start}
            data-tip={g.collapsed ? `Show Detail — ${what} ${span}` : `Hide Detail — ${what} ${span}`}
            style={row ? { left: c - BOX / 2, top: Math.round(mid - BOX / 2) } : { top: c - BOX / 2, left: Math.round(mid - BOX / 2) }}
            onMouseDown={(e) => e.stopPropagation()}
            onClick={() => dispatch({ op: 'outlineToggle', axis, level: g.level, start: g.start })}
          >
            <svg width="9" height="9" viewBox="0 0 9 9"><path d={g.collapsed ? 'M1.5 4.5h6M4.5 1.5v6' : 'M1.5 4.5h6'} /></svg>
          </div>
        );
      }
    }
    return (
      <div className={`sh-gutter ${row ? 'rows' : 'cols'}`} style={row ? { width: gutW, height: model.total.height } : { height: gutH, width: model.total.width }}>
        <svg className="sh-gutter-lines" width={row ? gutW : model.total.width} height={row ? model.total.height : gutH}>{lines}</svg>
        {boxes}
      </div>
    );
  };

  /** The level buttons, 1 2 3 …, in the corner: level n shows everything above it. */
  const levelButtons = () => {
    const out = [];
    const button = (axis, n, style) => (
      <button
        key={axis + n}
        type="button"
        className="sh-ol-level"
        data-axis={axis}
        data-level={n}
        data-tip={`${n} — show ${axis === 'row' ? 'rows' : 'columns'} down to outline level ${n}`}
        style={style}
        onMouseDown={(e) => e.stopPropagation()}
        onClick={() => dispatch({ op: 'outlineLevel', axis, level: n })}
      >
        {n}
      </button>
    );
    for (let n = 1; n <= rowLevels + 1 && rowLevels; n++) {
      out.push(button('row', n, { left: Math.round(lane(n) - 7), top: gutH + Math.round((model.headerHeight - 14) / 2) }));
    }
    for (let n = 1; n <= colLevels + 1 && colLevels; n++) {
      out.push(button('col', n, { top: Math.round(lane(n) - 7), left: gutW + Math.round((model.headerWidth - 14) / 2) }));
    }
    return out;
  };

  const colHead = (c, dx = 0) => (
    <div
      key={c.index}
      className={`sh-head${c.index >= (sel?.left ?? -1) && c.index <= (sel?.right ?? -2) ? ' active' : ''}`}
      // Both coordinates, always: an absolute heading with no top took its
      // static place, which the pinned wrapper in flow had moved down.
      style={{ left: c.x - dx, top: gutH, width: c.width, height: model.headerHeight }}
      onClick={(e) => dispatch({ op: 'selectColumn', col: c.index, extend: e.shiftKey, add: e.ctrlKey || e.metaKey })}
      onContextMenu={(e) => menu.open(e, menuItems(commands, ['sheet.insertCol', 'sheet.deleteCol', '-', 'sheet.sortAsc', 'sheet.sortDesc']))}
    >
      {c.label || colLabel(c.index)}
      <div
        className="sh-grip col"
        title="Drag to resize the column; double-click to fit its text"
        onMouseDown={(e) => startResize(e, 'col', c.index, c.width, c.x)}
        onClick={(e) => e.stopPropagation()}
        onDoubleClick={(e) => {
          e.stopPropagation();
          fitColumn(c.index);
        }}
      />
    </div>
  );

  const rowHead = (r, dy = 0) => (
    <div
      key={r.index}
      className={`sh-head${r.index >= (sel?.top ?? -1) && r.index <= (sel?.bottom ?? -2) ? ' active' : ''}`}
      style={{ top: r.y - dy, left: gutW, height: r.height, width: model.headerWidth }}
      onClick={(e) => dispatch({ op: 'selectRow', row: r.index, extend: e.shiftKey, add: e.ctrlKey || e.metaKey })}
      onContextMenu={(e) => menu.open(e, menuItems(commands, ['sheet.insertRow', 'sheet.deleteRow']))}
    >
      {r.label ?? r.index + 1}
      <div
        className="sh-grip row"
        title="Drag to resize the row; double-click for the default height"
        onMouseDown={(e) => startResize(e, 'row', r.index, r.height, r.y)}
        onClick={(e) => e.stopPropagation()}
        onDoubleClick={(e) => {
          e.stopPropagation();
          dispatch({ op: 'rowHeight', row: r.index, height: 20 });
        }}
      />
    </div>
  );

  /**
   * Formulas → Error Checking: the list pane, or nothing while the check is
   * off. `model.errors` is only ever an array (the check is on, even at zero
   * errors) or null (off) — see sheetModel in documents.js — so that alone
   * decides whether the pane shows.
   */
  const errorsPane = () => {
    const list = model?.errors;
    if (!Array.isArray(list)) return null;
    return (
      <div className="sh-errors">
        <div className="sh-errors-head">
          <strong>{list.length ? `${list.length} error${list.length === 1 ? '' : 's'}` : 'No errors found'}</strong>
          <Spacer />
          <Button className="sh-error-prev" icon="chevronUp" title="Previous error" disabled={!list.length} onClick={() => act('stepError', 'prev')} />
          <Button className="sh-error-next" icon="chevronDown" title="Next error" disabled={!list.length} onClick={() => act('stepError', 'next')} />
          <Button icon="close" title="Close — turns Error Checking off" onClick={() => act('errorCheck', false)} />
        </div>
        {list.length ? (
          <div className="sh-errors-list">
            {list.map((err) => (
              <div key={`${err.sheet}!${err.ref}`} className="sh-error" data-ref={err.ref} onClick={() => act('gotoError', err)}>
                <div className="sh-error-line">
                  {err.sheet !== model.activeSheet ? <span className="sh-error-sheet">{err.sheet}</span> : null}
                  <span className="sh-error-ref">{err.ref}</span>
                  <span className="sh-error-value">{err.value}</span>
                </div>
                {err.formula ? <div className="sh-error-formula">{err.formula}</div> : null}
                <div className="sh-error-reason">{err.reason}</div>
              </div>
            ))}
          </div>
        ) : null}
      </div>
    );
  };

  /**
   * Formulas → Watch Window: a pane docked along the bottom, one row per
   * watched cell — Sheet, Cell, Value, Formula — read fresh off the model
   * every render, exactly like the errors pane, so a value follows the
   * workbook as it recalculates. `model.watches` is the same two-state
   * shape as `model.errors` (an array once the pane is open, null while
   * it is closed), but the refs behind it live on the session for as long
   * as the workbook is open, closed pane or not — see sheetModel and the
   * watchAdd/watchRemove ops in documents.js. Which row is selected, for
   * Delete Watch, is this component's own state; it means nothing to the
   * engine.
   */
  const watchPane = () => {
    const list = model?.watches;
    if (!Array.isArray(list)) return null;
    const keyOf = (w) => `${w.sheet}!${w.ref}`;
    const selKey = watchSel ? keyOf(watchSel) : null;
    return (
      <div className="sh-watch">
        <div className="sh-watch-head">
          <strong>{list.length ? `${list.length} cell${list.length === 1 ? '' : 's'} watched` : 'No cells watched'}</strong>
          <Spacer />
          <Button icon="plus" label="Add Watch" title="Add Watch — every cell in the current selection" onClick={() => act('addWatch')} />
          <Button icon="trash" label="Delete Watch" title="Delete Watch — the selected row" disabled={!watchSel} onClick={() => act('deleteWatch')} />
          <Button icon="close" title="Close — turns the Watch Window off" onClick={() => act('watchOpen', false)} />
        </div>
        {list.length ? (
          <div className="sh-watch-list">
            <div className="sh-watch-row sh-watch-cols">
              <span>Sheet</span>
              <span>Cell</span>
              <span>Value</span>
              <span>Formula</span>
              <span />
            </div>
            {list.map((w) => (
              <div
                key={keyOf(w)}
                className={`sh-watch-row${selKey === keyOf(w) ? ' selected' : ''}`}
                data-sheet={w.sheet}
                data-ref={w.ref}
                onClick={() => setWatchSel(w)}
                onDoubleClick={() => act('gotoWatch', w)}
              >
                <span className="sh-watch-sheet">{w.sheet}</span>
                <span className="sh-watch-cell">{w.ref}</span>
                <span className="sh-watch-value">{w.value}</span>
                <span className="sh-watch-formula">{w.formula || ''}</span>
                <Button
                  className="sh-watch-x"
                  icon="close"
                  title="Remove this watch"
                  onClick={(e) => { e.stopPropagation(); if (selKey === keyOf(w)) setWatchSel(null); act('removeWatch', w); }}
                />
              </div>
            ))}
          </div>
        ) : null}
      </div>
    );
  };

  /**
   * View → Page Break Preview, drawn over the cells: what is not printed
   * greyed, the printed area and each page edged in blue, a break put by
   * hand solid and one the paper made dashed, and "Page 1", "Page 2"… across
   * each page. A break is a handle: dragged, it goes where it is dropped.
   */
  const breaksNode = (shift = null) => {
    const pb = model?.pageBreaks;
    if (!pb || !model) return null;
    const a = pb.area;
    const W = model.total.width;
    const H = model.total.height;
    const shade = (key, x, y, w, h) => (w > 0 && h > 0 ? <div key={key} className="sh-pb-shade" style={{ left: x, top: y, width: w, height: h }} /> : null);
    const startDrag = (e, axis, b) => {
      if (e.button !== 0) return;
      e.preventDefault();
      e.stopPropagation();
      const layer = gridRef.current?.querySelector('.sh-cells');
      if (!layer) return;
      const rect = layer.getBoundingClientRect();
      const z = view.zoom || 1;
      const at = (ev) => (axis === 'row' ? (ev.clientY - rect.top) / z : (ev.clientX - rect.left) / z);
      let pos = at(e);
      setBreakDrag({ axis, index: b.index, pos });
      const move = (ev) => { pos = at(ev); setBreakDrag({ axis, index: b.index, pos }); };
      const stop = () => {
        window.removeEventListener('mousemove', move);
        window.removeEventListener('mouseup', stop);
        setBreakDrag(null);
        // The row or column edge nearest the drop.
        const list = axis === 'row' ? (model.rows || []).map((r) => [r.index, r.y]) : (model.columns || []).map((c) => [c.index, c.x]);
        let best = null;
        for (const [index, edge] of list) if (!best || Math.abs(edge - pos) < Math.abs(best[1] - pos)) best = [index, edge];
        if (!best || best[0] === b.index) return;
        act('moveBreak', { axis, from: b.index, to: best[0], manual: b.manual });
      };
      window.addEventListener('mousemove', move);
      window.addEventListener('mouseup', stop);
    };
    return (
      // In a split pane the same marks are drawn where the pane has scrolled
      // to (`shift`), to be seen but not dragged; the handles are the main pane's.
      <div className={`sh-pb${shift ? ' mirror' : ''}`} data-pages={shift ? undefined : pb.count} style={shift ? { left: -shift.x, top: -shift.y } : undefined}>
        {shade('t', 0, 0, W, a.y)}
        {shade('b', 0, a.y + a.height, W, H - a.y - a.height)}
        {shade('l', 0, a.y, a.x, a.height)}
        {shade('r', a.x + a.width, a.y, W - a.x - a.width, a.height)}
        {pb.pages.map((p) => (
          <div key={'p' + p.n} className="sh-pb-page" data-page={p.n} style={{ left: p.x, top: p.y, width: p.width, height: p.height }}>
            <span style={{ fontSize: Math.max(18, Math.min(64, Math.round(Math.min(p.width / 4, p.height / 5)))) }}>Page {p.n}</span>
          </div>
        ))}
        <div className="sh-pb-area" style={{ left: a.x, top: a.y, width: a.width, height: a.height }} />
        {pb.rows.map((b) => (
          <div key={'r' + b.index} className={`sh-pb-break row${b.manual ? ' manual' : ''}`} data-index={b.index} data-tip={`${b.manual ? 'Page break put by hand' : 'Automatic page break'} above row ${b.index + 1} — drag to move it`}
            style={{ left: a.x, top: b.y - 4, width: a.width }} onMouseDown={shift ? undefined : (e) => startDrag(e, 'row', b)} />
        ))}
        {pb.cols.map((b) => (
          <div key={'c' + b.index} className={`sh-pb-break col${b.manual ? ' manual' : ''}`} data-index={b.index} data-tip={`${b.manual ? 'Page break put by hand' : 'Automatic page break'} left of column ${colLabel(b.index)} — drag to move it`}
            style={{ top: a.y, left: b.x - 4, height: a.height }} onMouseDown={shift ? undefined : (e) => startDrag(e, 'col', b)} />
        ))}
        {breakDrag ? (
          <div className={`sh-pb-guide ${breakDrag.axis}`} style={breakDrag.axis === 'row' ? { left: a.x, width: a.width, top: breakDrag.pos - 1 } : { top: a.y, height: a.height, left: breakDrag.pos - 1 }} />
        ) : null}
      </div>
    );
  };

  /**
   * View → Page Layout: each sheet of paper under its cells — white, with
   * its margins, a shadow on the desk, a ruler above and beside it when the
   * ruler is on, the header and footer where they print (a faint "Click to
   * add header" in the margin when there is none; a click opens Header &
   * Footer), and "Click to add data" on the blank pages past the printed
   * range. The engine has already put every cell where it lies on its page.
   */
  const pageLayoutNode = () => {
    const pl = model?.pageLayout;
    if (!pl) return null;
    const cm = (96 / 2.54) / (pl.scale || 1);
    const date = new Date();
    const fields = (text, n) => String(text || '')
      .replace(/&P/g, String(n ?? ''))
      .replace(/&N/g, String(pl.count))
      .replace(/&A/g, model.activeSheet || '')
      .replace(/&F/g, doc?.name || '')
      .replace(/&D/g, date.toLocaleDateString())
      .replace(/&T/g, date.toLocaleTimeString());
    const parts = (text, n) => {
      const s = String(text || '');
      if (!/&[LCR]/.test(s)) return { L: '', C: fields(s, n), R: '' };
      const out = { L: '', C: '', R: '' };
      let which = 'C';
      for (const piece of s.split(/(&[LCR])/)) {
        if (piece === '&L' || piece === '&C' || piece === '&R') which = piece[1];
        else out[which] += piece;
      }
      return { L: fields(out.L, n), C: fields(out.C, n), R: fields(out.R, n) };
    };
    const zone = (kind, p) => {
      const has = kind === 'head' ? Boolean(pl.header) : Boolean(pl.footer);
      const m = pl.margins;
      const width = p.box.width;
      const top = kind === 'head'
        ? (has ? p.box.y - p.y : Math.max(6, m.top / 2 - 12))
        : (has ? p.box.y - p.y + p.box.height - pl.foot : p.height - Math.max(6, m.bottom / 2 + 12));
      const height = has ? (kind === 'head' ? pl.head : pl.foot) : 24;
      const said = has && p.n ? parts(kind === 'head' ? pl.header : pl.footer, p.n) : null;
      const label = kind === 'head' ? 'Click to add header' : 'Click to add footer';
      return (
        <div
          className={`sh-pl-zone ${kind}${has ? ' set' : ''}`}
          data-zone={kind}
          data-tip={has ? (kind === 'head' ? 'Header — click to change it' : 'Footer — click to change it') : label}
          style={{ left: p.box.x - p.x, top, width, height }}
          onMouseDown={(e) => { e.stopPropagation(); }}
          onClick={() => act('headerFooter')}
        >
          {said ? (
            <>
              <span className="l">{said.L}</span>
              <span className="c">{said.C}</span>
              <span className="r">{said.R}</span>
            </>
          ) : has ? null : <span className="c ask">{label}</span>}
        </div>
      );
    };
    const ruler = (p, axis) => {
      const length = axis === 'x' ? p.width : p.height;
      const lo = axis === 'x' ? pl.margins.left : pl.margins.top;
      const hi = length - (axis === 'x' ? pl.margins.right : pl.margins.bottom);
      const ticks = [];
      for (let k = 0, at = 0; at <= length + 0.5; k++, at = (k * cm) / 2) {
        const whole = k % 2 === 0;
        const n = k / 2;
        ticks.push(axis === 'x'
          ? <line key={k} x1={at} x2={at} y1={whole ? 9 : 12} y2={16} />
          : <line key={k} y1={at} y2={at} x1={whole ? 9 : 12} x2={16} />);
        if (whole && n > 0 && n % 1 === 0 && at < length - 8) {
          ticks.push(axis === 'x'
            ? <text key={'t' + k} x={at} y={7.5} textAnchor="middle">{n}</text>
            : <text key={'t' + k} x={7.5} y={at + 3} textAnchor="middle">{n}</text>);
        }
      }
      return (
        <svg className={`sh-pl-ruler ${axis}`} width={axis === 'x' ? length : 16} height={axis === 'x' ? 16 : length}
          style={axis === 'x' ? { left: 0, top: -22 } : { left: -22, top: 0 }}>
          <rect className="margin" x={0} y={0} width={axis === 'x' ? lo : 16} height={axis === 'x' ? 16 : lo} />
          <rect className="margin" x={axis === 'x' ? hi : 0} y={axis === 'x' ? 0 : hi} width={axis === 'x' ? length - hi : 16} height={axis === 'x' ? 16 : length - hi} />
          {ticks}
        </svg>
      );
    };
    return (
      <div className="sh-pl" data-pages={pl.count}>
        {pl.pages.map((p) => (
          <div key={p.col + ':' + p.row} className={`sh-pl-page${p.blank ? ' blank' : ''}`} data-page={p.n ?? ''} style={{ left: p.x, top: p.y, width: p.width, height: p.height }}>
            {pl.ruler ? ruler(p, 'x') : null}
            {pl.ruler && p.col === 0 ? ruler(p, 'y') : null}
            <div className="sh-pl-box" style={{ left: p.box.x - p.x, top: p.box.y - p.y, width: p.box.width, height: p.box.height }} />
            {zone('head', p)}
            {zone('foot', p)}
            {p.blank ? <div className="sh-pl-blank" style={{ left: p.box.x - p.x, top: p.box.y - p.y, width: p.box.width, height: p.box.height }}><span>Click to add data</span></div> : null}
          </div>
        ))}
      </div>
    );
  };

  /**
   * View → Split: the bars between the panes, over the grid where the
   * panes meet. Dragged, a bar moves the split to the nearest row or column
   * edge; dragged to the window's edge, that split goes.
   */
  const splitBars = () => {
    if (!split || !model) return null;
    const z = view.zoom || 1;
    const el = gridRef.current;
    const barW = el ? el.offsetWidth - el.clientWidth : 0;
    const barH = el ? el.offsetHeight - el.clientHeight : 0;
    const start = (e, axis) => {
      if (e.button !== 0 || !el) return;
      e.preventDefault();
      const rect = el.getBoundingClientRect();
      const at = (ev) => (axis === 'h' ? (ev.clientY - rect.top) / z - headTop : (ev.clientX - rect.left) / z - headLeft);
      let pos = at(e);
      setSplitDrag({ axis, pos });
      const move = (ev) => { pos = at(ev); setSplitDrag({ axis, pos }); };
      const stop = () => {
        window.removeEventListener('mousemove', move);
        window.removeEventListener('mouseup', stop);
        setSplitDrag(null);
        const limit = axis === 'h' ? el.clientHeight / z - headTop - 8 : el.clientWidth / z - headLeft - 8;
        const size = pos <= 4 || pos >= limit ? 0 : pos;
        dispatch({ op: 'setSplit', height: axis === 'h' ? size : splitH, width: axis === 'v' ? size : splitW });
      };
      window.addEventListener('mousemove', move);
      window.addEventListener('mouseup', stop);
    };
    const hTop = (headTop + (splitDrag?.axis === 'h' ? splitDrag.pos : splitH)) * z;
    const vLeft = (headLeft + (splitDrag?.axis === 'v' ? splitDrag.pos : splitW)) * z;
    return (
      <>
        {splitH ? <div className={`sh-splitbar h${splitDrag?.axis === 'h' ? ' dragging' : ''}`} data-tip="Split — drag to move; drag to the edge to take it away" style={{ top: hTop - 2, left: 0, right: barW }} onMouseDown={(e) => start(e, 'h')} /> : null}
        {splitW ? <div className={`sh-splitbar v${splitDrag?.axis === 'v' ? ' dragging' : ''}`} data-tip="Split — drag to move; drag to the edge to take it away" style={{ left: vLeft - 2, top: 0, bottom: barH }} onMouseDown={(e) => start(e, 'v')} /> : null}
      </>
    );
  };

  /**
   * Review → comments: the card beside a cell — Excel 365's — with each
   * comment of the thread (who, when, what), a reply box, Resolve and
   * Reopen, and each comment's own Edit and Delete. A new thread starts in
   * an empty card. Keys typed in it stay in it: the grid never sees them.
   */
  const commentCard = () => {
    if (!card || !model) return null;
    const active = sel?.active;
    // The card belongs to its cell; moving away closes it (Excel's way).
    if (!active || active.row !== card.row || active.col !== card.col) return null;
    const thread = model.thread;
    const mode = card.mode === 'thread' && !thread ? null : card.mode;
    if (!mode) return null;
    const box = boxOf({ top: card.row, left: card.col, bottom: card.row, right: card.col });
    if (!box) return null;
    const me = model.commentAuthor || 'You';
    const initials = (name) => String(name || '?').trim().split(/\s+/).slice(0, 2).map((w) => w[0]?.toUpperCase() || '').join('') || '?';
    const hue = (name) => [...String(name || '')].reduce((h, ch) => (h * 31 + ch.charCodeAt(0)) % 360, 17);
    const avatar = (name) => <span className="sh-avatar" style={{ background: `hsl(${hue(name)} 42% 42%)` }}>{initials(name)}</span>;
    const when = (d) => {
      if (!d) return '';
      const t = new Date(/Z$|[+-]\d\d:?\d\d$/.test(d) ? d : d + 'Z');
      return Number.isNaN(t.getTime()) ? '' : t.toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' });
    };
    const post = (text) => act('postComment', { row: card.row, col: card.col, text });
    const composer = (placeholder, label) => (
      <form
        className="sh-card-compose"
        onSubmit={(e) => {
          e.preventDefault();
          const form = e.currentTarget;
          const text = form.elements.words.value.trim();
          // Posted, the box empties for the next reply.
          if (text) { post(text); form.reset(); }
        }}
      >
        {mode === 'new' ? <div className="sh-card-new-who">{avatar(me)}<span>{me}</span></div> : null}
        <textarea
          name="words"
          className="sh-card-input"
          placeholder={placeholder}
          autoFocus={mode !== 'thread'}
          rows={mode === 'new' ? 3 : 2}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); e.currentTarget.form.requestSubmit(); }
          }}
        />
        <div className="sh-card-actions">
          {mode !== 'thread' ? <Button label="Cancel" onClick={() => setCard(thread ? { ...card, mode: 'thread' } : null)} /> : null}
          <Button primary label={label} className="sh-card-post" type="submit" onClick={(e) => { e.preventDefault(); e.currentTarget.closest('form').requestSubmit(); }} />
        </div>
      </form>
    );
    return (
      <div
        className="sh-card"
        data-ref={refText(card.row, card.col)}
        style={{ left: box.right + 10, top: box.y - 4 }}
        onKeyDown={(e) => {
          e.stopPropagation();
          if (e.key === 'Escape') { e.preventDefault(); setCard(null); shRef.current?.focus(); }
        }}
      >
        {thread?.done ? (
          <div className="sh-card-resolved">
            <Icon name="check" size={13} />
            <span className="grow">Resolved</span>
            <Button label="Reopen" className="sh-card-reopen" onClick={() => act('resolveComment', { row: card.row, col: card.col, done: false })} />
          </div>
        ) : null}
        {thread ? (
          <div className="sh-card-list">
            {thread.comments.map((c, i) => (
              <div key={c.id} className="sh-comment" data-id={c.id}>
                <div className="sh-comment-head">
                  {avatar(c.author)}
                  <div className="sh-comment-who">
                    <div className="sh-comment-name">{c.author || 'Someone'}</div>
                    <div className="sh-comment-when">{when(c.date)}</div>
                  </div>
                  <div className="sh-comment-tools">
                    {i === 0 && !thread.done ? (
                      <Button icon="check" className="sh-card-resolve" title="Resolve thread — mark the conversation done; Reopen brings it back" onClick={() => act('resolveComment', { row: card.row, col: card.col, done: true })} />
                    ) : null}
                    <Button icon="more" className="sh-comment-more" title="More — edit or delete" onClick={(e) => menu.open(e, [
                      { label: 'Edit comment', icon: 'textbox', run: () => setCard({ ...card, mode: 'thread', editing: c.id }) },
                      i === 0
                        ? { label: 'Delete thread', icon: 'trash', run: () => act('deleteComment', { id: c.id, top: true }) }
                        : { label: 'Delete comment', icon: 'trash', run: () => act('deleteComment', { id: c.id }) },
                    ])} />
                  </div>
                </div>
                {card.editing === c.id ? (
                  <form
                    className="sh-card-compose"
                    style={{ padding: '8px 0 0 34px' }}
                    onSubmit={(e) => {
                      e.preventDefault();
                      const text = e.currentTarget.elements.words.value.trim();
                      if (text) act('editComment', { id: c.id, text });
                    }}
                  >
                    <textarea name="words" className="sh-comment-edit" defaultValue={c.text} autoFocus rows={3} onKeyDown={(e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); e.currentTarget.form.requestSubmit(); } }} />
                    <div className="sh-card-actions">
                      <Button label="Cancel" onClick={() => setCard({ ...card, editing: null })} />
                      <Button primary label="Save" className="sh-comment-save" onClick={(e) => { e.preventDefault(); e.currentTarget.closest('form').requestSubmit(); }} />
                    </div>
                  </form>
                ) : (
                  <div className="sh-comment-text">{c.text}</div>
                )}
              </div>
            ))}
          </div>
        ) : null}
        {thread?.done ? null : mode === 'new' ? composer('Start a conversation', 'Post') : composer('Reply…', 'Reply')}
      </div>
    );
  };

  /**
   * Review → Show Comments: every thread in the workbook, sheet by sheet,
   * the open or the resolved alone at a press. A thread pressed is gone to
   * and its card opened.
   */
  const commentsPane = () => {
    const all = model?.comments;
    if (!Array.isArray(all)) return null;
    const list = all.filter((t) => (commentFilter === 'open' ? !t.done : commentFilter === 'resolved' ? t.done : true));
    const open = all.filter((t) => !t.done).length;
    return (
      <div className="sh-comments">
        <div className="sh-comments-head">
          <strong>Comments</strong>
          <Spacer />
          <Button icon="plus" label="New" title="New — a comment on the active cell" onClick={() => act('newComment')} />
          <Button icon="close" title="Close — hides the Comments pane" onClick={() => act('commentsOpen', false)} />
        </div>
        <div className="sh-comments-filter">
          {[['all', `All ${all.length}`], ['open', `Open ${open}`], ['resolved', `Resolved ${all.length - open}`]].map(([key, label]) => (
            <button key={key} type="button" className={commentFilter === key ? 'on' : ''} data-filter={key} onClick={() => setCommentFilter(key)}>{label}</button>
          ))}
        </div>
        <div className="sh-comments-list">
          {list.length ? list.map((t) => {
            const here = t.sheet === model.activeSheet && sel?.active?.row === t.row && sel?.active?.col === t.col;
            const first = t.comments[0] || {};
            const replies = t.comments.length - 1;
            return (
              <div key={t.sheet + '!' + t.ref} className={`sh-comments-item${here ? ' here' : ''}${t.done ? ' resolved' : ''}`} data-ref={t.ref} data-sheet={t.sheet} onClick={() => act('gotoThread', t)}>
                <div className="where"><span className="ref">{t.ref}</span><span>{t.sheet}</span>{t.done ? <span className="done">Resolved</span> : null}</div>
                <div className="first"><b>{first.author || 'Someone'}</b><span>{first.text}</span></div>
                {replies ? <div className="more">{replies} {replies === 1 ? 'reply' : 'replies'}</div> : null}
              </div>
            );
          }) : <div className="sh-comments-empty">{all.length ? 'No comments in this view.' : 'No comments yet. New Comment starts one on the active cell.'}</div>}
        </div>
      </div>
    );
  };

  // Review → Check Accessibility and Spelling (sheets/review.js) — a hook,
  // so above the early return.
  const review = useSheetsReview({ shell, doc, model, dispatch, toast });

  if (error) {
    return (
      <AppFrame app={app} shell={shell} title="Worksheets" menu={appMenu}>
        <Empty icon={lockedOut ? 'lock' : 'sheets'} title={lockedOut ? 'This workbook is password-protected' : 'This file could not be opened'} action={lockedOut ? <LockedAction /> : null}>
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
  /**
   * Ops whose refusal a dialog shows in place — a wrong password, a range
   * that will not do — rather than as a toast: the sentence, or null.
   */
  const tryOps = async (...ops) => {
    if (!doc) return 'No workbook';
    try {
      const next = await shell.doc.apply({ id: doc.id, ops });
      setDoc(next);
      if (next.patch) setModel((m) => (m ? withSelection(m, next.patch) : m));
      else setModel(next.model);
      return null;
    } catch (err) {
      return String(err?.message || err);
    }
  };

  const act = async (name, arg, opts = {}) => {
    const at = sel?.active || { row: 0, col: 0 };
    const range = sel?.range || { top: at.row, left: at.col, bottom: at.row, right: at.col };
    switch (name) {
      case 'toggleGridlines': patchView((v) => ({ gridlines: v.gridlines === false })); return;
      case 'toggleHeadings': patchView((v) => ({ headings: v.headings === false })); return;
      case 'toggleFormulaBar': patchView((v) => ({ formulaBar: v.formulaBar === false })); return;
      case 'toggleFormulas': patchView((v) => ({ formulas: !v.formulas })); return;
      case 'page': {
        // Margins, Orientation, Size, Print Titles and Scale to Fit write the
        // file's own page setup — what the print dialog and Excel start from
        // — through the op the dialog uses; until 2026-09-21 they changed a
        // note in this window that nothing read. The setup is read first and
        // sent back whole, since the writer rebuilds it from what it is given.
        const current = await shell.doc.pageSetup({ id: doc.id });
        const next = { ...current };
        let said = 'Page setup saved with the file';
        if (arg.orientation) { next.orientation = arg.orientation; said = `Orientation: ${arg.orientation}`; }
        if (arg.size) { next.paper = arg.size; said = `Paper: ${arg.size}`; }
        if (arg.margins && MARGIN_PRESETS[arg.margins]) { next.margins = { ...MARGIN_PRESETS[arg.margins] }; said = `Margins: ${arg.margins}`; }
        if (arg.fit !== undefined) {
          next.fit = arg.fit;
          said = { none: 'No scaling', width: 'All the columns on one page across', height: 'All the rows on one page down', page: 'The sheet on one page' }[arg.fit] || 'Scaling set';
        }
        if (arg.scale !== undefined) { next.scale = arg.scale; next.fit = 'none'; said = `Scale: ${Math.round(arg.scale * 100)}%`; }
        if (arg.repeatRows !== undefined) {
          next.repeatRows = arg.repeatRows === 'selection'
            ? (sel && Number.isFinite(sel.top) && sel.top === 0 ? sel.bottom + 1 : 0)
            : Number(arg.repeatRows) || 0;
          said = next.repeatRows ? `Rows 1 to ${next.repeatRows} repeat at the top of every page` : 'No rows repeat';
          if (arg.repeatRows === 'selection' && !next.repeatRows) said = 'Select rows from row 1 to repeat them';
        }
        if (arg.header !== undefined || arg.footer !== undefined) {
          next.header = arg.header || null;
          next.footer = arg.footer || null;
          said = next.header || next.footer ? 'Header and footer saved with the file' : 'No header or footer';
        }
        if (arg.breaks) {
          // Excel's Breaks: a break goes above the cell's row and left of its
          // column (only one of them at the sheet's edge), comes off at the
          // cell, or they all go.
          const rowsSet = new Set(current.rowBreaks || []);
          const colsSet = new Set(current.colBreaks || []);
          const at = sel && Number.isFinite(sel.top) && Number.isFinite(sel.left) ? sel : null;
          const colName = (n) => { let s = ''; let k = n + 1; while (k > 0) { const m = (k - 1) % 26; s = String.fromCharCode(65 + m) + s; k = Math.floor((k - 1) / 26); } return s; };
          if (arg.breaks === 'reset') {
            rowsSet.clear(); colsSet.clear();
            said = 'All page breaks removed';
          } else if (!at) {
            return;
          } else if (arg.breaks === 'insert') {
            if (at.top > 0) rowsSet.add(at.top);
            if (at.left > 0) colsSet.add(at.left);
            said = at.top > 0 && at.left > 0 ? `Page break above row ${at.top + 1} and left of column ${colName(at.left)}`
              : at.top > 0 ? `Page break above row ${at.top + 1}`
              : at.left > 0 ? `Page break left of column ${colName(at.left)}`
              : 'A page break goes above the row and left of the column of the cell: pick one past A1';
          } else if (arg.breaks === 'remove') {
            const had = rowsSet.delete(at.top) | colsSet.delete(at.left);
            said = had ? 'Page break removed' : 'No page break at the cell';
          }
          next.rowBreaks = [...rowsSet].sort((a, b) => a - b);
          next.colBreaks = [...colsSet].sort((a, b) => a - b);
        }
        await dispatch({ op: 'setPageSetup', setup: next });
        patchView({ page: next });
        toast(said, { tone: 'good' });
        return;
      }
      // View → Normal / Page Break Preview, and the status bar's buttons.
      case 'view': {
        if (arg !== 'normal' && arg !== 'pageBreakPreview' && arg !== 'pageLayout') return;
        const was = model?.viewMode || 'normal';
        if (was === arg) return;
        // Each view keeps its own zoom, as Excel's do: the preview opens at
        // 60% — whole pages in sight — Page Layout at the zoom Normal had,
        // and each goes back to where it was.
        const keys = { normal: 'normalZoom', pageBreakPreview: 'previewZoom', pageLayout: 'layoutZoom' };
        const first = { normal: 1, pageBreakPreview: 0.6 };
        patchView((v) => ({ [keys[was]]: v.zoom ?? 1, zoom: v[keys[arg]] ?? first[arg] ?? v.zoom ?? 1 }));
        await dispatch({ op: 'setViewMode', mode: arg });
        gridRef.current?.scrollTo?.(0, 0);
        return;
      }
      // View → Ruler, in Page Layout.
      case 'toggleRuler':
        await dispatch({ op: 'setShowRuler', on: !(model?.showRuler !== false) });
        return;
      // View → Custom Views, greyed with Excel's reason in a workbook with a table.
      case 'customViews':
        if (model?.customViewsBlocked) { toast(`Custom Views are ${model.customViewsBlocked}`, { ms: 4000 }); return; }
        setDialog('customViews');
        return;
      // Page Layout → Background, or Delete Background when the sheet has one.
      case 'background': {
        if (model?.background) {
          if (await dispatch({ op: 'deleteBackground' })) toast('Background deleted', { tone: 'good' });
          return;
        }
        const [file] = await shell.dialog.open({
          title: 'Sheet Background',
          filters: [{ name: 'Pictures', extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp'] }],
        });
        if (!file) return;
        const { bytes, stat } = await shell.fs.read({ path: file });
        const ext = String(stat?.ext || file.split('.').pop()).replace('.', '').toLowerCase();
        const contentType = { png: 'image/png', gif: 'image/gif', webp: 'image/webp', bmp: 'image/bmp' }[ext] || 'image/jpeg';
        if (await dispatch({ op: 'setBackground', contentType, data: bytes })) toast('Background set — it is drawn behind the cells and not printed, as in Excel', { tone: 'good', ms: 4200 });
        return;
      }
      // View → Split, a toggle at the active cell.
      case 'split':
        await dispatch({ op: 'toggleSplit' });
        return;
      // A page break dragged in the preview: it becomes one put by hand where
      // it was dropped (Excel's way); dragged out of the printed area it goes.
      case 'moveBreak': {
        const { axis, from, to, manual } = arg;
        const current = await shell.doc.pageSetup({ id: doc.id });
        const key = axis === 'row' ? 'rowBreaks' : 'colBreaks';
        const set = new Set(current[key] || []);
        if (manual) set.delete(from);
        const area = model?.pageBreaks?.area;
        const inside = area && (axis === 'row' ? to > area.top && to <= area.bottom : to > area.left && to <= area.right);
        if (inside) set.add(to);
        const next = { ...current, [key]: [...set].sort((a, b) => a - b) };
        await dispatch({ op: 'setPageSetup', setup: next });
        patchView({ page: next });
        toast(inside
          ? (axis === 'row' ? `Page break above row ${to + 1}` : `Page break left of column ${colLabel(to)}`)
          : 'Page break removed', { tone: 'good', ms: 2400 });
        return;
      }
      case 'zoom': {
        // The grid alone is scaled, with CSS zoom on its scroll container. It
        // used to be the window's zoom — ribbon, status bar and the slider
        // under the pointer all grew with the cells, which made the slider
        // unusable (owner, 2026-09-24). Under CSS zoom the container's own
        // scroll and client sizes stay in its pixels, so the viewport it asks
        // for is right as it is; only pointer positions, which arrive in
        // screen pixels, are divided by the level.
        patchView({ zoom: Math.max(0.3, Math.min(3, Math.round((Number(arg) || 1) * 100) / 100)) });
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
      // Calculate Now (F9), Calculate Sheet (Shift+F9) and a full pass.
      case 'calculate': {
        const scope = arg || 'workbook';
        const next = await dispatch({ op: 'calculate', scope });
        if (!next) return;
        const n = Number(next.opResult) || 0;
        toast(`${scope === 'sheet' ? 'Calculated this sheet' : 'Calculated'}: ${n} formula${n === 1 ? '' : 's'} worked out`, { tone: 'good', ms: 2400 });
        return;
      }
      // Formulas → Calculation Options: written to the workbook.
      case 'calcMode': {
        await dispatch({ op: 'setCalcMode', mode: arg });
        toast({ auto: 'Calculation: automatic', autoNoTable: 'Calculation: automatic except for data tables', manual: 'Calculation: manual — F9 calculates' }[arg] || 'Calculation set', { tone: 'good', ms: 2600 });
        return;
      }
      // Formulas → Evaluate Formula: the active cell's formula, a part at a time.
      case 'evaluateFormula': {
        const cell = { row: at.row, col: at.col };
        try {
          const first = await shell.doc.evaluateFormula({ id: doc.id, ...cell, actions: [] });
          setEvaluating({ ...cell, first });
        } catch (err) {
          toast(String(err?.message || err), { tone: 'warn', ms: 4500 });
        }
        return;
      }
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
      case 'link': setDialog('link'); return;
      case 'note': setDialog('note'); return;
      case 'headerFooter': setDialog('headerFooter'); return;
      case 'sortDialog': setDialog('sort'); return;
      // Data → Outline. Whole rows group as rows and whole columns as
      // columns; anything else asks which, as Excel's Group dialog does.
      case 'group':
      case 'ungroup': {
        const axis = arg || sel?.whole;
        if (!axis) { setOutlineAsk(name); setDialog('outlineAxis'); return; }
        await dispatch({ op: name, axis });
        return;
      }
      case 'clearOutline': await dispatch({ op: 'clearOutline' }); return;
      // Data → Flash Fill (Ctrl+E), Advanced and Clear: what the engine says
      // when it will not is a note, not an alarm — a toast that says why.
      case 'flashFill':
      case 'clearFilter':
      case 'advancedFilter': {
        const op = name === 'flashFill' ? { op: 'flashFill' } : name === 'clearFilter' ? { op: 'clearAdvancedFilter' } : { op: 'advancedFilter', ...arg };
        try {
          const next = await shell.doc.apply({ id: doc.id, ops: [op] });
          setDoc(next);
          setModel(next.model);
          const n = Number(next.opResult) || 0;
          if (name === 'flashFill') toast(`Flash Fill filled ${n} cell${n === 1 ? '' : 's'}`, { tone: 'good' });
          else if (name === 'clearFilter') toast(`${n} row${n === 1 ? '' : 's'} shown again`, { tone: 'good' });
          else toast(arg.action === 'copy' ? `${n} row${n === 1 ? '' : 's'} copied to ${arg.copyTo}` : `${n} row${n === 1 ? '' : 's'} of the list pass the criteria`, { tone: 'good' });
        } catch (err) {
          toast(String(err?.message || err), { tone: 'warn', ms: 5500 });
        }
        return;
      }
      case 'advancedDialog': {
        const next = await dispatch({ op: 'listFields' });
        let info = null;
        try { info = next?.opResult ? JSON.parse(next.opResult) : null; } catch { info = null; }
        setListInfo(info || { ref: '', columns: [], filter: {} });
        setDialog('advanced');
        return;
      }
      case 'showDetail': await dispatch({ op: 'showDetail' }); return;
      case 'hideDetail': await dispatch({ op: 'hideDetail' }); return;
      case 'subtotalDialog': {
        const next = await dispatch({ op: 'listFields' });
        let info = null;
        try { info = next?.opResult ? JSON.parse(next.opResult) : null; } catch { info = null; }
        if (!info || info.bottom <= info.top) {
          toast('Select a cell in a list with a header row first — Subtotal groups the rows under it', { tone: 'warn', ms: 5000 });
          return;
        }
        setListInfo(info);
        setDialog('subtotal');
        return;
      }
      case 'removeNote': await dispatch({ op: 'removeNote', row: sel?.active?.row ?? 0, col: sel?.active?.col ?? 0 }); return;
      // Review → New Comment: the card on the active cell, its box ready —
      // a reply at the end of the thread when the cell has one already.
      case 'newComment': {
        if (model?.note) { toast('This cell has a note. Delete it, or edit it, to comment here instead.', { tone: 'warn', ms: 4500 }); return; }
        setCard({ row: at.row, col: at.col, mode: model?.thread ? 'reply' : 'new' });
        return;
      }
      case 'postComment': {
        const next = await dispatch({ op: 'addComment', row: arg.row, col: arg.col, text: arg.text });
        if (next) setCard({ row: arg.row, col: arg.col, mode: 'thread' });
        return;
      }
      case 'editComment': {
        const next = await dispatch({ op: 'editComment', id: arg.id, text: arg.text });
        if (next) setCard((c) => (c ? { ...c, editing: null } : c));
        return;
      }
      case 'deleteComment': {
        await dispatch({ op: 'deleteComment', id: arg.id });
        if (arg.top) setCard(null);
        return;
      }
      case 'deleteThread': {
        if (!model?.thread) { toast('There is no comment on this cell to delete.', { ms: 3000 }); return; }
        await dispatch({ op: 'deleteThread', row: at.row, col: at.col });
        setCard(null);
        return;
      }
      case 'resolveComment': {
        await dispatch({ op: 'resolveComment', row: arg.row, col: arg.col, done: arg.done });
        return;
      }
      // Previous and Next select the thread before or after the cell and open it.
      case 'stepComment': {
        const next = await dispatch({ op: 'stepComment', direction: arg });
        if (!next) return;
        if (!next.opResult) { toast('This workbook has no comments.', { ms: 3000 }); return; }
        const a = next.model?.selection?.active;
        if (a) setCard({ row: a.row, col: a.col, mode: 'thread' });
        return;
      }
      case 'commentsOpen':
        await dispatch({ op: 'commentsOpen', on: arg });
        return;
      case 'gotoThread': {
        const t = arg;
        const ops = [];
        if (t.sheet && t.sheet !== model?.activeSheet) ops.push({ op: 'sheet', name: t.sheet });
        ops.push({ op: 'select', row: t.row, col: t.col });
        const next = await dispatch(...ops);
        if (next) setCard({ row: t.row, col: t.col, mode: 'thread' });
        return;
      }
      // Format as Table: over the selection, or the block of data round the cell.
      case 'table': await dispatch({ op: 'formatAsTable', style: arg?.style, stripes: arg?.stripes !== false }); return;
      // Insert → PivotTable, and PivotChart: on a pivot a chart of it (its
      // kind from the menu), on data PivotChart & PivotTable — a dialog that
      // opens on the list round the cursor.
      case 'pivotTable':
      case 'pivotChart': {
        if (name === 'pivotChart' && arg?.kind) {
          await dispatch({ op: 'insertPivotChart', kind: arg.kind, fileName: doc?.name });
          return;
        }
        const next = await dispatch({ op: 'listFields' });
        let info = null;
        try { info = next?.opResult ? JSON.parse(next.opResult) : null; } catch { info = null; }
        setPivotAsk({ list: info && info.bottom > info.top ? info : null, chart: name === 'pivotChart' });
        return;
      }
      case 'themeGallery': setGallery(gallery?.kind === arg.kind ? null : arg); return;
      case 'arrange': {
        const ids = picked.filter((id) => (model?.drawings || []).some((d) => d.id === id) || (model?.objects || []).some((o) => o.id === id));
        if (arg?.op === 'pane') { setSelPane((v) => !v); return; }
        if (!ids.length) { toast('Select a picture, shape, chart or slicer first — click it on the sheet, or in the Selection Pane.', { tone: 'warn', ms: 4000 }); return; }
        if (arg.op === 'group') {
          const next = await dispatch({ op: 'groupDrawings', ids });
          if (next?.opResult) setPicked([next.opResult]);
          return;
        }
        if (arg.op === 'ungroup') { const next = await dispatch({ op: 'ungroupDrawings', ids }); if (next) setPicked([]); return; }
        const ops = {
          order: { op: 'reorderDrawings', ids, to: arg.to },
          align: { op: 'alignDrawings', ids, edge: arg.edge },
          distribute: { op: 'distributeDrawings', ids, axis: arg.axis },
          rotate: { op: 'rotateDrawings', ids, by: arg.by || 0, flip: arg.flip || null },
        };
        await dispatch(ops[arg.op]);
        return;
      }
      case 'slicer': {
        const next = await dispatch({ op: 'slicerSources' });
        let src = null;
        try { src = next?.opResult ? JSON.parse(next.opResult) : null; } catch { src = null; }
        if (!src?.kind) {
          toast(src?.reason || 'Put the cursor in a table or a pivot table — a slicer filters one of them.', { tone: 'warn', ms: 5000 });
          return;
        }
        setSlicerAsk(src);
        return;
      }
      case 'picture': await insertPicture(); return;
      // Page Layout → Print Area: the page setup is rebuilt from what it is
      // given, so the file's own setup is read first and sent back with the
      // area changed — what the print dialog does.
      case 'orientation': {
        const rotation = Number(arg) || 0;
        await dispatch({ op: 'setFormat', delta: { rotation } });
        if (!rotation || !sel || !Number.isFinite(sel.top)) return;
        // Excel grows a row to fit the turned words; measure each selected
        // cell's text at its font and ask for the tallest, never shorter.
        const ctx = document.createElement('canvas').getContext('2d');
        const rows = new Map();
        for (const c of model?.cells || []) {
          if (c.row < sel.top || c.row > sel.bottom || c.col < sel.left || c.col > sel.right || !c.text) continue;
          const font = c.style?.font || {};
          const px = font.sizePt ? (font.sizePt * 4) / 3 : 12.5;
          ctx.font = `${font.bold ? '700 ' : ''}${px}px ${font.family || 'Calibri, Arial, sans-serif'}`;
          const width = ctx.measureText(String(c.text)).width;
          const line = px * 1.25;
          const need = rotation === 255 ? String(c.text).length * line + 6
            : rotation === 90 || rotation === 180 ? width + 10
            : Math.abs(Math.sin(((rotation > 90 ? rotation - 90 : rotation) * Math.PI) / 180)) * width + line + 6;
          if (need > (c.height || 0)) rows.set(c.row, Math.max(rows.get(c.row) || 0, Math.ceil(need)));
        }
        for (const [row, height] of rows) await dispatch({ op: 'rowHeight', row, height });
        return;
      }
      case 'trace': {
        const at = sel?.active || { row: 0, col: 0 };
        const got = await shell.doc.trace({ id: doc.id, kind: arg, row: at.row, col: at.col }).catch(() => null);
        if (!got) return;
        if (!got.arrows.length) {
          toast(arg === 'dependents' ? 'No formula on this sheet reads this cell' : (got.elsewhere ? 'This formula reads other sheets or names only' : 'This cell reads no other cell'), { ms: 3500 });
          return;
        }
        setArrows((prev) => [...prev, ...got.arrows.map((a) => ({ ...a, kind: arg, at: { row: at.row, col: at.col } }))]);
        if (got.elsewhere) toast(`${got.elsewhere} reference${got.elsewhere === 1 ? '' : 's'} on other sheets or by name not drawn`, { ms: 3500 });
        return;
      }
      case 'removeArrows':
        setArrows([]);
        return;
      case 'namesFromSelection': {
        try {
          await dispatch({ op: 'namesFromSelection' });
          toast('A name for each column, from its header — the Name Manager lists them', { tone: 'good', ms: 4000 });
        } catch (err) {
          toast(String(err?.message || err), { tone: 'warn', ms: 5000 });
        }
        return;
      }
      case 'textToColumns': {
        // The engine refuses more than one column; its message is the toast.
        try {
          await dispatch({ op: 'textToColumns', delimiter: arg || 'comma' });
          toast('Split into the cells to the right', { tone: 'good' });
        } catch (err) {
          toast(String(err?.message || err), { tone: 'warn', ms: 5000 });
        }
        return;
      }
      case 'removeDuplicates': {
        const before = await shell.doc.model({ id: doc.id }).catch(() => null);
        const filledBefore = (before?.cells || []).filter((c) => c.text !== '').length;
        await dispatch({ op: 'removeDuplicates' });
        const after = await shell.doc.model({ id: doc.id }).catch(() => null);
        const filledAfter = (after?.cells || []).filter((c) => c.text !== '').length;
        toast(filledAfter < filledBefore ? 'Duplicate rows removed; the rest closed up' : 'No duplicate rows in the selection', { tone: 'good' });
        return;
      }
      case 'addSheet': {
        try {
          await dispatch({ op: 'addSheet' });
        } catch (err) {
          toast(String(err?.message || err), { tone: 'warn', ms: 5000 });
        }
        return;
      }
      // Review → Protect Sheet / Unprotect Sheet: a dialog to protect (the
      // password optional), the password asked for to unprotect when there is one.
      case 'protectSheet': {
        if (model?.protection?.sheet) {
          if (model.protection.hasPassword) {
            setPasswordAsk({ kind: 'sheet', title: 'Unprotect Sheet', message: 'This sheet is protected with a password. Type it to take the protection off.' });
            return;
          }
          if (await dispatch({ op: 'unprotect' })) toast('Sheet unprotected', { tone: 'good' });
          return;
        }
        setDialog('protectSheet');
        return;
      }
      // Review → Protect Workbook: the structure locked, or unlocked again.
      case 'protectWorkbook': {
        if (model?.workbookProtection?.structure) {
          if (model.workbookProtection.hasPassword) {
            setPasswordAsk({ kind: 'workbook', title: 'Unprotect Workbook', message: 'The workbook’s structure is protected with a password. Type it to take the protection off.' });
            return;
          }
          if (await dispatch({ op: 'unprotectWorkbook' })) toast('Workbook unprotected — sheets can be added, moved and renamed again', { tone: 'good' });
          return;
        }
        setDialog('protectWorkbook');
        return;
      }
      case 'editRanges': setDialog('editRanges'); return;
      // Data → Consolidate: the dialog opens on the sheet's last consolidation and the selection.
      case 'consolidateDialog': {
        const next = await dispatch({ op: 'consolidateInfo' });
        if (!next) return;
        setAnalysis(JSON.parse(next.opResult || 'null'));
        setDialog('consolidate');
        return;
      }
      // Data → Forecast Sheet: the dialog opens on the timeline and values round the cell.
      case 'forecastDialog': {
        const next = await dispatch({ op: 'forecastInfo' });
        if (!next) return;
        const info = JSON.parse(next.opResult || 'null');
        if (!info) { toast('Select a timeline and its values — dates in one column, numbers in the next — or put the cursor in them.', { ms: 5000 }); return; }
        setAnalysis(info);
        setDialog('forecast');
        return;
      }
      // The tabs' Hide, Unhide and Move: refused, in Excel's words, while the structure is locked.
      case 'hideSheet':
        await dispatch({ op: 'hideSheet', name: arg });
        return;
      case 'unhideSheet':
        await dispatch({ op: 'unhideSheet', name: arg });
        return;
      case 'moveSheet':
        await dispatch({ op: 'moveSheet', name: arg.name, to: arg.to });
        return;
      case 'printArea': {
        const current = await shell.doc.pageSetup({ id: doc.id });
        if (arg === 'clear') {
          await dispatch({ op: 'setPageSetup', setup: { ...current, area: '' } });
          toast('Print area cleared', { tone: 'good' });
          return;
        }
        // The frame's selection is flat: the range's top, bottom, left and right sit beside `active`.
        const r = sel && Number.isFinite(sel.top) && Number.isFinite(sel.left) ? sel : null;
        if (!r) return;
        const letters = (n) => { let s = ''; let k = n + 1; while (k > 0) { const m = (k - 1) % 26; s = String.fromCharCode(65 + m) + s; k = Math.floor((k - 1) / 26); } return s; };
        const area = `${letters(r.left)}${r.top + 1}:${letters(r.right)}${r.bottom + 1}`;
        await dispatch({ op: 'setPageSetup', setup: { ...current, area } });
        toast(`Print area: ${area}`, { tone: 'good' });
        return;
      }
      // A link is followed: an address opens outside the suite, a place in
      // the workbook (C12, Sheet2!B4, a name) is gone to.
      case 'follow': {
        const link = arg;
        if (!link) return;
        if (link.href) { shell.shell.openExternal({ url: link.href }); return; }
        const place = String(link.location || '');
        const m = /^(?:'([^']+)'|([^!]+))!(.+)$/.exec(place);
        const sheetName = m ? (m[1] || m[2]) : null;
        if (sheetName && sheetName !== model?.activeSheet) await dispatch({ op: 'sheet', name: sheetName });
        await act('goto', m ? m[3] : place);
        return;
      }
      // Formulas → Error Checking: the ribbon button toggles it (no arg);
      // the pane's own Close passes false so it always ends up off.
      case 'errorCheck':
        await dispatch({ op: 'errorCheck', on: arg });
        return;
      // A row in the error pane, or Next/Previous: select the cell it names,
      // switching sheet first when it is on another one — `select` scrolls
      // it into view itself (SheetView#ensureVisible).
      case 'gotoError': {
        const err = arg;
        if (!err) return;
        const ops = [];
        if (err.sheet && err.sheet !== model?.activeSheet) ops.push({ op: 'sheet', name: err.sheet });
        ops.push({ op: 'select', row: err.row, col: err.col });
        await dispatch(...ops);
        return;
      }
      case 'stepError': {
        const list = Array.isArray(model?.errors) ? model.errors : [];
        if (!list.length) return;
        const active = sel?.active;
        const at = list.findIndex((e) => e.sheet === model.activeSheet && active && e.row === active.row && e.col === active.col);
        const dir = arg === 'prev' ? -1 : 1;
        const next = at === -1 ? (dir > 0 ? 0 : list.length - 1) : (at + dir + list.length) % list.length;
        await act('gotoError', list[next]);
        return;
      }
      // Formulas → Watch Window: the ribbon button toggles it (no arg); the
      // pane's own Close passes false, the same shape as errorCheck.
      case 'watchOpen':
        await dispatch({ op: 'watchOpen', on: arg });
        return;
      // Add Watch: every cell of the current selection, not only the active
      // one — a person who wants one cell selects one cell. A selection
      // bigger than the cap (a whole row or column picked by its header)
      // adds only the first of it rather than flooding the pane.
      case 'addWatch': {
        const cap = 500;
        const all = [];
        for (let row = range.top; row <= range.bottom && all.length <= cap; row++) {
          for (let col = range.left; col <= range.right; col++) all.push({ sheet: model.activeSheet, row, col });
        }
        const refs = all.slice(0, cap);
        await dispatch({ op: 'watchAdd', refs });
        if (all.length > cap) toast(`Only the first ${cap} cells of the selection were added`, { tone: 'warn', ms: 4000 });
        return;
      }
      // A row's own × button: remove that one watch.
      case 'removeWatch':
        if (!arg) return;
        await dispatch({ op: 'watchRemove', ref: arg });
        return;
      // Delete Watch: the row selected in the pane, if any.
      case 'deleteWatch':
        if (!watchSel) return;
        await dispatch({ op: 'watchRemove', ref: watchSel });
        setWatchSel(null);
        return;
      // A double-clicked row: go to that cell, switching sheet first when
      // it is on another one — the same move gotoError makes.
      case 'gotoWatch': {
        const w = arg;
        if (!w) return;
        const ops = [];
        if (w.sheet && w.sheet !== model?.activeSheet) ops.push({ op: 'sheet', name: w.sheet });
        ops.push({ op: 'select', row: w.row, col: w.col });
        await dispatch(...ops);
        return;
      }
      case 'help': shell.shell.openExternal({ url: SITE.help }); return;
      case 'feedback': shell.shell.openExternal({ url: SITE.contact }); return;
      case 'about': shell.win.create({ app: 'home', query: { about: 1 } }); return;
      default:
        toast(`${name} is not wired yet.`, { ms: 3000 });
    }
  };

  actRef.current = act;

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
          arrange={{ picked: (model?.objects || []).filter((o) => picked.includes(o.id)), pane: selPane }}
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
          review={review}
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
          {review.status}
          <Chip>{model?.activeSheet || ''}</Chip>
          <Chip>{sel?.ref || ''}</Chip>
          {model?.link ? <Chip title="Ctrl+click the cell to open it">{model.link.href || model.link.location}</Chip> : null}
          {model?.notes ? <Chip title="Rest the pointer on a marked cell to read its note">{model.notes} {model.notes === 1 ? 'note' : 'notes'}</Chip> : null}
          {model?.threads ? <Chip title="Review → Show Comments lists them">{model.threads} {model.threads === 1 ? 'comment' : 'comments'}</Chip> : null}
          {model?.calc?.pending ? (
            // Excel's word for formulas that have not caught up with an edit
            // in manual mode; pressing it is Calculate Now.
            <button type="button" className="sh-calc-pending" data-tip="Calculate — formulas are waiting for Calculate Now (F9)" onClick={() => act('calculate', 'workbook')}>Calculate</button>
          ) : null}
          <span className="sh-viewbtns">
            <button type="button" className={`sh-viewbtn${(model?.viewMode || 'normal') === 'normal' ? ' on' : ''}`} data-view="normal" data-tip="Normal" onClick={() => act('view', 'normal')}>
              <svg width="14" height="14" viewBox="0 0 14 14"><rect x="1.5" y="1.5" width="11" height="11" rx="1" /><path d="M1.5 5.2h11M1.5 8.8h11M5.2 1.5v11M8.8 1.5v11" /></svg>
            </button>
            <button type="button" className={`sh-viewbtn${model?.viewMode === 'pageLayout' ? ' on' : ''}`} data-view="pageLayout" data-tip="Page Layout" onClick={() => act('view', 'pageLayout')}>
              <svg width="14" height="14" viewBox="0 0 14 14"><rect x="2.5" y="1.5" width="9" height="11" rx="0.8" /><path d="M4.3 4h5.4M4.3 6.2h5.4M4.3 8.4h5.4M4.3 10.6h3.2" /></svg>
            </button>
            <button type="button" className={`sh-viewbtn${model?.viewMode === 'pageBreakPreview' ? ' on' : ''}`} data-view="pageBreakPreview" data-tip="Page Break Preview" onClick={() => act('view', 'pageBreakPreview')}>
              <svg width="14" height="14" viewBox="0 0 14 14"><rect x="1.5" y="1.5" width="11" height="11" rx="1" /><path d="M7 1.5v11" strokeDasharray="1.6 1.4" /><path d="M1.5 7h11" /></svg>
            </button>
          </span>
          <ZoomSlider value={view.zoom ?? 1} onChange={(v) => act('zoom', v)} onReset={() => act('zoom', 1)} />
        </>
      }
    >
      {busy || !model ? (
        <div style={{ flex: 1, display: 'grid', placeItems: 'center' }}>
          <Spinner style={{ width: 22, height: 22 }} />
        </div>
      ) : (
        <div className={`sh${view.gridlines === false ? ' no-grid' : ''}${view.headings === false ? ' no-heads' : ''}${model.viewMode === 'pageLayout' ? ' pl' : ''}${backdrop && model.viewMode !== 'pageLayout' ? ' has-bg' : ''}`} onKeyDown={onKeyDown} tabIndex={0} ref={(el) => { shRef.current = el; if (el && !editing && document.activeElement === document.body) el.focus(); }}>
          <style>{CSS + OBJECTS_CSS + DESIGN_CSS + SHEET_DESIGN_CSS}</style>

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
          <div className="sh-body">
          <div className="sh-grid" ref={gridRef} style={{ zoom: view.zoom && Math.abs(view.zoom - 1) > 0.001 ? view.zoom : undefined }}>
            <div
              className="sh-canvas"
              style={{
                gridTemplateColumns: `${model.headerWidth + gutW}px ${model.total.width}px`,
                gridTemplateRows: `${model.headerHeight + gutH}px ${model.total.height}px`,
              }}
            >
              <div className={`sh-corner${gutW || gutH ? ' outlined' : ''}`}>
                {gutW || gutH ? levelButtons() : null}
              </div>

              <div className="sh-colheads" style={{ height: model.headerHeight + gutH }}>
                {gutH ? gutterNode('col') : null}
                {frozen.cols ? (
                  <div className="sh-pin sh-pin-colheads" style={{ left: headLeft, width: frozenW, height: model.headerHeight + gutH }}>
                    {gutH ? gutterNode('col', true) : null}
                    {(model.columns || []).filter((c) => c.index < frozen.cols).map((c) => colHead(c))}
                  </div>
                ) : null}
                {splitW ? (
                  <div className="sh-pin sh-pin-colheads split" style={{ left: headLeft, width: splitW, height: model.headerHeight + gutH }}>
                    {(model.columns || []).filter((c) => splitCols.has(c.index)).map((c) => colHead(c, split.leftX))}
                  </div>
                ) : null}
                {(model.columns || []).filter((c) => (split ? c.index >= (vpr?.firstCol ?? 0) : c.index >= frozen.cols)).map((c) => colHead(c))}
              </div>

              <div className="sh-rowheads" style={{ width: model.headerWidth + gutW }}>
                {gutW ? gutterNode('row') : null}
                {frozen.rows ? (
                  <div className="sh-pin sh-pin-rowheads" style={{ top: headTop, height: frozenH, width: model.headerWidth + gutW }}>
                    {gutW ? gutterNode('row', true) : null}
                    {(model.rows || []).filter((r) => r.index < frozen.rows).map((r) => rowHead(r))}
                  </div>
                ) : null}
                {splitH ? (
                  <div className="sh-pin sh-pin-rowheads split" style={{ top: headTop, height: splitH, width: model.headerWidth + gutW }}>
                    {(model.rows || []).filter((r) => splitRows.has(r.index)).map((r) => rowHead(r, split.topY))}
                  </div>
                ) : null}
                {(model.rows || []).filter((r) => (split ? r.index >= (vpr?.firstRow ?? 0) : r.index >= frozen.rows)).map((r) => rowHead(r))}
              </div>

              <div
                className="sh-cells"
                data-arrows={arrows.length}
                data-font={model.defaultFont || undefined}
                onMouseDownCapture={(e) => { if (picked.length && !e.target.closest('.sh-drawing, .sh-obj-handle, .sh-obj-rotate')) setPicked([]); }}
                style={{
                  ...(backdrop && model.viewMode !== 'pageLayout' ? { backgroundImage: `url("${backdrop.url}")`, backgroundRepeat: 'repeat', backgroundPosition: '0 0' } : {}),
                  // The workbook's default font, once it wears a theme of its own.
                  ...(model.defaultFont ? { fontFamily: `"${model.defaultFont}", ${/georgia|times|palatino|cambria|constantia|garamond/i.test(model.defaultFont) ? 'Georgia, serif' : '"Segoe UI", system-ui, sans-serif'}` } : {}),
                }}
                data-background={backdrop && model.viewMode !== 'pageLayout' ? backdrop.part : undefined}
                onMouseDown={(e) => {
                  // A press on a drawn cell is handled by the cell. Anywhere else
                  // is empty grid, and empty grid is still grid.
                  if (e.button !== 0 || e.target.closest('.sh-cell, .sh-editor, .sh-drawing, .sh-card, .sh-pl-zone')) return;
                  const at = cellAt(e);
                  if (at) dispatch({ op: 'select', row: at.row, col: at.col, extend: e.shiftKey, add: e.ctrlKey || e.metaKey });
                }}
                onDoubleClick={(e) => {
                  if (e.target.closest('.sh-cell, .sh-editor, .sh-drawing, .sh-card')) return;
                  const at = cellAt(e);
                  if (at) dispatch({ op: 'select', row: at.row, col: at.col }, { op: 'beginEdit' });
                }}
              >
                {frozen.rows ? (
                  <div className="sh-pin sh-pin-rows" style={{ top: headTop, height: frozenH }}>
                    {frozen.cols ? (
                      <div className="sh-pin sh-pin-corner" style={{ left: headLeft, width: frozenW, height: frozenH }}>
                        {model.cells.map((cell) => (pane(cell) === 'corner' ? cellNode(cell) : null))}
                        {editing && pane(editing) === 'corner' ? editorNode(0) : null}
                      </div>
                    ) : null}
                    {model.cells.map((cell) => (pane(cell) === 'rows' ? cellNode(cell) : null))}
                    {editing && pane(editing) === 'rows' ? editorNode(0) : null}
                  </div>
                ) : null}
                {frozen.cols ? (
                  <div className="sh-pin sh-pin-cols" data-oy={frozenH} style={{ left: headLeft, width: frozenW, height: Math.max(0, model.total.height - frozenH) }}>
                    {model.cells.map((cell) => (pane(cell) === 'cols' ? cellNode(cell, frozenH) : null))}
                    {editing && pane(editing) === 'cols' ? editorNode(frozenH) : null}
                  </div>
                ) : null}
                {splitH ? (
                  <div className="sh-pin sh-pin-rows split" data-oy={split.topY} style={{ top: headTop, height: splitH }}>
                    {splitW ? (
                      <div className="sh-pin sh-pin-corner split" data-ox={split.leftX} data-oy={split.topY} style={{ left: headLeft, width: splitW, height: splitH }}>
                        {model.cells.map((cell) => (splitRows.has(cell.row) && splitCols.has(cell.col) ? cellNode(cell, split.topY, split.leftX) : null))}
                        {breaksNode({ x: split.leftX, y: split.topY })}
                      </div>
                    ) : null}
                    {model.cells.map((cell) => (splitRows.has(cell.row) && cell.col >= (vpr?.firstCol ?? 0) ? cellNode(cell, split.topY) : null))}
                    {breaksNode({ x: 0, y: split.topY })}
                  </div>
                ) : null}
                {splitW ? (
                  <div className="sh-pin sh-pin-cols split" data-ox={split.leftX} data-oy={splitH} style={{ left: headLeft, width: splitW, height: Math.max(0, model.total.height - splitH) }}>
                    {model.cells.map((cell) => (splitCols.has(cell.col) && cell.row >= (vpr?.firstRow ?? 0) && cell.row <= (vpr?.lastRow ?? -1) ? cellNode(cell, splitH, split.leftX) : null))}
                    {breaksNode({ x: split.leftX, y: splitH })}
                  </div>
                ) : null}
                {(() => {
                  // The fill handle at the selection's corner, and the box a fill drag has reached.
                  // A selection in a frozen pane keeps its handle out of the way: the handle is
                  // drawn in the sliding layer, and a frozen cell is not there.
                  const active = model.cells.find((c) => c.active);
                  const source = sel?.range || (active ? { top: active.row, left: active.col, bottom: active.row, right: active.col } : null);
                  const box = source && !(source.top < frozen.rows || source.left < frozen.cols) ? boxOf(source) : null;
                  const reach = filling ? boxOf(filling.target) : null;
                  return (
                    <>
                      {box && !resizing ? (
                        <div className="sh-fill" title="Drag to fill the cells below or beside" style={{ left: box.right - 5, top: box.bottom - 5 }} onMouseDown={(e) => startFill(e, source)} />
                      ) : null}
                      {reach ? <div className="sh-fillguide" style={{ left: reach.x, top: reach.y, width: reach.right - reach.x, height: reach.bottom - reach.y }} /> : null}
                    </>
                  );
                })()}
                {resizing ? (
                  <div
                    className={`sh-guide ${resizing.kind}`}
                    style={resizing.kind === 'col' ? { left: resizing.start + resizing.size, top: 0, height: model.total.height } : { top: resizing.start + resizing.size, left: 0, width: model.total.width }}
                  />
                ) : null}
                {pageLayoutNode()}
                {model.cells.map((cell) => (inMain(cell) ? cellNode(cell) : null))}
                {arrowsNode()}
                {breaksNode()}

                {/*
                  What is drawn over the cells: the shapes, pictures and charts
                  the engine reads from the sheet's drawing part, each already
                  an SVG at its own size, anchored in the same coordinates as
                  the cells. The engine has produced these since charts were
                  built; the window never painted them, so a diagram drawn in
                  Excel opened as an empty grid with its captions.
                */}
                {drawingsNode()}

                {editing && pane(editing) === 'main' ? editorNode() : null}
                {commentCard()}
              </div>
            </div>
          </div>
          {splitBars()}
          {errorsPane()}
          {watchPane()}
          {commentsPane()}
          {review.pane ? (
            <Panel right width={300} resizable title={review.paneTitle} actions={<Button icon="close" title="Close the pane" onClick={review.close} />}>
              {review.paneNode}
            </Panel>
          ) : null}
          {selPane ? (
            <SelectionPane
              objects={model.objects || []}
              picked={picked}
              onPick={(id, add) => { setPicked((p) => (add ? (p.includes(id) ? p.filter((x) => x !== id) : [...p, id]) : [id])); shRef.current?.focus({ preventScroll: true }); }}
              onHidden={(id, hidden) => dispatch({ op: 'drawingHidden', id, hidden })}
              onRename={(id, name) => dispatch({ op: 'renameDrawing', id, name })}
              onAll={(hidden) => dispatch({ op: 'allDrawingsHidden', hidden })}
              onOrder={(to) => picked.length && dispatch({ op: 'reorderDrawings', ids: picked, to })}
              onClose={() => setSelPane(false)}
            />
          ) : null}
          </div>

          <div className={`sh-tabs${model.workbookProtection?.structure ? ' locked' : ''}`}>
            {(() => {
              // Hidden sheets have no tab; the menu's Unhide lists them. While
              // the workbook's structure is protected, everything that would
              // change the tabs is greyed, as Excel greys it.
              const hidden = new Set(model.hiddenSheets || []);
              const shown = (model.sheets || []).filter((n) => !hidden.has(n));
              const locked = Boolean(model.workbookProtection?.structure);
              const why = locked ? 'The workbook is protected — Review → Unprotect Workbook to change its sheets' : undefined;
              return shown.map((name) => {
                const at = (model.sheets || []).indexOf(name);
                const left = shown[shown.indexOf(name) - 1];
                const right = shown[shown.indexOf(name) + 1];
                return (
                  <button
                    key={name}
                    type="button"
                    className={`sh-tab${name === model.activeSheet ? ' active' : ''}`}
                    onClick={() => dispatch({ op: 'sheet', name })}
                    onContextMenu={(e) => menu.open(e, [
                      { label: 'Insert sheet', icon: 'plus', disabled: locked, title: why, run: () => act('addSheet') },
                      { label: 'Delete sheet…', icon: 'trash', disabled: locked || shown.length <= 1, title: why, run: () => { setSheetTarget(name); setDialog('deleteSheet'); } },
                      { label: 'Rename sheet…', icon: 'textbox', disabled: locked, title: why, run: () => { setSheetTarget(name); setDialog('renameSheet'); } },
                      { label: 'Move left', icon: 'chevronLeft', disabled: locked || !left, title: why, run: () => act('moveSheet', { name, to: (model.sheets || []).indexOf(left) }) },
                      { label: 'Move right', icon: 'chevronRight', disabled: locked || !right, title: why, run: () => act('moveSheet', { name, to: (model.sheets || []).indexOf(right) }) },
                      '-',
                      { label: 'Hide', icon: 'eye', disabled: locked || shown.length <= 1, title: why || (shown.length <= 1 ? 'A workbook must contain at least one visible worksheet' : undefined), run: () => act('hideSheet', name) },
                      ...(hidden.size
                        ? [...hidden].map((h) => ({ label: `Unhide "${h}"`, icon: 'eye', disabled: locked, title: why, run: () => act('unhideSheet', h) }))
                        : [{ label: 'Unhide…', icon: 'eye', disabled: true, title: 'No sheet is hidden' }]),
                    ])}
                    data-index={at}
                  >
                    {name}
                  </button>
                );
              });
            })()}
            <button type="button" className="sh-tab sh-tab-add" disabled={Boolean(model.workbookProtection?.structure)} data-tip={model.workbookProtection?.structure ? 'New sheet — the workbook is protected; Review → Unprotect Workbook first' : 'New sheet — at the end of the tabs'} onClick={() => act('addSheet')}>+</button>
            {model.workbookProtection?.structure ? <span className="sh-tabs-lock" data-tip="The workbook's structure is protected — Review → Unprotect Workbook"><Icon name="lock" size={12} /></span> : null}
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

      {review.dialogs}

      {dialog === 'goto' ? (
        <GoToDialog names={model?.names || []} onClose={() => setDialog(null)} onGo={async (ref) => { setDialog(null); await act('goto', ref); }} />
      ) : null}
      {dialog === 'link' ? (
        <LinkDialog
          current={model?.link || null}
          cellRef={sel?.ref || ''}
          onClose={() => setDialog(null)}
          onRemove={async () => { setDialog(null); await dispatch({ op: 'removeHyperlink', row: sel?.active?.row ?? 0, col: sel?.active?.col ?? 0 }); }}
          onSet={async (link) => { setDialog(null); await dispatch({ op: 'setHyperlink', row: sel?.active?.row ?? 0, col: sel?.active?.col ?? 0, ...link }); }}
        />
      ) : null}
      {dialog === 'note' ? (
        <NoteDialog
          current={model?.note || null}
          cellRef={sel?.ref || ''}
          onClose={() => setDialog(null)}
          onRemove={async () => { setDialog(null); await dispatch({ op: 'removeNote', row: sel?.active?.row ?? 0, col: sel?.active?.col ?? 0 }); }}
          onSet={async (note) => { setDialog(null); await dispatch({ op: 'setNote', row: sel?.active?.row ?? 0, col: sel?.active?.col ?? 0, ...note }); }}
        />
      ) : null}
      {dialog === 'sparklineLine' || dialog === 'sparklineColumn' ? (
        <SparklineDialog
          type={dialog === 'sparklineColumn' ? 'column' : 'line'}
          data={sel?.ref || ''}
          // The cell past the selection's last column, one per row — where a
          // person reaches for the sparkline to go once the numbers are picked.
          at={sel ? (sel.bottom > sel.top
            ? refText(sel.top, sel.right + 1) + ':' + refText(sel.bottom, sel.right + 1)
            : refText(sel.top, sel.right + 1)) : ''}
          onClose={() => setDialog(null)}
          onApply={async (spec) => { setDialog(null); await dispatch({ op: 'addSparklines', ...spec }); }}
        />
      ) : null}
      {dialog === 'renameSheet' && sheetTarget ? (
        <SheetNameDialog
          current={sheetTarget}
          onClose={() => setDialog(null)}
          onRename={async (to) => {
            setDialog(null);
            try { await dispatch({ op: 'renameSheet', from: sheetTarget, to }); } catch (err) { toast(String(err?.message || err), { tone: 'warn', ms: 5000 }); }
          }}
        />
      ) : null}
      {dialog === 'deleteSheet' && sheetTarget ? (
        <SheetDeleteDialog
          name={sheetTarget}
          onClose={() => setDialog(null)}
          onDelete={async () => {
            setDialog(null);
            try { await dispatch({ op: 'removeSheet', name: sheetTarget }); toast(`Sheet "${sheetTarget}" deleted`, { tone: 'good' }); } catch (err) { toast(String(err?.message || err), { tone: 'warn', ms: 5000 }); }
          }}
        />
      ) : null}
      {dialog === 'sort' ? (
        <SortDialog
          columns={(() => {
            // The columns to offer: the selection's when it spans more than one
            // cell, else the filled run round the active cell on its row; each
            // named by its heading when the row above reads as one.
            const r = sel && Number.isFinite(sel.left) ? sel : null;
            const active = model?.cells?.find((c) => c.active);
            const filled = (c) => model?.cells?.some((x) => x.col === c && x.text !== '');
            let left = r ? r.left : (active?.col ?? 0);
            let right = r ? r.right : (active?.col ?? 0);
            if (!r || (r.left === r.right && r.top === r.bottom)) {
              while (left > 0 && filled(left - 1)) left -= 1;
              while (filled(right + 1)) right += 1;
            }
            const top = r ? r.top : (active?.row ?? 0);
            const name = (c) => { let s = ''; let k = c + 1; while (k > 0) { const m = (k - 1) % 26; s = String.fromCharCode(65 + m) + s; k = Math.floor((k - 1) / 26); } return s; };
            const out = [];
            for (let c = left; c <= right; c++) {
              const head = model?.cells?.find((x) => x.col === c && x.row === Math.max(0, top - (r && r.top !== r.bottom ? 0 : 1)) && !x.numeric && x.text);
              out.push({ col: c, name: head ? `${name(c)} — ${head.text}` : `Column ${name(c)}` });
            }
            return out;
          })()}
          onClose={() => setDialog(null)}
          onSort={async (keys) => { setDialog(null); await dispatch({ op: 'sort', keys }); toast('Sorted', { tone: 'good' }); }}
        />
      ) : null}
      {dialog === 'outlineAxis' && outlineAsk ? (
        <OutlineAxisDialog
          verb={outlineAsk}
          onClose={() => setDialog(null)}
          onPick={async (axis) => { setDialog(null); await act(outlineAsk, axis); }}
        />
      ) : null}
      {dialog === 'subtotal' && listInfo ? (
        <SubtotalDialog
          list={listInfo}
          onClose={() => setDialog(null)}
          onApply={async (spec) => {
            setDialog(null);
            const next = await dispatch({ op: 'subtotal', ...spec });
            if (next) toast('Subtotals added — the buttons beside the row headings fold and open the groups', { tone: 'good', ms: 4000 });
          }}
          onRemoveAll={async () => {
            setDialog(null);
            const next = await dispatch({ op: 'removeSubtotals' });
            if (next) toast('Subtotals and their outline removed', { tone: 'good' });
          }}
        />
      ) : null}
      {dialog === 'advanced' && listInfo ? (
        <AdvancedFilterDialog
          list={listInfo}
          onClose={() => setDialog(null)}
          onApply={async (spec) => { setDialog(null); await act('advancedFilter', spec); }}
        />
      ) : null}
      {dialog === 'headerFooter' ? (
        <HeaderFooterDialog
          current={view.page}
          onClose={() => setDialog(null)}
          onSet={async (hf) => { setDialog(null); await act('page', hf); }}
        />
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

      {pivotAsk ? (
        <PivotDialog
          list={pivotAsk.list}
          sheet={model?.sheet || ''}
          chart={pivotAsk.chart}
          onClose={() => setPivotAsk(null)}
          onCreate={async (spec) => {
            const next = await dispatch({ op: 'pivot', ...spec, fileName: doc?.name });
            if (next) {
              setPivotAsk(null);
              toast(spec.chart ? 'PivotChart and PivotTable made — the chart follows the pivot' : 'PivotTable made under the data', { tone: 'good', ms: 3000 });
            }
          }}
        />
      ) : null}

      {gallery ? (
        <WorkbookGallery
          kind={gallery.kind}
          anchor={gallery.anchor}
          design={model?.design}
          onClose={() => setGallery(null)}
          onPick={async (spec) => {
            setGallery(null);
            await dispatch({ op: 'workbookTheme', ...spec });
          }}
          onCustomise={(kind) => { setGallery(null); setCustomise(kind); }}
        />
      ) : null}

      {customise === 'colours' ? (
        <CustomColoursDialog
          info={model?.design}
          hint="Saved into this workbook's theme: every cell, table, chart and shape that takes its colours from the theme follows, and Undo puts the old ones back."
          onClose={() => setCustomise(null)}
          onSave={async (colors, name) => { const next = await dispatch({ op: 'workbookTheme', colors, name }); if (next) setCustomise(null); }}
        />
      ) : null}
      {customise === 'fonts' ? (
        <CustomFontsDialog
          info={model?.design}
          onClose={() => setCustomise(null)}
          onSave={async (fonts, name) => { const next = await dispatch({ op: 'workbookTheme', fonts, name }); if (next) setCustomise(null); }}
        />
      ) : null}

      {slicerAsk ? (
        <InsertSlicersDialog
          source={slicerAsk}
          onClose={() => setSlicerAsk(null)}
          onInsert={async (fields) => {
            const next = await dispatch({ op: 'insertSlicers', fields });
            if (next) setSlicerAsk(null);
          }}
        />
      ) : null}

      {evaluating && doc ? (
        <EvaluateDialog
          initial={evaluating.first}
          load={(actions) => shell.doc.evaluateFormula({ id: doc.id, row: evaluating.row, col: evaluating.col, actions }).catch((err) => { toast(String(err?.message || err), { tone: 'warn' }); return null; })}
          onClose={() => setEvaluating(null)}
        />
      ) : null}

      {dialog === 'freeze' ? <FreezeDialog model={model} sel={sel} dispatch={dispatch} onClose={() => setDialog(null)} /> : null}

      {gate.node}
      {protection.node}

      {dialog === 'protectSheet' || dialog === 'protectWorkbook' ? (
        <ProtectDialog
          kind={dialog === 'protectWorkbook' ? 'workbook' : 'sheet'}
          onClose={() => setDialog(null)}
          onProtect={async ({ password }) => {
            const which = dialog;
            setDialog(null);
            const next = await dispatch({ op: which === 'protectWorkbook' ? 'protectWorkbook' : 'protect', password });
            if (next) {
              toast(which === 'protectWorkbook'
                ? `Workbook structure protected${password ? ' with a password' : ''} — no sheet can be added, deleted, renamed, moved or hidden`
                : `Sheet protected${password ? ' with a password' : ''}`, { tone: 'good', ms: 4000 });
            }
          }}
        />
      ) : null}

      {passwordAsk ? (
        <PasswordDialog
          key={passwordAsk.kind + (passwordAsk.range || '')}
          title={passwordAsk.title}
          message={passwordAsk.message}
          error={passwordAsk.error}
          onClose={() => { setPasswordAsk(null); shRef.current?.focus(); }}
          onSubmit={async (password) => {
            const ask = passwordAsk;
            const op = ask.kind === 'sheet' ? { op: 'unprotect', password }
              : ask.kind === 'workbook' ? { op: 'unprotectWorkbook', password }
              : { op: 'unlockRange', title: ask.range, password };
            const said = await tryOps(op);
            if (said) { setPasswordAsk({ ...ask, error: said }); return; }
            setPasswordAsk(null);
            shRef.current?.focus();
            if (ask.kind === 'range') {
              toast(`Range "${ask.range}" unlocked until the workbook is closed`, { tone: 'good', ms: 3200 });
              // The edit that was refused goes ahead: the key typed, or F2.
              if (ask.retry?.length) {
                const first = ask.retry.find((o) => o.op === 'beginEdit')?.initial;
                if (first) { startingRef.current = true; putDraft(first); }
                await dispatch(...ask.retry);
              }
            } else {
              toast(ask.kind === 'sheet' ? 'Sheet unprotected' : 'Workbook unprotected — sheets can be added, moved and renamed again', { tone: 'good' });
            }
          }}
        />
      ) : null}

      {dialog === 'consolidate' ? (
        <ConsolidateDialog
          info={analysis}
          onClose={() => setDialog(null)}
          onApply={async (spec) => {
            const said = await tryOps({ op: 'consolidate', ...spec });
            if (said) return said;
            setDialog(null);
            toast(spec.links ? 'Consolidated with links to the source data — the outline\'s 2 opens the detail' : 'Consolidated', { tone: 'good' });
            return null;
          }}
        />
      ) : null}

      {dialog === 'forecast' && doc ? (
        <ForecastDialog
          info={analysis}
          preview={async (spec) => {
            try {
              const next = await shell.doc.apply({ id: doc.id, ops: [{ op: 'forecastPreview', ...spec, width: 580, height: 250 }] });
              return JSON.parse(next.opResult || 'null');
            } catch (err) {
              return { svg: '', error: String(err?.message || err) };
            }
          }}
          onClose={() => setDialog(null)}
          onCreate={async (spec) => {
            const said = await tryOps({ op: 'forecastSheet', ...spec });
            if (said) return said;
            setDialog(null);
            toast('Forecast sheet made — the table and its chart, in front of the data', { tone: 'good', ms: 4000 });
            return null;
          }}
        />
      ) : null}

      {dialog === 'customViews' ? (
        <CustomViewsDialog
          views={model?.customViews || []}
          onClose={() => setDialog(null)}
          onAdd={(spec) => tryOps({ op: 'addCustomView', ...spec, zoom: view.zoom ?? 1, windowWidth: window.innerWidth, windowHeight: window.innerHeight })}
          onDelete={(name) => tryOps({ op: 'deleteCustomView', name })}
          onShow={async (name) => {
            setDialog(null);
            const next = await dispatch({ op: 'showCustomView', name });
            if (!next) return;
            let kept = null;
            try { kept = JSON.parse(next.opResult || 'null'); } catch { kept = null; }
            if (kept?.zoom) patchView({ zoom: Math.max(0.3, Math.min(3, kept.zoom)) });
            refreshPage(doc.id);
            toast(`Custom view "${name}" shown`, { tone: 'good' });
          }}
        />
      ) : null}

      {dialog === 'editRanges' ? (
        <EditRangesDialog
          ranges={model?.editRanges || []}
          sheetProtected={Boolean(model?.protection?.sheet)}
          selection={sel?.ref || ''}
          onClose={() => setDialog(null)}
          onSave={(spec) => tryOps({ op: 'setEditRange', ...spec })}
          onDelete={(title) => tryOps({ op: 'deleteEditRange', title })}
          onProtectSheet={() => setDialog('protectSheet')}
        />
      ) : null}
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
  display: flex; align-items: center; gap: 8px; padding: 5px 10px;
  background: var(--chrome); border-bottom: 1px solid var(--line-soft);
}
.sh-namebox {
  min-width: 82px; height: 26px; padding: 0 10px; font-size: 12px; font-weight: 600; font-variant-numeric: tabular-nums;
  border: 1px solid var(--line-soft); border-radius: var(--r-2); background: var(--surface); text-align: center;
  display: inline-flex; align-items: center; justify-content: center; color: var(--ink-2);
}
.sh-formula .rw-input {
  font-family: var(--mono); font-size: 12.5px; height: 26px; padding: 0 10px;
  border: 1px solid var(--line-soft); background: var(--surface);
}
.sh-formula .rw-input:focus { box-shadow: 0 0 0 3px var(--accent-soft); border-color: var(--accent); }

.sh-body { flex: 1; display: flex; position: relative; min-width: 0; min-height: 0; }
.sh-grid { flex: 1; overflow: auto; position: relative; min-width: 0; min-height: 0; background: var(--surface); }

/* Formulas → Error Checking: a floating pane over the grid, under the ribbon
   — there is no other right-hand pane in this app to match. */
.sh-errors {
  position: absolute; top: 10px; right: 10px; bottom: 10px; width: 320px; z-index: 15;
  display: flex; flex-direction: column; background: var(--surface); border: 1px solid var(--line);
  border-radius: var(--r-2); box-shadow: 0 8px 24px color-mix(in srgb, var(--ink) 20%, transparent);
  overflow: hidden;
}
.sh-errors-head {
  display: flex; align-items: center; gap: 4px; padding: 8px 10px;
  border-bottom: 1px solid var(--line-soft); font-size: 12.5px;
}
.sh-errors-list { overflow: auto; }
.sh-error { padding: 8px 10px; border-bottom: 1px solid var(--line-soft); cursor: pointer; font-size: 12px; }
.sh-error:hover { background: var(--selected); }
.sh-error-line { display: flex; align-items: baseline; gap: 6px; }
.sh-error-sheet { color: var(--ink-3); font-size: 11px; }
.sh-error-ref { font-weight: 600; font-variant-numeric: tabular-nums; }
.sh-error-value { color: var(--bad); font-weight: 600; }
.sh-error-formula { font-family: var(--mono); color: var(--ink-2); margin-top: 2px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sh-error-reason { color: var(--ink-3); margin-top: 2px; }

/* Formulas → Watch Window: docked along the bottom of the grid rather than
   the right, the way Excel's own Watch Window sits — a table, not a list,
   fits the four columns better and leaves the error pane's corner free for
   both to be open together. */
.sh-watch {
  position: absolute; left: 10px; right: 10px; bottom: 10px; max-height: 45%; z-index: 15;
  display: flex; flex-direction: column; background: var(--surface); border: 1px solid var(--line);
  border-radius: var(--r-2); box-shadow: 0 8px 24px color-mix(in srgb, var(--ink) 20%, transparent);
  overflow: hidden;
}
.sh-watch-head {
  display: flex; align-items: center; gap: 4px; padding: 8px 10px;
  border-bottom: 1px solid var(--line-soft); font-size: 12.5px;
}
.sh-watch-list { overflow: auto; }
.sh-watch-row {
  display: grid; grid-template-columns: 120px 70px 140px 1fr 26px; gap: 10px; align-items: center;
  padding: 6px 10px; border-bottom: 1px solid var(--line-soft); font-size: 12px; cursor: pointer;
}
.sh-watch-row:hover { background: var(--selected); }
.sh-watch-row.selected { background: var(--selected); }
.sh-watch-cols { font-weight: 600; color: var(--ink-3); cursor: default; }
.sh-watch-cols:hover { background: none; }
.sh-watch-cell { font-weight: 600; font-variant-numeric: tabular-nums; }
.sh-watch-value { font-variant-numeric: tabular-nums; }
.sh-watch-formula { font-family: var(--mono); color: var(--ink-2); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sh-watch-x { justify-self: end; }
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
/* Frozen panes: layers in normal flow that stick — the rows under the
   column headings, the columns beside the row headings, the corner at
   both — with their cells drawn at their own place inside the layer. */
/* The order of the layers, bottom up: the sliding cells, the frozen columns,
   the frozen rows (with the corner inside them), then the headings and
   their corner over everything — the headings come first in the tree, so
   they need the numbers to win. */
.sh-corner { z-index: 8; }
.sh-colheads, .sh-rowheads { z-index: 7; background: var(--surface-2); }
.sh-pin { position: sticky; background: var(--surface); box-sizing: border-box; }
.sh-pin-cols { z-index: 2; box-shadow: 1px 0 0 var(--line), 2px 0 5px color-mix(in srgb, var(--ink) 9%, transparent); }
.sh-pin-rows { z-index: 3; box-shadow: 0 1px 0 var(--line), 0 2px 5px color-mix(in srgb, var(--ink) 9%, transparent); }
.sh-pin-corner { z-index: 4; box-shadow: 1px 0 0 var(--line); }
.sh-pin-rowheads, .sh-pin-colheads { z-index: 2; background: var(--surface-2); }
.sh-pin-rowheads { box-shadow: 0 1px 0 var(--line); }
.sh-pin-colheads { box-shadow: 1px 0 0 var(--line); }
.sh-head {
  position: absolute; display: grid; place-items: center; font-size: 11.5px; color: var(--ink-2);
  background: var(--surface-2); border-right: 1px solid var(--line-soft); border-bottom: 1px solid var(--line);
  font-variant-numeric: tabular-nums; user-select: none;
  transition: background var(--fast), color var(--fast);
}
.sh-head.active { background: var(--selected); color: var(--accent); font-weight: 600; }
/* The heading of the active column carries a bar along its edge, the row's likewise. */
.sh-colheads .sh-head.active { box-shadow: inset 0 -2px 0 var(--accent); }
.sh-rowheads .sh-head.active { box-shadow: inset -2px 0 0 var(--accent); }
/* The outline gutter, drawn as Excel draws it: a lane per level beside the
   row headings (above the column headings), a bracket along each open group
   ending in its − box at the summary, a + box where a group is folded, and
   the level buttons 1 2 3 in the corner. Hairlines on the pixel grid. */
.sh-gutter { position: absolute; left: 0; top: 0; pointer-events: none; box-sizing: border-box; }
.sh-gutter.rows { border-right: 1px solid var(--line-soft); }
.sh-gutter.cols { border-bottom: 1px solid var(--line-soft); }
.sh-pin > .sh-gutter.rows { height: 100% !important; overflow: hidden; }
.sh-pin > .sh-gutter.cols { width: 100% !important; overflow: hidden; }
.sh-gutter-lines { position: absolute; left: 0; top: 0; overflow: visible; }
.sh-gutter-lines polyline { stroke: var(--ink-3); stroke-width: 1; shape-rendering: crispEdges; }
.sh-ol-box {
  position: absolute; width: 11px; height: 11px; box-sizing: border-box; display: grid; place-items: center;
  border: 1px solid var(--ink-3); border-radius: 1px; background: var(--surface); cursor: pointer; pointer-events: auto;
}
.sh-ol-box svg { display: block; }
.sh-ol-box path { stroke: var(--ink); stroke-width: 1.3; shape-rendering: crispEdges; }
.sh-ol-box:hover { border-color: var(--accent); background: var(--selected); }
.sh-ol-box:hover path { stroke: var(--accent); }
.sh-corner.outlined { background: var(--surface-2); }
.sh-ol-level {
  position: absolute; width: 14px; height: 14px; padding: 0; box-sizing: border-box;
  border: 1px solid var(--ink-3); border-radius: 2px; background: var(--surface); color: var(--ink-2);
  font-size: 9.5px; font-weight: 600; line-height: 12px; text-align: center; cursor: pointer;
  font-variant-numeric: tabular-nums;
}
.sh-ol-level:hover { border-color: var(--accent); color: var(--accent); background: var(--selected); }
.sh-sub-cols { display: flex; flex-direction: column; gap: 5px; max-height: 132px; overflow: auto; padding: 6px 8px; border: 1px solid var(--line-soft); border-radius: var(--r-2); }
/* The edge of a heading is a handle: drag it and the column or row follows. */
.sh-grip { position: absolute; z-index: 4; }
.sh-grip.col { top: 0; bottom: 0; right: -4px; width: 8px; cursor: col-resize; }
.sh-grip.row { left: 0; right: 0; bottom: -4px; height: 8px; cursor: row-resize; }
.sh-grip:hover { background: color-mix(in srgb, var(--accent) 35%, transparent); }
.sh-guide { position: absolute; z-index: 6; pointer-events: none; background: var(--accent); }
/* The fill handle and the box a fill drag has reached. */
.sh-fill { position: absolute; width: 9px; height: 9px; background: var(--accent); border: 1.5px solid #fff; border-radius: 1px; z-index: 5; cursor: crosshair; box-sizing: border-box; }
.sh-fillguide { position: absolute; border: 1.5px dashed var(--accent); pointer-events: none; z-index: 5; box-sizing: border-box; }
.sh-guide.col { width: 2px; margin-left: -1px; }
.sh-guide.row { height: 2px; margin-top: -1px; }
.sh-cell {
  position: absolute; display: flex; align-items: center; padding: 0 5px;
  border-right: 1px solid var(--line-soft); border-bottom: 1px solid var(--line-soft);
  font-size: 12.5px; overflow: hidden; white-space: nowrap; background: var(--surface);
}
.sh-cell.sel { background: var(--selected); }
/* A link: the accent, underlined, a hand; Ctrl+click follows it. */
.sh-cell.link { color: var(--accent); text-decoration: underline; text-decoration-color: color-mix(in srgb, var(--accent) 55%, transparent); cursor: pointer; }
/* A note: Excel's red corner, and the note itself on hover through the tip layer. */
.sh-cell.noted::after { content: ''; position: absolute; top: 0; right: 0; border: 4px solid transparent; border-top-color: #d0362f; border-right-color: #d0362f; }
/* View → Page Break Preview: the unprinted greyed, the pages edged in blue,
   a break by hand solid and one by the paper dashed, the page numbers
   faint across each page. The breaks are handles. */
.sh-pb { position: absolute; left: 0; top: 0; width: 0; height: 0; z-index: 3; }
.sh-pb-shade { position: absolute; background: color-mix(in srgb, #6b6f7a 30%, transparent); pointer-events: none; }
.sh-pb-area { position: absolute; box-sizing: border-box; border: 4px solid #2152c8; pointer-events: none; }
.sh-pb-page { position: absolute; display: grid; place-items: center; pointer-events: none; overflow: hidden; }
/* Over the cells, but multiplied in, so it reads as lying under their words as Excel's does. */
.sh-pb-page span { color: #c9ccd3; mix-blend-mode: multiply; font-weight: 700; letter-spacing: 0.02em; white-space: nowrap; user-select: none; }
.sh-pb-break { position: absolute; box-sizing: border-box; }
.sh-pb-break.row { height: 9px; cursor: row-resize; }
.sh-pb-break.col { width: 9px; cursor: col-resize; }
.sh-pb-break.row::after { content: ''; position: absolute; left: 0; right: 0; top: 2px; border-top: 4px dashed #2152c8; }
.sh-pb-break.col::after { content: ''; position: absolute; top: 0; bottom: 0; left: 2px; border-left: 4px dashed #2152c8; }
.sh-pb-break.manual.row::after { border-top-style: solid; }
.sh-pb-break.manual.col::after { border-left-style: solid; }
.sh-pb-break:hover::after { border-color: var(--accent); }
.sh-pb-guide { position: absolute; background: var(--accent); pointer-events: none; z-index: 4; }
.sh-pb-guide.row { height: 3px; }
.sh-pb-guide.col { width: 3px; }
/* View → Split: the bars where the panes meet; the panes themselves are the
   pinned layers, with no frozen shadow of their own. */
/* clip, not hidden: a hidden overflow is a scroll container, and the corner's own stickiness would then be measured from its pane. */
.sh-pin.split { box-shadow: none; overflow: clip; }
.sh-pb.mirror .sh-pb-break { pointer-events: none; }
/* The split panes cover the main pane's preview marks where they lie over it. */
.sh-pin-cols.split { z-index: 5; }
.sh-pin-rows.split { z-index: 6; }
.sh-splitbar { position: absolute; z-index: 12; background: color-mix(in srgb, var(--ink) 24%, var(--surface)); }
.sh-splitbar.h { height: 5px; cursor: row-resize; border-top: 1px solid var(--line); border-bottom: 1px solid var(--line); }
.sh-splitbar.v { width: 5px; cursor: col-resize; border-left: 1px solid var(--line); border-right: 1px solid var(--line); }
.sh-splitbar:hover, .sh-splitbar.dragging { background: var(--accent); border-color: var(--accent); }
/* The status bar's view buttons, beside the zoom as Excel has them. */
.sh-viewbtns { display: inline-flex; gap: 2px; margin-right: 4px; }
.sh-viewbtn { width: 24px; height: 22px; display: grid; place-items: center; border: 1px solid transparent; border-radius: var(--r-2); background: transparent; color: var(--ink-2); cursor: pointer; padding: 0; }
.sh-viewbtn svg { fill: none; stroke: currentColor; stroke-width: 1.1; }
.sh-viewbtn:hover { background: var(--hover); }
.sh-viewbtn.on { background: var(--selected); color: var(--accent); border-color: color-mix(in srgb, var(--accent) 35%, var(--line)); }
/* A comment thread: Excel's purple corner — grey once the thread is resolved —
   and the corner itself is what a click opens the thread from. */
.sh-cell.threaded::before { content: ''; position: absolute; top: 0; right: 0; border: 5px solid transparent; border-top-color: #7a3db8; border-right-color: #7a3db8; pointer-events: none; }
.sh-cell.threaded.resolved::before { border-top-color: #9a93a5; border-right-color: #9a93a5; }
.sh-thread-mark { position: absolute; top: 0; right: 0; width: 12px; height: 12px; cursor: pointer; z-index: 1; }
/* The comment card: beside its cell, over the grid, scrolling with it. */
.sh-card {
  position: absolute; z-index: 9; width: 300px; box-sizing: border-box; cursor: default;
  background: var(--surface); border: 1px solid var(--line); border-radius: var(--r-3, 8px);
  box-shadow: 0 10px 28px color-mix(in srgb, var(--ink) 22%, transparent); font-size: 12.5px; color: var(--ink);
  white-space: normal;
}
.sh-card::before { content: ''; position: absolute; left: -6px; top: 12px; width: 10px; height: 10px; background: var(--surface); border-left: 1px solid var(--line); border-bottom: 1px solid var(--line); transform: rotate(45deg); }
.sh-card-resolved { display: flex; align-items: center; gap: 6px; padding: 7px 12px; border-bottom: 1px solid var(--line-soft); background: var(--surface-2); color: var(--ink-2); font-weight: 600; font-size: 11.5px; border-radius: var(--r-3, 8px) var(--r-3, 8px) 0 0; }
.sh-card-resolved .grow { flex: 1; }
.sh-card-list { max-height: 320px; overflow: auto; }
.sh-comment { padding: 10px 12px 8px; border-bottom: 1px solid var(--line-soft); position: relative; }
.sh-comment-head { display: flex; align-items: center; gap: 8px; }
.sh-avatar { width: 26px; height: 26px; border-radius: 50%; display: grid; place-items: center; color: #fff; font-size: 10.5px; font-weight: 700; flex: none; letter-spacing: 0.02em; }
.sh-comment-who { flex: 1; min-width: 0; line-height: 1.25; }
.sh-comment-name { font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sh-comment-when { font-size: 11px; color: var(--ink-3); }
.sh-comment-text { margin: 6px 0 0 34px; line-height: 1.45; white-space: pre-wrap; word-break: break-word; }
.sh-comment-tools { display: flex; gap: 2px; }
.sh-comment-tools .rw-btn { min-width: 24px; height: 24px; padding: 0 4px; }
.sh-card-compose { padding: 10px 12px 12px; display: flex; flex-direction: column; gap: 8px; }
.sh-card-compose textarea, .sh-comment textarea {
  font: inherit; font-size: 12.5px; width: 100%; box-sizing: border-box; resize: vertical; min-height: 56px;
  border: 1px solid var(--line); border-radius: var(--r-2); padding: 6px 8px; background: var(--surface); color: var(--ink); outline: none;
}
.sh-card-compose textarea:focus, .sh-comment textarea:focus { border-color: var(--accent); box-shadow: 0 0 0 3px var(--accent-soft); }
.sh-card-actions { display: flex; justify-content: flex-end; gap: 6px; }
.sh-card-new-who { display: flex; align-items: center; gap: 8px; font-weight: 600; }
/* Review → Show Comments: every thread in the workbook, beside the grid. */
.sh-comments {
  position: absolute; top: 10px; right: 10px; bottom: 10px; width: 330px; z-index: 15;
  display: flex; flex-direction: column; background: var(--surface); border: 1px solid var(--line);
  border-radius: var(--r-2); box-shadow: 0 8px 24px color-mix(in srgb, var(--ink) 20%, transparent); overflow: hidden;
}
.sh-comments-head { display: flex; align-items: center; gap: 4px; padding: 8px 10px; border-bottom: 1px solid var(--line-soft); font-size: 12.5px; }
.sh-comments-filter { display: flex; gap: 4px; padding: 8px 10px; border-bottom: 1px solid var(--line-soft); }
.sh-comments-filter button { border: 1px solid var(--line-soft); background: var(--surface); color: var(--ink-2); font: inherit; font-size: 11.5px; padding: 3px 10px; border-radius: 999px; cursor: pointer; }
.sh-comments-filter button.on { background: var(--selected); color: var(--accent); border-color: color-mix(in srgb, var(--accent) 40%, var(--line)); font-weight: 600; }
.sh-comments-list { overflow: auto; flex: 1; }
.sh-comments-item { padding: 9px 12px; border-bottom: 1px solid var(--line-soft); cursor: pointer; }
.sh-comments-item:hover, .sh-comments-item.here { background: var(--selected); }
.sh-comments-item .where { display: flex; align-items: center; gap: 6px; font-size: 11px; color: var(--ink-3); margin-bottom: 3px; }
.sh-comments-item .where .ref { font-weight: 700; color: var(--ink-2); font-variant-numeric: tabular-nums; }
.sh-comments-item .where .done { margin-left: auto; color: var(--ink-3); font-weight: 600; }
.sh-comments-item .first { display: flex; gap: 6px; align-items: baseline; }
.sh-comments-item .first b { flex: none; }
.sh-comments-item .first span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--ink-2); }
.sh-comments-item .more { font-size: 11px; color: var(--accent); margin-top: 2px; }
.sh-comments-item.resolved .first span, .sh-comments-item.resolved .first b { color: var(--ink-3); }
.sh-comments-empty { padding: 18px 12px; color: var(--ink-3); font-size: 12px; text-align: center; }
/* A sparkline: drawn under the cell's own text, which is normally empty. */
.sh-spark { position: absolute; left: 0; top: 0; pointer-events: none; }
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
  display: flex; align-items: stretch; gap: 3px; padding: 4px 8px;
  background: var(--chrome); border-top: 1px solid var(--line-soft); overflow-x: auto;
  scrollbar-width: none;
}
.sh-tabs::-webkit-scrollbar { display: none; }
.sh-tab {
  border: 0; background: transparent; color: var(--ink-2); padding: 4px 13px;
  font-size: 12px; font-weight: 500; border-radius: var(--r-2); white-space: nowrap;
  transition: background var(--fast), color var(--fast);
}
.sh-tab:hover { background: var(--hover); }
.sh-tab.active { background: var(--surface); color: var(--accent); font-weight: 600; box-shadow: var(--shadow-1); }
/* The status bar's "Calculate": formulas waiting in manual mode; a press calculates. */
.sh-calc-pending {
  border: 1px solid color-mix(in srgb, var(--accent) 45%, var(--line)); background: var(--surface); color: var(--accent);
  font: inherit; font-size: 11.5px; font-weight: 600; padding: 1px 10px; border-radius: 999px; cursor: pointer;
}
.sh-calc-pending:hover { background: var(--selected); }
/* Formulas → Evaluate Formula: reference beside evaluation, a box per level. */
.sh-eval { display: flex; flex-direction: column; gap: 6px; }
.sh-eval-head, .sh-eval-level { display: grid; grid-template-columns: 150px 14px 1fr; gap: 8px; align-items: start; }
.sh-eval-head { font-size: 11.5px; font-weight: 600; color: var(--ink-3); }
.sh-eval-head span:last-child { grid-column: 3; }
.sh-eval-ref { font-size: 12px; font-variant-numeric: tabular-nums; padding-top: 7px; color: var(--ink-2); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sh-eval-eq { padding-top: 6px; color: var(--ink-3); text-align: center; }
.sh-eval-text {
  font-family: var(--mono); font-size: 12.5px; line-height: 1.55; min-height: 56px; max-height: 150px; overflow: auto;
  padding: 6px 9px; border: 1px solid var(--line-soft); border-radius: var(--r-2); background: var(--surface-2);
  white-space: pre-wrap; word-break: break-all; color: var(--ink-2);
}
.sh-eval-level.current .sh-eval-text { background: var(--surface); border-color: var(--line); color: var(--ink); }
.sh-eval-next { text-decoration: underline; text-decoration-thickness: 1.5px; text-underline-offset: 3px; text-decoration-color: var(--accent); background: var(--selected); border-radius: 2px; }
.sh-eval-recent { font-style: italic; color: var(--accent); }
.sh-eval-message { margin: 6px 0 0; font-size: 12px; color: var(--ink-3); }
/* The + at the end of the tabs: a new sheet, as every spreadsheet has it. */
.sh-tab-add { min-width: 28px; font-weight: 600; color: var(--ink-2); }
/* View → Page Layout: the paper on a desk. The cells sit on the pages the
   engine put them on; the pages sit under them, each with a soft shadow. */
.sh.pl { --pl-desk: color-mix(in srgb, var(--ink) 11%, var(--surface)); }
.sh.pl .sh-grid { background: var(--pl-desk); }
.sh.pl .sh-cells { background: transparent; }
.sh.pl .sh-colheads, .sh.pl .sh-rowheads, .sh.pl .sh-corner { background: var(--pl-desk); }
.sh.pl .sh-head { background: color-mix(in srgb, var(--surface-2) 80%, var(--pl-desk)); }
.sh-pl { position: absolute; left: 0; top: 0; width: 0; height: 0; z-index: 0; }
.sh-pl-page { position: absolute; background: var(--surface); box-shadow: 0 1px 2px color-mix(in srgb, #000 16%, transparent), 0 6px 18px color-mix(in srgb, #000 10%, transparent); border-radius: 1px; }
.sh-pl-page.blank { background: color-mix(in srgb, var(--surface) 92%, var(--pl-desk)); }
/* The printable box: a hairline where the margins end, as Excel shows it on hover. */
.sh-pl-box { position: absolute; pointer-events: none; outline: 1px dashed transparent; }
.sh-pl-page:hover .sh-pl-box { outline-color: color-mix(in srgb, var(--ink) 14%, transparent); }
.sh-pl-blank { position: absolute; display: grid; place-items: center; pointer-events: none; }
.sh-pl-blank span { font-size: 15px; color: color-mix(in srgb, var(--ink) 30%, transparent); letter-spacing: 0.01em; }
.sh-pl-zone {
  position: absolute; z-index: 2; display: grid; grid-template-columns: 1fr 1fr 1fr; align-items: center; gap: 6px;
  box-sizing: border-box; padding: 0 6px; border: 1px dashed transparent; border-radius: 2px; cursor: text;
  font: 9pt Calibri, Arial, sans-serif; color: #444;
}
.sh-pl-zone .l { text-align: left; } .sh-pl-zone .c { text-align: center; } .sh-pl-zone .r { text-align: right; }
.sh-pl-zone span { overflow: hidden; white-space: nowrap; text-overflow: ellipsis; }
.sh-pl-zone .ask { grid-column: 1 / 4; color: color-mix(in srgb, var(--ink) 34%, transparent); font-size: 12px; transition: color var(--fast); }
.sh-pl-zone:hover { border-color: color-mix(in srgb, var(--accent) 55%, transparent); background: color-mix(in srgb, var(--accent) 6%, transparent); }
.sh-pl-zone:hover .ask { color: var(--accent); }
.sh-pl-ruler { position: absolute; overflow: visible; pointer-events: none; }
.sh-pl-ruler rect.margin { fill: color-mix(in srgb, var(--ink) 9%, transparent); }
.sh-pl-ruler line { stroke: color-mix(in srgb, var(--ink) 42%, transparent); stroke-width: 1; shape-rendering: crispEdges; }
.sh-pl-ruler text { font: 8.5px Calibri, Arial, sans-serif; fill: color-mix(in srgb, var(--ink) 58%, transparent); }
.sh-pl-ruler.x { border-bottom: 1px solid color-mix(in srgb, var(--ink) 18%, transparent); background: var(--surface); }
.sh-pl-ruler.y { border-right: 1px solid color-mix(in srgb, var(--ink) 18%, transparent); background: var(--surface); }
/* Page Layout → Background: the picture tiled behind the cells; a cell with
   no fill of its own lets it through, as Excel's do. */
.sh.has-bg .sh-cell { background-color: transparent; }
.sh.has-bg .sh-cell.sel { background-color: color-mix(in srgb, var(--selected) 70%, transparent); }
.sh-tab-add:disabled { color: var(--ink-4, #b0b4bc); cursor: default; background: transparent; }
/* Review → Protect Workbook on: a padlock at the end of the tabs, which cannot change. */
.sh-tabs-lock { display: inline-grid; place-items: center; padding: 0 6px; color: var(--ink-3); }
/* The protection dialogs. */
.sh-protect-note { margin: 0; font-size: 12.5px; line-height: 1.5; color: var(--ink-2); }
.sh-protect-warn { margin: 0; font-size: 12px; color: var(--bad); }
.sh-ranges { display: flex; flex-direction: column; gap: 10px; }
.sh-ranges-body { display: grid; grid-template-columns: 1fr auto; gap: 12px; align-items: start; }
.sh-ranges-list { border: 1px solid var(--line); border-radius: var(--r-2); background: var(--surface); min-height: 132px; max-height: 220px; overflow: auto; display: flex; flex-direction: column; }
.sh-ranges-head, .sh-ranges-row { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; padding: 5px 10px; font-size: 12.5px; text-align: left; }
.sh-ranges-head { position: sticky; top: 0; background: var(--surface-2); border-bottom: 1px solid var(--line-soft); font-size: 11.5px; font-weight: 600; color: var(--ink-3); }
.sh-ranges-row { border: 0; background: transparent; color: var(--ink); font: inherit; font-size: 12.5px; cursor: default; }
.sh-ranges-row:hover { background: var(--hover); }
.sh-ranges-row.on { background: var(--selected); color: var(--accent); }
.sh-ranges-row .t { display: inline-flex; align-items: center; gap: 5px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sh-ranges-row .c { font-variant-numeric: tabular-nums; color: var(--ink-2); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sh-ranges-empty { padding: 26px 12px; font-size: 12px; color: var(--ink-3); text-align: center; }
.sh-ranges-buttons { display: flex; flex-direction: column; gap: 6px; min-width: 96px; }
.sh-ranges-buttons .rw-btn { justify-content: center; min-height: 28px; }
.sh-ranges-buttons .rw-btn:not(.primary) { border: 1px solid var(--line); background: var(--surface); }
.sh-ranges-buttons .rw-btn:not(.primary):hover:not(:disabled) { background: var(--hover); }
.sh-ranges-row.one { grid-template-columns: 1fr; }
.sh-cons-add { border: 1px solid var(--line); background: var(--surface); min-width: 64px; justify-content: center; }
.sh-cons-options { display: flex; justify-content: space-between; align-items: flex-end; gap: 12px; }
.sh-cons-options > div { display: flex; flex-direction: column; gap: 6px; }
.sh-cons-head { font-size: 11.5px; font-weight: 600; color: var(--ink-3); margin-top: 2px; }
/* Data → Forecast Sheet: the chart kinds, the preview, the options in two columns. */
.sh-fc { display: flex; flex-direction: column; gap: 10px; }
.sh-fc-kinds { display: flex; gap: 6px; }
.sh-fc-kind { width: 40px; height: 30px; display: grid; place-items: center; border: 1px solid var(--line); border-radius: var(--r-2); background: var(--surface); color: var(--ink-2); cursor: pointer; }
.sh-fc-kind svg { fill: none; stroke: currentColor; stroke-width: 2; stroke-linecap: round; stroke-linejoin: round; }
.sh-fc-kind.on { border-color: var(--accent); background: var(--selected); color: var(--accent); }
.sh-fc-preview { height: 250px; border: 1px solid var(--line-soft); border-radius: var(--r-2); background: var(--surface); display: grid; place-items: center; overflow: hidden; }
.sh-fc-preview > div, .sh-fc-preview svg { width: 100%; height: 100%; display: block; }
.sh-fc-preview p { margin: 0; padding: 0 24px; font-size: 12.5px; color: var(--ink-3); text-align: center; }
.sh-fc-row { display: flex; align-items: flex-end; justify-content: space-between; gap: 12px; }
.sh-fc-row .rw-field { flex: 0 0 200px; }
.sh-fc-toggle { border: 0; background: transparent; color: var(--accent); font: inherit; font-size: 12.5px; cursor: pointer; padding: 4px 2px; }
.sh-fc-options { display: grid; grid-template-columns: 1fr 1fr; gap: 12px 20px; padding-top: 8px; border-top: 1px solid var(--line-soft); }
.sh-fc-col { display: flex; flex-direction: column; gap: 8px; }
.sh-ranges-locked { margin: 0; display: flex; align-items: center; gap: 6px; font-size: 12px; color: var(--ink-3); }
`;
