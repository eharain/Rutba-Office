// Rutba PDF: read a PDF, and put its pages in order.
//
// The window keeps a list of pages, each naming the file it comes from, its
// page there and its turn (pages.js); rotating, deleting, moving, inserting
// from another file and merging are new lists, so undo and redo are the
// lists before and after. pdf.js draws the pages (pdfjs.js), with their
// words over them to select, copy and find. Save, Extract and Split hand
// the list to the main process (main/pdf.js), which writes a new PDF from
// the pages named, each copied byte for byte, nothing drawn again.
//
// Recognise Text (ocr.js) finds the words in scanned pages and has the
// main process append them to the file as an invisible text layer; the
// file with its layer is a new source, and the pages are pointed at it as
// one more list in the history, so Undo takes the layer away again.
//
// Comments (annots.js) are written into the file as they are made: each
// highlight, note, drawing, reply, resolution or deletion is an incremental
// update the main process appends (@rutba/pdf/annotations), the file with
// it a new source the pages are pointed at, so Undo takes it away as it
// takes a recognised text layer away, and pdf.js draws the comment from the
// appearance written, as any other reader will.
//
// A form's fields are filled the same way (fields.js): the controls over
// the page commit each value as an update with the appearances other
// readers show, and Fill & Sign places a typed or drawn signature as a
// stamp, which the window says plainly is a picture, not a digital
// signature.
//
// Redaction (redact.js) marks words, areas or every match of a search with
// Redact annotations, written like comments, then applies them: the main
// process removes what the marks cover from each page's content and writes
// the window's pages whole as a new file, which the window opens.
//
// Saving is careful, as a page tool's must be: Save asks before it replaces
// the file the pages came from (once a window), and offers to save a copy.
// While there is unsaved work, a copy is kept in the recovery folder for
// the launcher to offer back after a crash, as Video and Image do.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  Ribbon, Group, Rows, Button, Spacer, Chip, Empty, Spinner, Dialog, Input, ZoomSlider, useToast, useCommands, useMenu, useWheelZoom,
  basename, dirname, formatBytes, t, tn,
} from '@rutba/office-ui';
import { findAll } from '@rutba/editing/find';
import { AppFrame, useAppMenu, useDirtyGuard, useFileDrop } from '../../shell.js';
import { openDocument, pageText, pageImage } from './pdfjs.js';
import {
  pagesOf, rotatePages, deletePages, insertPages, movePages, rangeKeys, rangeText, chosenKeys, shownSize, splitGroups,
} from './pages.js';
import { ThumbStrip, PageView, PX_PER_PT } from './view.js';
import { CSS } from './css.js';
import { RecogniseDialog, recognisePages } from './ocr.js';
import { OCR_CSS } from '../../ocr-ui.js';
import {
  AnnotLayer, CommentsPane, NoteEditor, commentItems, selectionQuads, swatch, cssColour,
  MARKUP_TOOLS, TOOL_TYPES, SWATCHES, DEFAULT_TOOL_COLOURS,
} from './annots.js';
import { FieldLayer, SignatureDialog, savedSignature, signatureStamp } from './fields.js';
import { searchQuads, decodePlan, ApplyRedactionsDialog } from './redact.js';
import { BookmarksPane, labelOf } from './nav.js';
import { ExportPicturesDialog, CompressDialog, pageBytes, encodePlan } from './export.js';

const AUTOSAVE_FIRST_MS = 1500;
const AUTOSAVE_EVERY_MS = 30000;
const ZOOM_MIN = 0.1;
const ZOOM_MAX = 6.4;
const PDF_FILTER = () => [{ name: t('PDF Document (.pdf)'), extensions: ['pdf'] }];

const newCopyKey = () => `p${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;
const stemOf = (p) => basename(p || '').replace(/\.pdf$/i, '');
const joinPath = (dir, name) => (dir ? `${dir}${dir.includes('\\') ? '\\' : '/'}${name}` : name);

/** Split: files of so many pages, or a new file at each picked page. */
function SplitDialog({ count, picked, onSplit, onClose }) {
  const [mode, setMode] = useState(picked > 0 ? 'picked' : 'every');
  const [every, setEvery] = useState(1);
  const files = mode === 'every' ? Math.ceil(count / Math.max(1, every || 1)) : picked + 1;
  return (
    <Dialog
      title={t('Split')}
      onClose={onClose}
      actions={(
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('Split')} data-role="split-go" disabled={files < 2} onClick={() => onSplit(mode === 'every' ? { every: Math.max(1, every || 1) } : { picked: true })} />
        </>
      )}
    >
      <div className="pdf-split-form">
        <label className="pdf-split-choice">
          <input type="radio" name="pdf-split" checked={mode === 'every'} onChange={() => setMode('every')} />
          <span>{t('Into files of')}</span>
          <Input type="number" min={1} max={count} value={every} data-role="split-every" onChange={(e) => { setMode('every'); setEvery(Math.max(1, Number(e.target.value) || 1)); }} />
          <span>{t('pages each')}</span>
        </label>
        <label className="pdf-split-choice">
          <input type="radio" name="pdf-split" checked={mode === 'picked'} disabled={!picked} onChange={() => setMode('picked')} />
          <span>{t('A new file at each selected page')}</span>
        </label>
        <div className="pdf-split-note">{tn(files, 'This makes {count} file, in a folder you choose. This document is not changed.', 'This makes {count} files, in a folder you choose. This document is not changed.')}</div>
      </div>
    </Dialog>
  );
}

export default function PdfApp({ app, shell, boot }) {
  const toast = useToast();
  const menu = useMenu();
  /** id -> { id, name, path, info, doc }: every file the pages come from. */
  const [sources, setSources] = useState({});
  const sourcesRef = useRef(sources);
  sourcesRef.current = sources;
  const [pages, setPages] = useState([]);
  const pagesRef = useRef(pages);
  pagesRef.current = pages;
  const history = useRef({ undo: [], redo: [] });
  const [, setHistoryTick] = useState(0);
  const [selected, setSelected] = useState([]);
  const selectedRef = useRef(selected);
  selectedRef.current = selected;
  const anchor = useRef(null);
  const [current, setCurrent] = useState(null);
  const currentRef = useRef(current);
  currentRef.current = current;
  const [path, setPath] = useState(null);
  const [name, setName] = useState(null);
  const [dirty, setDirty] = useState(false);
  // Save over the file the pages came from: asked once a window.
  const confirmed = useRef(null);
  const [tab, setTab] = useState('home');
  const [zoom, setZoom] = useState(1);
  // How the zoom follows the window: 'auto' (as wide as the window, never past 100%), 'width', 'page', or null for a level chosen by hand.
  const [fit, setFit] = useState('auto');
  const [layout, setLayout] = useState('continuous');
  const [strip, setStrip] = useState(true);
  // What the side pane shows: the pages, or the bookmarks.
  const [side, setSide] = useState('pages');
  const [stripWidth, setStripWidth] = useState(168);
  const [busy, setBusy] = useState(null);
  const [failed, setFailed] = useState(null);
  const [locked, setLocked] = useState(false);
  const [findOpen, setFindOpen] = useState(false);
  const [query, setQuery] = useState('');
  const [hits, setHits] = useState([]);
  const [hitAt, setHitAt] = useState(-1);
  const [splitOpen, setSplitOpen] = useState(false);
  const [printing, setPrinting] = useState(null);
  // Recognise Text: null, { } while its dialog asks, { progress } while it reads.
  const [ocr, setOcr] = useState(null);
  const ocrStop = useRef(false);
  const ocrJob = useRef(null);
  // Comments: the tool in hand, each tool's colour, the comment picked, the pane, and a note being typed.
  const [tool, setTool] = useState(null);
  const toolRef = useRef(tool);
  toolRef.current = tool;
  const [toolColours, setToolColours] = useState(DEFAULT_TOOL_COLOURS);
  const [picked, setPicked] = useState(null);
  const [commentsOpen, setCommentsOpen] = useState(false);
  const [noteDraft, setNoteDraft] = useState(null);
  // Fill & Sign: fields tinted, the signature dialog, and a signature made, in hand to place.
  const [tintFields, setTintFields] = useState(true);
  const [signing, setSigning] = useState(null);
  const [signature, setSignature] = useState(null);
  // Redaction: the Apply dialog.
  const [applying, setApplying] = useState(false);
  // Export as Pictures and Reduce File Size: their dialogs.
  const [exporting, setExporting] = useState(false);
  const [compressing, setCompressing] = useState(false);
  const openFindRef = useRef(null);
  const viewports = useRef(new Map()).current;
  const viewRef = useRef(null);
  const findRef = useRef(null);
  const stripRef = useRef(null);
  const copyKey = useRef(newCopyKey());
  const openFileRef = useRef(null);
  const scrollToRef = useRef(null);

  const hasDoc = pages.length > 0;
  const protectedFile = Object.values(sources).some((s) => s.info.encrypted);

  /* ── sources ────────────────────────────────────────────────────────── */

  /** Read a PDF into the window: its structure in the main process, its pages for drawing here. */
  const addSource = useCallback(async (file, { bytes = null, label = null } = {}) => {
    const raw = bytes || (await shell.fs.read({ path: file })).bytes;
    const info = await shell.pdf.open(bytes ? { bytes: raw, name: label || basename(file || '') || t('Recovered PDF') } : { path: file });
    let doc;
    try {
      doc = await openDocument(raw);
    } catch (err) {
      shell.pdf.close({ id: info.id }).catch(() => {});
      if (err.password) {
        const e = new Error(t('{name} is protected with a password. Opening a PDF that needs a password to be read is not built yet.', { name: info.name }));
        e.password = true;
        throw e;
      }
      throw new Error(t('{name} could not be drawn: {reason}', { name: info.name, reason: err.message }));
    }
    const [annots, fields, nav] = await Promise.all([
      shell.pdf.annotations({ id: info.id }).catch(() => []),
      shell.pdf.fields({ id: info.id }).catch(() => []),
      shell.pdf.navigation({ id: info.id }).catch(() => null),
    ]);
    const source = { id: info.id, name: info.name, path: file || null, info, doc, annots, fields, nav };
    setSources((s) => ({ ...s, [info.id]: source }));
    sourcesRef.current = { ...sourcesRef.current, [info.id]: source };
    return source;
  }, [shell]);

  const resetHistory = () => {
    history.current = { undo: [], redo: [] };
    setHistoryTick((n) => n + 1);
  };

  /** Open a PDF as this window's document. */
  const openPdf = useCallback(async (file, { bytes = null, recovered = null } = {}) => {
    setBusy(t('Opening…'));
    setFailed(null);
    try {
      const source = await addSource(file, { bytes, label: recovered?.name });
      const list = pagesOf(source.id, source.info);
      setPages(list);
      pagesRef.current = list;
      setSelected([]);
      setCurrent(list[0]?.key || null);
      anchor.current = null;
      resetHistory();
      const where = recovered ? recovered.path : file;
      setPath(where || null);
      setName(where ? basename(where) : (recovered?.name || source.name));
      confirmed.current = null;
      setDirty(Boolean(recovered));
      setLocked(Boolean(source.info.encrypted));
      setFit('auto');
      if (!recovered && file) shell.app.addRecent({ path: file, app: 'pdf' }).catch(() => {});
      if (source.info.repaired) toast(t('{name} was damaged, and was read by finding its pages. Save it to keep a whole copy.', { name: source.name }), { ms: 7000 });
    } catch (err) {
      setFailed(err.message);
      setLocked(Boolean(err.password));
    } finally {
      setBusy(null);
    }
  }, [addSource, shell, toast]);

  const pickPdfs = useCallback(async (multiple) => {
    const paths = await shell.dialog.open({ title: multiple ? t('Choose PDFs') : t('Open'), filters: [...PDF_FILTER(), { name: t('All files'), extensions: ['*'] }], multiple });
    return (paths || []).filter(Boolean);
  }, [shell]);

  const openFile = useCallback(async () => {
    const [file] = await pickPdfs(false);
    if (!file) return;
    if (pagesRef.current.length) shell.win.create({ app: 'pdf', file });
    else openPdf(file);
  }, [pickPdfs, openPdf, shell]);
  openFileRef.current = openFile;

  /* ── the list and its history ────────────────────────────────────────── */

  const edit = useCallback((next, { select = null, current: now = null } = {}) => {
    const before = pagesRef.current;
    if (next === before) return false;
    history.current.undo.push({ pages: before, selected: selectedRef.current, current: currentRef.current });
    if (history.current.undo.length > 200) history.current.undo.shift();
    history.current.redo = [];
    pagesRef.current = next;
    setPages(next);
    if (select) setSelected(select);
    const keep = now || (next.some((p) => p.key === currentRef.current) ? currentRef.current : next[0]?.key || null);
    setCurrent(keep);
    // The page the change is about is brought into view.
    if (now) requestAnimationFrame(() => scrollToRef.current?.(now));
    setDirty(true);
    setHistoryTick((n) => n + 1);
    return true;
  }, []);

  const step = useCallback((from, to) => {
    const h = history.current;
    const entry = h[from].pop();
    if (!entry) return;
    h[to].push({ pages: pagesRef.current, selected: selectedRef.current, current: currentRef.current });
    pagesRef.current = entry.pages;
    setPages(entry.pages);
    setSelected(entry.selected.filter((k) => entry.pages.some((p) => p.key === k)));
    setCurrent(entry.pages.some((p) => p.key === entry.current) ? entry.current : entry.pages[0]?.key || null);
    setDirty(true);
    setHistoryTick((n) => n + 1);
  }, []);
  const undo = useCallback(() => step('undo', 'redo'), [step]);
  const redo = useCallback(() => step('redo', 'undo'), [step]);

  const chosen = useCallback(() => chosenKeys(pagesRef.current, selectedRef.current, currentRef.current), []);
  const numbersOf = (keys) => keys.map((k) => pagesRef.current.findIndex((p) => p.key === k) + 1);

  const guardLocked = useCallback(() => {
    if (!protectedFile) return false;
    toast(t('This PDF is protected, so its pages can be read but not changed or saved yet.'), { tone: 'bad', ms: 6000 });
    return true;
  }, [protectedFile, toast]);

  const rotate = useCallback((by) => {
    if (guardLocked()) return;
    const keys = chosen();
    if (keys.length) edit(rotatePages(pagesRef.current, keys, by));
  }, [chosen, edit, guardLocked]);

  const removePages = useCallback(() => {
    if (guardLocked()) return;
    const keys = chosen();
    if (!keys.length) return;
    if (keys.length >= pagesRef.current.length) {
      toast(t('A PDF needs at least one page. Select fewer pages to delete.'), { tone: 'bad', ms: 5000 });
      return;
    }
    const list = pagesRef.current;
    const first = list.findIndex((p) => keys.includes(p.key));
    const next = deletePages(list, keys);
    edit(next, { select: [], current: next[Math.min(first, next.length - 1)]?.key });
    toast(tn(keys.length, '{count} page deleted', '{count} pages deleted'), { ms: 2500 });
  }, [chosen, edit, guardLocked, toast]);

  const movePicked = useCallback((to, key) => {
    if (guardLocked()) return;
    const sel = selectedRef.current;
    const keys = sel.includes(key) ? pagesRef.current.filter((p) => sel.includes(p.key)).map((p) => p.key) : [key];
    edit(movePages(pagesRef.current, keys, to), { select: keys, current: key });
  }, [edit, guardLocked]);

  /** Pages of other files: after the last picked page, or at the end. */
  const insertFiles = useCallback(async (files, { at = null, announce = true } = {}) => {
    if (guardLocked() || !files.length) return;
    setBusy(t('Reading…'));
    const added = [];
    try {
      for (const file of files) {
        const source = await addSource(file);
        if (source.info.encrypted) {
          toast(t('{name} is protected, and its pages cannot be inserted yet.', { name: source.name }), { tone: 'bad', ms: 6000 });
          continue;
        }
        added.push(...pagesOf(source.id, source.info));
      }
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 7000 });
    } finally {
      setBusy(null);
    }
    if (!added.length) return;
    const list = pagesRef.current;
    let where = at;
    if (where === null) {
      const keys = selectedRef.current.length ? selectedRef.current : [];
      const last = Math.max(-1, ...keys.map((k) => list.findIndex((p) => p.key === k)));
      where = last >= 0 ? last + 1 : list.length;
    }
    edit(insertPages(list, where, added), { select: added.map((p) => p.key), current: added[0].key });
    if (announce) toast(tn(added.length, '{count} page inserted', '{count} pages inserted'), { ms: 3000 });
  }, [addSource, edit, guardLocked, toast]);

  const insertFromFile = useCallback(async () => {
    const files = await pickPdfs(true);
    if (files.length) insertFiles(files);
  }, [pickPdfs, insertFiles]);

  /** Merge Files: the PDFs chosen, one after another, at the end (or as a new document). */
  const mergeFiles = useCallback(async (given = null) => {
    const files = given || await pickPdfs(true);
    if (!files.length) return;
    if (!pagesRef.current.length) {
      await openPdf(files[0]);
      if (files.length > 1) await insertFiles(files.slice(1), { at: Number.MAX_SAFE_INTEGER, announce: false });
      if (files.length > 1) {
        setPath(null);
        setName(t('Combined PDF'));
        setDirty(true);
      }
      toast(tn(files.length, '{count} file combined', '{count} files combined'), { ms: 3000 });
      return;
    }
    await insertFiles(files, { at: pagesRef.current.length, announce: false });
    toast(tn(files.length, '{count} file added at the end', '{count} files added at the end'), { ms: 3000 });
  }, [pickPdfs, openPdf, insertFiles, toast]);

  /* ── recognising text ──────────────────────────────────────────────── */

  const recognise = useCallback(async ({ scope, languages }) => {
    if (guardLocked()) return;
    const list = pagesRef.current;
    const keys = new Set(scope === 'picked' ? selectedRef.current : scope === 'current' ? [currentRef.current] : list.map((p) => p.key));
    const seen = new Set();
    const targets = [];
    for (const p of list) {
      const id = `${p.source}:${p.page}`;
      if (!keys.has(p.key) || seen.has(id)) continue;
      seen.add(id);
      targets.push({ source: p.source, page: p.page });
    }
    ocrStop.current = false;
    ocrJob.current = `pdf-${Date.now().toString(36)}`;
    setOcr({ progress: { done: 0, total: 0 } });
    try {
      const out = await recognisePages({
        shell,
        sources: sourcesRef.current,
        targets,
        languages,
        job: ocrJob.current,
        cancelled: () => ocrStop.current,
        onProgress: (progress) => setOcr({ progress }),
      });
      if (ocrStop.current) {
        toast(t('Recognise Text was cancelled. Nothing was changed.'), { ms: 4000 });
        return;
      }
      if (!out.scanned) {
        toast(t('Every page chosen already has text, so there was nothing to recognise.'), { ms: 5000 });
        return;
      }
      const swap = {};
      for (const [id, written] of out.bySource) {
        if (!written.length) continue;
        const { bytes, written: pagesWritten, ...info } = await shell.pdf.addText({ id, pages: written });
        const doc = await openDocument(bytes);
        const source = { id: info.id, name: info.name, path: sourcesRef.current[id]?.path || null, info, doc, annots: sourcesRef.current[id]?.annots || [], fields: sourcesRef.current[id]?.fields || [], nav: sourcesRef.current[id]?.nav || null };
        sourcesRef.current = { ...sourcesRef.current, [info.id]: source };
        setSources((s) => ({ ...s, [info.id]: source }));
        swap[id] = info.id;
      }
      if (Object.keys(swap).length) edit(pagesRef.current.map((p) => (swap[p.source] ? { ...p, source: swap[p.source] } : p)));
      if (out.words) toast(tn(out.done, 'Text recognised on {count} page. Find, select and copy now work on it.', 'Text recognised on {count} pages. Find, select and copy now work on them.'), { tone: 'good', ms: 5000 });
      else toast(t('No words were found on the scanned pages.'), { ms: 5000 });
    } catch (err) {
      if (err?.code === 'OCR_CANCELLED' || ocrStop.current) toast(t('Recognise Text was cancelled. Nothing was changed.'), { ms: 4000 });
      else toast(t('Text could not be recognised: {reason}', { reason: err.message }), { tone: 'bad', ms: 8000 });
    } finally {
      setOcr(null);
    }
  }, [guardLocked, shell, edit, toast]);

  const cancelRecognise = useCallback(() => {
    ocrStop.current = true;
    if (ocrJob.current) shell.ocr.cancel({ job: ocrJob.current }).catch(() => {});
  }, [shell]);

  const askRecognise = useCallback(() => {
    if (guardLocked() || !pagesRef.current.length) return;
    setOcr({});
  }, [guardLocked]);

  /* ── comments ───────────────────────────────────────────────────────── */

  /**
   * Comments changed in one or more sources, as one step of the history:
   * [{ source, changes }] for comments, [{ source, values }] for a form's
   * values. Each source's file gets an incremental update and becomes a new
   * source; the pages that came from it point at the new one.
   */
  // Changes are written one after another, each to the newest version of its
  // file: a field left while the one before it is still being written (Tab,
  // Tab, Tab) would otherwise be written to the file as it was, and lost.
  const writing = useRef(Promise.resolve());
  const newer = useRef(new Map());
  const writeRef = useRef(null);
  const changeComments = useCallback((list, opts) => {
    const run = () => writeRef.current(list, opts);
    const next = writing.current.then(run, run);
    writing.current = next.catch(() => null);
    return next;
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const writeChanges = async (list, { pick = false } = {}) => {
    if (guardLocked()) return null;
    const swap = {};
    let first = null;
    try {
      for (const { source: asked, changes, values } of list) {
        // The version the pages show now: what this change was asked of, unless a change before it replaced it.
        let id = asked;
        while (!pagesRef.current.some((p) => p.source === id) && newer.current.has(id)) id = newer.current.get(id);
        const { bytes, added, annotations, fields, ...info } = values
          ? await shell.pdf.fill({ id, values })
          : await shell.pdf.annotate({ id, changes });
        const doc = await openDocument(bytes);
        const old = sourcesRef.current[id];
        const source = { id: info.id, name: info.name, path: old?.path || null, info, doc, annots: annotations, fields, nav: old?.nav || null };
        sourcesRef.current = { ...sourcesRef.current, [info.id]: source };
        setSources((s) => ({ ...s, [info.id]: source }));
        swap[id] = info.id;
        newer.current.set(id, info.id);
        if (!first && added?.length) first = added[0];
      }
    } catch (err) {
      toast(list.some((x) => x.values) ? t('The form could not be filled: {reason}', { reason: err.message }) : t('The comment could not be written: {reason}', { reason: err.message }), { tone: 'bad', ms: 8000 });
      return null;
    }
    if (Object.keys(swap).length) edit(pagesRef.current.map((p) => (swap[p.source] ? { ...p, source: swap[p.source] } : p)));
    if (pick && first) setPicked(first);
    return first;
  };
  writeRef.current = writeChanges;

  const colourOf = useCallback((which) => swatch(toolColours[which] || DEFAULT_TOOL_COLOURS[which]).rgb, [toolColours]);

  /** The words selected marked with a highlight, an underline, a strikethrough or for redaction. Answers whether there were any. */
  const markSelection = useCallback((which) => {
    const quads = selectionQuads(viewRef.current, viewports);
    const bySource = new Map();
    for (const [key, list] of Object.entries(quads)) {
      const entry = pagesRef.current.find((p) => p.key === key);
      if (!entry || !list.length) continue;
      if (!bySource.has(entry.source)) bySource.set(entry.source, []);
      const spec = which === 'redact-text'
        ? { page: entry.page, type: 'Redact', quads: list }
        : { page: entry.page, type: TOOL_TYPES[which], quads: list, color: colourOf(which) };
      bySource.get(entry.source).push(spec);
    }
    if (!bySource.size) return false;
    window.getSelection()?.removeAllRanges();
    changeComments([...bySource].map(([source, add]) => ({ source, changes: { add } })), { pick: which !== 'redact-text' });
    return true;
  }, [viewports, changeComments, colourOf]);

  /** A markup button: the words already selected are marked at once; otherwise the tool stays in hand for the next selection. */
  const pressTool = useCallback((which) => {
    if (guardLocked()) return;
    if (MARKUP_TOOLS.has(which) && markSelection(which)) return;
    setTool((now) => (now === which ? null : which));
    setPicked(null);
  }, [guardLocked, markSelection]);

  // With a markup tool in hand, letting go of a selection marks it.
  const onViewMouseUp = useCallback(() => {
    const now = toolRef.current;
    if (!now || !MARKUP_TOOLS.has(now)) return;
    setTimeout(() => markSelection(now), 0);
  }, [markSelection]);

  /** A shape drawn, or a place chosen for a note, on a page. */
  const onCreate = useCallback((d) => {
    if (guardLocked()) return;
    if (d.tool === 'note' || d.tool === 'text') { setNoteDraft({ ...d, colour: colourOf(d.tool) }); return; }
    if (d.tool === 'stamp') {
      if (!signature) return;
      setTool(null);
      signatureStamp(signature, d).then((spec) => {
        if (spec) changeComments([{ source: d.source, changes: { add: [{ ...spec, page: d.page }] } }], { pick: true });
      });
      return;
    }
    if (d.tool === 'date') {
      // Today, as the words of a text box where the page was clicked, top left at the click.
      const words = new Intl.DateTimeFormat('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }).format(new Date());
      const ctx2d = document.createElement('canvas').getContext('2d');
      ctx2d.font = '11px Helvetica, Arial, sans-serif';
      const wide = Math.ceil(ctx2d.measureText(words).width) + 10;
      changeComments([{ source: d.source, changes: { add: [{ page: d.page, type: 'FreeText', rect: [d.point[0], d.point[1] - 18, d.point[0] + wide, d.point[1]], contents: words, fontSize: 11, color: [0.06, 0.14, 0.36], subject: 'Date' }] } }], { pick: true });
      return;
    }
    if (d.tool === 'check') {
      // A tick, 14 points square, centred on the click.
      const [cx, cy] = d.point;
      const tick = [[cx - 5, cy, cx - 1.5, cy - 4, cx + 5.5, cy + 4.5]];
      changeComments([{ source: d.source, changes: { add: [{ page: d.page, type: 'Stamp', rect: [cx - 7, cy - 7, cx + 7, cy + 7], paths: tick, width: 1.8, icon: 'Check', subject: 'Checkmark', color: [0.06, 0.14, 0.36] }] } }], { pick: true });
      return;
    }
    if (d.tool === 'redact-area') {
      changeComments([{ source: d.source, changes: { add: [{ page: d.page, type: 'Redact', rect: d.rect }] } }]);
      return;
    }
    const spec = { page: d.page, type: TOOL_TYPES[d.tool], color: colourOf(d.tool) };
    if (d.ink) spec.ink = d.ink;
    if (d.line) spec.line = d.line;
    if (d.rect) spec.rect = d.rect;
    changeComments([{ source: d.source, changes: { add: [spec] } }], { pick: true });
  }, [guardLocked, colourOf, changeComments, signature]);

  const fillField = useCallback((sourceId, field, value) => {
    if (guardLocked()) return;
    changeComments([{ source: sourceId, values: { [field.id]: value } }]);
  }, [guardLocked, changeComments]);

  /* ── redaction ──────────────────────────────────────────────────────── */

  /** The redaction marks on the window's pages, and how many pages carry them. */
  const redactionMarks = useCallback(() => {
    let marks = 0;
    const on = new Set();
    for (const p of pagesRef.current) {
      const list = (sourcesRef.current[p.source]?.annots || []).filter((a) => a.type === 'Redact' && a.page === p.page);
      marks += list.length;
      if (list.length) on.add(p.key);
    }
    return { marks, pages: on.size };
  }, []);

  /** Every match of the words in Find, on every page, marked for redaction. */
  const markSearchResults = useCallback(async () => {
    if (guardLocked()) return;
    const q = query.trim();
    if (!findOpen || !q) {
      openFindRef.current?.();
      toast(t('Type the words to find, then press Mark Matches again to mark every match.'), { ms: 5000 });
      return;
    }
    setBusy(t('Finding…'));
    const bySource = new Map();
    let count = 0;
    try {
      const seen = new Set();
      for (const p of pagesRef.current) {
        const id = `${p.source}:${p.page}`;
        if (seen.has(id)) continue;
        seen.add(id);
        const quads = await searchQuads(sourcesRef.current[p.source].doc, p.page, q);
        if (!quads.length) continue;
        count += quads.length;
        if (!bySource.has(p.source)) bySource.set(p.source, []);
        bySource.get(p.source).push({ page: p.page, type: 'Redact', quads });
      }
    } finally {
      setBusy(null);
    }
    if (!bySource.size) { toast(t('No matches to mark.'), { ms: 4000 }); return; }
    await changeComments([...bySource].map(([source, add]) => ({ source, changes: { add } })));
    toast(tn(count, '{count} match marked for redaction. Apply Redactions removes it.', '{count} matches marked for redaction. Apply Redactions removes them.'), { ms: 5000 });
  }, [guardLocked, query, findOpen, toast, changeComments]);

  const askApply = useCallback(() => {
    if (guardLocked()) return;
    if (!redactionMarks().marks) {
      toast(t('Nothing is marked for redaction yet. Mark words, an area or the matches of a search first.'), { ms: 5000 });
      return;
    }
    setApplying(true);
  }, [guardLocked, redactionMarks, toast]);

  /** The marks applied: the pages written whole as a new file, which then opens here. */
  const applyRedactions = useCallback(async () => {
    setApplying(false);
    const list = pagesRef.current;
    const folder = path ? dirname(path) : (await shell.app.paths().catch(() => ({}))).documents || '';
    const stem = stemOf(path || name || '') || t('Untitled');
    const target = await shell.dialog.save({ title: t('Save the redacted PDF as'), defaultPath: joinPath(folder, `${stem}_Redacted.pdf`), filters: PDF_FILTER() });
    if (!target) return;
    const out = /\.pdf$/i.test(target) ? target : `${target}.pdf`;
    if (path && out.toLowerCase() === path.toLowerCase()) {
      toast(t('Save the redacted PDF under a new name: the original is kept as it is.'), { tone: 'bad', ms: 6000 });
      return;
    }
    setBusy(t('Redacting…'));
    try {
      const sources = [...new Set(list.map((p) => p.source))];
      const plan = await shell.pdf.redactPlan({ sources });
      const images = await decodePlan(plan);
      const { removed } = await shell.pdf.redact({ pages: entriesOf(list), path: out, images });
      setBusy(null);
      await openPdf(out);
      const parts = [tn(removed.pages, 'Redacted {count} page.', 'Redacted {count} pages.')];
      if (removed.images) parts.push(tn(removed.images, '{count} picture under a mark could not be painted out and was left off its page.', '{count} pictures under marks could not be painted out and were left off their pages.'));
      parts.push(t('Saved as {name}.', { name: basename(out) }));
      toast(parts.join(' '), { tone: 'good', ms: 7000 });
    } catch (err) {
      toast(t('The redactions could not be applied: {reason}', { reason: err.message }), { tone: 'bad', ms: 8000 });
    } finally {
      setBusy(null);
    }
  }, [path, name, shell, toast, openPdf]);

  /** Fill & Sign: the signature or initials kept from last time, or the dialog to make them. */
  const sign = useCallback((kind, { fresh = false } = {}) => {
    if (guardLocked()) return;
    const kept = fresh ? null : savedSignature(kind);
    if (kept) { setSignature(kept); setTool('stamp'); setPicked(null); return; }
    setSigning(kind);
  }, [guardLocked]);

  const postNote = useCallback((text) => {
    const d = noteDraft;
    setNoteDraft(null);
    if (!d) return;
    const spec = d.tool === 'note'
      ? { page: d.page, type: 'Text', rect: [d.point[0], d.point[1] - 20, d.point[0] + 20, d.point[1]], contents: text, color: d.colour }
      : { page: d.page, type: 'FreeText', rect: d.rect, contents: text, fontSize: 12, color: d.colour };
    changeComments([{ source: d.source, changes: { add: [spec] } }], { pick: true });
  }, [noteDraft, changeComments]);

  const replyTo = useCallback((item, text) => changeComments([{ source: item.source, changes: { reply: [{ to: item.annot.id, contents: text }] } }]), [changeComments]);
  const resolveComment = useCallback((item, resolved) => changeComments([{ source: item.source, changes: { resolve: [{ id: item.annot.id, resolved }] } }]), [changeComments]);
  const editComment = useCallback((item, text) => changeComments([{ source: item.source, changes: { edit: [{ id: item.annot.id, contents: text }] } }]), [changeComments]);
  const deleteComment = useCallback((item) => {
    setPicked(null);
    changeComments([{ source: item.source, changes: { remove: [item.annot.id] } }]);
  }, [changeComments]);
  const deletePicked = useCallback(() => {
    if (!picked) return;
    const item = commentItems(pagesRef.current, sourcesRef.current).find((it) => it.annot.id === picked);
    if (item) deleteComment(item);
  }, [picked, deleteComment]);

  const pickComment = useCallback((id, pageKey) => {
    setPicked(id);
    setTool(null);
    if (pageKey && pageKey !== currentRef.current) {
      setCurrent(pageKey);
      requestAnimationFrame(() => scrollToRef.current?.(pageKey));
    }
    // The comment on its page brought into view.
    requestAnimationFrame(() => {
      const el = viewRef.current?.querySelector(`.pdf-page[data-key="${pageKey}"] .pdf-annot[data-id="${id}"]`);
      const view = viewRef.current;
      if (!el || !view) return;
      const r = el.getBoundingClientRect();
      const box = view.getBoundingClientRect();
      if (r.top < box.top + 20 || r.bottom > box.bottom - 20) view.scrollTop += r.top - box.top - box.height / 3;
    });
  }, []);

  /** A link on a page: to a page of this file, where that page now stands, or out to a web address. */
  const followLink = useCallback((link, entry) => {
    if (link.url) {
      shell.shell.openExternal({ url: link.url }).catch(() => {});
      return;
    }
    const to = pagesRef.current.find((p) => p.source === entry.source && p.page === link.target);
    if (to) goToRef.current?.(to.key);
    else toast(t('The page this link goes to is not in the document now.'), { ms: 4000 });
  }, [shell, toast]);
  const goToRef = useRef(null);

  const overlay = useCallback((entry, { source, scale }) => (
    <>
      <FieldLayer entry={entry} source={source} scale={scale} fields={source?.fields} highlight={tintFields} onFill={fillField} />
      <AnnotLayer
      entry={entry}
      source={source}
      scale={scale}
      annots={source?.annots}
      tool={tool}
      colour={tool ? colourOf(tool) : null}
      picked={picked}
      onPick={pickComment}
      onCreate={onCreate}
      viewports={viewports}
      onLink={followLink}
      />
    </>
  ), [tool, colourOf, picked, pickComment, onCreate, viewports, tintFields, fillField, followLink]);

  /* ── writing ────────────────────────────────────────────────────────── */

  const entriesOf = (list) => list.map((p) => ({ source: p.source, page: p.page, rotate: p.rotate }));

  const writeTo = useCallback(async (target, list, { title } = {}) => {
    const result = await shell.pdf.assemble({ pages: entriesOf(list), path: target, title });
    return result;
  }, [shell]);

  const saveTo = useCallback(async (target) => {
    if (guardLocked()) return false;
    setBusy(t('Saving…'));
    try {
      await writeTo(target, pagesRef.current);
      setPath(target);
      setName(basename(target));
      confirmed.current = target;
      setDirty(false);
      shell.doc.dropCopy({ key: copyKey.current }).catch(() => {});
      shell.app.addRecent({ path: target, app: 'pdf' }).catch(() => {});
      toast(t('Saved {name}', { name: basename(target) }), { tone: 'good' });
      return true;
    } catch (err) {
      toast(t('The PDF could not be saved: {reason}', { reason: err.message }), { tone: 'bad', ms: 8000 });
      return false;
    } finally {
      setBusy(null);
    }
  }, [guardLocked, writeTo, shell, toast]);

  const saveAs = useCallback(async () => {
    if (!pagesRef.current.length || guardLocked()) return false;
    const folder = path ? dirname(path) : (await shell.app.paths().catch(() => ({}))).documents || '';
    const stem = path ? stemOf(path) : stemOf(name || '') || t('Untitled');
    const target = await shell.dialog.save({ title: t('Save as'), defaultPath: joinPath(folder, `${stem}.pdf`), filters: PDF_FILTER() });
    if (!target) return false;
    return saveTo(/\.pdf$/i.test(target) ? target : `${target}.pdf`);
  }, [path, name, shell, saveTo, guardLocked]);

  /** Save: to the file it came from only once the person has said so; Save As otherwise. */
  const save = useCallback(async () => {
    if (!pagesRef.current.length || guardLocked()) return false;
    if (!path) return saveAs();
    if (confirmed.current !== path) {
      const { response } = await shell.dialog.message({
        type: 'question',
        message: t('Replace {name}?', { name: basename(path) }),
        detail: t('Saving writes the pages as they are now over the original file. Save a Copy keeps the original as it is.'),
        buttons: [t('Save a Copy…'), t('Replace the Original'), t('Cancel')],
        defaultId: 0,
        cancelId: 2,
      });
      if (response === 2) return false;
      if (response === 0) return saveAs();
      confirmed.current = path;
    }
    return saveTo(path);
  }, [path, shell, saveAs, saveTo, guardLocked]);

  const extract = useCallback(async () => {
    if (guardLocked()) return;
    const keys = chosen();
    if (!keys.length) return;
    const list = pagesRef.current.filter((p) => keys.includes(p.key));
    const range = rangeText(numbersOf(keys));
    const folder = path ? dirname(path) : (await shell.app.paths().catch(() => ({}))).documents || '';
    const target = await shell.dialog.save({
      title: t('Extract pages {range}', { range }),
      defaultPath: joinPath(folder, `${stemOf(path || name || '') || t('Untitled')} ${t('pages {range}', { range })}.pdf`),
      filters: PDF_FILTER(),
    });
    if (!target) return;
    setBusy(t('Saving…'));
    try {
      const out = /\.pdf$/i.test(target) ? target : `${target}.pdf`;
      await writeTo(out, list);
      toast(tn(list.length, '{count} page extracted to {name}', '{count} pages extracted to {name}', { name: basename(out) }), { tone: 'good', ms: 4000 });
    } catch (err) {
      toast(t('The pages could not be extracted: {reason}', { reason: err.message }), { tone: 'bad', ms: 8000 });
    } finally {
      setBusy(null);
    }
  }, [chosen, path, name, shell, writeTo, toast, guardLocked]);

  const split = useCallback(async ({ every = null, picked = false }) => {
    setSplitOpen(false);
    const list = pagesRef.current;
    const starts = picked ? list.map((p, i) => (selectedRef.current.includes(p.key) ? i : -1)).filter((i) => i > 0) : null;
    const groups = splitGroups(list.length, { every, starts });
    if (groups.length < 2) return;
    const [folder] = await shell.dialog.open({ title: t('Choose a folder for the files'), directory: true, defaultPath: path ? dirname(path) : undefined }) || [];
    if (!folder) return;
    const stem = stemOf(path || name || '') || t('Untitled');
    setBusy(t('Saving…'));
    const written = [];
    try {
      for (let g = 0; g < groups.length; g++) {
        let target = joinPath(folder, `${stem}_Part${g + 1}.pdf`);
        for (let n = 2; await shell.fs.exists({ path: target }); n++) target = joinPath(folder, `${stem}_Part${g + 1} (${n}).pdf`);
        await writeTo(target, groups[g].map((i) => list[i]));
        written.push(target);
      }
      toast(tn(written.length, 'Split into {count} file in {folder}', 'Split into {count} files in {folder}', { folder: basename(folder) || folder }), { tone: 'good', ms: 5000 });
    } catch (err) {
      toast(t('The PDF could not be split: {reason}', { reason: err.message }), { tone: 'bad', ms: 8000 });
    } finally {
      setBusy(null);
    }
  }, [path, name, shell, writeTo, toast]);

  /** Export as Pictures: each page chosen written as a PNG or a JPEG into a folder. */
  const exportPictures = useCallback(async ({ scope, format, dpi }) => {
    setExporting(false);
    const list = pagesRef.current;
    const keys = new Set(scope === 'picked' ? selectedRef.current : scope === 'current' ? [currentRef.current] : list.map((p) => p.key));
    const chosen = list.map((p, i) => ({ p, n: i + 1 })).filter(({ p }) => keys.has(p.key));
    if (!chosen.length) return;
    const [folder] = await shell.dialog.open({ title: t('Choose a folder for the pictures'), directory: true, defaultPath: path ? dirname(path) : undefined }) || [];
    if (!folder) return;
    const stem = stemOf(path || name || '') || t('Untitled');
    const ext = format === 'jpeg' ? 'jpg' : 'png';
    setBusy(t('Exporting…'));
    let written = 0;
    try {
      for (const { p, n } of chosen) {
        const { bytes } = await pageBytes(sourcesRef.current[p.source].doc, p.page, { rotate: p.rotate, dpi, type: format });
        let target = joinPath(folder, `${stem}_Page_${n}.${ext}`);
        for (let k = 2; await shell.fs.exists({ path: target }); k++) target = joinPath(folder, `${stem}_Page_${n} (${k}).${ext}`);
        await shell.fs.write({ path: target, bytes });
        written += 1;
      }
      toast(tn(written, '{count} page exported as a picture to {folder}', '{count} pages exported as pictures to {folder}', { folder: basename(folder) || folder }), { tone: 'good', ms: 5000 });
    } catch (err) {
      toast(t('The pages could not be exported: {reason}', { reason: err.message }), { tone: 'bad', ms: 8000 });
    } finally {
      setBusy(null);
    }
  }, [path, name, shell, toast]);

  /** Reduce File Size: the pictures drawn smaller, written whole as a new file. */
  const compressFile = useCallback(async (quality) => {
    setCompressing(false);
    if (guardLocked()) return;
    const folder = path ? dirname(path) : (await shell.app.paths().catch(() => ({}))).documents || '';
    const stem = stemOf(path || name || '') || t('Untitled');
    const target = await shell.dialog.save({ title: t('Save the smaller PDF as'), defaultPath: joinPath(folder, `${stem}_Reduced.pdf`), filters: PDF_FILTER() });
    if (!target) return;
    const out = /\.pdf$/i.test(target) ? target : `${target}.pdf`;
    setBusy(t('Making the file smaller…'));
    try {
      const { token, plan } = await shell.pdf.compressPlan({ pages: entriesOf(pagesRef.current), dpi: quality.dpi });
      const images = await encodePlan(plan);
      const { before, after, replaced } = await shell.pdf.compress({ token, dpi: quality.dpi, images, path: out });
      if (!replaced) toast(t('The pictures were already small enough. The copy is saved as {name}.', { name: basename(out) }), { ms: 6000 });
      else toast(t('Saved {name}: {before} down to {after}.', { name: basename(out), before: formatBytes(before), after: formatBytes(after) }), { tone: 'good', ms: 7000 });
    } catch (err) {
      toast(t('The file could not be made smaller: {reason}', { reason: err.message }), { tone: 'bad', ms: 8000 });
    } finally {
      setBusy(null);
    }
  }, [guardLocked, path, name, shell, toast]);

  /** Print: every page drawn at print resolution, handed to the system's print dialog. */
  const print = useCallback(async () => {
    const list = pagesRef.current;
    if (!list.length || printing) return;
    setBusy(t('Preparing to print…'));
    const images = [];
    try {
      for (const p of list) {
        const source = sourcesRef.current[p.source];
        images.push(await pageImage(source.doc, p.page, { rotate: p.rotate, dpi: 150 }));
      }
      setPrinting(images);
      await new Promise((r) => setTimeout(r, 120));
      await shell.print.print({ silent: false });
    } catch (err) {
      toast(t('The pages could not be printed: {reason}', { reason: err.message }), { tone: 'bad', ms: 8000 });
    } finally {
      setBusy(null);
      setPrinting(null);
      for (const img of images) URL.revokeObjectURL(img.url);
    }
  }, [printing, shell, toast]);

  /* ── selection and the page in view ──────────────────────────────────── */

  const scrollToPage = useCallback((key, block = 'start') => {
    const el = viewRef.current?.querySelector(`.pdf-page[data-key="${key}"]`);
    const view = viewRef.current;
    if (!el || !view) return;
    // Only the page view scrolls: scrollIntoView would move the window's frame too.
    const r = el.getBoundingClientRect();
    const box = view.getBoundingClientRect();
    view.scrollTop += block === 'start' ? r.top - box.top - 12 : 0;
    if (r.left < box.left || r.right > box.right) view.scrollLeft += r.left - box.left - 12;
  }, []);

  scrollToRef.current = scrollToPage;

  const goTo = useCallback((key, { scroll = true } = {}) => {
    if (!key) return;
    setCurrent(key);
    if (scroll) requestAnimationFrame(() => scrollToPage(key));
  }, [scrollToPage]);
  goToRef.current = goTo;

  const onPick = useCallback((key, { ctrl, shift, press, release }) => {
    const sel = selectedRef.current;
    if (press) {
      if (!ctrl && !shift && !sel.includes(key)) {
        setSelected([key]);
        anchor.current = key;
        goTo(key);
      }
      return;
    }
    if (!release) return;
    if (ctrl) {
      setSelected(sel.includes(key) ? sel.filter((k) => k !== key) : [...sel, key]);
      anchor.current = key;
    } else if (shift) {
      setSelected(rangeKeys(pagesRef.current, anchor.current || currentRef.current || key, key));
    } else {
      setSelected([key]);
      anchor.current = key;
    }
    goTo(key);
  }, [goTo]);

  const onCurrent = useCallback((key) => setCurrent((k) => (k === key ? k : key)), []);

  const stepPage = useCallback((by) => {
    const list = pagesRef.current;
    const at = list.findIndex((p) => p.key === currentRef.current);
    const next = list[Math.max(0, Math.min(list.length - 1, (at < 0 ? 0 : at) + by))];
    if (next) goTo(next.key);
  }, [goTo]);

  const onStripKey = useCallback((e) => {
    const list = pagesRef.current;
    if (!list.length) return;
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      const at = list.findIndex((p) => p.key === currentRef.current);
      const next = list[Math.max(0, Math.min(list.length - 1, at + (e.key === 'ArrowDown' ? 1 : -1)))];
      if (!next) return;
      if (e.shiftKey) setSelected(rangeKeys(list, anchor.current || currentRef.current, next.key));
      else { setSelected([next.key]); anchor.current = next.key; }
      goTo(next.key);
    } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'a') {
      e.preventDefault();
      setSelected(list.map((p) => p.key));
    } else if (e.key === 'Delete') {
      e.preventDefault();
      removePages();
    }
  }, [goTo, removePages]);

  /* ── zoom ───────────────────────────────────────────────────────────── */

  const sizeOf = useCallback((p) => shownSize(sourcesRef.current[p.source]?.info.pages[p.page], p.rotate), []);
  const fitZoom = useCallback((how) => {
    const el = viewRef.current;
    const list = pagesRef.current;
    if (!el || !list.length) return null;
    const p = list.find((x) => x.key === currentRef.current) || list[0];
    const size = sizeOf(p);
    const across = layout === 'two' ? 2 : 1;
    const roomW = el.clientWidth - 48 - (across - 1) * 14 - 16;
    const roomH = el.clientHeight - 36;
    const byWidth = roomW / (size.width * across * PX_PER_PT);
    const byPage = Math.min(byWidth, roomH / (size.height * PX_PER_PT));
    const z = how === 'page' ? byPage : how === 'auto' ? Math.min(1, byWidth) : byWidth;
    return Math.max(ZOOM_MIN, Math.min(ZOOM_MAX, Math.floor(z * 100) / 100));
  }, [layout, sizeOf]);

  // Fit Width and Fit Page follow the window as it is resized.
  useEffect(() => {
    if (!fit || !hasDoc) return undefined;
    const apply = () => { const z = fitZoom(fit); if (z) setZoom(z); };
    apply();
    const ro = new ResizeObserver(apply);
    if (viewRef.current) ro.observe(viewRef.current);
    return () => ro.disconnect();
  }, [fit, fitZoom, hasDoc, strip, stripWidth, layout]);

  const setZoomByHand = useCallback((z) => { setFit(null); setZoom(Math.max(ZOOM_MIN, Math.min(ZOOM_MAX, z))); }, []);
  useWheelZoom({ shell, zoom, setZoom: setZoomByHand, min: ZOOM_MIN, max: ZOOM_MAX, area: '.pdf-view', target: '.pdf-view-inner', reset: 1 });

  /* ── find ───────────────────────────────────────────────────────────── */

  const findSeq = useRef(0);
  useEffect(() => {
    const q = query;
    const seq = ++findSeq.current;
    if (!findOpen || !q.trim()) { setHits([]); setHitAt(-1); return; }
    (async () => {
      const found = [];
      for (const p of pagesRef.current) {
        const source = sourcesRef.current[p.source];
        if (!source) continue;
        let text = '';
        try { text = await pageText(source.doc, p.page); } catch { text = ''; }
        if (seq !== findSeq.current) return;
        findAll(text, q).forEach((m, n) => found.push({ key: p.key, n, ...m }));
      }
      if (seq !== findSeq.current) return;
      setHits(found);
      const at = found.findIndex((h) => pagesRef.current.findIndex((p) => p.key === h.key) >= pagesRef.current.findIndex((p) => p.key === currentRef.current));
      setHitAt(found.length ? Math.max(0, at) : -1);
    })();
  }, [query, findOpen, pages]);

  const [textTick, setTextTick] = useState(0);
  const onText = useCallback(() => setTextTick((n) => n + 1), []);

  // The matches on the pages drawn, marked with the browser's own highlights.
  useEffect(() => {
    const registry = typeof CSS !== 'undefined' ? globalThis.CSS?.highlights : null;
    if (!registry) return;
    registry.delete('pdf-find');
    registry.delete('pdf-find-now');
    if (!findOpen || !hits.length) return;
    const all = [];
    let now = null;
    const byKey = new Map();
    hits.forEach((h, i) => { if (!byKey.has(h.key)) byKey.set(h.key, []); byKey.get(h.key).push({ ...h, i }); });
    for (const [key, list] of byKey) {
      const layer = viewRef.current?.querySelector(`.pdf-page[data-key="${key}"] .textLayer[data-ready="yes"]`);
      if (!layer) continue;
      // The layer's words in order, a line break where pdf.js put one: the same text find searched.
      const nodes = [];
      let text = '';
      const walker = document.createTreeWalker(layer, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        if (n.nodeType === 3) { nodes.push({ node: n, at: text.length }); text += n.data; } else if (n.nodeName === 'BR') text += '\n';
      }
      const local = findAll(text, query);
      const locate = (offset) => {
        for (let k = nodes.length - 1; k >= 0; k--) if (nodes[k].at <= offset) return { node: nodes[k].node, offset: Math.min(offset - nodes[k].at, nodes[k].node.data.length) };
        return null;
      };
      for (const h of list) {
        const m = local[h.n];
        if (!m) continue;
        const a = locate(m.start);
        const b = locate(m.end);
        if (!a || !b) continue;
        const range = document.createRange();
        try { range.setStart(a.node, a.offset); range.setEnd(b.node, b.offset); } catch { continue; }
        all.push(range);
        if (h.i === hitAt) now = range;
      }
    }
    if (all.length) registry.set('pdf-find', new Highlight(...all));
    if (now) {
      registry.set('pdf-find-now', new Highlight(now));
      const r = now.getBoundingClientRect();
      const box = viewRef.current?.getBoundingClientRect();
      if (box && (r.top < box.top + 40 || r.bottom > box.bottom - 20)) viewRef.current.scrollTop += r.top - box.top - box.height / 3;
    }
  }, [hits, hitAt, textTick, findOpen, query]);

  const nextHit = useCallback((by) => {
    if (!hits.length) return;
    const at = (hitAt + by + hits.length) % hits.length;
    setHitAt(at);
    const key = hits[at].key;
    if (key !== currentRef.current || layout === 'single') goTo(key, { scroll: layout !== 'single' });
    else setCurrent(key);
  }, [hits, hitAt, goTo, layout]);

  // A page coming to be the match's moves the view there first.
  useEffect(() => {
    if (hitAt < 0 || !hits[hitAt]) return;
    const key = hits[hitAt].key;
    if (key !== currentRef.current) goTo(key);
    // A new search whose first match is again match 0 moves too: it can be on another page.
  }, [hitAt, hits]); // eslint-disable-line react-hooks/exhaustive-deps

  const openFind = useCallback(() => {
    setFindOpen(true);
    requestAnimationFrame(() => { findRef.current?.focus(); findRef.current?.select(); });
  }, []);
  openFindRef.current = openFind;

  /* ── start-up, drops, the close guard, recovery ──────────────────────── */

  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const recover = params.get('recover');
    if (recover) {
      (async () => {
        try {
          const copy = await shell.doc.readCopy({ file: recover });
          const m = /^window-([A-Za-z0-9_-]+)\./.exec(recover);
          if (m) copyKey.current = m[1];
          await openPdf(copy.path || null, { bytes: copy.bytes, recovered: copy });
          toast(t('Recovered unsaved work. Save it to keep it.'), { ms: 6000 });
        } catch (err) {
          toast(err.message, { tone: 'bad' });
        }
      })();
    } else if (boot.file) {
      openPdf(boot.file);
    } else if (params.get('combine')) {
      mergeFiles();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useFileDrop(useCallback((files) => {
    const pdfs = files.filter((f) => /\.pdf$/i.test(f));
    if (!pdfs.length) return;
    if (!pagesRef.current.length) openPdf(pdfs[0]);
    else for (const f of pdfs) shell.win.create({ app: 'pdf', file: f });
  }, [openPdf, shell]));

  useDirtyGuard({ shell, dirty: hasDoc && dirty, name: name || t('Untitled'), onSave: save });

  const copied = useRef({ at: 0, pages: null });
  useEffect(() => {
    if (!hasDoc || !dirty || protectedFile) return undefined;
    const since = Date.now() - copied.current.at;
    const wait = copied.current.at ? Math.max(0, AUTOSAVE_EVERY_MS - since) : AUTOSAVE_FIRST_MS;
    const timer = setTimeout(async () => {
      if (copied.current.pages === pagesRef.current) return;
      copied.current = { at: Date.now(), pages: pagesRef.current };
      try {
        const { bytes } = await shell.pdf.assemble({ pages: entriesOf(pagesRef.current) });
        await shell.doc.keepCopy({ key: copyKey.current, kind: 'pdf', name: name || t('Untitled'), path, bytes });
      } catch {
        /* a copy that could not be made is made at the next change */
      }
    }, wait);
    return () => clearTimeout(timer);
  }, [pages, dirty, hasDoc, name, path, shell, protectedFile]);
  useEffect(() => {
    const away = () => { shell.doc.dropCopy({ key: copyKey.current }).catch(() => {}); };
    window.addEventListener('pagehide', away);
    return () => window.removeEventListener('pagehide', away);
  }, [shell]);

  /* ── commands ───────────────────────────────────────────────────────── */

  const canEdit = hasDoc && !protectedFile;
  const commands = useMemo(() => ({
    'file.new': { label: t('New window'), icon: 'new', key: 'Mod+N', run: () => shell.win.create({ app: 'pdf' }) },
    'file.open': { label: t('Open…'), icon: 'open', key: 'Mod+O', run: openFile },
    'file.save': { label: t('Save'), icon: 'save', key: 'Mod+S', when: () => canEdit, run: save },
    'file.saveAs': { label: t('Save As…'), icon: 'save', key: 'Mod+Shift+S', when: () => canEdit, run: saveAs },
    'file.print': { label: t('Print…'), icon: 'print', key: 'Mod+P', when: () => hasDoc, run: print },
    'edit.undo': { label: t('Undo'), icon: 'undo', key: 'Mod+Z', when: () => history.current.undo.length > 0, run: undo },
    'edit.redo': { label: t('Redo'), icon: 'redo', key: 'Mod+Y', when: () => history.current.redo.length > 0, run: redo },
    'edit.find': { label: t('Find'), icon: 'find', key: 'Mod+F', global: true, when: () => hasDoc, run: openFind },
    'find.next': { label: t('Next match'), key: 'F3', global: true, when: () => findOpen, run: () => nextHit(1) },
    'find.previous': { label: t('Previous match'), key: 'Shift+F3', global: true, when: () => findOpen, run: () => nextHit(-1) },
    'find.close': { label: t('Close Find'), key: 'Escape', global: true, when: () => findOpen, run: () => { setFindOpen(false); setQuery(''); } },
    'page.rotateRight': { label: t('Rotate Right'), icon: 'rotate', key: 'Mod+R', when: () => canEdit, run: () => rotate(90) },
    'page.rotateLeft': { label: t('Rotate Left'), icon: 'rotateLeft', key: 'Mod+L', when: () => canEdit, run: () => rotate(-90) },
    'page.delete': { label: t('Delete Pages'), icon: 'trash', when: () => canEdit, run: removePages },
    'page.extract': { label: t('Extract Pages…'), icon: 'export', when: () => canEdit, run: extract },
    'page.insert': { label: t('Insert from File…'), icon: 'import', when: () => canEdit, run: insertFromFile },
    'page.recognise': { label: t('Recognise Text…'), icon: 'scanText', when: () => canEdit && !ocr, run: askRecognise },
    'comment.cancel': { label: t('Put the tool down'), key: 'Escape', global: true, when: () => Boolean(tool) && !findOpen && !noteDraft, run: () => setTool(null) },
    'comment.unpick': { label: t('Unpick the comment'), key: 'Escape', when: () => Boolean(picked) && !tool && !findOpen, run: () => setPicked(null) },
    'comment.delete': { label: t('Delete the comment'), key: 'Delete', when: () => Boolean(picked) && canEdit, run: deletePicked },
    'page.next': { label: t('Next page'), key: 'PageDown', when: () => hasDoc && layout === 'single', run: () => stepPage(1) },
    'page.previous': { label: t('Previous page'), key: 'PageUp', when: () => hasDoc && layout === 'single', run: () => stepPage(-1) },
  }), [shell, openFile, save, saveAs, print, undo, redo, openFind, findOpen, nextHit, rotate, removePages, extract, insertFromFile, stepPage, hasDoc, canEdit, layout, ocr, askRecognise, tool, picked, noteDraft, deletePicked]);
  useCommands(commands, [commands]);

  const appMenu = useAppMenu({
    shell,
    appKey: 'pdf',
    onNew: () => shell.win.create({ app: 'pdf' }),
    onOpen: () => openFileRef.current?.(),
    extra: [{ label: t('Merge Files…'), icon: 'pdf', run: () => mergeFiles() }],
  });

  const thumbMenu = useCallback((e, key) => {
    e.preventDefault();
    if (!selectedRef.current.includes(key)) { setSelected([key]); anchor.current = key; setCurrent(key); }
    menu.open(e, [
      { label: t('Rotate Right'), icon: 'rotate', run: () => rotate(90) },
      { label: t('Rotate Left'), icon: 'rotateLeft', run: () => rotate(-90) },
      '-',
      { label: t('Extract Pages…'), icon: 'export', run: extract },
      { label: t('Insert from File…'), icon: 'import', run: insertFromFile },
      { label: t('Recognise Text…'), icon: 'scanText', run: askRecognise },
      '-',
      { label: t('Delete Pages'), icon: 'trash', run: removePages },
    ]);
  }, [menu, rotate, extract, insertFromFile, removePages, askRecognise]);

  /* ── the window ─────────────────────────────────────────────────────── */

  const pickedCount = selected.length;
  const pickedLabel = pickedCount ? tn(pickedCount, '{count} page selected', '{count} pages selected') : null;
  const currentIndex = pages.findIndex((p) => p.key === current);
  const currentSize = currentIndex >= 0 ? sizeOf(pages[currentIndex]) : null;
  const currentLabel = currentIndex >= 0 ? labelOf(sources, pages[currentIndex]) : null;

  const ribbon = (
    <Ribbon
      tabs={[
        { id: 'home', label: t('Home') },
        { id: 'pages', label: t('Pages') },
        { id: 'comment', label: t('Comment') },
        { id: 'fill', label: t('Fill & Sign') },
        { id: 'redact', label: t('Redact') },
        { id: 'view', label: t('View') },
      ]}
      active={tab}
      onTab={(id) => { setTab(id); if (id === 'comment' && hasDoc) setCommentsOpen(true); if (!['comment', 'fill', 'redact'].includes(id)) setTool(null); }}
      quick={(
        <>
          <Button icon="save" title={t('Save (Ctrl+S)')} disabled={!canEdit} onClick={save} />
          <Button icon="undo" title={t('Undo (Ctrl+Z)')} disabled={!history.current.undo.length} onClick={undo} />
          <Button icon="redo" title={t('Redo (Ctrl+Y)')} disabled={!history.current.redo.length} onClick={redo} />
        </>
      )}
    >
      {tab === 'home' ? (
        <>
          <Group label={t('File')}>
            <Button tall icon="open" label={t('Open')} title={t('Open: a PDF in this window, or in a new one when this one has a document (Ctrl+O)')} onClick={openFile} />
            <Button tall icon="save" label={t('Save')} title={t('Save: the pages as they are now; the original is replaced only after you say so (Ctrl+S)')} disabled={!canEdit} onClick={save} />
            <Rows>
              <Button icon="save" label={t('Save As')} title={t('Save As: the pages as a new PDF (Ctrl+Shift+S)')} disabled={!canEdit} onClick={saveAs} />
              <Button icon="print" label={t('Print')} title={t('Print: every page, through the printer dialog (Ctrl+P)')} disabled={!hasDoc} onClick={print} />
            </Rows>
          </Group>
          <Group label={t('Zoom')}>
            <Rows>
              <Button icon="zoomIn" label={t('Zoom In')} title={t('Zoom In (Ctrl++)')} disabled={!hasDoc} onClick={() => setZoomByHand(zoom * 1.25)} />
              <Button icon="zoomOut" label={t('Zoom Out')} title={t('Zoom Out (Ctrl+-)')} disabled={!hasDoc} onClick={() => setZoomByHand(zoom / 1.25)} />
            </Rows>
            <Rows>
              <Button icon="pagesAcross" label={t('Fit Width')} title={t('Fit Width: the page as wide as the window')} pressed={fit === 'width'} disabled={!hasDoc} onClick={() => setFit('width')} />
              <Button icon="file" label={t('Fit Page')} title={t('Fit Page: the whole page in the window')} pressed={fit === 'page'} disabled={!hasDoc} onClick={() => setFit('page')} />
            </Rows>
          </Group>
          <Group label={t('Find')}>
            <Button tall icon="find" label={t('Find')} title={t('Find: words on every page (Ctrl+F)')} disabled={!hasDoc} onClick={openFind} />
          </Group>
          <Group label={t('Scans')}>
            <Button tall icon="scanText" label={t('Recognise Text')} title={t('Recognise Text: the words in scanned pages found, so they can be searched, selected and copied; the pages look as they did')} disabled={!canEdit || Boolean(ocr)} onClick={askRecognise} data-role="recognise-text" />
          </Group>
          <Group label={t('Pages')}>
            <Rows>
              <Button icon="rotate" label={t('Rotate Right')} title={t('Rotate Right: the selected pages a quarter turn clockwise (Ctrl+R)')} disabled={!canEdit} onClick={() => rotate(90)} />
              <Button icon="trash" label={t('Delete')} title={t('Delete: the selected pages (Delete)')} disabled={!canEdit} onClick={removePages} />
            </Rows>
          </Group>
        </>
      ) : null}
      {tab === 'pages' ? (
        <>
          <Group label={t('Rotate')}>
            <Button tall icon="rotateLeft" label={t('Rotate Left')} title={t('Rotate Left: the selected pages a quarter turn anticlockwise (Ctrl+L)')} disabled={!canEdit} onClick={() => rotate(-90)} data-role="rotate-left" />
            <Button tall icon="rotate" label={t('Rotate Right')} title={t('Rotate Right: the selected pages a quarter turn clockwise (Ctrl+R)')} disabled={!canEdit} onClick={() => rotate(90)} data-role="rotate-right" />
          </Group>
          <Group label={t('Organise')}>
            <Button tall icon="trash" label={t('Delete')} title={t('Delete: the selected pages (Delete)')} disabled={!canEdit} onClick={removePages} data-role="delete-pages" />
            <Button tall icon="export" label={t('Extract')} title={t('Extract: the selected pages as a new PDF, this one left as it is')} disabled={!canEdit} onClick={extract} data-role="extract-pages" />
          </Group>
          <Group label={t('Scans')}>
            <Button tall icon="scanText" label={t('Recognise Text')} title={t('Recognise Text: the words in scanned pages found, so they can be searched, selected and copied; the pages look as they did')} disabled={!canEdit || Boolean(ocr)} onClick={askRecognise} />
          </Group>
          <Group label={t('Insert')}>
            <Button tall icon="import" label={t('Insert from File')} title={t('Insert from File: the pages of other PDFs, after the selected page')} disabled={!canEdit} onClick={insertFromFile} data-role="insert-file" />
          </Group>
          <Group label={t('Export')}>
            <Button tall icon="picture" label={t('Export as Pictures')} title={t('Export as Pictures: pages written as PNG or JPEG pictures, at the resolution you choose')} disabled={!hasDoc} onClick={() => setExporting(true)} data-role="export-pictures" />
            <Button tall icon="export" label={t('Reduce File Size')} title={t('Reduce File Size: large pictures drawn again at the detail the page needs, saved as a new file')} disabled={!canEdit} onClick={() => setCompressing(true)} data-role="compress" />
          </Group>
          <Group label={t('Split and Merge')}>
            <Button tall icon="scissors" label={t('Split')} title={t('Split: this PDF into several files, by a number of pages or at the selected pages')} disabled={!canEdit || pages.length < 2} onClick={() => setSplitOpen(true)} data-role="split" />
            <Button tall icon="pdf" label={t('Merge Files')} title={t('Merge Files: other PDFs added at the end, or combined into a new one')} disabled={protectedFile} onClick={() => mergeFiles()} data-role="merge" />
          </Group>
        </>
      ) : null}
      {tab === 'comment' ? (
        <>
          <Group label={t('Markup')}>
            <Button tall icon="highlighter" label={t('Highlight')} title={t('Highlight: the words selected, or select words with it in hand')} pressed={tool === 'highlight'} disabled={!canEdit} onClick={() => pressTool('highlight')} data-role="tool-highlight" />
            <Rows>
              <Button icon="underline" label={t('Underline')} title={t('Underline: the words selected, or select words with it in hand')} pressed={tool === 'underline'} disabled={!canEdit} onClick={() => pressTool('underline')} data-role="tool-underline" />
              <Button icon="strike" label={t('Strikethrough')} title={t('Strikethrough: a line through the words selected')} pressed={tool === 'strike'} disabled={!canEdit} onClick={() => pressTool('strike')} data-role="tool-strike" />
            </Rows>
          </Group>
          <Group label={t('Notes')}>
            <Button tall icon="note" label={t('Sticky Note')} title={t('Sticky Note: a note where you click on the page')} pressed={tool === 'note'} disabled={!canEdit} onClick={() => pressTool('note')} data-role="tool-note" />
            <Button tall icon="textbox" label={t('Text Box')} title={t('Text Box: words written on the page, in a box you drag out or click to place')} pressed={tool === 'text'} disabled={!canEdit} onClick={() => pressTool('text')} data-role="tool-text" />
          </Group>
          <Group label={t('Drawing')}>
            <Button tall icon="pen" label={t('Draw')} title={t('Draw: freehand lines on the page')} pressed={tool === 'ink'} disabled={!canEdit} onClick={() => pressTool('ink')} data-role="tool-ink" />
            <Rows>
              <Button icon="mkRect" label={t('Rectangle')} title={t('Rectangle: drag one out on the page')} pressed={tool === 'rect'} disabled={!canEdit} onClick={() => pressTool('rect')} data-role="tool-rect" />
              <Button icon="mkEllipse" label={t('Ellipse')} title={t('Ellipse: drag one out on the page')} pressed={tool === 'ellipse'} disabled={!canEdit} onClick={() => pressTool('ellipse')} data-role="tool-ellipse" />
            </Rows>
            <Rows>
              <Button icon="mkLine" label={t('Line')} title={t('Line: drag one out on the page')} pressed={tool === 'line'} disabled={!canEdit} onClick={() => pressTool('line')} data-role="tool-line" />
              <Button icon="mkArrow" label={t('Arrow')} title={t('Arrow: drag one out on the page, pointing where you let go')} pressed={tool === 'arrow'} disabled={!canEdit} onClick={() => pressTool('arrow')} data-role="tool-arrow" />
            </Rows>
          </Group>
          <Group label={t('Colour')}>
            <div className="pdf-swatches" role="radiogroup" aria-label={t('Colour')}>
              {SWATCHES.map((s) => {
                const which = tool && DEFAULT_TOOL_COLOURS[tool] ? tool : 'highlight';
                const on = (toolColours[which] || DEFAULT_TOOL_COLOURS[which]) === s.key;
                return (
                  <button
                    key={s.key}
                    type="button"
                    role="radio"
                    aria-checked={on}
                    className={`pdf-swatch${on ? ' on' : ''}`}
                    data-swatch={s.key}
                    data-tip={t('{colour}: the colour of the tool in hand', { colour: s.label() })}
                    aria-label={s.label()}
                    style={{ '--swatch': cssColour(s.rgb) }}
                    onClick={() => setToolColours((c) => ({ ...c, [which]: s.key }))}
                  />
                );
              })}
            </div>
          </Group>
          <Group label={t('Show')}>
            <Button tall icon="comment" label={t('Comments')} title={t('Comments: the pane listing every comment, to reply, resolve or delete')} pressed={commentsOpen} disabled={!hasDoc} onClick={() => setCommentsOpen((v) => !v)} data-role="comments-toggle" />
          </Group>
        </>
      ) : null}
      {tab === 'fill' ? (
        <>
          <Group label={t('Form')}>
            <Button tall icon="formField" label={t('Highlight Fields')} title={t('Highlight Fields: the form\'s fields tinted, so they are easy to find')} pressed={tintFields} disabled={!hasDoc} onClick={() => setTintFields((v) => !v)} data-role="tint-fields" />
          </Group>
          <Group label={t('Sign')}>
            <Button tall icon="signature" label={t('Signature')} title={t('Signature: your signature, typed, drawn or from a picture, placed where you click; not a digital signature')} pressed={tool === 'stamp' && signature?.kind === 'signature'} disabled={!canEdit} onClick={() => sign('signature')} data-role="sign-signature" />
            <Rows>
              <Button icon="signature" label={t('Initials')} title={t('Initials: your initials, typed, drawn or from a picture, placed where you click')} pressed={tool === 'stamp' && signature?.kind === 'initials'} disabled={!canEdit} onClick={() => sign('initials')} data-role="sign-initials" />
              <Button icon="pen" label={t('Change Signature')} title={t('Change Signature: type, draw or choose a picture of your signature again')} disabled={!canEdit} onClick={() => sign('signature', { fresh: true })} data-role="sign-change" />
            </Rows>
          </Group>
          <Group label={t('Text')}>
            <Button tall icon="textbox" label={t('Add Text')} title={t('Add Text: words written anywhere on the page, where the form has no field for them')} pressed={tool === 'text'} disabled={!canEdit} onClick={() => pressTool('text')} data-role="fill-text" />
            <Rows>
              <Button icon="calendar" label={t('Date')} title={t('Date: today\'s date written where you click')} pressed={tool === 'date'} disabled={!canEdit} onClick={() => pressTool('date')} data-role="fill-date" />
              <Button icon="check" label={t('Checkmark')} title={t('Checkmark: a tick where you click, for a box the form did not make a field')} pressed={tool === 'check'} disabled={!canEdit} onClick={() => pressTool('check')} data-role="fill-check" />
            </Rows>
          </Group>
        </>
      ) : null}
      {tab === 'redact' ? (
        <>
          <Group label={t('Mark')}>
            <Button tall icon="redact" label={t('Mark Text')} title={t('Mark Text: the words selected, or select words with it in hand, marked for redaction')} pressed={tool === 'redact-text'} disabled={!canEdit} onClick={() => pressTool('redact-text')} data-role="redact-text" />
            <Button tall icon="marquee" label={t('Mark Area')} title={t('Mark Area: drag out an area of the page, words and pictures, to redact')} pressed={tool === 'redact-area'} disabled={!canEdit} onClick={() => pressTool('redact-area')} data-role="redact-area" />
            <Button tall icon="find" label={t('Mark Matches')} title={t('Mark Matches: every match of the words in Find, on every page, marked for redaction')} disabled={!canEdit} onClick={markSearchResults} data-role="redact-search" />
          </Group>
          <Group label={t('Apply')}>
            <Button tall icon="check" label={t('Apply Redactions')} title={t('Apply Redactions: what the marks cover removed from the pages, saved as a new file')} disabled={!canEdit} onClick={askApply} data-role="redact-apply" />
          </Group>
          <Group label={t('Review')}>
            <Button tall icon="comment" label={t('Marks')} title={t('Marks: the Comments pane, where each mark can be found and deleted')} pressed={commentsOpen} disabled={!hasDoc} onClick={() => setCommentsOpen((v) => !v)} data-role="redact-marks" />
          </Group>
        </>
      ) : null}
      {tab === 'view' ? (
        <>
          <Group label={t('Page Display')}>
            <Button tall icon="file" label={t('Single Page')} title={t('Single Page: one page at a time (Page Up and Page Down)')} pressed={layout === 'single'} disabled={!hasDoc} onClick={() => setLayout('single')} data-role="layout-single" />
            <Button tall icon="pagesAcross" label={t('Two Pages')} title={t('Two Pages: pages side by side')} pressed={layout === 'two'} disabled={!hasDoc} onClick={() => setLayout('two')} data-role="layout-two" />
            <Button tall icon="pagesDown" label={t('Continuous')} title={t('Continuous: every page in one column')} pressed={layout === 'continuous'} disabled={!hasDoc} onClick={() => setLayout('continuous')} data-role="layout-continuous" />
          </Group>
          <Group label={t('Show')}>
            <Button tall icon="grid" label={t('Thumbnails')} title={t('Thumbnails: the strip of pages beside the document')} pressed={strip} onClick={() => setStrip((v) => !v)} data-role="thumbnails" />
          </Group>
          <Group label={t('Zoom')}>
            <Rows>
              <Button icon="pagesAcross" label={t('Fit Width')} title={t('Fit Width: the page as wide as the window')} pressed={fit === 'width'} disabled={!hasDoc} onClick={() => setFit('width')} />
              <Button icon="file" label={t('Fit Page')} title={t('Fit Page: the whole page in the window')} pressed={fit === 'page'} disabled={!hasDoc} onClick={() => setFit('page')} />
            </Rows>
            <Button tall icon="zoomIn" label={t('Actual Size')} title={t('Actual Size: the page at 100% (Ctrl+0)')} disabled={!hasDoc} onClick={() => setZoomByHand(1)} />
          </Group>
        </>
      ) : null}
    </Ribbon>
  );

  const status = (
    <>
      <span data-role="page-status">{hasDoc && currentIndex >= 0 ? (currentLabel && currentLabel !== String(currentIndex + 1) ? t('Page {label} ({page} of {count})', { label: currentLabel, page: currentIndex + 1, count: pages.length }) : t('Page {page} of {count}', { page: currentIndex + 1, count: pages.length })) : path || name || t('No PDF open')}</span>
      {hasDoc ? (
        <>
          <Button ghost icon="chevronUp" title={t('Previous page')} disabled={currentIndex <= 0} onClick={() => stepPage(-1)} />
          <Button ghost icon="chevronDown" title={t('Next page')} disabled={currentIndex >= pages.length - 1} onClick={() => stepPage(1)} />
        </>
      ) : null}
      <Spacer />
      {pickedLabel ? <Chip>{pickedLabel}</Chip> : null}
      {currentSize ? <Chip>{t('{width} × {height} mm', { width: Math.round(currentSize.width / 72 * 25.4), height: Math.round(currentSize.height / 72 * 25.4) })}</Chip> : null}
      {protectedFile ? <Chip>{t('Protected: read only')}</Chip> : null}
      {hasDoc ? <ZoomSlider value={zoom} min={ZOOM_MIN} max={4} onChange={setZoomByHand} onReset={() => setZoomByHand(1)} /> : null}
    </>
  );

  return (
    <AppFrame app={app} shell={shell} title={name || app.name} dirty={hasDoc && dirty} menu={appMenu} ribbon={ribbon} status={status}>
      <style>{CSS + OCR_CSS}</style>
      {!hasDoc ? (
        failed ? (
          <Empty icon={locked ? 'lock' : 'pdf'} title={locked ? t('This PDF is protected') : t('This file could not be opened')} action={<div className="pdf-empty-actions"><Button primary icon="open" label={t('Open another')} onClick={openFile} /></div>}>
            {failed}
          </Empty>
        ) : (
          <div className="pdf-empty" style={{ flex: 1, display: 'grid', placeItems: 'center' }}>
            <Empty
              icon="pdf"
              title={busy || t('No PDF open')}
              action={busy ? <Spinner /> : (
                <div className="pdf-empty-actions">
                  <Button primary icon="open" label={t('Open a PDF')} onClick={openFile} />
                  <Button icon="pdf" label={t('Merge Files')} title={t('Merge Files: several PDFs combined into one')} onClick={() => mergeFiles()} />
                </div>
              )}
            >
              {busy ? null : t('Open a PDF to read it, or to rotate, move, delete, extract and insert pages. Merge Files combines several into one. Your file is changed only when you save over it.')}
            </Empty>
          </div>
        )
      ) : (
        <div className="pdf-app" data-pages={pages.length} data-layout={layout}>
          {strip ? (
            <div className="pdf-side" style={{ width: stripWidth }}>
              <div className="pdf-side-head">
                <div className="pdf-side-tabs" role="tablist">
                  <button type="button" role="tab" aria-selected={side === 'pages'} className={side === 'pages' ? 'on' : ''} data-role="side-pages" onClick={() => setSide('pages')}>{t('Pages')}</button>
                  <button type="button" role="tab" aria-selected={side === 'bookmarks'} className={side === 'bookmarks' ? 'on' : ''} data-role="side-bookmarks" onClick={() => setSide('bookmarks')}>{t('Bookmarks')}</button>
                </div>
                <span className="spacer" />
                {side === 'pages' ? <span className="pdf-side-count">{pages.length}</span> : null}
              </div>
              {side === 'bookmarks' ? (
                <BookmarksPane pages={pages} sources={sources} current={current} onGo={(key) => goTo(key)} onLink={(url) => shell.shell.openExternal({ url }).catch(() => {})} />
              ) : (
              <ThumbStrip
                pages={pages}
                sources={sources}
                selected={selected}
                current={current}
                width={stripWidth}
                onPick={onPick}
                onMove={movePicked}
                onMenu={thumbMenu}
                onKey={onStripKey}
              />
              )}
            </div>
          ) : null}
          {strip ? (
            <div
              className="pdf-split"
              ref={stripRef}
              onMouseDown={(e) => {
                const x0 = e.clientX;
                const w0 = stripWidth;
                const rtl = getComputedStyle(document.documentElement).direction === 'rtl';
                const move = (ev) => setStripWidth(Math.max(120, Math.min(360, w0 + (rtl ? -1 : 1) * (ev.clientX - x0))));
                const up = () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
                window.addEventListener('mousemove', move);
                window.addEventListener('mouseup', up);
              }}
            />
          ) : null}
          <div className="pdf-main">
            {protectedFile ? <div className="pdf-notice">{t('This PDF is protected, so its pages can be read but not changed or saved yet.')}</div> : null}
            {tab === 'redact' && !protectedFile ? (
              <div className="pdf-notice info" data-role="redact-notice">
                {t('Marks show what will be removed; nothing is removed until you apply them. Applying saves a new file and leaves the original as it is.')}
              </div>
            ) : null}
            {tab === 'fill' && !protectedFile ? (
              <div className="pdf-notice info" data-role="sign-notice">
                {tool === 'stamp' ? t('Click the page where your {kind} goes, or drag out the box it fills. Press Esc to put it down.', { kind: signature?.kind === 'initials' ? t('initials') : t('signature') }) : t('Type into the form\'s fields, tick its boxes and choose from its lists. A signature placed here is a picture of your signature, not a digital signature: it does not prove who signed, or that the PDF has not changed since.')}
              </div>
            ) : null}
            {findOpen ? (
              <div className="pdf-find" role="search">
                <Input
                  ref={findRef}
                  autoFocus
                  value={query}
                  placeholder={t('Find in the document')}
                  aria-label={t('Find in the document')}
                  data-role="find-input"
                  onChange={(e) => setQuery(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') { e.preventDefault(); nextHit(e.shiftKey ? -1 : 1); }
                  }}
                />
                <span className="pdf-find-count" data-count={hits.length} data-at={hitAt} data-query={query.trim() ? 'yes' : 'no'}>
                  {!query.trim() ? '' : hits.length ? t('{at} of {count}', { at: hitAt + 1, count: hits.length }) : t('No matches')}
                </span>
                <Button ghost icon="chevronUp" title={t('Previous match (Shift+F3)')} disabled={!hits.length} onClick={() => nextHit(-1)} />
                <Button ghost icon="chevronDown" title={t('Next match (F3)')} disabled={!hits.length} onClick={() => nextHit(1)} />
                <Button ghost icon="close" title={t('Close Find (Esc)')} onClick={() => { setFindOpen(false); setQuery(''); }} />
              </div>
            ) : null}
            <PageView
              pages={pages}
              sources={sources}
              zoom={zoom}
              layout={layout}
              current={current}
              onCurrent={onCurrent}
              onText={onText}
              scrollerRef={viewRef}
              overlay={overlay}
              onMouseUp={onViewMouseUp}
            />
            {busy ? <div className="pdf-busy"><Spinner /></div> : null}
          </div>
          {commentsOpen ? (
            <CommentsPane
              items={commentItems(pages, sources)}
              picked={picked}
              onPick={pickComment}
              onReply={replyTo}
              onResolve={resolveComment}
              onDelete={deleteComment}
              onEdit={editComment}
              onClose={() => setCommentsOpen(false)}
            />
          ) : null}
        </div>
      )}
      {noteDraft ? <NoteEditor draft={noteDraft} onPost={postNote} onCancel={() => setNoteDraft(null)} /> : null}
      {exporting ? <ExportPicturesDialog count={pages.length} picked={selected.length} onExport={exportPictures} onClose={() => setExporting(false)} /> : null}
      {compressing ? <CompressDialog size={formatBytes(Object.values(sources).filter((s) => pages.some((p) => p.source === s.id)).reduce((n, s) => n + (s.info.size || 0), 0))} onCompress={compressFile} onClose={() => setCompressing(false)} /> : null}
      {applying ? (
        <ApplyRedactionsDialog {...redactionMarks()} onApply={applyRedactions} onClose={() => setApplying(false)} />
      ) : null}
      {signing ? (
        <SignatureDialog
          kind={signing}
          shell={shell}
          onClose={() => setSigning(null)}
          onDone={(made) => { setSigning(null); setSignature(made); setTool('stamp'); setPicked(null); }}
        />
      ) : null}
      {splitOpen ? <SplitDialog count={pages.length} picked={selected.filter((k) => pages.findIndex((p) => p.key === k) > 0).length} onSplit={split} onClose={() => setSplitOpen(false)} /> : null}
      {ocr ? (
        <RecogniseDialog
          count={pages.length}
          picked={selected.length}
          current={currentIndex + 1}
          progress={ocr.progress || null}
          onStart={recognise}
          onCancel={cancelRecognise}
          onClose={() => setOcr(null)}
        />
      ) : null}
      {printing ? (
        <PrintPages images={printing} />
      ) : null}
      {menu.node}
    </AppFrame>
  );
}

/** The pages as pictures, outside the window's own root, for the printer and nothing else. */
function PrintPages({ images }) {
  const [host] = useState(() => {
    const div = document.createElement('div');
    div.className = 'pdf-print-root';
    document.body.append(div);
    return div;
  });
  useEffect(() => () => host.remove(), [host]);
  useEffect(() => {
    host.replaceChildren(...images.map((img) => {
      const el = document.createElement('img');
      el.src = img.url;
      el.alt = '';
      return el;
    }));
  }, [host, images]);
  return null;
}
