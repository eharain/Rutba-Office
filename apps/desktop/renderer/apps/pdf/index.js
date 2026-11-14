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
// A PDF protected with a password asks for it (the Password dialog, as the
// other apps ask), pdf.js draws with it and the main process reads the file
// with it, both holding it in memory for that file only. Everything written
// keeps the protection, and what the file's permissions do not allow is
// refused in a sentence, until its owner password is given.
//
// Saving is careful, as a page tool's must be: Save asks before it replaces
// the file the pages came from (once a window), and offers to save a copy.
// While there is unsaved work, a copy is kept in the recovery folder for
// the launcher to offer back after a crash, as Video and Image do.

import { RecentFiles } from '../../start-page.js';
import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { createReader } from '../../read-aloud/reader.js';
import { ReadAloudBar } from '../../read-aloud/bar.js';
import {
  Ribbon, Group, Rows, Button, Spacer, Chip, Empty, Spinner, Dialog, Input, ZoomSlider, useToast, useCommands, useMenu, useWheelZoom,
  basename, dirname, formatBytes, t, tn,
} from '@rutba/office-ui';
import { findAll } from '@rutba/editing/find';
import { AppFrame, useAppMenu, useDirtyGuard, useFileDrop } from '../../shell.js';
import { usePasswordGate } from '../../protect.js';
import { openDocument, pageText, pageImage } from '../../pdfjs.js';
import {
  pagesOf, rotatePages, deletePages, insertPages, movePages, rangeKeys, rangeText, chosenKeys, shownSize, splitGroups,
} from './pages.js';
import { ThumbStrip, PageView, PX_PER_PT } from './view.js';
import { CSS } from './css.js';
import { RecogniseDialog, recognisePages } from './ocr.js';
import { OCR_CSS } from '../../ocr-ui.js';
import { FieldsPane, FoundMark, FOUND_CSS } from './found.js';
import { RECORD_FIELDS } from '@rutba/ocr/record';
import {
  AnnotLayer, CommentsPane, NoteEditor, commentItems, selectionQuads, swatch, cssColour,
  MARKUP_TOOLS, TOOL_TYPES, SWATCHES, DEFAULT_TOOL_COLOURS,
} from './annots.js';
import { FieldLayer, SignatureDialog, savedSignature, signatureStamp } from './fields.js';
import { searchQuads, decodePlan, ApplyRedactionsDialog } from './redact.js';
import { BookmarksPane, labelOf } from './nav.js';

/** A tool's name as its button says it. */
function toolName(tool) {
  switch (tool) {
    case 'highlight': return t('common.highlight');
    case 'underline': return t('common.underline');
    case 'strike': return t('common.strikethrough');
    case 'note': return t('pdf.sticky-note');
    case 'text': return t('common.text-box');
    case 'ink': return t('common.draw');
    case 'rect': return t('common.rectangle');
    case 'ellipse': return t('common.ellipse');
    case 'line': return t('common.line');
    case 'arrow': return t('common.arrow');
    case 'redact-area': return t('pdf.mark-area');
    case 'redact-text': return t('pdf.mark-text');
    case 'stamp': return t('common.signature');
    case 'date': return t('common.date');
    case 'check': return t('pdf.checkmark');
    default: return t('pdf.the-tool');
  }
}
import { ExportPicturesDialog, CompressDialog, pageBytes, encodePlan } from './export.js';

const AUTOSAVE_FIRST_MS = 1500;
const AUTOSAVE_EVERY_MS = 30000;
const ZOOM_MIN = 0.1;
const ZOOM_MAX = 6.4;
const PDF_FILTER = () => [{ name: t('common.pdf-document-pdf'), extensions: ['pdf'] }];

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
      title={t('common.split')}
      onClose={onClose}
      actions={(
        <>
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary label={t('common.split')} data-role="split-go" disabled={files < 2} onClick={() => onSplit(mode === 'every' ? { every: Math.max(1, every || 1) } : { picked: true })} />
        </>
      )}
    >
      <div className="pdf-split-form">
        <label className="pdf-split-choice">
          <input type="radio" name="pdf-split" checked={mode === 'every'} onChange={() => setMode('every')} />
          <span>{t('pdf.into-files-of')}</span>
          <Input type="number" min={1} max={count} value={every} data-role="split-every" onChange={(e) => { setMode('every'); setEvery(Math.max(1, Number(e.target.value) || 1)); }} />
          <span>{t('pdf.pages-each')}</span>
        </label>
        <label className="pdf-split-choice">
          <input type="radio" name="pdf-split" checked={mode === 'picked'} disabled={!picked} onChange={() => setMode('picked')} />
          <span>{t('pdf.new-file-each')}</span>
        </label>
        <div className="pdf-split-note">{tn(files, 'pdf.makes-count-file')}</div>
      </div>
    </Dialog>
  );
}

export default function PdfApp({ app, shell, boot }) {
  const toast = useToast();
  const menu = useMenu();
  const gate = usePasswordGate();
  const { ask: askPassword, done: passwordAsked } = gate;
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
  // The words the hits were found for, on the count as data-for: a check waits for its own search to finish rather than reading the last one's count.
  const [hitsFor, setHitsFor] = useState('');
  const [splitOpen, setSplitOpen] = useState(false);
  const [printing, setPrinting] = useState(null);
  // Recognise Text: null, { } while its dialog asks, { progress } while it reads.
  const [ocr, setOcr] = useState(null);
  const ocrStop = useRef(false);
  const ocrJob = useRef(null);
  // Find Fields: what Recognise Text found on each page (by the page's key), the pane, and the field marked on its page.
  const [found, setFound] = useState({});
  const [fieldsOpen, setFieldsOpen] = useState(false);
  const [foundMark, setFoundMark] = useState(null);
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
  // Protection: a file read with its password, one that could not be (locked: read here, never written), and what its permissions allow.
  const sourceList = Object.values(sources);
  const protectedFile = sourceList.some((s) => s.info.encrypted);
  const lockedFile = sourceList.some((s) => s.info.locked);
  const allowed = (kind) => !lockedFile && sourceList.every((s) => s.info.permissions?.[kind] !== false);
  const restricted = protectedFile && !lockedFile && !['assemble', 'change', 'comment', 'fill'].every(allowed);

  /* ── sources ────────────────────────────────────────────────────────── */

  /**
   * Read a PDF into the window: its structure in the main process, its pages
   * for drawing here. A file that needs a password asks for it (the Password
   * dialog, again while it is not right); the password is kept with the
   * source, in memory only, for drawing that file again. Pages being
   * inserted into another document (`insert`) are refused from a protected
   * file before anything is asked.
   */
  const addSource = useCallback(async (file, { bytes = null, label = null, insert = false, password: known = null } = {}) => {
    const raw = bytes || (await shell.fs.read({ path: file })).bytes;
    let info = await shell.pdf.open(bytes ? { bytes: raw, name: label || basename(file || '') || t('pdf.recovered-pdf') } : { path: file });
    if (insert && info.encrypted) {
      shell.pdf.close({ id: info.id }).catch(() => {});
      const e = new Error(t('pdf.name-protected-pages', { name: info.name }));
      e.protected = true;
      throw e;
    }
    let doc;
    let password = known;
    try {
      doc = await openDocument(raw, {
        password: known,
        ask: insert ? null : async ({ wrong }) => {
          password = await askPassword({ name: info.name, wrong });
          return password;
        },
      });
    } catch (err) {
      shell.pdf.close({ id: info.id }).catch(() => {});
      if (err.password) {
        const e = new Error(t('common.name-protected-password', { name: info.name }));
        e.password = true;
        throw e;
      }
      throw new Error(t('pdf.name-could-not', { name: info.name, reason: err.message }));
    } finally {
      passwordAsked();
    }
    // The main process reads the file with the password pdf.js took, for its comments, fields, bookmarks and every change.
    if (info.locked && password !== null) {
      const unlocked = await shell.pdf.unlock({ id: info.id, password });
      if (!unlocked.wrong) info = unlocked;
    }
    // Read here but not there (a protection it cannot read): the pages as pdf.js finds them, to read only.
    if (!info.pages.length && doc.numPages) {
      const pages = [];
      for (let i = 0; i < doc.numPages; i++) {
        const page = await doc.getPage(i + 1);
        const [x0, y0, x1, y1] = page.view;
        pages.push({ width: Math.abs(x1 - x0), height: Math.abs(y1 - y0), rotate: page.rotate || 0 });
      }
      info = { ...info, pages };
    }
    const [annots, fields, nav] = await Promise.all([
      shell.pdf.annotations({ id: info.id }).catch(() => []),
      shell.pdf.fields({ id: info.id }).catch(() => []),
      shell.pdf.navigation({ id: info.id }).catch(() => null),
    ]);
    const source = { id: info.id, name: info.name, path: file || null, info, doc, annots, fields, nav, password };
    setSources((s) => ({ ...s, [info.id]: source }));
    sourcesRef.current = { ...sourcesRef.current, [info.id]: source };
    return source;
  }, [shell, askPassword, passwordAsked]);

  const resetHistory = () => {
    history.current = { undo: [], redo: [] };
    setHistoryTick((n) => n + 1);
  };

  // What was last opened, for the protected screen's Enter Password to try again.
  const lastOpen = useRef(null);

  /** Open a PDF as this window's document. */
  const openPdf = useCallback(async (file, { bytes = null, recovered = null, password = null } = {}) => {
    lastOpen.current = { file, bytes, recovered };
    setBusy(t('common.opening'));
    setFailed(null);
    try {
      const source = await addSource(file, { bytes, label: recovered?.name, password });
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
      if (source.info.repaired) toast(t('pdf.name-damaged-read', { name: source.name }), { ms: 7000 });
    } catch (err) {
      setFailed(err.message);
      setLocked(Boolean(err.password));
    } finally {
      setBusy(null);
    }
  }, [addSource, shell, toast]);

  const pickPdfs = useCallback(async (multiple) => {
    const paths = await shell.dialog.open({ title: multiple ? t('pdf.choose-pdfs') : t('common.open'), filters: [...PDF_FILTER(), { name: t('common.all-files'), extensions: ['*'] }], multiple });
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

  /**
   * Whether a change is refused, said in a sentence: a file that could not
   * be read with its password is read only; a protected file's permissions
   * name what its owner allows: 'assemble' (pages moved, turned, deleted,
   * inserted, extracted), 'change' (its content), 'comment', 'fill' (its
   * form), and 'save', which any of them allows.
   */
  const guardLocked = useCallback((kind = 'change') => {
    const list = Object.values(sourcesRef.current);
    if (list.some((s) => s.info.locked)) {
      toast(t('pdf.pdf-protected-way'), { tone: 'bad', ms: 6000 });
      return true;
    }
    if (kind === 'save' || list.every((s) => s.info.permissions?.[kind] !== false)) return false;
    if (kind === 'comment') toast(t('pdf.whoever-protected-pdf-does'), { tone: 'bad', ms: 7000 });
    else if (kind === 'fill') toast(t('pdf.whoever-protected-pdf-does-not-allow'), { tone: 'bad', ms: 7000 });
    else toast(t('pdf.whoever-protected-pdf-does-not'), { tone: 'bad', ms: 7000 });
    return true;
  }, [toast]);

  /** The owner password, for a protected file whose permissions keep some changes back: asked for, and every change allowed once it is right. */
  const enterOwnerPassword = useCallback(async () => {
    const list = Object.values(sourcesRef.current).filter((s) => s.info.encrypted && !s.info.locked && s.info.access !== 'owner');
    if (!list.length) return;
    let wrong = false;
    try {
      for (;;) {
        const password = await askPassword({ name: list[0].name, wrong, sub: t('pdf.allows-only-some') });
        if (password === null) return;
        const answers = await Promise.all(list.map((s) => shell.pdf.unlock({ id: s.id, password })));
        if (answers.every((a) => a.access === 'owner')) {
          const next = { ...sourcesRef.current };
          list.forEach((s, i) => { next[s.id] = { ...s, info: answers[i], password }; });
          sourcesRef.current = next;
          setSources(next);
          toast(t('pdf.owner-password-every'), { tone: 'good', ms: 4000 });
          return;
        }
        wrong = true;
      }
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 6000 });
    } finally {
      passwordAsked();
    }
  }, [askPassword, passwordAsked, shell, toast]);

  const rotate = useCallback((by) => {
    if (guardLocked('assemble')) return;
    const keys = chosen();
    if (keys.length) edit(rotatePages(pagesRef.current, keys, by));
  }, [chosen, edit, guardLocked]);

  const removePages = useCallback(() => {
    if (guardLocked('assemble')) return;
    const keys = chosen();
    if (!keys.length) return;
    if (keys.length >= pagesRef.current.length) {
      toast(t('pdf.pdf-needs-least'), { tone: 'bad', ms: 5000 });
      return;
    }
    const list = pagesRef.current;
    const first = list.findIndex((p) => keys.includes(p.key));
    const next = deletePages(list, keys);
    edit(next, { select: [], current: next[Math.min(first, next.length - 1)]?.key });
    toast(tn(keys.length, 'pdf.count-page-deleted'), { ms: 2500 });
  }, [chosen, edit, guardLocked, toast]);

  const movePicked = useCallback((to, key) => {
    if (guardLocked('assemble')) return;
    const sel = selectedRef.current;
    const keys = sel.includes(key) ? pagesRef.current.filter((p) => sel.includes(p.key)).map((p) => p.key) : [key];
    const next = movePages(pagesRef.current, keys, to);
    if (next === pagesRef.current) {
      // Dropped among the pages being moved: nothing to do, and the person is told why.
      if (keys.length > 1) toast(t('pdf.selected-pages-already'), { ms: 5000 });
      return;
    }
    edit(next, { select: keys, current: key });
  }, [edit, guardLocked, toast]);

  /** Pages of other files: after the last picked page, or at the end. */
  const insertFiles = useCallback(async (files, { at = null, announce = true } = {}) => {
    if (guardLocked('assemble') || !files.length) return;
    setBusy(t('pdf.reading'));
    const added = [];
    try {
      for (const file of files) {
        let source;
        try {
          source = await addSource(file, { insert: true });
        } catch (err) {
          if (!err.protected) throw err;
          toast(err.message, { tone: 'bad', ms: 6000 });
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
    if (announce) toast(tn(added.length, 'pdf.count-page-inserted'), { ms: 3000 });
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
        setName(t('pdf.combined-pdf'));
        setDirty(true);
      }
      toast(tn(files.length, 'pdf.count-file-combined'), { ms: 3000 });
      return;
    }
    await insertFiles(files, { at: pagesRef.current.length, announce: false });
    toast(tn(files.length, 'pdf.count-file-added'), { ms: 3000 });
  }, [pickPdfs, openPdf, insertFiles, toast]);

  /* ── recognising text ──────────────────────────────────────────────── */

  const recognise = useCallback(async ({ scope, languages, fields = false }) => {
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
        fields,
      });
      if (ocrStop.current) {
        toast(t('pdf.recognise-text-cancelled'), { ms: 4000 });
        return;
      }
      if (!out.scanned) {
        toast(t('pdf.every-page-chosen'), { ms: 5000 });
        return;
      }
      // What each page offered, by the keys of the pages showing it (a page's key outlives its source).
      const offered = {};
      for (const f of out.found || []) {
        for (const p of pagesRef.current) if (p.source === f.source && p.page === f.page) offered[p.key] = f.found;
      }
      const swap = {};
      for (const [id, written] of out.bySource) {
        if (!written.length) continue;
        const { bytes, written: pagesWritten, ...info } = await shell.pdf.addText({ id, pages: written });
        const password = sourcesRef.current[id]?.password || null;
        const doc = await openDocument(bytes, { password });
        const source = { id: info.id, name: info.name, path: sourcesRef.current[id]?.path || null, info, doc, annots: sourcesRef.current[id]?.annots || [], fields: sourcesRef.current[id]?.fields || [], nav: sourcesRef.current[id]?.nav || null, password };
        sourcesRef.current = { ...sourcesRef.current, [info.id]: source };
        setSources((s) => ({ ...s, [info.id]: source }));
        swap[id] = info.id;
      }
      if (Object.keys(swap).length) edit(pagesRef.current.map((p) => (swap[p.source] ? { ...p, source: swap[p.source] } : p)));
      const offeredCount = Object.keys(offered).length;
      if (offeredCount) {
        setFound((all) => ({ ...all, ...offered }));
        setFoundMark(null);
        setCommentsOpen(false);
        setFieldsOpen(true);
      }
      if (offeredCount) toast(tn(offeredCount, 'pdf.text-recognised-fields'), { tone: 'good', ms: 6000 });
      else if (out.words) toast(tn(out.done, 'pdf.text-recognised-count'), { tone: 'good', ms: 5000 });
      else toast(t('pdf.no-words-found'), { ms: 5000 });
    } catch (err) {
      if (err?.code === 'OCR_CANCELLED' || ocrStop.current) toast(t('pdf.recognise-text-cancelled'), { ms: 4000 });
      else toast(t('common.text-could-not', { reason: err.message }), { tone: 'bad', ms: 8000 });
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
    // Check boxes made and redaction marks change the file; a value fills its form; the rest are comments.
    const redacting = list.some((x) => (x.changes?.add || []).some((a) => a.type === 'Redact'));
    if (guardLocked(list.some((x) => x.checkboxes) || redacting ? 'change' : list.some((x) => x.values) ? 'fill' : 'comment')) return null;
    const swap = {};
    let first = null;
    try {
      for (const { source: asked, changes, values, checkboxes } of list) {
        // The version the pages show now: what this change was asked of, unless a change before it replaced it.
        let id = asked;
        while (!pagesRef.current.some((p) => p.source === id) && newer.current.has(id)) id = newer.current.get(id);
        const { bytes, added, annotations, fields, ...info } = checkboxes
          ? await shell.pdf.addCheckboxes({ id, boxes: checkboxes })
          : values
            ? await shell.pdf.fill({ id, values })
            : await shell.pdf.annotate({ id, changes });
        const old = sourcesRef.current[id];
        const doc = await openDocument(bytes, { password: old?.password || null });
        const source = { id: info.id, name: info.name, path: old?.path || null, info, doc, annots: annotations, fields, nav: old?.nav || null, password: old?.password || null };
        sourcesRef.current = { ...sourcesRef.current, [info.id]: source };
        setSources((s) => ({ ...s, [info.id]: source }));
        swap[id] = info.id;
        newer.current.set(id, info.id);
        if (!first && added?.length) first = added[0];
      }
    } catch (err) {
      toast(list.some((x) => x.checkboxes) ? t('pdf.check-boxes-could', { reason: err.message }) : list.some((x) => x.values) ? t('pdf.form-could-not', { reason: err.message }) : t('pdf.comment-could-not', { reason: err.message }), { tone: 'bad', ms: 8000 });
      return null;
    }
    if (Object.keys(swap).length) edit(pagesRef.current.map((p) => (swap[p.source] ? { ...p, source: swap[p.source] } : p)));
    if (pick && first) setPicked(first);
    return first;
  };
  writeRef.current = writeChanges;

  /* ── fields found ───────────────────────────────────────────────────── */

  const toggleFields = useCallback(() => {
    setCommentsOpen(false);
    setFieldsOpen((v) => !v);
    setFoundMark(null);
  }, []);

  /** A field or a tick box pressed in the pane: marked on its page, the page brought into view. */
  const markFound = useCallback((mark) => {
    setFoundMark(mark);
    goToRef.current?.(mark.key);
  }, []);

  /** An invoice's details as one row a sheet takes cell by cell. */
  const copyFoundRow = useCallback(async (item) => {
    const byField = Object.fromEntries(item.found.fields.map((x) => [x.field, x.text]));
    const cells = RECORD_FIELDS.map((name) => byField[name] ?? '');
    const esc = (v) => String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    await shell.clipboard.write({ text: cells.join('\t'), html: `<table><tr>${cells.map((v) => `<td>${esc(v)}</td>`).join('')}</tr></table>` });
    toast(t('pdf.copied-row-paste'), { tone: 'good' });
  }, [shell, toast]);

  /** A form's tick boxes made check box fields, ticked as they were read: one step Undo takes back. */
  const makeCheckboxes = useCallback(async (item, boxes) => {
    const entry = pagesRef.current.find((p) => p.key === item.key);
    if (!entry || !boxes.length) return;
    const first = await changeComments([{ source: entry.source, checkboxes: boxes.map((c) => ({ page: entry.page, rect: c.rect, name: c.label, checked: c.checked })) }]);
    if (first) toast(tn(boxes.length, 'pdf.count-check-box'), { tone: 'good', ms: 5000 });
  }, [changeComments, toast]);

  const colourOf = useCallback((which) => swatch(toolColours[which] || DEFAULT_TOOL_COLOURS[which]).rgb, [toolColours]);

  /** The words selected marked with a highlight, an underline, a strikethrough or for redaction. Answers whether there were any. */
  const markSelection = useCallback((which) => {
    const quads = selectionQuads(viewRef.current, viewports);
    // A redaction mark records the words it covers, so applying it takes them out of the rest of the file too.
    const words = which === 'redact-text' ? (window.getSelection()?.toString() || '').replace(/\s+/g, ' ').trim() : '';
    const bySource = new Map();
    for (const [key, list] of Object.entries(quads)) {
      const entry = pagesRef.current.find((p) => p.key === key);
      if (!entry || !list.length) continue;
      if (!bySource.has(entry.source)) bySource.set(entry.source, []);
      const spec = which === 'redact-text'
        ? { page: entry.page, type: 'Redact', quads: list, contents: words }
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
    if (guardLocked(/^redact/.test(which) ? 'change' : 'comment')) return;
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
    if (guardLocked(/^redact/.test(d.tool) ? 'change' : ['stamp', 'date', 'check'].includes(d.tool) ? 'fill' : 'comment')) return;
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
      // One date a click: the tool goes back once it is placed, as Sign's does.
      setTool(null);
      const words = new Intl.DateTimeFormat('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }).format(new Date());
      const ctx2d = document.createElement('canvas').getContext('2d');
      ctx2d.font = '11px Helvetica, Arial, sans-serif';
      const wide = Math.ceil(ctx2d.measureText(words).width) + 10;
      changeComments([{ source: d.source, changes: { add: [{ page: d.page, type: 'FreeText', rect: [d.point[0], d.point[1] - 18, d.point[0] + wide, d.point[1]], contents: words, fontSize: 11, color: [0.06, 0.14, 0.36], subject: 'Date' }] } }], { pick: true });
      return;
    }
    if (d.tool === 'check') {
      // A tick, 14 points square, centred on the click; the tool goes back once it is placed.
      setTool(null);
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
    if (guardLocked('fill')) return;
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
      toast(t('pdf.type-words-find'), { ms: 5000 });
      return;
    }
    setBusy(t('pdf.finding'));
    const bySource = new Map();
    let count = 0;
    try {
      const seen = new Set();
      for (const p of pagesRef.current) {
        const id = `${p.source}:${p.page}`;
        if (seen.has(id)) continue;
        seen.add(id);
        const quads = await searchQuads(sourcesRef.current[p.source].doc, p.page, q, () => shell.pdf.matchQuads({ source: p.source, page: p.page, query: q }));
        if (!quads.length) continue;
        count += quads.length;
        if (!bySource.has(p.source)) bySource.set(p.source, []);
        bySource.get(p.source).push({ page: p.page, type: 'Redact', quads, contents: q });
      }
    } finally {
      setBusy(null);
    }
    if (!bySource.size) { toast(t('pdf.no-matches-mark'), { ms: 4000 }); return; }
    await changeComments([...bySource].map(([source, add]) => ({ source, changes: { add } })));
    toast(tn(count, 'pdf.count-match-marked'), { ms: 5000 });
  }, [guardLocked, query, findOpen, toast, changeComments]);

  /**
   * Apply Redactions, first step: the pictures under the marks decoded here,
   * and what applying would take out (besides the pages' words) and leave
   * off which pages, worked out before the dialog asks.
   */
  const askApply = useCallback(async () => {
    if (guardLocked()) return;
    if (!redactionMarks().marks) {
      toast(t('pdf.nothing-marked-redaction'), { ms: 5000 });
      return;
    }
    setBusy(t('pdf.checking-marks-cover'));
    try {
      const list = pagesRef.current;
      const sources = [...new Set(list.map((p) => p.source))];
      const plan = await shell.pdf.redactPlan({ sources });
      const images = await decodePlan(plan);
      const preview = await shell.pdf.redactPreview({ pages: entriesOf(list), images });
      setApplying({ images, preview });
    } catch (err) {
      toast(t('pdf.redactions-could-not', { reason: err.message }), { tone: 'bad', ms: 8000 });
    } finally {
      setBusy(null);
    }
  }, [guardLocked, redactionMarks, toast, shell]);

  // Save, for Apply Redactions to call: it is made further down.
  const saveRef = useRef(null);

  /**
   * The marks applied: the pages written whole as a new file, which then
   * opens here. Marks not yet saved are asked about first, as closing the
   * window would ask: saved, left, or kept open here with the redacted copy
   * in a window of its own.
   */
  const applyRedactions = useCallback(async () => {
    const images = applying?.images || {};
    setApplying(false);
    const list = pagesRef.current;
    const folder = path ? dirname(path) : (await shell.app.paths().catch(() => ({}))).documents || '';
    const stem = stemOf(path || name || '') || t('common.untitled');
    const target = await shell.dialog.save({ title: t('pdf.save-redacted-pdf'), defaultPath: joinPath(folder, `${stem}_Redacted.pdf`), filters: PDF_FILTER() });
    if (!target) return;
    const out = /\.pdf$/i.test(target) ? target : `${target}.pdf`;
    if (path && out.toLowerCase() === path.toLowerCase()) {
      toast(t('pdf.save-redacted-pdf-under'), { tone: 'bad', ms: 6000 });
      return;
    }
    setBusy(t('pdf.redacting'));
    try {
      const { removed } = await shell.pdf.redact({ pages: entriesOf(list), path: out, images });
      setBusy(null);
      let here = true;
      if (dirty) {
        const { response } = await shell.dialog.message({
          type: 'question',
          message: t('pdf.save-marks-name', { name: name || t('common.untitled') }),
          detail: t('pdf.redacted-copy-saved', { copy: basename(out), name: name || t('common.untitled') }),
          buttons: [t('common.save'), t('pdf.dont-save'), t('pdf.keep-it-open')],
          defaultId: 0,
          cancelId: 2,
        });
        if (response === 2) here = false;
        else if (response === 0 && !(await saveRef.current?.())) here = false;
      }
      if (here) await openPdf(out, { password: Object.values(sourcesRef.current).find((s) => s.password)?.password || null });
      else await shell.win.create({ app: 'pdf', file: out });
      const parts = [tn(removed.pages, 'pdf.redacted-count-page')];
      if (removed.images) parts.push(tn(removed.images, 'pdf.count-picture-under'));
      parts.push(t('pdf.saved-as-name', { name: basename(out) }));
      toast(parts.join(' '), { tone: 'good', ms: 7000 });
    } catch (err) {
      toast(t('pdf.redactions-could-not', { reason: err.message }), { tone: 'bad', ms: 8000 });
    } finally {
      setBusy(null);
    }
  }, [path, name, shell, toast, openPdf, applying, dirty]);

  /** Fill & Sign: the signature or initials kept from last time, or the dialog to make them. */
  const sign = useCallback((kind, { fresh = false } = {}) => {
    if (guardLocked('fill')) return;
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

  const pickComment = useCallback((id, pageKey, { onPage = false } = {}) => {
    setPicked(id);
    const had = toolRef.current;
    if (had) toast(t('pdf.tool-put-down', { tool: toolName(had) }), { ms: 4000 });
    setTool(null);
    if (pageKey && pageKey !== currentRef.current) {
      setCurrent(pageKey);
      if (!onPage) requestAnimationFrame(() => scrollToRef.current?.(pageKey));
    }
    // The comment on its page brought into view (one pressed on the page is in view, and may be being dragged).
    if (onPage) return;
    requestAnimationFrame(() => {
      const el = viewRef.current?.querySelector(`.pdf-page[data-key="${pageKey}"] .pdf-annot[data-id="${id}"]`);
      const view = viewRef.current;
      if (!el || !view) return;
      const r = el.getBoundingClientRect();
      const box = view.getBoundingClientRect();
      if (r.top < box.top + 20 || r.bottom > box.bottom - 20) view.scrollTop += r.top - box.top - box.height / 3;
    });
  }, [toast]);

  /** A signature, tick, date or note dragged to a new place, or resized by a corner. */
  const moveComment = useCallback(({ id, source, rect }) => changeComments([{ source, changes: { move: [{ id, rect }] } }]), [changeComments]);

  /** A link on a page: to a page of this file, where that page now stands, or out to a web address. */
  const followLink = useCallback((link, entry) => {
    if (link.url) {
      shell.shell.openExternal({ url: link.url }).catch(() => {});
      return;
    }
    const to = pagesRef.current.find((p) => p.source === entry.source && p.page === link.target);
    if (to) goToRef.current?.(to.key);
    else toast(t('pdf.page-link-goes'), { ms: 4000 });
  }, [shell, toast]);
  const goToRef = useRef(null);

  const overlay = useCallback((entry, { source, scale }) => (
    <>
      <FoundMark entry={entry} source={source} scale={scale} mark={foundMark} />
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
      onMove={moveComment}
      viewports={viewports}
      onLink={followLink}
      />
    </>
  ), [tool, colourOf, picked, pickComment, onCreate, moveComment, viewports, tintFields, fillField, followLink, foundMark]);

  /* ── writing ────────────────────────────────────────────────────────── */

  const entriesOf = (list) => list.map((p) => ({ source: p.source, page: p.page, rotate: p.rotate }));

  const writeTo = useCallback(async (target, list, { title } = {}) => {
    const result = await shell.pdf.assemble({ pages: entriesOf(list), path: target, title });
    return result;
  }, [shell]);

  const saveTo = useCallback(async (target) => {
    if (guardLocked('save')) return false;
    setBusy(t('common.saving'));
    try {
      await writeTo(target, pagesRef.current);
      setPath(target);
      setName(basename(target));
      confirmed.current = target;
      setDirty(false);
      shell.doc.dropCopy({ key: copyKey.current }).catch(() => {});
      shell.app.addRecent({ path: target, app: 'pdf' }).catch(() => {});
      toast(t('common.saved-name', { name: basename(target) }), { tone: 'good' });
      return true;
    } catch (err) {
      toast(t('pdf.pdf-could-not', { reason: err.message }), { tone: 'bad', ms: 8000 });
      return false;
    } finally {
      setBusy(null);
    }
  }, [guardLocked, writeTo, shell, toast]);

  const saveAs = useCallback(async () => {
    if (!pagesRef.current.length || guardLocked('save')) return false;
    const folder = path ? dirname(path) : (await shell.app.paths().catch(() => ({}))).documents || '';
    const stem = path ? stemOf(path) : stemOf(name || '') || t('common.untitled');
    const target = await shell.dialog.save({ title: t('common.save-as-2'), defaultPath: joinPath(folder, `${stem}.pdf`), filters: PDF_FILTER() });
    if (!target) return false;
    return saveTo(/\.pdf$/i.test(target) ? target : `${target}.pdf`);
  }, [path, name, shell, saveTo, guardLocked]);

  /** Save: to the file it came from only once the person has said so; Save As otherwise. */
  const save = useCallback(async () => {
    if (!pagesRef.current.length || guardLocked('save')) return false;
    if (!path) return saveAs();
    if (confirmed.current !== path) {
      const { response } = await shell.dialog.message({
        type: 'question',
        message: t('pdf.replace-name', { name: basename(path) }),
        detail: t('pdf.saving-writes-pages'),
        buttons: [t('pdf.save-a-copy'), t('pdf.replace-the-original'), t('common.cancel')],
        defaultId: 0,
        cancelId: 2,
      });
      if (response === 2) return false;
      if (response === 0) return saveAs();
      confirmed.current = path;
    }
    return saveTo(path);
  }, [path, shell, saveAs, saveTo, guardLocked]);
  saveRef.current = save;

  const extract = useCallback(async () => {
    if (guardLocked('assemble')) return;
    const keys = chosen();
    if (!keys.length) return;
    const list = pagesRef.current.filter((p) => keys.includes(p.key));
    const range = rangeText(numbersOf(keys));
    const folder = path ? dirname(path) : (await shell.app.paths().catch(() => ({}))).documents || '';
    const target = await shell.dialog.save({
      title: t('pdf.extract-pages-range', { range }),
      defaultPath: joinPath(folder, `${stemOf(path || name || '') || t('common.untitled')} ${t('pdf.pages-range', { range })}.pdf`),
      filters: PDF_FILTER(),
    });
    if (!target) return;
    setBusy(t('common.saving'));
    try {
      const out = /\.pdf$/i.test(target) ? target : `${target}.pdf`;
      await writeTo(out, list);
      toast(tn(list.length, 'pdf.count-page-extracted', { name: basename(out) }), { tone: 'good', ms: 4000 });
    } catch (err) {
      toast(t('pdf.pages-could-not-extracted', { reason: err.message }), { tone: 'bad', ms: 8000 });
    } finally {
      setBusy(null);
    }
  }, [chosen, path, name, shell, writeTo, toast, guardLocked]);

  const split = useCallback(async ({ every = null, picked = false }) => {
    setSplitOpen(false);
    if (guardLocked('assemble')) return;
    const list = pagesRef.current;
    const starts = picked ? list.map((p, i) => (selectedRef.current.includes(p.key) ? i : -1)).filter((i) => i > 0) : null;
    const groups = splitGroups(list.length, { every, starts });
    if (groups.length < 2) return;
    const [folder] = await shell.dialog.open({ title: t('pdf.choose-folder-files'), directory: true, defaultPath: path ? dirname(path) : undefined }) || [];
    if (!folder) return;
    const stem = stemOf(path || name || '') || t('common.untitled');
    setBusy(t('common.saving'));
    const written = [];
    try {
      for (let g = 0; g < groups.length; g++) {
        let target = joinPath(folder, `${stem}_Part${g + 1}.pdf`);
        for (let n = 2; await shell.fs.exists({ path: target }); n++) target = joinPath(folder, `${stem}_Part${g + 1} (${n}).pdf`);
        await writeTo(target, groups[g].map((i) => list[i]));
        written.push(target);
      }
      toast(tn(written.length, 'pdf.split-count-file', {
        folder: basename(folder) || folder, first: basename(written[0]), last: basename(written[written.length - 1]),
      }), { tone: 'good', ms: 12000, action: { label: t('common.show-in'), run: () => shell.shell.showInFolder({ path: written[0] }) } });
    } catch (err) {
      toast(t('pdf.pdf-could-not-split', { reason: err.message }), { tone: 'bad', ms: 8000 });
    } finally {
      setBusy(null);
    }
  }, [path, name, shell, writeTo, toast, guardLocked]);

  /** Export as Pictures: each page chosen written as a PNG or a JPEG into a folder. */
  const exportPictures = useCallback(async ({ scope, format, dpi }) => {
    setExporting(false);
    const list = pagesRef.current;
    const keys = new Set(scope === 'picked' ? selectedRef.current : scope === 'current' ? [currentRef.current] : list.map((p) => p.key));
    const chosen = list.map((p, i) => ({ p, n: i + 1 })).filter(({ p }) => keys.has(p.key));
    if (!chosen.length) return;
    const [folder] = await shell.dialog.open({ title: t('pdf.choose-folder-pictures'), directory: true, defaultPath: path ? dirname(path) : undefined }) || [];
    if (!folder) return;
    const stem = stemOf(path || name || '') || t('common.untitled');
    const ext = format === 'jpeg' ? 'jpg' : 'png';
    setBusy(t('pdf.exporting'));
    let written = 0;
    try {
      for (const { p, n } of chosen) {
        const { bytes } = await pageBytes(sourcesRef.current[p.source].doc, p.page, { rotate: p.rotate, dpi, type: format });
        let target = joinPath(folder, `${stem}_Page_${n}.${ext}`);
        for (let k = 2; await shell.fs.exists({ path: target }); k++) target = joinPath(folder, `${stem}_Page_${n} (${k}).${ext}`);
        await shell.fs.write({ path: target, bytes });
        written += 1;
      }
      toast(tn(written, 'pdf.count-page-exported', { folder: basename(folder) || folder }), { tone: 'good', ms: 5000 });
    } catch (err) {
      toast(t('pdf.pages-could-not-exported', { reason: err.message }), { tone: 'bad', ms: 8000 });
    } finally {
      setBusy(null);
    }
  }, [path, name, shell, toast]);

  /** Reduce File Size: the pictures drawn smaller, written whole as a new file. */
  const compressFile = useCallback(async (quality) => {
    setCompressing(false);
    if (guardLocked()) return;
    const folder = path ? dirname(path) : (await shell.app.paths().catch(() => ({}))).documents || '';
    const stem = stemOf(path || name || '') || t('common.untitled');
    const target = await shell.dialog.save({ title: t('pdf.save-smaller-pdf'), defaultPath: joinPath(folder, `${stem}_Reduced.pdf`), filters: PDF_FILTER() });
    if (!target) return;
    const out = /\.pdf$/i.test(target) ? target : `${target}.pdf`;
    setBusy(t('pdf.making-file-smaller'));
    try {
      const { token, plan } = await shell.pdf.compressPlan({ pages: entriesOf(pagesRef.current), dpi: quality.dpi });
      const images = await encodePlan(plan);
      const { before, after, replaced } = await shell.pdf.compress({ token, dpi: quality.dpi, images, path: out });
      if (!replaced) toast(t('pdf.pictures-already-small', { name: basename(out) }), { ms: 6000 });
      else toast(t('pdf.saved-name-before', { name: basename(out), before: formatBytes(before), after: formatBytes(after) }), { tone: 'good', ms: 7000 });
    } catch (err) {
      toast(t('pdf.file-could-not', { reason: err.message }), { tone: 'bad', ms: 8000 });
    } finally {
      setBusy(null);
    }
  }, [guardLocked, path, name, shell, toast]);

  /** Print: every page drawn at print resolution, handed to the system's print dialog. */
  const print = useCallback(async () => {
    const list = pagesRef.current;
    if (!list.length || printing) return;
    setBusy(t('pdf.preparing-to-print'));
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
      toast(t('pdf.pages-could-not', { reason: err.message }), { tone: 'bad', ms: 8000 });
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

  // Read Aloud (C8): the text layer, from this page to the last, each page shown as it is read.
  const readerRef = useRef(null);
  const [readState, setReadState] = useState(null);
  const [readingAloud, setReadingAloud] = useState(false);
  const reader = () => (readerRef.current ||= createReader({ shell, onState: setReadState }));
  useEffect(() => () => readerRef.current?.stop(), []);
  const readAloud = useCallback(async () => {
    if (readingAloud) {
      reader().stop();
      setReadingAloud(false);
      return;
    }
    const list = pagesRef.current;
    const from = Math.max(0, list.findIndex((p) => p.key === currentRef.current));
    const pieces = [];
    for (const p of list.slice(from)) {
      const source = sourcesRef.current[p.source];
      if (!source) continue;
      let text = '';
      try { text = await pageText(source.doc, p.page); } catch { text = ''; }
      if (text.trim()) pieces.push({ text: text.replace(/-\n(?=\p{Ll})/gu, '').replace(/\s*\n\s*/g, ' '), source: { key: p.key }, language: null });
    }
    if (!pieces.length) {
      toast(t('pdf.no-words-read'), { ms: 5000 });
      return;
    }
    setReadingAloud(true);
    reader().read(pieces, {
      onSentence: (it) => { if (it.source.key !== currentRef.current) goTo(it.source.key); },
      onEnd: () => setReadingAloud(false),
    });
  }, [readingAloud, goTo, toast]);

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
    if (!findOpen || !q.trim()) { setHits([]); setHitAt(-1); setHitsFor(q); return; }
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
      setHitsFor(q);
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
          toast(t('common.recovered-unsaved-work'), { ms: 6000 });
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

  useDirtyGuard({ shell, dirty: hasDoc && dirty, name: name || t('common.untitled'), onSave: save });

  const copied = useRef({ at: 0, pages: null });
  useEffect(() => {
    if (!hasDoc || !dirty || lockedFile) return undefined;
    const since = Date.now() - copied.current.at;
    const wait = copied.current.at ? Math.max(0, AUTOSAVE_EVERY_MS - since) : AUTOSAVE_FIRST_MS;
    const timer = setTimeout(async () => {
      if (copied.current.pages === pagesRef.current) return;
      copied.current = { at: Date.now(), pages: pagesRef.current };
      try {
        const { bytes } = await shell.pdf.assemble({ pages: entriesOf(pagesRef.current) });
        await shell.doc.keepCopy({ key: copyKey.current, kind: 'pdf', name: name || t('common.untitled'), path, bytes });
      } catch {
        /* a copy that could not be made is made at the next change */
      }
    }, wait);
    return () => clearTimeout(timer);
  }, [pages, dirty, hasDoc, name, path, shell, lockedFile]);
  useEffect(() => {
    const away = () => { shell.doc.dropCopy({ key: copyKey.current }).catch(() => {}); };
    window.addEventListener('pagehide', away);
    return () => window.removeEventListener('pagehide', away);
  }, [shell]);

  /* ── commands ───────────────────────────────────────────────────────── */

  const canEdit = hasDoc && !lockedFile;
  const commands = useMemo(() => ({
    'file.new': { label: t('pdf.new-window'), icon: 'new', key: 'Mod+N', run: () => shell.win.create({ app: 'pdf' }) },
    'file.open': { label: t('common.open-2'), icon: 'open', key: 'Mod+O', run: openFile },
    'file.save': { label: t('common.save'), icon: 'save', key: 'Mod+S', when: () => canEdit, run: save },
    'file.saveAs': { label: t('common.save-as-3'), icon: 'save', key: 'Mod+Shift+S', when: () => canEdit, run: saveAs },
    'file.print': { label: t('common.print-2'), icon: 'print', key: 'Mod+P', when: () => hasDoc, run: print },
    'edit.undo': { label: t('common.undo'), icon: 'undo', key: 'Mod+Z', when: () => history.current.undo.length > 0, run: undo },
    'edit.redo': { label: t('common.redo'), icon: 'redo', key: 'Mod+Y', when: () => history.current.redo.length > 0, run: redo },
    'edit.find': { label: t('common.find'), icon: 'find', key: 'Mod+F', global: true, when: () => hasDoc, run: openFind },
    'find.next': { label: t('common.next-match'), key: 'F3', global: true, when: () => findOpen, run: () => nextHit(1) },
    'find.previous': { label: t('common.previous-match'), key: 'Shift+F3', global: true, when: () => findOpen, run: () => nextHit(-1) },
    'find.close': { label: t('pdf.close-find'), key: 'Escape', global: true, when: () => findOpen, run: () => { setFindOpen(false); setQuery(''); } },
    'page.rotateRight': { label: t('common.rotate-right'), icon: 'rotate', key: 'Mod+R', when: () => canEdit, run: () => rotate(90) },
    'page.rotateLeft': { label: t('common.rotate-left'), icon: 'rotateLeft', key: 'Mod+L', when: () => canEdit, run: () => rotate(-90) },
    'page.delete': { label: t('pdf.delete-pages'), icon: 'trash', when: () => canEdit, run: removePages },
    'page.extract': { label: t('pdf.extract-pages'), icon: 'export', when: () => canEdit, run: extract },
    'page.insert': { label: t('pdf.insert-from-file'), icon: 'import', when: () => canEdit, run: insertFromFile },
    'page.recognise': { label: t('pdf.recognise-text-2'), icon: 'scanText', when: () => canEdit && !ocr, run: askRecognise },
    'comment.cancel': { label: t('pdf.put-tool-down'), key: 'Escape', global: true, when: () => Boolean(tool) && !findOpen && !noteDraft, run: () => setTool(null) },
    'comment.unpick': { label: t('pdf.unpick-the-comment'), key: 'Escape', when: () => Boolean(picked) && !tool && !findOpen, run: () => setPicked(null) },
    'comment.delete': { label: t('pdf.delete-the-comment'), key: 'Delete', when: () => Boolean(picked) && canEdit, run: deletePicked },
    'page.next': { label: t('common.next-page'), key: 'PageDown', when: () => hasDoc && layout === 'single', run: () => stepPage(1) },
    'page.previous': { label: t('common.previous-page'), key: 'PageUp', when: () => hasDoc && layout === 'single', run: () => stepPage(-1) },
  }), [shell, openFile, save, saveAs, print, undo, redo, openFind, findOpen, nextHit, rotate, removePages, extract, insertFromFile, stepPage, hasDoc, canEdit, layout, ocr, askRecognise, tool, picked, noteDraft, deletePicked]);
  useCommands(commands, [commands]);

  const appMenu = useAppMenu({
    shell,
    appKey: 'pdf',
    onNew: () => shell.win.create({ app: 'pdf' }),
    onOpen: () => openFileRef.current?.(),
    // Merging other files into one PDF is this app's export.
    exports: [{ label: t('pdf.merge-files-2'), icon: 'pdf', run: () => mergeFiles() }],
  });

  const thumbMenu = useCallback((e, key) => {
    e.preventDefault();
    if (!selectedRef.current.includes(key)) { setSelected([key]); anchor.current = key; setCurrent(key); }
    menu.open(e, [
      { label: t('common.rotate-right'), icon: 'rotate', run: () => rotate(90) },
      { label: t('common.rotate-left'), icon: 'rotateLeft', run: () => rotate(-90) },
      '-',
      { label: t('pdf.extract-pages'), icon: 'export', run: extract },
      { label: t('pdf.insert-from-file'), icon: 'import', run: insertFromFile },
      { label: t('pdf.recognise-text-2'), icon: 'scanText', run: askRecognise },
      '-',
      { label: t('pdf.delete-pages'), icon: 'trash', run: removePages },
    ]);
  }, [menu, rotate, extract, insertFromFile, removePages, askRecognise]);

  /* ── the window ─────────────────────────────────────────────────────── */

  const pickedCount = selected.length;
  const pickedLabel = pickedCount ? tn(pickedCount, 'pdf.count-page-selected') : null;
  const currentIndex = pages.findIndex((p) => p.key === current);
  const currentSize = currentIndex >= 0 ? sizeOf(pages[currentIndex]) : null;
  const currentLabel = currentIndex >= 0 ? labelOf(sources, pages[currentIndex]) : null;

  const ribbon = (
    <Ribbon
      tabs={[
        { id: 'home', label: t('common.home') },
        { id: 'pages', label: t('common.pages') },
        { id: 'comment', label: t('common.comment') },
        { id: 'fill', label: t('pdf.fill-and') },
        { id: 'redact', label: t('common.redact') },
        { id: 'view', label: t('common.view') },
      ]}
      active={tab}
      onTab={(id) => { setTab(id); if (id === 'comment' && hasDoc) setCommentsOpen(true); if (!['comment', 'fill', 'redact'].includes(id)) setTool(null); }}
      quick={(
        <>
          <Button icon="save" title={t('common.save-ctrl-s')} disabled={!canEdit} onClick={save} />
          <Button icon="undo" title={t('common.undo-ctrl-z')} disabled={!history.current.undo.length} onClick={undo} />
          <Button icon="redo" title={t('common.redo-ctrl-y')} disabled={!history.current.redo.length} onClick={redo} />
        </>
      )}
    >
      {tab === 'home' ? (
        <>
          <Group label={t('common.file')}>
            <Button tall icon="open" label={t('common.open')} title={t('pdf.open-pdf-window')} onClick={openFile} />
            <Button tall icon="save" label={t('common.save')} title={t('pdf.save-pages-now')} disabled={!canEdit} onClick={save} />
            <Rows>
              <Button icon="save" label={t('common.save-as')} title={t('pdf.save-pages-new')} disabled={!canEdit} onClick={saveAs} />
              <Button icon="print" label={t('common.print')} title={t('pdf.print-every-page')} disabled={!hasDoc} onClick={print} />
            </Rows>
          </Group>
          <Group label={t('common.zoom')}>
            <Rows>
              <Button icon="zoomIn" label={t('common.zoom-in')} title={t('pdf.zoom-in-ctrl')} disabled={!hasDoc} onClick={() => setZoomByHand(zoom * 1.25)} />
              <Button icon="zoomOut" label={t('common.zoom-out')} title={t('pdf.zoom-out-ctrl')} disabled={!hasDoc} onClick={() => setZoomByHand(zoom / 1.25)} />
            </Rows>
            <Rows>
              <Button icon="pagesAcross" label={t('pdf.fit-width')} title={t('pdf.fit-width-page')} pressed={fit === 'width'} disabled={!hasDoc} onClick={() => setFit('width')} />
              <Button icon="file" label={t('pdf.fit-page')} title={t('pdf.fit-page-whole')} pressed={fit === 'page'} disabled={!hasDoc} onClick={() => setFit('page')} />
            </Rows>
          </Group>
          <Group label={t('common.find')}>
            <Button tall icon="find" label={t('common.find')} title={t('pdf.find-words-every')} disabled={!hasDoc} onClick={openFind} />
          </Group>
          <Group label={t('word.ribbon.speech')}>
            <Button tall icon="volume" label={t('word.ribbon.read-aloud')} title={t('pdf.read-aloud-words')} pressed={readingAloud} disabled={!hasDoc} onClick={readAloud} />
          </Group>
          <Group label={t('pdf.scans')}>
            <Button tall icon="scanText" label={t('pdf.recognise-text')} title={t('pdf.recognise-text-words')} disabled={!canEdit || Boolean(ocr)} onClick={askRecognise} data-role="recognise-text" />
            <Button tall icon="scanReceipt" label={t('common.fields')} title={t('pdf.fields-details-invoice')} pressed={fieldsOpen} disabled={!hasDoc} onClick={toggleFields} data-role="fields-toggle" />
          </Group>
          <Group label={t('common.pages')}>
            <Rows>
              <Button icon="rotate" label={t('common.rotate-right')} title={t('pdf.rotate-right-selected')} disabled={!canEdit} onClick={() => rotate(90)} />
              <Button icon="trash" label={t('common.delete')} title={t('pdf.delete-selected-pages')} disabled={!canEdit} onClick={removePages} />
            </Rows>
          </Group>
        </>
      ) : null}
      {tab === 'pages' ? (
        <>
          <Group label={t('common.rotate')}>
            <Button tall icon="rotateLeft" label={t('common.rotate-left')} title={t('pdf.rotate-left-selected')} disabled={!canEdit} onClick={() => rotate(-90)} data-role="rotate-left" />
            <Button tall icon="rotate" label={t('common.rotate-right')} title={t('pdf.rotate-right-selected')} disabled={!canEdit} onClick={() => rotate(90)} data-role="rotate-right" />
          </Group>
          <Group label={t('pdf.organise')}>
            <Button tall icon="trash" label={t('common.delete')} title={t('pdf.delete-selected-pages')} disabled={!canEdit} onClick={removePages} data-role="delete-pages" />
            <Button tall icon="export" label={t('common.extract')} title={t('pdf.extract-selected-pages')} disabled={!canEdit} onClick={extract} data-role="extract-pages" />
          </Group>
          <Group label={t('pdf.scans')}>
            <Button tall icon="scanText" label={t('pdf.recognise-text')} title={t('pdf.recognise-text-words')} disabled={!canEdit || Boolean(ocr)} onClick={askRecognise} />
            <Button tall icon="scanReceipt" label={t('common.fields')} title={t('pdf.fields-details-invoice')} pressed={fieldsOpen} disabled={!hasDoc} onClick={toggleFields} />
          </Group>
          <Group label={t('common.insert')}>
            <Button tall icon="import" label={t('pdf.insert-from')} title={t('pdf.insert-file-pages')} disabled={!canEdit} onClick={insertFromFile} data-role="insert-file" />
          </Group>
          <Group label={t('common.export')}>
            <Button tall icon="picture" label={t('pdf.export-as-pictures')} title={t('pdf.export-pictures-pages')} disabled={!hasDoc} onClick={() => setExporting(true)} data-role="export-pictures" />
            <Button tall icon="export" label={t('pdf.reduce-file-size')} title={t('pdf.reduce-file-size-large')} disabled={!canEdit} onClick={() => setCompressing(true)} data-role="compress" />
          </Group>
          <Group label={t('pdf.split-and')}>
            <Button tall icon="scissors" label={t('common.split')} title={t('pdf.split-pdf-several')} disabled={!canEdit || pages.length < 2} onClick={() => setSplitOpen(true)} data-role="split" />
            <Button tall icon="pdf" label={t('pdf.merge-files')} title={t('pdf.merge-files-other')} disabled={!allowed('assemble')} onClick={() => mergeFiles()} data-role="merge" />
          </Group>
        </>
      ) : null}
      {tab === 'comment' ? (
        <>
          <Group label={t('pdf.markup')}>
            <Button tall icon="highlighter" label={t('common.highlight')} title={t('pdf.highlight-words-selected')} pressed={tool === 'highlight'} disabled={!canEdit} onClick={() => pressTool('highlight')} data-role="tool-highlight" />
            <Rows>
              <Button icon="underline" label={t('common.underline')} title={t('pdf.underline-words-selected')} pressed={tool === 'underline'} disabled={!canEdit} onClick={() => pressTool('underline')} data-role="tool-underline" />
              <Button icon="strike" label={t('common.strikethrough')} title={t('pdf.strikethrough-line')} pressed={tool === 'strike'} disabled={!canEdit} onClick={() => pressTool('strike')} data-role="tool-strike" />
            </Rows>
          </Group>
          <Group label={t('common.notes')}>
            <Button tall icon="note" label={t('pdf.sticky-note')} title={t('pdf.sticky-note-note')} pressed={tool === 'note'} disabled={!canEdit} onClick={() => pressTool('note')} data-role="tool-note" />
            <Button tall icon="textbox" label={t('common.text-box')} title={t('pdf.text-box-words')} pressed={tool === 'text'} disabled={!canEdit} onClick={() => pressTool('text')} data-role="tool-text" />
          </Group>
          <Group label={t('common.drawing')}>
            <Button tall icon="pen" label={t('common.draw')} title={t('pdf.draw-freehand-lines')} pressed={tool === 'ink'} disabled={!canEdit} onClick={() => pressTool('ink')} data-role="tool-ink" />
            <Rows>
              <Button icon="mkRect" label={t('common.rectangle')} title={t('pdf.rectangle-drag-one')} pressed={tool === 'rect'} disabled={!canEdit} onClick={() => pressTool('rect')} data-role="tool-rect" />
              <Button icon="mkEllipse" label={t('common.ellipse')} title={t('pdf.ellipse-drag-one')} pressed={tool === 'ellipse'} disabled={!canEdit} onClick={() => pressTool('ellipse')} data-role="tool-ellipse" />
            </Rows>
            <Rows>
              <Button icon="mkLine" label={t('common.line')} title={t('pdf.line-drag-one')} pressed={tool === 'line'} disabled={!canEdit} onClick={() => pressTool('line')} data-role="tool-line" />
              <Button icon="mkArrow" label={t('common.arrow')} title={t('pdf.arrow-drag-one')} pressed={tool === 'arrow'} disabled={!canEdit} onClick={() => pressTool('arrow')} data-role="tool-arrow" />
            </Rows>
          </Group>
          <Group label={t('common.colour')}>
            <div className="pdf-swatches" role="radiogroup" aria-label={t('common.colour')}>
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
                    data-tip={t('pdf.colour-colour-tool', { colour: s.label() })}
                    aria-label={s.label()}
                    style={{ '--swatch': cssColour(s.rgb) }}
                    onClick={() => setToolColours((c) => ({ ...c, [which]: s.key }))}
                  />
                );
              })}
            </div>
          </Group>
          <Group label={t('common.show')}>
            <Button tall icon="comment" label={t('common.comments')} title={t('pdf.comments-pane-listing')} pressed={commentsOpen} disabled={!hasDoc} onClick={() => { setFieldsOpen(false); setCommentsOpen((v) => !v); }} data-role="comments-toggle" />
          </Group>
        </>
      ) : null}
      {tab === 'fill' ? (
        <>
          <Group label={t('pdf.form')}>
            <Button tall icon="formField" label={t('pdf.highlight')} title={t('pdf.highlight-fields-forms')} pressed={tintFields} disabled={!hasDoc} onClick={() => setTintFields((v) => !v)} data-role="tint-fields" />
          </Group>
          <Group label={t('common.sign')}>
            <Button tall icon="signature" label={t('common.signature')} title={t('pdf.signature-signature')} pressed={tool === 'stamp' && signature?.kind === 'signature'} disabled={!canEdit} onClick={() => sign('signature')} data-role="sign-signature" />
            <Rows>
              <Button icon="signature" label={t('pdf.initials')} title={t('pdf.initials-initials-typed')} pressed={tool === 'stamp' && signature?.kind === 'initials'} disabled={!canEdit} onClick={() => sign('initials')} data-role="sign-initials" />
              <Button icon="pen" label={t('pdf.change')} title={t('pdf.change-signature-type')} disabled={!canEdit} onClick={() => sign('signature', { fresh: true })} data-role="sign-change" />
            </Rows>
          </Group>
          <Group label={t('common.text')}>
            <Button tall icon="textbox" label={t('common.add-text')} title={t('pdf.add-text-words')} pressed={tool === 'text'} disabled={!canEdit} onClick={() => pressTool('text')} data-role="fill-text" />
            <Rows>
              <Button icon="calendar" label={t('common.date')} title={t('pdf.date-todays-date')} pressed={tool === 'date'} disabled={!canEdit} onClick={() => pressTool('date')} data-role="fill-date" />
              <Button icon="check" label={t('pdf.checkmark')} title={t('pdf.checkmark-tick-click')} pressed={tool === 'check'} disabled={!canEdit} onClick={() => pressTool('check')} data-role="fill-check" />
            </Rows>
          </Group>
        </>
      ) : null}
      {tab === 'redact' ? (
        <>
          <Group label={t('common.mark')}>
            <Button tall icon="redact" label={t('pdf.mark-text')} title={t('pdf.mark-text-words')} pressed={tool === 'redact-text'} disabled={!canEdit} onClick={() => pressTool('redact-text')} data-role="redact-text" />
            <Button tall icon="marquee" label={t('pdf.mark-area')} title={t('pdf.mark-area-drag')} pressed={tool === 'redact-area'} disabled={!canEdit} onClick={() => pressTool('redact-area')} data-role="redact-area" />
            <Button tall icon="find" label={t('pdf.mark-matches')} title={t('pdf.mark-matches-every')} disabled={!canEdit} onClick={markSearchResults} data-role="redact-search" />
          </Group>
          <Group label={t('common.apply')}>
            <Button tall icon="check" label={t('common.apply-redactions')} title={t('pdf.apply-redactions-marks')} disabled={!canEdit} onClick={askApply} data-role="redact-apply" />
          </Group>
          <Group label={t('common.review')}>
            <Button tall icon="comment" label={t('common.marks')} title={t('pdf.marks-comments-pane')} pressed={commentsOpen} disabled={!hasDoc} onClick={() => setCommentsOpen((v) => !v)} data-role="redact-marks" />
          </Group>
        </>
      ) : null}
      {tab === 'view' ? (
        <>
          <Group label={t('pdf.page-display')}>
            <Button tall icon="file" label={t('pdf.single-page')} title={t('pdf.single-page-one')} pressed={layout === 'single'} disabled={!hasDoc} onClick={() => setLayout('single')} data-role="layout-single" />
            <Button tall icon="pagesAcross" label={t('pdf.two-pages')} title={t('pdf.two-pages-pages')} pressed={layout === 'two'} disabled={!hasDoc} onClick={() => setLayout('two')} data-role="layout-two" />
            <Button tall icon="pagesDown" label={t('common.continuous')} title={t('pdf.continuous-every-page')} pressed={layout === 'continuous'} disabled={!hasDoc} onClick={() => setLayout('continuous')} data-role="layout-continuous" />
          </Group>
          <Group label={t('common.show')}>
            <Button tall icon="grid" label={t('pdf.thumbnails')} title={t('pdf.thumbnails-strip-pages')} pressed={strip} onClick={() => setStrip((v) => !v)} data-role="thumbnails" />
          </Group>
          <Group label={t('common.zoom')}>
            <Rows>
              <Button icon="pagesAcross" label={t('pdf.fit-width')} title={t('pdf.fit-width-page')} pressed={fit === 'width'} disabled={!hasDoc} onClick={() => setFit('width')} />
              <Button icon="file" label={t('pdf.fit-page')} title={t('pdf.fit-page-whole')} pressed={fit === 'page'} disabled={!hasDoc} onClick={() => setFit('page')} />
            </Rows>
            <Button tall icon="zoomIn" label={t('common.actual-size')} title={t('pdf.actual-size-page')} disabled={!hasDoc} onClick={() => setZoomByHand(1)} />
          </Group>
        </>
      ) : null}
    </Ribbon>
  );

  const status = (
    <>
      <span data-role="page-status">{hasDoc && currentIndex >= 0 ? (currentLabel && currentLabel !== String(currentIndex + 1) ? t('pdf.page-label-page', { label: currentLabel, page: currentIndex + 1, count: pages.length }) : t('pdf.page-page-count', { page: currentIndex + 1, count: pages.length })) : path || name || t('pdf.no-pdf-open')}</span>
      {hasDoc ? (
        <>
          <Button ghost icon="chevronUp" title={t('common.previous-page')} disabled={currentIndex <= 0} onClick={() => stepPage(-1)} />
          <Button ghost icon="chevronDown" title={t('common.next-page')} disabled={currentIndex >= pages.length - 1} onClick={() => stepPage(1)} />
        </>
      ) : null}
      <Spacer />
      {pickedLabel ? <Chip>{pickedLabel}</Chip> : null}
      {currentSize ? <Chip>{t('pdf.width-height-mm', { width: Math.round(currentSize.width / 72 * 25.4), height: Math.round(currentSize.height / 72 * 25.4) })}</Chip> : null}
      {lockedFile ? <Chip>{t('pdf.protected-read-only')}</Chip> : protectedFile ? <Chip>{t('pdf.protected-password')}</Chip> : null}
      {hasDoc ? <ZoomSlider value={zoom} min={ZOOM_MIN} max={4} onChange={setZoomByHand} onReset={() => setZoomByHand(1)} /> : null}
    </>
  );

  return (
    <AppFrame app={app} shell={shell} title={name || app.name} dirty={hasDoc && dirty} menu={appMenu} ribbon={ribbon} status={status}>
      <style>{CSS + OCR_CSS + FOUND_CSS}</style>
      {!hasDoc ? (
        failed ? (
          <Empty
            icon={locked ? 'lock' : 'pdf'}
            title={locked ? t('pdf.pdf-protected') : t('common.file-could-not')}
            action={(
              <div className="pdf-empty-actions">
                {locked && lastOpen.current ? <Button primary icon="lock" label={t('common.enter')} data-role="pdf-password-retry" onClick={() => { const o = lastOpen.current; openPdf(o.file, { bytes: o.bytes, recovered: o.recovered }); }} /> : null}
                <Button primary={!locked} icon="open" label={t('common.open-another')} onClick={openFile} />
              </div>
            )}
          >
            {failed}
          </Empty>
        ) : (
          <div className="pdf-empty" style={{ flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', overflow: 'auto', padding: '24px 16px' }}>
            <Empty
              icon="pdf"
              title={busy || t('pdf.no-pdf-open')}
              action={busy ? <Spinner /> : (
                <div className="pdf-empty-actions">
                  <Button primary icon="open" label={t('pdf.open-a-pdf')} onClick={openFile} />
                  <Button icon="pdf" label={t('pdf.merge-files')} title={t('pdf.merge-files-several')} onClick={() => mergeFiles()} />
                </div>
              )}
            >
              {busy ? null : t('pdf.open-pdf-read')}
            </Empty>
            {busy ? null : <RecentFiles shell={shell} app="pdf" open={false} embedded />}
          </div>
        )
      ) : (
        <div className="pdf-app" data-pages={pages.length} data-layout={layout}>
          {strip ? (
            <div className="pdf-side" style={{ width: stripWidth }}>
              <div className="pdf-side-head">
                <div className="pdf-side-tabs" role="tablist">
                  <button type="button" role="tab" aria-selected={side === 'pages'} className={side === 'pages' ? 'on' : ''} data-role="side-pages" onClick={() => setSide('pages')}>{t('common.pages')}</button>
                  <button type="button" role="tab" aria-selected={side === 'bookmarks'} className={side === 'bookmarks' ? 'on' : ''} data-role="side-bookmarks" onClick={() => setSide('bookmarks')}>{t('pdf.bookmarks')}</button>
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
            {lockedFile ? <div className="pdf-notice">{t('pdf.pdf-protected-way')}</div> : null}
            {restricted ? (
              <div className="pdf-notice" data-role="pdf-restricted">
                {t('pdf.whoever-protected-pdf')}
                <Button icon="lock" label={t('pdf.enter-owner')} data-role="pdf-owner-password" onClick={enterOwnerPassword} />
              </div>
            ) : null}
            {tab === 'redact' && allowed('change') ? (
              <div className="pdf-notice info" data-role="redact-notice">
                {t('pdf.marks-show-removed')}
              </div>
            ) : null}
            {tab === 'fill' && allowed('fill') ? (
              <div className="pdf-notice info" data-role="sign-notice">
                {tool === 'stamp' ? t('pdf.click-page-kind', { kind: signature?.kind === 'initials' ? t('pdf.initials-2') : t('pdf.signature') }) : t('pdf.type-forms-fields')}
              </div>
            ) : null}
            {readingAloud ? <ReadAloudBar reader={reader()} state={readState} onClose={readAloud} /> : null}
            {findOpen ? (
              <div className="pdf-find" role="search">
                <Input
                  ref={findRef}
                  autoFocus
                  value={query}
                  placeholder={t('pdf.find-document')}
                  aria-label={t('pdf.find-document')}
                  data-role="find-input"
                  onChange={(e) => setQuery(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') { e.preventDefault(); nextHit(e.shiftKey ? -1 : 1); }
                  }}
                />
                <span className="pdf-find-count" data-count={hits.length} data-at={hitAt} data-for={hitsFor} data-query={query.trim() ? 'yes' : 'no'}>
                  {!query.trim() ? '' : hits.length ? t('pdf.at-of-count', { at: hitAt + 1, count: hits.length }) : t('pdf.no-matches')}
                </span>
                <Button ghost icon="chevronUp" title={t('pdf.previous-match-shift')} disabled={!hits.length} onClick={() => nextHit(-1)} />
                <Button ghost icon="chevronDown" title={t('pdf.next-match-f3')} disabled={!hits.length} onClick={() => nextHit(1)} />
                <Button ghost icon="close" title={t('pdf.close-find-esc')} onClick={() => { setFindOpen(false); setQuery(''); }} />
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
          {fieldsOpen ? (
            <FieldsPane
              items={pages.map((p, i) => (found[p.key] ? { key: p.key, number: i + 1, source: p.source, page: p.page, found: found[p.key] } : null)).filter(Boolean)}
              sources={sources}
              mark={foundMark}
              onMark={markFound}
              onCopyRow={copyFoundRow}
              onMake={makeCheckboxes}
              onClose={() => { setFieldsOpen(false); setFoundMark(null); }}
            />
          ) : null}
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
        <ApplyRedactionsDialog {...redactionMarks()} preview={applying.preview} onApply={applyRedactions} onClose={() => setApplying(false)} />
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
      {gate.node}
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
