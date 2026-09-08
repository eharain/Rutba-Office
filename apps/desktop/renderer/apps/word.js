// Rutba Word.
//
// The page is a view of a model that lives in the backend, and every keystroke
// becomes a named operation on it. The surface is contenteditable — that is the
// only way to get a native caret, native text selection, an IME that works for
// Chinese or Arabic, and the platform's own spell checker — but the browser is
// never allowed to actually change the text. Every `beforeinput` is cancelled
// and turned into an operation, the engine applies it, and the page re-renders
// from what came back.
//
// That inversion is what makes undo, formatting, tables and the file format all
// behave consistently: there is one model, and the DOM is only ever a picture
// of it.

import React, { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { Button, Icon, Spacer, Chip, Empty, Spinner, useToast, useMenu, useCommands, menuItems } from '@rutba/office-ui';
import { AppFrame, useAppMenu, pickOpen, pickSave, confirmDiscard, useFileDrop, openInApp , useDirtyGuard } from '../shell.js';
import { SITE } from '@rutba/office-formats/registry';
import WordRibbon from './word/ribbon.js';
import { NavigationPane, Ruler, installWordStyles } from './word/panes.js';

installWordStyles();
import {
  LinkDialog, TableDialog, BandDialog, CommentDialog, CommentsDialog, FindDialog, WordCountDialog,
  DateTimeDialog, SymbolDialog, PropertiesDialog, ShortcutsDialog, TrackedDialog,
} from './word/dialogs.js';

/**
 * Character offset of a DOM position within its block element.
 *
 * A selection position is a node and an offset, and the offset means two
 * different things depending on the node. In a text node it is a character
 * index; in an element it is a *child index*. Treating the second as the first
 * is why selecting a whole paragraph and pressing Ctrl+B did nothing: the
 * anchor landed on the paragraph element with offset 0, which was read as "the
 * end of the text", collapsing the selection onto its own focus before the
 * engine ever saw it.
 */
function offsetIn(blockEl, node, offset) {
  if (!blockEl || !node) return 0;

  // An element position: count the text in the children before it.
  if (node.nodeType !== Node.TEXT_NODE) {
    const children = [...node.childNodes].slice(0, offset);
    return children.reduce((n, child) => n + (child.textContent?.length ?? 0), 0);
  }

  let total = 0;
  const walker = document.createTreeWalker(blockEl, NodeFilter.SHOW_TEXT);
  let current = walker.nextNode();
  while (current) {
    if (current === node) return total + offset;
    total += current.nodeValue.length;
    current = walker.nextNode();
  }
  // A text node outside this block: clamp to the block's end rather than guess.
  return total;
}

/** The DOM position for a character offset inside a block. */
function pointIn(blockEl, offset) {
  if (!blockEl) return null;
  const walker = document.createTreeWalker(blockEl, NodeFilter.SHOW_TEXT);
  let seen = 0;
  let node = walker.nextNode();
  while (node) {
    const len = node.nodeValue.length;
    if (seen + len >= offset) return { node, offset: Math.max(0, Math.min(len, offset - seen)) };
    seen += len;
    node = walker.nextNode();
  }
  return { node: blockEl, offset: blockEl.childNodes.length };
}

/**
 * Put the selection back where the engine says it is.
 *
 * A *range*, not a caret. Restoring only the focus point collapsed whatever the
 * reader had selected, so selecting a sentence and pressing Ctrl+B bolded
 * nothing: the selection was destroyed by the act of telling the engine about
 * it, and the engine then applied the format to an empty caret.
 */
function placeSelection(page, anchor, focus) {
  if (!page || !focus) return;
  const focusEl = page.querySelector(`[data-block="${focus.block}"]`);
  if (!focusEl) return;
  const anchorEl = anchor ? page.querySelector(`[data-block="${anchor.block}"]`) : focusEl;

  const start = pointIn(anchorEl || focusEl, (anchor ?? focus).offset);
  const end = pointIn(focusEl, focus.offset);
  if (!start || !end) return;

  const sel = window.getSelection();
  sel.removeAllRanges();
  const range = document.createRange();
  range.setStart(start.node, start.offset);
  range.setEnd(end.node, end.offset);
  // A backwards selection has its end before its start, which a Range refuses.
  if (range.collapsed && (start.node !== end.node || start.offset !== end.offset)) {
    range.setStart(end.node, end.offset);
    range.setEnd(start.node, start.offset);
  }
  sel.addRange(range);
}

const HEADING_SIZES = { Title: 28, Heading1: 21, Heading2: 17, Heading3: 15, Heading4: 14 };

/**
 * Word's highlighter colours are named, and half the names are not CSS names.
 * The ones that are — yellow, green, cyan, magenta, red, blue — pass through.
 */
const HIGHLIGHT_CSS = {
  darkBlue: '#00008b', darkCyan: '#008b8b', darkGreen: '#006400', darkMagenta: '#8b008b',
  darkRed: '#8b0000', darkYellow: '#8b8b00', darkGray: '#a9a9a9', lightGray: '#d3d3d3', black: '#000000',
};

/**
 * The lines already in a header or footer, ready to edit.
 *
 * OOXML has three of each — default, first page and even pages — and stores
 * them as parts full of paragraphs. The dialog edits the default one, which is
 * the one every document has and the only one most documents want.
 */
function bandLines(model, band, fallback = []) {
  const part = model?.bands?.[band === 'header' ? 'headers' : 'footers']?.default;
  if (!part) return fallback;
  const lines = (part.paragraphs || []).map((p) => p.text ?? (p.runs || []).map((r) => r.text ?? '').join(''));
  return lines.length ? lines : fallback;
}

export default function Word({ app, shell, boot }) {
  const toast = useToast();
  const [doc, setDoc] = useState(null);
  const [model, setModel] = useState(null);
  const [busy, setBusy] = useState(true);
  const [error, setError] = useState(null);
  const [tab, setTab] = useState('home');
  // One name at a time, the way the spreadsheet does it.
  const [dialog, setDialog] = useState(null);

  // What was selected when a dialog opened. Read at the moment of the press,
  // because a dialog takes focus and a selection read after that is empty.
  const selectionText = useRef('');
  const openDialog = useCallback((name) => {
    selectionText.current = window.getSelection()?.toString() || '';
    setDialog(name);
  }, []);

  // A picture from a file on disk: the bytes, its type from the extension, and
  // a size that fits a page. The engine takes the bytes and writes the part.
  const insertPictureRef = useRef(null);

  /**
   * How the page is shown. None of this is in the document: the view mode,
   * the panes, the marks, whether the system's spell checker underlines, and
   * whether the format painter is loaded. Word keeps the same things outside
   * the file, which is why they live here and not in the engine.
   */
  const [view, setView] = useState({
    mode: 'print', marks: false, ruler: false, navigation: false, focus: false, spell: true, reading: false, painting: null,
  });
  const patchView = useCallback((patch) => setView((v) => ({ ...v, ...(typeof patch === 'function' ? patch(v) : patch) })), []);
  const actRef = useRef(null);
  const [find, setFind] = useState(null);
  const pageRef = useRef(null);
  const pendingCaret = useRef(null);
  const menu = useMenu();
  const openFileRef = useRef(null);
  const appMenu = useAppMenu({
    shell,
    appKey: 'word',
    onNew: () => shell.win.create({ app: 'word' }),
    onOpen: () => openFileRef.current?.(),
  });

  const apply = useCallback(
    async (...ops) => {
      if (!doc) return null;
      // A selection sync is a statement about where the caret already is; it
      // must not then be put back, because that collapses a range the reader
      // has just made.
      const movesCaret = ops.some((op) => op.op !== 'setSelection');
      try {
        const next = await shell.doc.apply({ id: doc.id, ops });
        setDoc(next);

        if (next.patch) {
          // A splice, not a replacement: the blocks that did not change keep
          // their identity, so React re-renders the one paragraph that did.
          const { from, removed, blocks: inserted, ...rest } = next.patch;
          setModel((current) => {
            if (!current) return current;
            const blocks = current.blocks.slice();
            blocks.splice(from, removed, ...inserted);
            // Block indices are positions; renumber whatever the splice moved.
            for (let i = from; i < blocks.length; i++) {
              if (blocks[i].index !== i) blocks[i] = { ...blocks[i], index: i };
            }
            return { ...current, ...rest, blocks };
          });
          pendingCaret.current = movesCaret ? next.patch.selection ?? null : null;
        } else {
          setModel(next.model);
          pendingCaret.current = movesCaret ? next.model?.selection ?? null : null;
        }
        return next;
      } catch (err) {
        toast(err.message, { tone: 'bad' });
        return null;
      }
    },
    [doc, shell, toast]
  );

  /* ── opening and saving ──────────────────────────────────────────────── */

  useEffect(() => {
    const template = new URLSearchParams(location.search).get('template');
    const run = async () => {
      setBusy(true);
      try {
        const opened = boot.file
          ? await shell.doc.open({ path: boot.file })
          : await shell.doc.new({ kind: 'word', template: template && template !== 'blank' ? template : 'doc' });
        setDoc(opened);
        setModel(opened.model);
        if (opened.path) shell.app.addRecent({ path: opened.path, app: 'word' }).catch(() => {});
        if (opened.converted?.from) toast(`Opened from ${opened.converted.from.toUpperCase()}. Saving will write a .docx.`, { ms: 5200 });
      } catch (err) {
        setError(err.message);
      } finally {
        setBusy(false);
      }
    };
    run();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const save = useCallback(
    async (as = false) => {
      if (!doc) return false;
      let target = doc.path;
      if (as || !target) {
        target = await pickSave(shell, 'word', doc.path || doc.name);
        // A cancelled Save As is not a save; the caller must know.
        if (!target) return false;
      }
      try {
        const saved = await shell.doc.save({ id: doc.id, path: target });
        setDoc((d) => ({ ...d, ...saved, dirty: false }));
        shell.app.addRecent({ path: saved.path, app: 'word' }).catch(() => {});
        toast(`Saved ${saved.path.split(/[\\/]/).pop()}`, { tone: 'good' });
        return true;
      } catch (err) {
        toast(err.message, { tone: 'bad' });
        return false;
      }
    },
    [doc, shell, toast]
  );

  // Closing a window with unsaved work must ask, not discard.
  useDirtyGuard({ shell, dirty: doc?.dirty, name: doc?.name, onSave: () => save(false) });
  const openFile = useCallback(async () => {
    const file = await pickOpen(shell, 'word');
    if (file) openInApp(shell, file, 'word');
  }, [shell]);
  openFileRef.current = openFile;

  const exportAs = useCallback(
    async (format) => {
      if (!doc) return;
      const target = await shell.dialog.save({
        title: `Export as ${format.toUpperCase()}`,
        defaultPath: (doc.path || doc.name).replace(/\.[^.]+$/, `.${format}`),
        filters: [{ name: format.toUpperCase(), extensions: [format] }],
      });
      // A cancelled Save As is not a save; the caller must know.
      if (!target) return false;
      try {
        await shell.doc.export({ id: doc.id, format, path: target });
        toast(`Exported ${target.split(/[\\/]/).pop()}`, { tone: 'good' });
      } catch (err) {
        toast(err.message, { tone: 'bad' });
        return false;
      }
    },
    [doc, shell, toast]
  );

  useFileDrop(useCallback((files) => files.forEach((f) => openInApp(shell, f, 'word')), [shell]));

  /* ── the editing surface ─────────────────────────────────────────────── */

  const blockOf = (node) => {
    let el = node?.nodeType === 3 ? node.parentElement : node;
    while (el && !el.dataset?.block) el = el.parentElement;
    return el;
  };

  const currentPosition = useCallback(() => {
    const sel = window.getSelection();
    if (!sel || !sel.focusNode) return null;
    const el = blockOf(sel.focusNode);
    if (!el) return null;
    const index = Number(el.dataset.block);
    return {
      anchor: (() => {
        const a = blockOf(sel.anchorNode);
        return a ? { block: Number(a.dataset.block), offset: offsetIn(a, sel.anchorNode, sel.anchorOffset) } : null;
      })(),
      focus: { block: index, offset: offsetIn(el, sel.focusNode, sel.focusOffset) },
    };
  }, []);

  /**
   * The caret moved. This is deliberately *not* sent to the backend on every
   * key release: an edit carries its own position, so the only reason to tell
   * the engine separately is a click or an arrow key, and doing it per keystroke
   * doubled the round trips for no gain.
   */
  const syncSelection = useCallback(() => {
    const pos = currentPosition();
    if (!pos?.focus) return;
    const ops = [{ op: 'setSelection', anchor: pos.anchor || pos.focus, focus: pos.focus }];

    // The format painter: a loaded brush is applied to whatever is selected
    // next, then put down. The value properties are set; the toggles are
    // toggled only where the target differs, because a toggle on text that is
    // already bold would make it plain.
    const brush = view.painting;
    const selected = pos.anchor && (pos.anchor.block !== pos.focus.block || pos.anchor.offset !== pos.focus.offset);
    if (brush && selected) {
      const here = model?.format || {};
      ops.push({ op: 'setRunFormat', delta: { fontName: brush.fontName ?? null, fontSize: brush.fontSize ?? null, fontColour: brush.fontColour ?? null, highlight: brush.highlight ?? null } });
      for (const key of ['bold', 'italic', 'underline', 'strike']) {
        if (Boolean(brush[key]) !== Boolean(here[key])) ops.push({ op: 'toggleFormat', tag: key });
      }
      patchView({ painting: null });
    }
    apply(...ops);
  }, [apply, currentPosition, view.painting, model, patchView]);

  const handleBeforeInput = useCallback(
    (e) => {
      // Nothing the browser does to the DOM is kept; the engine decides.
      e.preventDefault();
      const pos = currentPosition();
      const ops = [];
      if (pos?.focus) ops.push({ op: 'setSelection', anchor: pos.anchor || pos.focus, focus: pos.focus });

      switch (e.inputType) {
        case 'insertText':
        case 'insertCompositionText':
          if (e.data) ops.push({ op: 'insertText', text: e.data });
          break;
        case 'insertParagraph':
        case 'insertLineBreak':
          ops.push({ op: 'splitParagraph' });
          break;
        case 'deleteContentBackward':
          ops.push({ op: 'deleteBackward' });
          break;
        case 'deleteContentForward':
          ops.push({ op: 'deleteForward' });
          break;
        case 'deleteByCut':
          ops.push({ op: 'deleteSelection' });
          break;
        case 'insertFromPaste':
        case 'insertFromDrop': {
          // The native event carries the payload on `dataTransfer`; a paste
          // routed through the menu arrives with neither, so the clipboard is
          // read as a fallback.
          const text = e.dataTransfer?.getData('text/plain');
          if (text) ops.push({ op: 'pasteText', text });
          else {
            shell.clipboard.readText().then((clip) => clip && apply({ op: 'pasteText', text: clip }));
            return;
          }
          break;
        }
        case 'insertReplacementText': {
          // A spell-check correction: the engine replaces the current word.
          const text = e.dataTransfer?.getData('text/plain') || e.data;
          if (text) ops.push({ op: 'insertText', text });
          break;
        }
        case 'formatBold':
          ops.push({ op: 'toggleFormat', tag: 'bold' });
          break;
        case 'formatItalic':
          ops.push({ op: 'toggleFormat', tag: 'italic' });
          break;
        case 'formatUnderline':
          ops.push({ op: 'toggleFormat', tag: 'underline' });
          break;
        default:
          return;
      }
      if (ops.length > 1 || (ops.length === 1 && ops[0].op !== 'setSelection')) apply(...ops);
    },
    [apply, currentPosition]
  );

  /**
   * `beforeinput` is attached natively, not through React.
   *
   * React's `onBeforeInput` is not this event. It is a synthetic event
   * react-dom assembles from `keypress`, `textInput` and `compositionend` — a
   * polyfill older than the standard — and it carries no `inputType`. An editor
   * wired to it cancels the browser's insertion and then asks what to do about
   * `undefined`, which is nothing: typing does absolutely nothing, silently.
   * Backspace never arrives at all, because `keypress` does not fire for it.
   *
   * The native event carries `inputType` and is cancellable, which is the whole
   * mechanism this editor runs on. `word-input.test.js` pins it.
   */
  useEffect(() => {
    const el = pageRef.current;
    if (!el) return undefined;
    el.addEventListener('beforeinput', handleBeforeInput);
    return () => el.removeEventListener('beforeinput', handleBeforeInput);
  }, [handleBeforeInput, model]);

  // The engine's caret is authoritative; after every render the DOM caret is
  // put back where the engine says it is.
  useLayoutEffect(() => {
    const target = pendingCaret.current;
    if (!target || !pageRef.current) return;
    placeSelection(pageRef.current, target.anchor, target.focus);
    pendingCaret.current = null;
  }, [model]);

  /* ── commands ────────────────────────────────────────────────────────── */

  const format = model?.format || {};

  const insertPicture = useCallback(async () => {
    const [file] = await shell.dialog.open({
      title: 'Insert picture',
      filters: [{ name: 'Pictures', extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp'] }],
    });
    if (!file) return;
    const { bytes, stat } = await shell.fs.read({ path: file });
    const ext = String(stat?.ext || file.split('.').pop()).replace('.', '').toLowerCase();
    const contentType = { png: 'image/png', gif: 'image/gif', webp: 'image/webp', bmp: 'image/bmp' }[ext] || 'image/jpeg';
    // Sized to the text width and a sensible height; the engine keeps the
    // aspect from the bytes if it can read them, and this is the fallback.
    const page = model?.section;
    const width = page ? Math.round(page.widthPx - page.margins.left - page.margins.right) : 600;
    await apply({ op: 'insertImage', name: stat?.name || file.split(/[\\/]/).pop(), contentType, data: bytes, widthPx: Math.min(width, 480), heightPx: Math.round(Math.min(width, 480) * 0.66) });
  }, [shell, apply, model]);
  insertPictureRef.current = insertPicture;

  /**
   * The ribbon's verbs that are not one engine operation.
   *
   * A view mode, a pane, reading aloud, the format painter, a change of case,
   * a cover page, a table of contents built from the headings — each is a few
   * operations in a row, or no operation at all, and the ribbon should not
   * have to know which. It names the verb; this does it.
   */
  const act = useCallback(
    async (name, arg) => {
      const blocks = model?.blocks || [];
      const sel = model?.selection;
      const at = sel?.focus?.block ?? 0;

      switch (name) {
        case 'mode':
          patchView({ mode: arg });
          return;
        case 'toggleMarks':
          patchView((v) => ({ marks: !v.marks }));
          return;
        case 'toggleRuler':
          patchView((v) => ({ ruler: !v.ruler }));
          return;
        case 'toggleNavigation':
          patchView((v) => ({ navigation: !v.navigation }));
          return;
        case 'toggleSpell':
          patchView((v) => ({ spell: !v.spell }));
          return;
        case 'focus': {
          const on = !view.focus;
          patchView({ focus: on });
          shell.win.fullscreen({ on }).catch(() => {});
          return;
        }
        case 'zoom': {
          // The window's zoom is additive on a factor; reset first so a chosen
          // level is that level and not that level times the last one.
          await shell.win.zoom({ reset: true });
          const page = model?.section?.widthPx || 794;
          const target = arg === 'width' ? (window.innerWidth - 120) / page : arg === 'page' ? (window.innerHeight - 200) / (model?.section?.heightPx || 1123) : arg === 'pages' ? 0.5 : Number(arg) || 1;
          if (Math.abs(target - 1) > 0.01) await shell.win.zoom({ delta: Math.max(0.3, Math.min(3, target)) - 1 });
          return;
        }
        case 'newWindow':
          if (!doc?.path) return toast('Save the document first, so a second window can open the same file.', { ms: 5000 });
          shell.win.create({ app: 'word', file: doc.path });
          return;
        case 'open': {
          const urls = { help: SITE.help, contact: SITE.contact, releases: SITE.releases };
          if (arg === 'about') shell.win.create({ app: 'home', query: { about: 1 } });
          else shell.shell.openExternal({ url: urls[arg] || SITE.home });
          return;
        }
        case 'readAloud': {
          const synth = window.speechSynthesis;
          if (!synth) return toast('Speech is not available on this system.', { tone: 'bad' });
          if (view.reading) {
            synth.cancel();
            patchView({ reading: false });
            return;
          }
          const text = blocks.slice(at).map((b) => b.text || (b.runs || []).map((r) => r.text).join('')).filter(Boolean).join('\n');
          if (!text.trim()) return toast('Nothing to read from here.');
          const utterance = new SpeechSynthesisUtterance(text);
          utterance.onend = () => patchView({ reading: false });
          utterance.onerror = () => patchView({ reading: false });
          patchView({ reading: true });
          synth.cancel();
          synth.speak(utterance);
          return;
        }
        case 'formatPainter': {
          if (view.painting) return patchView({ painting: null });
          const f = model?.format || {};
          patchView({ painting: { bold: f.bold, italic: f.italic, underline: f.underline, strike: f.strike, fontName: f.fontName, fontSize: f.fontSize, fontColour: f.fontColour, highlight: f.highlight } });
          toast('Formatting copied — select the words to paint it onto.', { ms: 4000 });
          return;
        }
        case 'changeCase': {
          if (!sel || sel.collapsed) return toast('Select some words first.');
          const from = sel.from ?? sel.anchor;
          const to = sel.to ?? sel.focus;
          if (from.block !== to.block) return toast('Change case works within one paragraph at a time.');
          const text = (blocks[from.block]?.text || '').slice(from.offset, to.offset);
          if (!text) return;
          const cased = {
            upper: text.toUpperCase(),
            lower: text.toLowerCase(),
            title: text.replace(/\b(\p{L})(\p{L}*)/gu, (_, a, b) => a.toUpperCase() + b.toLowerCase()),
            sentence: text.toLowerCase().replace(/(^\s*\p{L}|[.!?]\s+\p{L})/gu, (m) => m.toUpperCase()),
            toggle: [...text].map((c) => (c === c.toUpperCase() ? c.toLowerCase() : c.toUpperCase())).join(''),
          }[arg] ?? text;
          await apply({ op: 'deleteSelection' }, { op: 'insertText', text: cased }, { op: 'setSelection', anchor: from, focus: { block: from.block, offset: from.offset + cased.length } });
          return;
        }
        case 'coverPage':
          // A title paragraph at the very front, on a page of its own.
          await apply(
            { op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } },
            { op: 'insertText', text: doc?.name?.replace(/\.[^.]+$/, '') || 'Title' },
            { op: 'splitParagraph' },
            { op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } },
            { op: 'setParagraphFormat', delta: { styleId: 'Title', align: 'center' } },
            { op: 'setSelection', anchor: { block: 1, offset: 0 }, focus: { block: 1, offset: 0 } },
            { op: 'insertPageBreak' }
          );
          return;
        case 'blankPage':
          await apply({ op: 'insertPageBreak' }, { op: 'insertPageBreak' });
          return;
        case 'tableOfContents': {
          const headings = blocks
            .map((b) => ({ level: Number((/^Heading(\d)$/.exec(b.style || '') || [])[1] || 0), text: b.text || (b.runs || []).map((r) => r.text).join('') }))
            .filter((h) => h.level >= 1 && h.level <= 3 && h.text.trim());
          if (!headings.length) return toast('No headings yet. Use Heading 1, 2 and 3 on the paragraphs you want listed.', { ms: 6000 });
          const ops = [{ op: 'insertText', text: 'Contents' }, { op: 'setParagraphFormat', delta: { styleId: 'Heading1' } }, { op: 'splitParagraph' }];
          for (const h of headings) {
            ops.push({ op: 'setParagraphFormat', delta: { styleId: null } });
            ops.push({ op: 'insertText', text: h.text });
            if (h.level > 1) ops.push({ op: 'setParagraphFormat', delta: { indentDelta: h.level - 1 } });
            ops.push({ op: 'splitParagraph' });
          }
          await apply(...ops);
          toast(`Contents listed ${headings.length} heading${headings.length === 1 ? '' : 's'}. Insert it again after the headings change.`, { ms: 6000 });
          return;
        }
        case 'nextNote': {
          // The next paragraph carrying a footnote or endnote reference — the
          // references are in body order, so the notes are too.
          const list = blocks.filter((b) => (b.runs || []).some((r) => r.noteRef)).map((b) => b.index);
          if (!list.length) return toast('No footnotes in this document.', { ms: 4000 });
          const next = arg > 0 ? list.find((n) => n > at) ?? list[0] : [...list].reverse().find((n) => n < at) ?? list[list.length - 1];
          await apply({ op: 'setSelection', anchor: { block: next, offset: 0 }, focus: { block: next, offset: 0 } });
          pageRef.current?.querySelector(`[data-block="${next}"]`)?.scrollIntoView({ block: 'center', behavior: 'smooth' });
          return;
        }
        case 'showNotes': {
          const notes = pageRef.current?.querySelector('.wd-notes');
          if (!notes) return toast('No footnotes in this document.', { ms: 4000 });
          notes.scrollIntoView({ block: 'start', behavior: 'smooth' });
          return;
        }
        case 'comment': {
          const list = (model?.comments || []).map((c) => c.blockIndex).filter((n) => Number.isFinite(n)).sort((a, b) => a - b);
          if (!list.length) return;
          const next = arg > 0 ? list.find((n) => n > at) ?? list[0] : [...list].reverse().find((n) => n < at) ?? list[list.length - 1];
          await apply({ op: 'setSelection', anchor: { block: next, offset: 0 }, focus: { block: next, offset: 0 } });
          pageRef.current?.querySelector(`[data-block="${next}"]`)?.scrollIntoView({ block: 'center', behavior: 'smooth' });
          return;
        }
        case 'goto': {
          await apply({ op: 'setSelection', anchor: { block: arg, offset: 0 }, focus: { block: arg, offset: 0 } });
          pageRef.current?.querySelector(`[data-block="${arg}"]`)?.scrollIntoView({ block: 'start', behavior: 'smooth' });
          return;
        }
        default:
          return;
      }
    },
    [model, view, apply, shell, toast, doc, patchView]
  );
  actRef.current = act;

  const commands = useMemo(
    () => ({
      'file.new': { label: 'New', icon: 'new', key: 'Mod+N', run: () => shell.win.create({ app: 'word' }) },
      'file.open': { label: 'Open…', icon: 'open', key: 'Mod+O', run: openFile },
      'file.save': { label: 'Save', icon: 'save', key: 'Mod+S', global: true, run: () => save(false) },
      'file.saveAs': { label: 'Save as…', icon: 'save', key: 'Mod+Shift+S', global: true, run: () => save(true) },
      'file.pdf': { label: 'Export as PDF…', icon: 'pdf', run: () => exportAs('pdf') },
      'edit.undo': { label: 'Undo', icon: 'undo', key: 'Mod+Z', global: true, run: async () => { const n = await shell.doc.undo({ id: doc.id }); setDoc(n); setModel(n.model); } },
      'edit.redo': { label: 'Redo', icon: 'redo', key: 'Mod+Y', global: true, run: async () => { const n = await shell.doc.redo({ id: doc.id }); setDoc(n); setModel(n.model); } },
      'edit.find': { label: 'Find and replace…', icon: 'find', key: 'Mod+F', global: true, run: () => setFind({ find: '', replace: '' }) },
      'format.bold': { label: 'Bold', icon: 'bold', key: 'Mod+B', global: true, run: () => apply({ op: 'toggleFormat', tag: 'bold' }) },
      'format.italic': { label: 'Italic', icon: 'italic', key: 'Mod+I', global: true, run: () => apply({ op: 'toggleFormat', tag: 'italic' }) },
      'format.underline': { label: 'Underline', icon: 'underline', key: 'Mod+U', global: true, run: () => apply({ op: 'toggleFormat', tag: 'underline' }) },
      'format.clear': { label: 'Clear formatting', icon: 'close', run: () => apply({ op: 'clearFormat' }) },
      'insert.table': { label: 'Table', icon: 'table', run: () => apply({ op: 'insertTable', rows: 3, cols: 3 }) },
      'insert.break': { label: 'Page break', icon: 'file', run: () => apply({ op: 'insertPageBreak' }) },
      'insert.image': { label: 'Picture…', icon: 'picture', run: () => insertPictureRef.current?.() },
    }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [doc, apply, save, openFile, exportAs, shell]
  );

  useCommands(commands, [doc, model]);

  const align = (value) => apply({ op: 'setParagraphFormat', delta: { align: value } });
  const setStyle = (style) => apply({ op: 'setParagraphFormat', delta: { style: style || null } });

  if (error) {
    return (
      <AppFrame app={app} shell={shell} title="Word" menu={appMenu}>
        <Empty icon="word" title="This file could not be opened">{error}</Empty>
      </AppFrame>
    );
  }

  const section = model?.section;

  return (
    <AppFrame
      app={app}
      shell={shell}
      title={doc?.name || 'Word'}
      subtitle={doc?.converted ? `from ${doc.converted.from.toUpperCase()}` : null}
      dirty={doc?.dirty}
      menu={appMenu}
      ribbon={
        <WordRibbon
          tab={tab}
          setTab={setTab}
          doc={doc}
          model={model}
          dispatch={apply}
          commands={commands}
          shell={shell}
          menu={menu}
          save={save}
          openFile={openFile}
          exportAs={exportAs}
          openDialog={openDialog}
          insertPicture={insertPicture}
          act={act}
          view={view}
        />
      }
      status={
        <>
          <span>{doc?.path || 'Not saved yet'}</span>
          <Spacer />
          <Chip>{model?.wordCount ?? 0} words</Chip>
          <Chip>{model?.characterCount ?? 0} characters</Chip>
          <Chip>{model?.blocks?.length ?? 0} paragraphs</Chip>
        </>
      }
    >
      {busy || !model ? (
        <div style={{ flex: 1, display: 'grid', placeItems: 'center' }}>
          <Spinner style={{ width: 22, height: 22 }} />
        </div>
      ) : (
        <div className={`wd mode-${view.mode || 'print'}${view.focus ? ' focus' : ''}`}>
          <style>{CSS}</style>
          {view.navigation ? (
            <NavigationPane blocks={model.blocks} at={model.selection?.focus?.block ?? -1} onGo={(i) => act('goto', i)} onClose={() => act('toggleNavigation')} />
          ) : null}
          <div className="wd-scroll">
            {view.ruler ? <Ruler section={section} /> : null}
            <div
              className={`wd-page${view.marks ? ' marks' : ''}`}
              ref={pageRef}
              contentEditable
              suppressContentEditableWarning
              spellCheck={view.spell !== false}
              onMouseUp={syncSelection}
              onKeyDown={(e) => {
                // Arrow keys and Home/End move the caret without an edit, so the
                // engine is told where it landed — after the browser has moved
                // it, which is why this waits a tick.
                if (/^(Arrow|Home|End|Page)/.test(e.key)) setTimeout(syncSelection, 0);
              }}
              onContextMenu={(e) => menu.open(e, menuItems(commands, ['edit.undo', 'edit.redo', '-', 'format.bold', 'format.italic', 'format.underline', '-', 'edit.find']))}
              style={{
                // The document's own base font, from its stylesheet's defaults.
                fontFamily: model.resolvedStyles?.['*default*']?.fontName || undefined,
                fontSize: model.resolvedStyles?.['*default*']?.sizePx ? `${model.resolvedStyles['*default*'].sizePx}px` : undefined,
                width: section ? Math.round(section.widthPx) : 794,
                minHeight: section ? Math.round(section.heightPx) : 1123,
                paddingTop: section?.margins.top ?? 96,
                paddingRight: section?.margins.right ?? 96,
                paddingBottom: section?.margins.bottom ?? 96,
                paddingLeft: section?.margins.left ?? 96,
                // The bands read the margins too, to line up with the body.
                '--wd-margin-left': `${section?.margins.left ?? 96}px`,
                '--wd-margin-right': `${section?.margins.right ?? 96}px`,
              }}
            >
              {model.bands?.watermark ? (
                <div className="wd-watermark" contentEditable={false} aria-hidden="true" style={{ color: model.bands.watermark.colour || 'silver', transform: `rotate(${model.bands.watermark.rotation ?? 315}deg)` }}>
                  {model.bands.watermark.text}
                </div>
              ) : null}
              <Band kind="header" bands={model.bands} section={section} onEdit={() => setDialog('header')} />
              {groupTables(model.blocks).map((item) =>
                item.table ? (
                  <TableGroup key={`t${item.table.id}`} table={item.table} labels={model.listLabels} styles={model.resolvedStyles} />
                ) : (
                  <Block key={item.index} block={item} labels={model.listLabels} styles={model.resolvedStyles} />
                )
              )}
              <Notes notes={model.footnotes} kind="footnotes" styles={model.resolvedStyles} />
              <Notes notes={model.endnotes} kind="endnotes" styles={model.resolvedStyles} />
              <Band kind="footer" bands={model.bands} section={section} onEdit={() => setDialog('footer')} />
            </div>
          </div>
          {menu.node}
          {find ? (
            <FindPanel
              state={find}
              onChange={setFind}
              onClose={() => setFind(null)}
              onReplaceAll={async () => {
                await apply({ op: 'replaceAll', find: find.find, replace: find.replace });
                toast('Replaced', { tone: 'good' });
              }}
            />
          ) : null}
        </div>
      )}

      {dialog === 'link' ? (
        <LinkDialog
          current={model?.format?.link || null}
          selectedText={selectionText.current || null}
          onClose={() => setDialog(null)}
          onApply={async (url) => {
            await apply({ op: 'setLink', url });
            setDialog(null);
          }}
          onRemove={async () => {
            await apply({ op: 'setLink', url: null });
            setDialog(null);
          }}
        />
      ) : null}

      {dialog === 'table' ? (
        <TableDialog
          onClose={() => setDialog(null)}
          onInsert={async (spec) => {
            // The engine takes a size; a header row is a formatting decision it
            // does not model yet, so it is not pretended.
            await apply({ op: 'insertTable', rows: spec.rows, cols: spec.cols });
            setDialog(null);
          }}
        />
      ) : null}

      {dialog === 'header' || dialog === 'footer' ? (
        <BandDialog
          band={dialog}
          current={bandLines(model, dialog)}
          onClose={() => setDialog(null)}
          onApply={async (lines) => {
            await apply({ op: 'setBand', band: dialog, lines });
            setDialog(null);
          }}
        />
      ) : null}

      {dialog === 'pageNumber' ? (
        <BandDialog
          band="footer"
          current={bandLines(model, 'footer', ['{PAGE} of {PAGES}'])}
          onClose={() => setDialog(null)}
          onApply={async (lines) => {
            await apply({ op: 'setBand', band: 'footer', lines });
            setDialog(null);
          }}
        />
      ) : null}

      {dialog === 'comment' ? (
        <CommentDialog
          onClose={() => setDialog(null)}
          onAdd={async (text) => {
            await apply({ op: 'addComment', text });
            setDialog(null);
            toast('Comment added', { tone: 'good' });
          }}
        />
      ) : null}

      {dialog === 'comments' ? (
        <CommentsDialog comments={model?.comments || []} onClose={() => setDialog(null)} onGoto={() => setDialog(null)} />
      ) : null}

      {dialog === 'find' ? (
        <FindDialog
          onClose={() => setDialog(null)}
          onReplaceAll={async (find, replace, matchCase) => {
            const next = await apply({ op: 'replaceAll', find, replace, matchCase });
            return next ? 'Replaced every match.' : 'Nothing matched.';
          }}
        />
      ) : null}

      {dialog === 'wordCount' ? <WordCountDialog blocks={model?.blocks || []} onClose={() => setDialog(null)} /> : null}

      {dialog === 'dateTime' ? (
        <DateTimeDialog
          onClose={() => setDialog(null)}
          onInsert={async (text) => {
            await apply({ op: 'insertText', text });
            setDialog(null);
          }}
        />
      ) : null}

      {dialog === 'symbol' ? <SymbolDialog onClose={() => setDialog(null)} onInsert={(c) => apply({ op: 'insertText', text: c })} /> : null}

      {dialog === 'properties' ? <PropertiesDialog doc={doc} model={model} onClose={() => setDialog(null)} /> : null}

      {dialog === 'shortcuts' ? <ShortcutsDialog onClose={() => setDialog(null)} /> : null}

      {dialog === 'tracked' ? (
        <TrackedDialog
          blocks={model?.blocks || []}
          onClose={() => setDialog(null)}
          onGoto={(index) => {
            setDialog(null);
            act('goto', index);
          }}
        />
      ) : null}
    </AppFrame>
  );
}

// Memoised: with the patch above, a keystroke changes one block, and only
// that one should re-render. Without this the saving is thrown away in
// reconciliation.
/**
 * Consecutive cell paragraphs of one table, gathered into a table.
 *
 * The engine's frame is a flat list of paragraphs, and a table's cells are
 * paragraphs whose `container` names their cell — `t1024:r2:c0`. Painting them
 * one under another is how a three-by-three table looked like nine lines and
 * "Insert table" looked broken. This groups a run of blocks that share a table
 * into rows and cells, keeping each cell paragraph as its own editable
 * [data-block] so the caret, selection and typing keep working inside it.
 */
function groupTables(blocks) {
  const out = [];
  let current = null;
  const flush = () => {
    if (current) out.push({ table: current });
    current = null;
  };
  for (const block of blocks) {
    const at = /^(t\d+):r(\d+):c(\d+)$/.exec(String(block.container || ''));
    if (!at) {
      flush();
      out.push(block);
      continue;
    }
    if (!current || current.id !== at[1]) {
      flush();
      current = { id: at[1], rows: new Map() };
    }
    const row = Number(at[2]);
    const cell = Number(at[3]);
    if (!current.rows.has(row)) current.rows.set(row, new Map());
    const cells = current.rows.get(row);
    if (!cells.has(cell)) cells.set(cell, []);
    cells.get(cell).push(block);
  }
  flush();
  return out;
}

function TableGroup({ table, labels, styles }) {
  const rows = [...table.rows.entries()].sort((a, b) => a[0] - b[0]);
  return (
    <table className="wd-table">
      <tbody>
        {rows.map(([r, cells]) => (
          <tr key={r}>
            {[...cells.entries()]
              .sort((a, b) => a[0] - b[0])
              .map(([c, paragraphs]) => (
                <td key={c}>
                  {paragraphs.map((block) => (
                    <Block key={block.index} block={block} labels={labels} styles={styles} />
                  ))}
                </td>
              ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

/**
 * The header or footer the page shows, drawn in the margin where Word draws
 * it and greyed the way Word greys it while the body has the caret. The page
 * here is one continuous sheet, so this is the band for page 1 (the first-page
 * band when the section asks for one), with the page fields resolved the only
 * way a continuous sheet can: page 1 of the count the engine reports. Not
 * editable in place — a double-click opens the band's dialog, as in Word.
 */
function Band({ kind, bands, section, onEdit }) {
  const set = bands?.[kind === 'header' ? 'headers' : 'footers'];
  if (!set) return null;
  // With a title page, page 1 gets the FIRST-page band — and nothing at all
  // when the file defines none, which is how a cover page has no header.
  const band = section?.titlePage ? set.first ?? null : set.default || set.even || set.first;
  const paragraphs = band?.paragraphs || [];
  if (!paragraphs.length) return null;
  const of = section?.pageCount || null;
  // w:pgMar's header/footer distance: from the page edge to the band, 48 px by default.
  const distance = Math.round(section?.margins?.[kind] ?? 48);
  return (
    <div
      className={`wd-band wd-${kind}`}
      contentEditable={false}
      style={kind === 'header' ? { top: distance } : { bottom: distance }}
      onDoubleClick={onEdit}
      title={`Double-click to edit the ${kind}`}
    >
      {paragraphs.map((p, i) => (
        <p key={i} className="wd-band-line" style={{ textAlign: p.align === 'both' ? 'justify' : p.align || undefined }}>
          {(p.runs || []).map((r, j) => {
            const text = r.field === 'PAGE' ? '1' : r.field === 'NUMPAGES' && of ? String(of) : r.text ?? '';
            return (
              <span key={j} style={{ fontWeight: r.bold ? 700 : undefined, fontStyle: r.italic ? 'italic' : undefined, textDecoration: r.underline ? 'underline' : undefined }}>
                {withTabs(text)}
              </span>
            );
          })}
        </p>
      ))}
    </div>
  );
}

/** Word's default tab stops: every half inch, 48 px at 96 dpi. */
const DEFAULT_TAB_PX = 48;

/**
 * A run's text, with each tab as its own span so the layout pass below can
 * give it the width that reaches the next stop. The span still CONTAINS the
 * tab character, so the paragraph's text — and every caret offset computed
 * from it — is exactly what the engine holds.
 */
function withTabs(text) {
  if (!text || !text.includes('\t')) return text;
  const out = [];
  const parts = text.split('\t');
  parts.forEach((part, i) => {
    if (part) out.push(part);
    if (i < parts.length - 1) out.push(<span key={`t${i}`} className="wd-tab">{'\t'}</span>);
  });
  return out;
}

/**
 * Size each tab span so the text after it lands on the next stop — the
 * paragraph's own stops first, Word's half-inch defaults after the last of
 * them. A centre or right stop measures the text that follows the tab and
 * pulls it back by that much; a decimal stop aligns on the point. Read live,
 * in document order, because each width moves everything after it.
 */
function sizeTabs(p, stops) {
  const tabs = p.querySelectorAll('.wd-tab');
  if (!tabs.length) return;
  const paddingLeft = parseFloat(getComputedStyle(p).paddingLeft) || 0;
  const left = p.getBoundingClientRect().left + paddingLeft;
  const custom = (stops || []).filter((s) => s && s.posPx > 0);
  const image = p.querySelector('.wd-image');
  for (const span of tabs) {
    span.style.width = '';
    const x = span.getBoundingClientRect().left - left;
    const stop = custom.find((s) => s.posPx > x + 1) || { posPx: (Math.floor(x / DEFAULT_TAB_PX) + 1) * DEFAULT_TAB_PX, align: 'left' };
    let width = stop.posPx - x;
    if (stop.align === 'center' || stop.align === 'right' || stop.align === 'decimal') {
      // How wide the text between this tab and the next is — the amount a
      // right stop pulls it back, half of it for a centre stop.
      const range = document.createRange();
      range.setStartAfter(span);
      let next = span.nextSibling;
      let end = null;
      const following = [];
      while (next) {
        if (next.nodeType === 1 && (next.classList.contains('wd-tab') || next === image)) { end = next; break; }
        if (next.nodeType === 1 && next.querySelector('.wd-tab')) { end = next.querySelector('.wd-tab'); break; }
        following.push(next);
        next = next.nextSibling;
      }
      if (end) range.setEndBefore(end); else range.setEnd(p, p.childNodes.length);
      let w = range.getBoundingClientRect().width;
      if (stop.align === 'decimal') {
        // The point's position, taken as its share of the text's width — an
        // estimate, but a decimal stop lines up a column of figures, and the
        // figures are the same width.
        const text = range.toString();
        const dot = text.search(/[.,]/);
        if (dot >= 0) w = w * (dot / Math.max(1, text.length));
      }
      width -= stop.align === 'center' ? w / 2 : w;
    }
    span.style.width = `${Math.max(2, Math.round(width))}px`;
    span.dataset.leader = stop.leader || '';
  }
}

/** Paragraph borders as CSS: one line per side, in the file's colour and weight. */
const BORDER_CSS = { single: 'solid', thick: 'solid', double: 'double', dotted: 'dotted', dashed: 'dashed', dashSmallGap: 'dashed', dotDash: 'dashed', dotDotDash: 'dashed', wave: 'solid', thinThickSmallGap: 'double', thickThinSmallGap: 'double' };
function borderStyle(borders) {
  if (!borders) return {};
  const out = {};
  for (const side of ['top', 'right', 'bottom', 'left']) {
    const b = borders[side];
    if (!b) continue;
    const key = `border${side[0].toUpperCase()}${side.slice(1)}`;
    out[key] = `${b.widthPx}px ${BORDER_CSS[b.style] || 'solid'} ${b.colour}`;
    // Word keeps `space` points between the text and the line.
    out[`padding${side[0].toUpperCase()}${side.slice(1)}`] = `${Math.round((b.spacePt || 1) * (96 / 72))}px`;
  }
  return out;
}

/**
 * A paragraph's CSS from what the engine says about it. The paragraph's named
 * style, resolved by the engine from the document's own stylesheet, gives the
 * font, size, colour and spacing; direct formatting on the paragraph beats
 * it, exactly as in Word. A document with no stylesheet (RTF, Markdown) falls
 * back to the app's own heading sizes. Shared by the body's blocks and the
 * paragraphs inside a text box, so the two cannot drift apart.
 */
function paragraphCss(block, styles) {
  const named = styles ? styles[block.style] ?? styles['*default*'] ?? null : null;
  const heading = Boolean(block.style && /Title|Heading/.test(block.style));
  const align = block.align || named?.align || null;
  const hanging = block.hangingPx ?? named?.hangingPx ?? null;
  const firstLine = block.firstLinePx ?? named?.firstLinePx ?? null;
  return {
    fontFamily: named?.fontName || undefined,
    fontSize: named?.sizePx ? `${Math.round(named.sizePx * 100) / 100}px` : HEADING_SIZES[block.style] ? `${HEADING_SIZES[block.style]}px` : undefined,
    fontWeight: named ? (named.bold ? 700 : undefined) : heading ? 600 : undefined,
    fontStyle: named?.italic ? 'italic' : undefined,
    textDecoration: named?.underline ? 'underline' : undefined,
    textTransform: named?.caps ? 'uppercase' : undefined,
    fontVariant: named?.smallCaps ? 'small-caps' : undefined,
    color: named?.colour || undefined,
    // 'both' is OOXML for justified; the other three are CSS already.
    textAlign: align === 'both' ? 'justify' : align || undefined,
    marginTop: named?.spaceBeforePx != null ? Math.round(named.spaceBeforePx) : heading ? '1.1em' : undefined,
    marginBottom: named?.spaceAfterPx != null ? Math.round(named.spaceAfterPx) : undefined,
    marginLeft: block.indentLevel ? block.indentLevel * 24 : block.indent ? block.indent * 24 : named?.indentPx ? Math.round(named.indentPx) : undefined,
    marginRight: block.rightPx ? Math.round(block.rightPx) : undefined,
    // A first-line indent pushes the first line in; a hanging one pulls it out
    // and the rest of the paragraph in by the same amount, the way a list does.
    textIndent: firstLine ? Math.round(firstLine) : hanging ? -Math.round(hanging) : undefined,
    paddingLeft: hanging ? Math.round(hanging) : undefined,
    lineHeight: block.lineSpacing ? block.lineSpacing * 1.2 : named?.lineFactor ? named.lineFactor * 1.2 : undefined,
    // Shading and borders: the paragraph's own, or its style's, drawn edge to
    // edge like Word.
    backgroundColor: block.shading || named?.shading || undefined,
    ...borderStyle(block.borders),
  };
}

/** The tab stops that apply: the paragraph's own, else its style's. */
const tabStops = (block, styles) => block.tabs ?? (styles ? (styles[block.style] ?? styles['*default*'])?.tabs : null) ?? null;

/** One run of text: its direct formatting, its link, its tabs. */
function RunSpan({ run }) {
  // A footnote/endnote reference, or the mark at the head of the note: the
  // number is drawn by CSS from `data-n`, so the span holds no text and the
  // engine's caret offsets stay exactly right.
  if (run.noteRef || run.noteMark) {
    const n = run.noteRef ? run.noteRef.n : run.noteMark?.n;
    return <span className={run.noteRef ? 'wd-noteref' : 'wd-notemark'} data-n={n ?? '?'} contentEditable={false} title={run.noteRef ? `${run.noteRef.kind} ${n}` : undefined} />;
  }
  return (
    <span
      style={{
        fontWeight: run.bold ? 700 : undefined,
        fontStyle: run.italic ? 'italic' : undefined,
        textDecoration: [run.underline ? 'underline' : '', run.strike ? 'line-through' : ''].filter(Boolean).join(' ') || undefined,
        // The engine reports `fontColour` as bare hex, the way the file
        // stores it; the size is in points, the way Word means it. The
        // painter used to read `colour` and paint the size in pixels, so
        // a colour never showed and 12 pt drew at two-thirds size.
        color: run.fontColour ? `#${run.fontColour}` : run.link != null ? '#0563C1' : undefined,
        backgroundColor: run.highlight ? HIGHLIGHT_CSS[run.highlight] || run.highlight : undefined,
        fontFamily: run.fontName || undefined,
        fontSize: run.fontSize ? `${run.fontSize}pt` : undefined,
        // A hyperlink is blue and underlined unless the run says otherwise;
        // a footnote reference sits superscript, a chemical formula sub.
        ...(run.link != null && !run.underline && !run.fontColour ? { textDecoration: 'underline' } : {}),
        ...(run.vertAlign === 'superscript' ? { verticalAlign: 'super', fontSize: '0.65em' } : run.vertAlign === 'subscript' ? { verticalAlign: 'sub', fontSize: '0.65em' } : {}),
        textTransform: run.caps ? 'uppercase' : undefined,
        fontVariant: run.smallCaps ? 'small-caps' : undefined,
      }}
    >
      {withTabs(run.text)}
    </span>
  );
}

/**
 * A text box anchored in a paragraph — a cover page's title block, a pull
 * quote, a sidebar. Drawn in flow under its paragraph at the size the file
 * gives it, filled and outlined as the shape says, its paragraphs painted
 * exactly like the body's. Not editable: the engine cannot rebuild a box, so
 * the caret is kept out of it rather than allowed to make edits that vanish.
 */
function TextBox({ box, styles }) {
  const ref = React.useRef(null);
  React.useLayoutEffect(() => {
    if (!ref.current) return;
    for (const p of ref.current.querySelectorAll('.wd-box-p')) {
      if (p.querySelector('.wd-tab')) sizeTabs(p, p._tabs || null);
    }
  });
  const style = {
    width: box.widthPx ? Math.min(box.widthPx, 720) : undefined,
    minHeight: box.heightPx ? Math.min(box.heightPx, 900) : undefined,
    backgroundColor: box.fill || undefined,
    border: box.line ? `1px solid ${box.line}` : undefined,
    margin: box.hAlign === 'center' ? '6px auto' : box.hAlign === 'right' ? '6px 0 6px auto' : '6px 0',
  };
  return (
    <div ref={ref} className="wd-textbox" contentEditable={false} style={style} title={box.name || undefined}>
      {box.paragraphs.map((p, i) => (
        <p key={i} className="wd-box-p" style={paragraphCss(p, styles)} ref={(el) => { if (el) el._tabs = tabStops(p, styles); }}>
          {(p.runs || []).length ? p.runs.map((run, j) => <RunSpan key={j} run={run} />) : <br />}
          {(p.images || []).map((image, j) => (
            <img key={j} className="wd-image" src={image.href} alt={image.name || ''} draggable={false} style={{ width: image.widthPx ? Math.min(image.widthPx, 640) : undefined, height: 'auto', maxWidth: '100%', display: 'block', margin: '4px 0' }} />
          ))}
        </p>
      ))}
    </div>
  );
}

/**
 * The footnotes (or endnotes) under the body: a short rule, then each note's
 * paragraphs painted like the body's, its number drawn by the mark at its
 * head. Read-only, like the references that point at them.
 */
function Notes({ notes, kind, styles }) {
  if (!notes?.length) return null;
  return (
    <div className={`wd-notes wd-${kind}`} contentEditable={false}>
      {notes.map((note) => (
        <div key={note.id} className="wd-note" id={`wd-${kind}-${note.n}`}>
          {note.paragraphs.map((p, i) => (
            <p key={i} className="wd-box-p" style={paragraphCss(p, styles)}>
              {(p.runs || []).length ? p.runs.map((run, j) => <RunSpan key={j} run={run} />) : <br />}
            </p>
          ))}
        </div>
      ))}
    </div>
  );
}

const Block = React.memo(function Block({ block, labels, styles }) {
  const ref = React.useRef(null);
  const hasTabs = (block.runs || []).some((r) => r.text && r.text.includes('\t'));
  React.useLayoutEffect(() => {
    if (hasTabs && ref.current) sizeTabs(ref.current, tabStops(block, styles));
  });
  const style = paragraphCss(block, styles);

  // A list marker arrives as `{ label, indentPx, bullet }` — the text to draw,
  // how far in it sits, and whether it is a bullet or a number. Drawing the
  // object itself is a React crash, and it was one: pressing "Bulleted list"
  // took the whole page down.
  const mark = labels?.[block.index];
  const label = typeof mark === 'string' ? mark : mark?.label ?? null;
  const markerIndent = typeof mark === 'object' && mark?.indentPx ? mark.indentPx : null;
  return (
    <p ref={ref} className="wd-block" data-block={block.index} data-style={block.style || 'Normal'} style={markerIndent ? { ...style, marginLeft: markerIndent } : style}>
      {label ? <span className="wd-marker" contentEditable={false}>{label}</span> : null}
      {(block.runs || []).length ? block.runs.map((run, i) => <RunSpan key={i} run={run} />) : <br />}
      {/*
        Pictures, charts and shapes sit under the paragraph's text as blocks —
        the engine's own honest simplification of float layout. Not editable:
        the caret has no business inside a picture, and letting the browser
        put it there is how an image gets deleted by a stray Backspace.
      */}
      {(block.images || []).map((image, i) => (
        <img
          key={i}
          className="wd-image"
          contentEditable={false}
          src={image.href}
          alt={image.name || ''}
          draggable={false}
          style={{ width: image.widthPx ? Math.min(image.widthPx, 640) : undefined, height: 'auto', maxWidth: '100%', display: 'block', margin: '6px 0' }}
        />
      ))}
      {(block.textBoxes || []).map((box, i) => (
        <TextBox key={i} box={box} styles={styles} />
      ))}
    </p>
  );
});

function FindPanel({ state, onChange, onClose, onReplaceAll }) {
  return (
    <div className="wd-find">
      <Icon name="find" size={14} />
      <input
        autoFocus
        className="rw-input"
        placeholder="Find"
        value={state.find}
        onChange={(e) => onChange({ ...state, find: e.target.value })}
      />
      <input
        className="rw-input"
        placeholder="Replace with"
        value={state.replace}
        onChange={(e) => onChange({ ...state, replace: e.target.value })}
      />
      <Button label="Replace all" primary onClick={onReplaceAll} disabled={!state.find} />
      <Button icon="close" title="Close" onClick={onClose} />
    </div>
  );
}

const CSS = `
.wd { flex: 1; display: flex; flex-direction: column; min-height: 0; position: relative; }
.wd-scroll { flex: 1; overflow: auto; padding: 26px 0 40px; display: flex; justify-content: center; background: var(--window); }
.wd-page {
  background: #fff; color: #111; box-shadow: var(--shadow-2); border-radius: 2px;
  outline: none; font-family: Calibri, "Segoe UI", system-ui, sans-serif; font-size: 15px;
  line-height: 1.5; caret-color: var(--accent); position: relative;
}
/* Headers and footers sit in the margins, greyed while the body has the caret. */
.wd-band { position: absolute; left: 0; right: 0; color: #777; font-size: 13px; line-height: 1.35; user-select: none; cursor: default; }
.wd-band .wd-band-line { margin: 0; padding: 0 var(--wd-margin-right, 96px) 0 var(--wd-margin-left, 96px); min-height: 1.2em; white-space: pre-wrap; }
.wd-band:hover { color: #333; }
.wd.mode-web .wd-band, .wd.mode-draft .wd-band, .wd.mode-read .wd-band, .wd.mode-outline .wd-band { display: none; }
:root[data-theme='dark'] .wd-page { background: #f7f7f5; }
/* pre-wrap: a run of spaces, a line break inside a paragraph (w:br) and a tab
   all mean what they meant in Word, instead of collapsing to one space. */
.wd-block { margin: 0 0 0.55em; min-height: 1.2em; white-space: pre-wrap; }
.wd-tab { display: inline-block; white-space: pre; tab-size: 0; overflow: hidden; vertical-align: baseline; min-width: 2px; }
.wd-tab[data-leader="dot"] { background: radial-gradient(circle, currentColor 0.6px, transparent 0.9px) 0 calc(100% - 3px) / 4px 2px repeat-x; }
.wd-tab[data-leader="hyphen"] { background: linear-gradient(currentColor, currentColor) 0 calc(100% - 3px) / 3px 1px repeat-x; }
.wd-tab[data-leader="underscore"] { border-bottom: 1px solid currentColor; }
.wd-marker { color: #555; margin-right: 6px; user-select: none; }
/* A watermark: the header's WordArt, drawn behind the body as Word does. */
.wd-watermark {
  position: absolute; left: 0; right: 0; top: 380px; text-align: center; pointer-events: none; user-select: none;
  font-family: Calibri, "Segoe UI", sans-serif; font-size: 150px; font-weight: 400; opacity: .35; letter-spacing: .02em; z-index: 0;
}
.wd-page > .wd-block, .wd-page > .wd-table, .wd-page > .wd-notes { position: relative; z-index: 1; }
/* Footnote references and the notes themselves. */
.wd-noteref::after, .wd-notemark::after { content: attr(data-n); vertical-align: super; font-size: 0.65em; line-height: 0; }
.wd-notemark::after { margin-right: 3px; }
.wd-notes { margin-top: 28px; padding-top: 6px; border-top: 1px solid #333; width: 33%; min-width: 220px; font-size: 0.85em; user-select: none; }
.wd-notes .wd-note { width: 300%; }
.wd-notes .wd-box-p { margin: 0 0 3px; white-space: pre-wrap; }
/* A text box: in flow under its paragraph, the caret kept out of it. */
.wd-textbox { display: block; box-sizing: border-box; padding: 4px 8px; max-width: 100%; overflow: hidden; user-select: none; }
.wd-textbox .wd-box-p { margin: 0; min-height: 1.2em; white-space: pre-wrap; }
.wd-table { border-collapse: collapse; width: 100%; margin: 0.6em 0; }
.wd-table td { border: 1px solid #bbb; padding: 4px 7px; vertical-align: top; }
.wd-find {
  position: absolute; top: 12px; right: 22px; display: flex; align-items: center; gap: 7px;
  background: var(--surface); border: 1px solid var(--line); border-radius: var(--r-3);
  padding: 7px 10px; box-shadow: var(--shadow-2); z-index: 20;
}
.wd-find .rw-input { width: 160px; }
`;
