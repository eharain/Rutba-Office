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
import { flushSync } from 'react-dom';
import { Button, Icon, Spacer, Chip, Empty, Spinner, ZoomSlider, useToast, useMenu, useCommands, menuItems, formatWhen } from '@rutba/office-ui';
import { AppFrame, useAppMenu, pickOpen, pickSave, confirmDiscard, useFileDrop, openInApp , useDirtyGuard } from '../shell.js';
import { SITE } from '@rutba/office-formats/registry';
import WordRibbon from './word/ribbon.js';
import { NavigationPane, installWordStyles } from './word/panes.js';
import { Ruler, TableGrips, installRulerStyles } from './word/ruler.js';
import { selectionToSend } from './word/caret.js';
import { PrintDialog, defaultPrintOptions } from '../print.js';
import { geometryOf, layPages, clearPages, sliceRuns, pageOfElement, columnBoxesOf, pageTopOf, pageHeightOf, pageIndexAt } from './word/pages.js';

installWordStyles();
installRulerStyles();
import {
  LinkDialog, TableDialog, BandDialog, CommentDialog, CommentsDialog, FindDialog, WordCountDialog,
  DateTimeDialog, SymbolDialog, PropertiesDialog, ShortcutsDialog, TrackedDialog, NoteDialog, WatermarkDialog,
  BookmarkDialog, CrossReferenceDialog, CaptionDialog,
} from './word/dialogs.js';
import { lineBoxes, rectOf } from './word/pages.js';
import { MathRun, mathHostOf, EQUATION_CSS, EquationDialog, clipOf, CLIP_TYPE } from './word/equations.js';
import { useMailings, installMailingsStyles } from './word/mailings.js';
import { useEnvelopesLabels, installEnvelopeStyles } from './word/envelopes.js';
import { MERGE_KINDS } from '@rutba/ooxml/mailmerge';

installMailingsStyles();
installEnvelopeStyles();

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
  // The second part of a paragraph split across pages starts partway in — and
  // a part that holds only the paragraph's pictures starts past its text,
  // where the caret has no business: it goes to the end of the words.
  const at = offsetWithin(blockEl, node, offset) + Number(blockEl.dataset?.from || 0);
  const length = Number(blockEl.dataset?.length);
  return Number.isFinite(length) ? Math.min(at, length) : at;
}

function offsetWithin(blockEl, node, offset) {

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
    if (seen + len >= offset) {
      const at = Math.max(0, Math.min(len, offset - seen));
      // An equation's one character is not drawn (its MathML is, in the
      // host's shadow), so the caret goes beside the host instead: before it
      // at its start, after it at its end.
      const host = mathHostOf(node);
      if (host && host.parentNode) {
        const index = [...host.parentNode.childNodes].indexOf(host);
        return { node: host.parentNode, offset: index + (at > 0 ? 1 : 0) };
      }
      return { node, offset: at };
    }
    seen += len;
    node = walker.nextNode();
  }
  return { node: blockEl, offset: blockEl.childNodes.length };
}

/**
 * The element that draws a block at a character offset. One element for
 * most paragraphs; a paragraph split across pages has one per page, each
 * marked with the offset it starts at, and the caret goes in the part that
 * holds the offset — the second part when the offset is exactly the split,
 * which is the head of the next page, as in Word.
 */
function partFor(page, index, offset) {
  let best = null;
  let bestFrom = -1;
  for (const el of page.querySelectorAll(`[data-block="${index}"]`)) {
    const from = Number(el.dataset.from || 0);
    if (from <= offset && from > bestFrom) {
      best = el;
      bestFrom = from;
    }
  }
  return best;
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
  const focusEl = partFor(page, focus.block, focus.offset);
  if (!focusEl) return;
  const anchorEl = anchor ? partFor(page, anchor.block, anchor.offset) : focusEl;

  const within = (el, offset) => pointIn(el, offset - Number(el.dataset?.from || 0));
  const start = within(anchorEl || focusEl, (anchor ?? focus).offset);
  const end = within(focusEl, focus.offset);
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
  // The equation editor: null, or what it opened on — a new equation, or
  // one being edited at `block`/`offset`.
  const [equation, setEquation] = useState(null);
  // How much of the flow is mounted. A specification of three thousand
  // paragraphs and two thousand table cells reached the engine in half a
  // second and then took six more to mount, all before the first paint.
  // The first screens mount at once and the rest follows in slices while
  // the person is already reading; block indices are stable, so nothing
  // moves under the caret.
  const [mounted, setMounted] = useState(MOUNT_FIRST);

  // The note dialog carries its own state: which kind, and — when editing —
  // which note and its current words.
  const [noteDialog, setNoteDialog] = useState(null);

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
    mode: 'print', marks: false, ruler: true, navigation: false, focus: false, spell: true, reading: false, painting: null, zoom: 1,
    // How tracked changes are shown — a VIEW preference, never saved in the
    // file (only whether recording is ON is, in settings.xml). Word's own
    // default for a document that already carries changes is Simple Markup:
    // final text, with a change bar in the margin, not every insertion and
    // deletion inline.
    markupMode: 'simple',
  });
  const patchView = useCallback((patch) => setView((v) => ({ ...v, ...(typeof patch === 'function' ? patch(v) : patch) })), []);
  const actRef = useRef(null);
  const [find, setFind] = useState(null);
  // The picture the reader clicked — the one Wrap Text and Position act on.
  // A click on a picture says so through a DOM event from the memoised
  // paragraph; a caret move takes the pick away, as in Word.
  const [picked, setPicked] = useState(null);
  // True while a picture handle is being dragged: the mouseup that ends
  // the drag lands wherever the pointer is, and must not take the pick away.
  const pictureDrag = useRef(false);
  const pageRef = useRef(null);
  const pendingCaret = useRef(null);
  // Mailings → Preview Results is on: the page is a record's words.
  const previewRef = useRef(false);
  previewRef.current = Boolean(model?.mailMerge?.preview);
  // The caret as this editor last left it: sent with an edit whose answer is
  // not painted yet, and placed from the engine's answer. A keystroke that
  // finds the caret exactly there does not resend it — see word/caret.js.
  const sentCaret = useRef(null);
  const placedCaret = useRef(null);
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
        // A recovery copy the launcher offered: opened as the document it
        // came from, dirty, because what is on screen is not what is on disk.
        const recover = new URLSearchParams(location.search).get('recover');
        // Finish & Merge → Edit Individual Documents: the merged letters are
        // a session already, made for this window to take over.
        const adopt = new URLSearchParams(location.search).get('session');
        const opened = adopt
          ? await shell.doc.adopt({ id: adopt })
          : recover
          ? await shell.doc.recover({ file: recover })
          : boot.file
            ? await shell.doc.open({ path: boot.file, kind: 'doc' })
            : await shell.doc.new({ kind: 'word', template: template && template !== 'blank' ? template : 'doc' });
        if (recover) toast('Recovered unsaved work. Save it to keep it.', { ms: 6000 });
        setDoc(opened);
        setModel(opened.model);
        if (opened.path) shell.app.addRecent({ path: opened.path, app: 'word' }).catch(() => {});
        // What saving will actually do, which is not one answer: a .md
        // opened here saves as .md, and an .rtf cannot be saved at all
        // until it is given a new name. Both used to be promised a .docx.
        if (opened.converted?.from) {
          const was = opened.converted.from.toUpperCase();
          toast(
            opened.converted.writesBack
              ? `Opened from ${was}. Saving writes the ${was} back.`
              : `Opened from ${was}. This build cannot write ${was} — Save as will write a .docx.`,
            { ms: 5200 }
          );
        }
      } catch (err) {
        setError(err.message);
      } finally {
        setBusy(false);
      }
    };
    run();
    // A window an earlier build zoomed stays zoomed across restarts — the
    // level is kept per origin — and the page carries the zoom now, so the
    // window itself goes back to 100%.
    shell.win.zoom({ reset: true }).catch(() => {});
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
        // and a workbook or a deck must be laid out before it is one. The
        // same door for all three, so what is printed and what is exported
        // can never drift apart.
        if (format === 'pdf') await shell.print.pdf({ id: doc.id, path: target, options: options || defaultPrintOptions('doc') });
        else await shell.doc.export({ id: doc.id, format, path: target });
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

  /**
   * A DOM selection end as a block position. Inside a paragraph it is the
   * character offset. A selection dragged into the margin, a triple-click or
   * Ctrl+A ends on the page itself (or a sheet behind the words) with a child
   * index, not in any paragraph — and reading that as no position at all
   * meant the engine never learnt the range: Delete took one character and a
   * paste landed at the old caret, with the selected words untouched. Such an
   * end is the start of the first paragraph at or after that index, or the
   * end of the last paragraph before it.
   */
  const pointOf = (node, offset) => {
    const el = blockOf(node);
    if (el) return { block: Number(el.dataset.block), offset: offsetIn(el, node, offset) };
    if (!node || node.nodeType !== Node.ELEMENT_NODE || !pageRef.current) return null;
    const blocks = [...pageRef.current.querySelectorAll('[data-block]')];
    if (!blocks.length) return null;
    const marker = node.childNodes[offset] || null;
    if (marker) {
      const after = blocks.find((b) => marker === b || marker.contains(b) || Boolean(marker.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING));
      if (after) return { block: Number(after.dataset.block), offset: Number(after.dataset.from || 0) };
    }
    const before = [...blocks].reverse().find((b) => node.contains(b) ? Boolean(node.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_CONTAINED_BY) && (!marker || Boolean(marker.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_PRECEDING)) : Boolean(node.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_PRECEDING));
    if (before) return { block: Number(before.dataset.block), offset: offsetIn(before, before, before.childNodes.length) };
    return { block: Number(blocks[0].dataset.block), offset: 0 };
  };

  const currentPosition = useCallback(() => {
    const sel = window.getSelection();
    if (!sel || !sel.focusNode) return null;
    const focus = pointOf(sel.focusNode, sel.focusOffset);
    if (!focus) return null;
    return {
      anchor: sel.anchorNode ? pointOf(sel.anchorNode, sel.anchorOffset) : null,
      focus,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /**
   * The caret moved. This is deliberately *not* sent to the backend on every
   * key release: an edit carries its own position, so the only reason to tell
   * the engine separately is a click or an arrow key, and doing it per keystroke
   * doubled the round trips for no gain.
   */
  const syncSelection = useCallback(() => {
    setPicked(null);
    // While a record is previewed the words on the page are the record's,
    // not the file's, and a caret measured in them would land in the wrong
    // place: the engine keeps the caret it had.
    if (previewRef.current) return;
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
      ops.push({ op: 'setRunFormat', delta: {
        fontName: brush.fontName ?? null, fontSize: brush.fontSize ?? null, fontColour: brush.fontColour ?? null, highlight: brush.highlight ?? null,
        outline: Boolean(brush.outline), shadow: Boolean(brush.shadow), glow: brush.glow ?? null,
      } });
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
      if (previewRef.current) {
        toast('Preview Results is on — these are a recipient\'s words. Turn it off to edit the letter.', { ms: 4500 });
        return;
      }
      const pos = currentPosition();
      const ops = [];
      const selection = selectionToSend(pos, { sent: sentCaret.current, placed: placedCaret.current });
      if (selection) ops.push(selection);

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
      if (ops.some((op) => op.op !== 'setSelection')) {
        sentCaret.current = pos;
        apply(...ops);
      }
    },
    [apply, currentPosition, toast]
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
  useEffect(() => {
    const el = pageRef.current;
    if (!el) return undefined;
    const pick = (e) => setPicked(e.detail);
    el.addEventListener('wd-pick', pick);
    return () => el.removeEventListener('wd-pick', pick);
  }, [busy, model === null]);

  // The engine's caret is authoritative; after every render the DOM caret is
  // put back where the engine says it is.
  useLayoutEffect(() => {
    // The answer to whatever was sent has been painted by the time this runs,
    // so nothing is in flight any more — whether or not the answer moved the
    // caret. Clearing it only when it did left a stale position standing
    // after a formatting edit, and a keystroke that happened to land on
    // exactly that spot then sent no selection at all.
    sentCaret.current = null;
    const target = pendingCaret.current;
    if (!target || !pageRef.current) return;
    placeSelection(pageRef.current, target.anchor, target.focus);
    placedCaret.current = target;
    pendingCaret.current = null;
  }, [model]);

  // An equation drawn in a shadow root takes no part in the browser's own
  // selection highlight, so one inside the engine's selection is marked by
  // hand — the blue Word puts over a selected equation.
  useLayoutEffect(() => {
    const page = pageRef.current;
    if (!page) return;
    const s = model?.selection;
    const before = (a, b) => a.block < b.block || (a.block === b.block && a.offset <= b.offset);
    for (const host of page.querySelectorAll('.wd-math')) {
      const block = Number(host.closest('[data-block]')?.dataset.block);
      const at = Number(host.dataset.at);
      const on = Boolean(s && !s.collapsed && Number.isFinite(block) && Number.isFinite(at)
        && before(s.from, { block, offset: at }) && before({ block, offset: at + 1 }, s.to));
      host.classList.toggle('sel', on);
    }
  }, [model]);

  /**
   * Copy (or cut) a selection that holds an equation: plain text for other
   * programs, with each equation in its linear form, and the equations
   * themselves on a clipboard type of the suite's own for a paste here. A
   * selection with no equation is left to the browser.
   */
  const copyEquations = (e, cut) => {
    const pos = currentPosition();
    if (!pos?.focus || !model?.blocks) return;
    const a = pos.anchor || pos.focus;
    const b = pos.focus;
    const [from, to] = a.block < b.block || (a.block === b.block && a.offset <= b.offset) ? [a, b] : [b, a];
    if (from.block === to.block && from.offset === to.offset) return;
    const clip = clipOf(model.blocks, from, to);
    if (!clip.hasMath) return;
    e.preventDefault();
    e.clipboardData.setData('text/plain', clip.text);
    e.clipboardData.setData(CLIP_TYPE, JSON.stringify({ lines: clip.lines }));
    if (cut) apply({ op: 'setSelection', anchor: from, focus: to }, { op: 'deleteSelection' });
  };

  /* ── commands ────────────────────────────────────────────────────────── */

  const format = model?.format || {};
  const section = model?.section;

  /* ── pages ─────────────────────────────────────────────────────────── */
  //
  // Print layout draws the flow on sheets. Which sheet each block lands on
  // is decided by measuring the drawn page after every commit — see
  // word/pages.js — and the split points that decision produces are the
  // only part of it React needs to know: a split paragraph renders as parts.
  const [pages, setPages] = useState(NO_PAGES);
  const pagesRef = useRef(pages);
  pagesRef.current = pages;
  /** Line numbers, measured from the drawn lines after every layout: [{ top, height, n }], page-relative. */
  const [lineNos, setLineNos] = useState(null);
  const lineNosKey = useRef('');
  const paged = (view.mode || 'print') === 'print' && Boolean(section);
  // An envelope in front of the letter (Mailings → Envelopes) is the first
  // sheet, at the envelope's own size — see pages.js `geometryOf`.
  const envelopeKey = model?.envelope ? `${model.envelope.widthPx}:${model.envelope.heightPx}:${model.envelope.endsAt}:${JSON.stringify(model.envelope.margins)}` : '';
  const geo = useMemo(() => geometryOf(section, model?.envelope || null), [section, envelopeKey]);
  // The tallest a picture may be drawn on a page: the sheet's inside, less a
  // little for the paragraph's own spacing.
  const inner = paged && geo ? geo.H - geo.top - geo.bottom - 12 : null;
  // Layout → Columns: the editable page is ONE contenteditable flow — the
  // caret, the IME, the spell checker all depend on that — and cannot show
  // true side-by-side columns without a different editor entirely. So on
  // screen it does what Word's own Draft view does: lay the single flow at
  // the width of the FIRST column, narrower than the page, and be honest
  // about the difference in the status bar. Print and PDF flow the words
  // into the real columns — see paginate.js — because those are not typed
  // into live.
  const columnBoxes = useMemo(() => columnBoxesOf(section), [section]);
  const multiColumn = Boolean(columnBoxes && columnBoxes.length > 1);
  const passes = useRef(0);
  const repaginate = useCallback(() => {
    const page = pageRef.current;
    if (!page) return;
    const state = pagesRef.current;
    if (!paged || !geo) {
      clearPages(page);
      if (state !== NO_PAGES) setPages(NO_PAGES);
      return;
    }
    // A layout that will not settle — a paragraph that fits only when it is
    // not split, say — stops here rather than looping. Sixteen is far more
    // than any document has needed; the count resets with every edit.
    if (passes.current > 16) return;
    const laid = layPages(page, geo, state);
    const focus = model?.selection?.focus;
    const at = focus ? pageOfElement(partFor(page, focus.block, focus.offset), geo) + 1 : 1;
    if (laid.changed) {
      passes.current += 1;
      // Synchronously, so the parts are drawn before the browser paints the
      // frame; otherwise the unsplit paragraph shows over the gap for a frame.
      flushSync(() => setPages({ splits: laid.splits, tableSplits: laid.tableSplits, notes: laid.notes, frames: laid.frames, count: laid.count, at }));
    } else if (laid.count !== state.count || at !== state.at) {
      setPages({ ...state, count: laid.count, at });
    }
  }, [paged, geo, model]);
  const repaginateRef = useRef(repaginate);
  repaginateRef.current = repaginate;
  useLayoutEffect(() => {
    passes.current = 0;
  }, [model, mounted, paged, geo]);
  // Every commit of the page lays it out again — after the tab widths, which
  // are a microtask queued during the commit, and before the paint.
  useLayoutEffect(() => {
    let live = true;
    queueMicrotask(() => {
      if (live) repaginateRef.current();
    });
    return () => {
      live = false;
    };
  }, [model, mounted, paged, geo, pages.splits, pages.tableSplits, pages.notes, pages.frames]);
  // The numbers down the margin, from the line boxes the browser drew — after
  // the layout pass, in the same commit. Only the body's paragraphs count,
  // as in Word: not a table's cells, not the notes.
  useLayoutEffect(() => {
    const page = pageRef.current;
    const spec = section?.lineNumbers;
    if (!page || !spec) {
      if (lineNosKey.current !== '') { lineNosKey.current = ''; setLineNos(null); }
      return;
    }
    const pageTop = rectOf(page).top;
    const stride = paged && geo ? geo.H + geo.G : Infinity;
    const pageAt = (top) => (paged && geo ? pageIndexAt(geo, top) : 0);
    const items = [];
    let n = spec.start || 1;
    let lastPage = 0;
    for (const el of page.querySelectorAll(':scope > .wd-block')) {
      for (const box of lineBoxes(el)) {
        const top = box.top - pageTop;
        const k = Number.isFinite(stride) ? pageAt(top) : 0;
        if (spec.restart === 'newPage' && k !== lastPage) { n = spec.start || 1; lastPage = k; }
        const mine = n++;
        if (mine % spec.countBy === 0) items.push({ top: Math.round(top), height: Math.max(8, Math.round(box.bottom - box.top)), n: mine });
      }
    }
    const key = JSON.stringify(items);
    if (key !== lineNosKey.current) { lineNosKey.current = key; setLineNos(items); }
  }, [model, mounted, paged, geo, pages, section]);
  // Pictures that arrive and fonts that load change heights without a render.
  const hasPage = Boolean(!busy && model);
  useEffect(() => {
    const page = pageRef.current;
    if (!page || typeof ResizeObserver === 'undefined') return undefined;
    let raf = 0;
    const ro = new ResizeObserver(() => {
      cancelAnimationFrame(raf);
      raf = requestAnimationFrame(() => repaginateRef.current());
    });
    ro.observe(page);
    // A picture's bytes arrive after the pass that placed its paragraph, and
    // the picture grows where the pass left it — over the page's edge when
    // the paragraph sat near the foot. Its load is a fresh start for the
    // passes, whatever the observer saw meanwhile.
    const loaded = (e) => {
      if (!e.target?.classList?.contains('wd-image')) return;
      passes.current = 0;
      cancelAnimationFrame(raf);
      raf = requestAnimationFrame(() => repaginateRef.current());
    };
    page.addEventListener('load', loaded, true);
    return () => {
      ro.disconnect();
      page.removeEventListener('load', loaded, true);
      cancelAnimationFrame(raf);
    };
  }, [hasPage]);

  const flowItems = useMemo(() => (model?.blocks ? groupTables(model.blocks) : []), [model?.blocks]);
  // A new document starts with the first screens; the rest mounts in slices
  // once the browser has painted, and a document already fully mounted stays
  // so through every edit.
  useEffect(() => {
    setMounted(MOUNT_FIRST);
  }, [doc?.id]);
  useEffect(() => {
    if (mounted >= flowItems.length) return undefined;
    const schedule = window.requestIdleCallback || ((fn) => setTimeout(fn, 16));
    const cancel = window.cancelIdleCallback || clearTimeout;
    const handle = schedule(() => setMounted((m) => Math.min(flowItems.length, m + MOUNT_STEP)));
    return () => cancel(handle);
  }, [mounted, flowItems.length]);


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

  // The ruler and the grips on a table: where the caret is, which table
  // that is in, and what a dragged border asks of the engine.
  const at = model?.selection?.focus?.block ?? 0;
  const tableAt = useMemo(() => {
    const key = /^(t\d+):r\d+:c\d+$/.exec(String(model?.blocks?.[at]?.container || ''));
    if (!key) return null;
    const group = flowItems.find((item) => item.table?.id === key[1]);
    return group?.table?.gridPx ? { id: key[1], gridPx: group.table.gridPx } : null;
  }, [model, at, flowItems]);
  const resizeColumn = useCallback((id, k, dx) => {
    const grid = flowItems.find((item) => item.table?.id === id)?.table?.gridPx;
    if (!grid || !(k >= 1 && k <= grid.length)) return;
    // The column at the left of the border grows by the drag and the one at
    // its right shrinks by it, so the table keeps its width, as in Word; the
    // last border moves only the last column, and the table with it.
    const MIN = 20;
    const left = k - 1;
    const right = k < grid.length ? k : null;
    const d = Math.max(-(grid[left] - MIN), right != null ? Math.min(grid[right] - MIN, dx) : dx);
    const widths = { [left]: Math.round((grid[left] + d) * 15) };
    if (right != null) widths[right] = Math.round((grid[right] - d) * 15);
    apply({ op: 'setTableColumnWidths', table: Number(id.slice(1)), widths });
  }, [flowItems, apply]);
  const resizeRow = useCallback((id, row, heightPx) => {
    apply({ op: 'setTableRowHeight', table: Number(id.slice(1)), row, twips: Math.round(heightPx * 15) });
  }, [apply]);

  /**
   * The on-screen page (1-based) of each heading, in the order given — what
   * `insertTableOfContents`/`updateTableOfContents` send as `pages`, since
   * the engine cannot paginate the screen itself (see `pages.js`). A heading
   * not currently drawn (view mode without pages, or the block not yet
   * painted) gets `undefined`, which the engine reads as "leave it blank".
   */
  const headingPages = useCallback(
    (headings) => headings.map((h) => {
      const el = pageRef.current?.querySelector(`[data-block="${h.index}"]`);
      return el ? pageOfElement(el, geo) + 1 : undefined;
    }),
    [geo]
  );

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
          // The page alone is scaled, with CSS zoom on the page element. It
          // used to be the window's zoom — the ribbon, the status bar and
          // the slider under the pointer all grew with the page, which made
          // the slider unusable (owner, 2026-09-24). The paginator, the ruler
          // and the picture handles read their rects through pages.js, which
          // divides them by the level, so the layout is the same at any zoom.
          const pageWidth = model?.section?.widthPx || 794;
          const scroll = document.querySelector('.wd-scroll');
          const roomW = (scroll?.clientWidth || window.innerWidth) - 48;
          const roomH = (scroll?.clientHeight || window.innerHeight) - 60;
          const target = arg === 'width' ? roomW / pageWidth : arg === 'page' ? roomH / (model?.section?.heightPx || 1123) : arg === 'pages' ? 0.5 : Number(arg) || 1;
          const level = Math.max(0.3, Math.min(3, target));
          patchView({ zoom: Math.round(level * 100) / 100 });
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
          patchView({ painting: {
            bold: f.bold, italic: f.italic, underline: f.underline, strike: f.strike, fontName: f.fontName, fontSize: f.fontSize, fontColour: f.fontColour, highlight: f.highlight,
            outline: f.outline, shadow: f.shadow, glow: f.glow,
          } });
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
        case 'wrap': {
          // How the text treats the picked picture: Word's Wrap Text menu.
          if (!picked) return toast('Click a picture first, then choose how the text wraps round it.', { ms: 4500 });
          const img = blocks[picked.block]?.images?.[picked.image];
          await apply({ op: 'setImageLayout', block: picked.block, image: picked.image, wrap: arg, hAlign: img?.hAlign || 'left' });
          return;
        }
        case 'position': {
          // Where the picked picture sits: at the left or right with the words
          // round it, or centred with the words above and below.
          if (!picked) return toast('Click a picture first, then choose where it sits.', { ms: 4500 });
          const img = blocks[picked.block]?.images?.[picked.image];
          const floating = img?.anchored && img.wrap !== 'none' ? img.wrap : null;
          const wrap = arg === 'center' ? 'topAndBottom' : floating && floating !== 'topAndBottom' ? floating : 'square';
          await apply({ op: 'setImageLayout', block: picked.block, image: picked.image, wrap, hAlign: arg });
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
          const headings = blocks.filter((b) => /^Heading[1-3]$/.test(b.style || '') && (b.text || '').trim());
          if (!headings.length) return toast('No headings yet. Use Heading 1, 2 and 3 on the paragraphs you want listed.', { ms: 6000 });
          const pages = headingPages(headings);
          await apply({ op: 'insertTableOfContents', pages });
          toast(`Table of contents inserted — ${headings.length} heading${headings.length === 1 ? '' : 's'}. Update Table after the headings change.`, { ms: 6000 });
          return;
        }
        case 'updateTableOfContents': {
          if (!model?.tableOfContents) return toast('No table of contents in this document yet.', { ms: 4000 });
          const headings = blocks.filter((b) => /^Heading[1-3]$/.test(b.style || '') && (b.text || '').trim());
          const pages = headingPages(headings);
          await apply({ op: 'updateTableOfContents', pages });
          toast('Table of contents updated.', { ms: 4000 });
          return;
        }
        case 'removeTableOfContents': {
          if (!model?.tableOfContents) return;
          await apply({ op: 'removeTableOfContents' });
          return;
        }
        case 'equation': {
          // Insert → Equation. A paragraph with nothing in it takes a display
          // equation, on a line of its own; one with words takes it inline,
          // in the words — Word's own rule. A pick from the built-in gallery
          // goes straight in; otherwise the editor opens.
          const here = blocks[sel?.focus?.block ?? 0];
          const display = !here || (here.text || '').replace(/\s/g, '') === '';
          if (arg?.linear) {
            await apply({ op: 'insertEquation', linear: arg.linear, display });
            return;
          }
          setEquation({ initial: '', display, editing: false });
          return;
        }
        case 'editEquation': {
          // An equation opened again — a double-click, or Enter on a selected
          // one: the editor holds its linear form, and OK replaces it.
          const b = blocks[arg?.block];
          let o = 0;
          const run = (b?.runs || []).find((r) => { const hit = o === arg.at && r.math; o += (r.text || '').length; return hit; });
          if (!run) return;
          setEquation({ initial: run.math.linear || '', display: Boolean(run.math.display), editing: true, block: arg.block, offset: arg.at });
          return;
        }
        case 'insertNote':
          setNoteDialog({ kind: arg === 'endnote' ? 'endnote' : 'footnote' });
          return;
        case 'editNote':
          setNoteDialog(arg);
          return;
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
          const notes = pageRef.current?.querySelector('.wd-pagenotes .wd-notes, .wd-notes:not(.wd-notes-measure)');
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
        case 'toggleTrackChanges': {
          const next = await apply({ op: 'toggleTrackChanges', on: !model?.trackRevisions });
          if (next && !model?.trackRevisions) toast('Track Changes is on — typing and deleting are recorded.', { ms: 5000 });
          return;
        }
        case 'markupMode':
          // A view preference only — never saved in the file (see view.js).
          patchView({ markupMode: arg });
          return;
        case 'acceptChanges':
          await apply({ op: 'acceptChanges', all: arg === 'all' });
          return;
        case 'rejectChanges':
          await apply({ op: 'rejectChanges', all: arg === 'all' });
          return;
        case 'nextChange': {
          const list = blocks.filter((b) => b.tracked).map((b) => b.index);
          if (!list.length) return toast('No tracked changes in this document.', { ms: 4000 });
          const next = arg > 0 ? list.find((n) => n > at) ?? list[0] : [...list].reverse().find((n) => n < at) ?? list[list.length - 1];
          await apply({ op: 'setSelection', anchor: { block: next, offset: 0 }, focus: { block: next, offset: 0 } });
          pageRef.current?.querySelector(`[data-block="${next}"]`)?.scrollIntoView({ block: 'center', behavior: 'smooth' });
          return;
        }
        default:
          return;
      }
    },
    [model, view, apply, shell, toast, doc, patchView, picked, headingPages]
  );
  actRef.current = act;

  // Mailings: envelopes and labels (word/envelopes.js), and the merge's
  // verbs and dialogs (word/mailings.js), which the first join.
  const readSelection = useCallback(() => window.getSelection()?.toString() || '', []);
  const envelopes = useEnvelopesLabels({ shell, doc, model, apply, toast, selectionText: readSelection });
  const mailings = useMailings({ shell, doc, model, apply, toast, extra: envelopes.extra });

  const commands = useMemo(
    () => ({
      'file.new': { label: 'New', icon: 'new', key: 'Mod+N', run: () => shell.win.create({ app: 'word' }) },
      'file.open': { label: 'Open…', icon: 'open', key: 'Mod+O', run: openFile },
      'file.save': { label: 'Save', icon: 'save', key: 'Mod+S', global: true, run: () => save(false) },
      'file.print': { label: 'Print…', icon: 'print', key: 'Mod+P', global: true, run: () => setDialog('print') },
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
      // References → Update Fields, or F9: every REF's words refreshed from
      // its bookmark. The count comes back as `opResult` (documents.js's
      // `apply`), which is the only way this toast can say how many.
      'field.update': {
        label: 'Update Fields', icon: 'refresh', key: 'F9', global: true,
        run: async () => {
          // A table of contents is a field too (see `Document#refreshRefFields`);
          // updateFields runs first so its numeric count — the only thing
          // `opResult` carries back — is not shadowed by the second op's `this`.
          const ops = [{ op: 'updateFields' }];
          if (model?.tableOfContents) ops.push({ op: 'updateTableOfContents' });
          const next = await apply(...ops);
          if (!next) return;
          const n = next.opResult ?? 0;
          toast(`${n} field${n === 1 ? '' : 's'} updated`, { tone: 'good' });
        },
      },
    }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [doc, apply, save, openFile, exportAs, shell, toast, model]
  );

  useCommands(commands, [doc, model]);

  const align = (value) => apply({ op: 'setParagraphFormat', delta: { align: value } });
  const setStyle = (style) => apply({ op: 'setParagraphFormat', delta: { style: style || null } });

  if (error) {
    return (
      <AppFrame app={app} shell={shell} title="Rutba Word" menu={appMenu}>
        <Empty icon="word" title="This file could not be opened">{error}</Empty>
      </AppFrame>
    );
  }

  return (
    <AppFrame
      app={app}
      shell={shell}
      title={doc?.name || 'Rutba Word'}
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
          picked={picked}
          mailings={mailings}
        />
      }
      status={
        <>
          <span>{doc?.path || 'Not saved yet'}</span>
          <Spacer />
          {paged ? <Chip>{`Page ${pages.at} of ${pages.count}`}</Chip> : null}
          {multiColumn ? (
            <span className="wd-columns-chip">
              <Chip title="One flow on screen for a native caret; the print and PDF layout actually splits it into columns">{`${columnBoxes.length} columns — laid as one on screen, flowed into columns in print`}</Chip>
            </span>
          ) : null}
          {model?.mailMerge?.type ? (
            <Chip title="Mailings — the kind of mail merge document and its recipients">{mergeChip(model.mailMerge)}</Chip>
          ) : null}
          <Chip>{model?.wordCount ?? 0} words</Chip>
          <Chip>{model?.characterCount ?? 0} characters</Chip>
          <Chip>{model?.blocks?.length ?? 0} paragraphs</Chip>
          <ZoomSlider value={view.zoom ?? 1} onChange={(v) => act('zoom', v)} onReset={() => act('zoom', 1)} />
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
          <style>{EQUATION_CSS}</style>
          {view.navigation ? (
            <NavigationPane blocks={model.blocks} at={model.selection?.focus?.block ?? -1} onGo={(i) => act('goto', i)} onClose={() => act('toggleNavigation')} />
          ) : null}
          <div className="wd-scroll">
            {view.ruler ? (
              <Ruler
                section={section}
                zoom={view.zoom ?? 1}
                page={pageRef}
                model={model}
                at={at}
                tableId={tableAt?.id ?? null}
                gridPx={tableAt?.gridPx ?? null}
                onParagraph={(delta) => apply({ op: 'setParagraphFormat', delta })}
                onMargin={(side, px) => apply({ op: 'setPageSetup', spec: { margins: { [side]: Math.round(px * 15) } } })}
                onColumn={resizeColumn}
              />
            ) : null}
            <div
              className={`wd-page${view.marks ? ' marks' : ''}${paged ? ' paged' : ''}${mailings.highlight ? ' wd-mm-hl' : ''}${model.mailMerge?.preview ? ' wd-mm-preview' : ''}`}
              ref={pageRef}
              contentEditable
              suppressContentEditableWarning
              spellCheck={view.spell !== false}
              onMouseDown={(e) => {
                // A press on an equation selects it whole — the one character
                // it is — so Delete, Backspace, typing and copy all act on
                // the equation, and a second press opens it in the editor.
                const host = e.button === 0 ? e.target.closest?.('.wd-math') : null;
                if (!host) return;
                const blockEl = host.closest('[data-block]');
                const at = Number(host.dataset.at);
                if (!blockEl || !Number.isFinite(at)) return;
                e.preventDefault();
                const block = Number(blockEl.dataset.block);
                pageRef.current?.focus({ preventScroll: true });
                const range = document.createRange();
                range.setStartBefore(host);
                range.setEndAfter(host);
                const sel = window.getSelection();
                sel.removeAllRanges();
                sel.addRange(range);
                apply({ op: 'setSelection', anchor: { block, offset: at }, focus: { block, offset: at + 1 } });
                if (e.detail >= 2) act('editEquation', { block, at });
              }}
              onMouseUp={(e) => {
                // A press on a picture or its handles is a pick, not a caret move;
                // one on an equation selected it already.
                if (pictureDrag.current || e.target.closest?.('.wd-handles, .wd-image, .wd-math')) return;
                // Ctrl+click (Cmd+click on a Mac) a REF field to go to the
                // bookmark it names — Word's own way into a cross-reference.
                const field = e.target.closest?.('.wd-field');
                if (field && (e.ctrlKey || e.metaKey) && field.dataset.name) {
                  apply({ op: 'gotoBookmark', name: field.dataset.name });
                  return;
                }
                // Ctrl+click a table of contents entry the same way — its
                // hyperlink points at the heading's own `_Toc` bookmark.
                const link = e.target.closest?.('.wd-link');
                if (link && (e.ctrlKey || e.metaKey) && link.dataset.link?.startsWith('#')) {
                  apply({ op: 'gotoBookmark', name: link.dataset.link.slice(1) });
                  return;
                }
                syncSelection();
              }}
              onKeyDown={(e) => {
                // A picked picture goes with Delete or Backspace — its block
                // with it when the block held nothing else.
                if ((e.key === 'Delete' || e.key === 'Backspace') && picked) {
                  e.preventDefault();
                  const target = picked;
                  setPicked(null);
                  apply({ op: 'removeImage', block: target.block, image: target.image });
                  return;
                }
                // Arrow keys and Home/End move the caret without an edit, so the
                // engine is told where it landed — after the browser has moved
                // it, which is why this waits a tick.
                if (/^(Arrow|Home|End|Page)/.test(e.key)) setTimeout(syncSelection, 0);
                // Alt+= — Word's own shortcut for a new equation.
                if (e.altKey && !e.ctrlKey && (e.key === '=' || e.code === 'Equal')) {
                  e.preventDefault();
                  act('equation');
                  return;
                }
                // Enter on a selected equation opens it, as Word does.
                const s = model?.selection;
                if (e.key === 'Enter' && s && !s.collapsed && s.from.block === s.to.block && s.to.offset - s.from.offset === 1) {
                  let o = 0;
                  const hit = (model.blocks[s.from.block]?.runs || []).some((r) => { const at = o; o += (r.text || '').length; return at === s.from.offset && r.math; });
                  if (hit) {
                    e.preventDefault();
                    act('editEquation', { block: s.from.block, at: s.from.offset });
                  }
                }
              }}
              onCopy={(e) => copyEquations(e, false)}
              onCut={(e) => copyEquations(e, true)}
              onPaste={(e) => {
                // A copy made in the suite that held equations: they come back
                // as themselves, not as their linear form.
                const data = e.clipboardData?.getData(CLIP_TYPE);
                if (!data) return;
                let lines = null;
                try { lines = JSON.parse(data).lines; } catch { lines = null; }
                if (!Array.isArray(lines)) return;
                e.preventDefault();
                const pos = currentPosition();
                const ops = [];
                if (pos?.focus) ops.push({ op: 'setSelection', anchor: pos.anchor || pos.focus, focus: pos.focus });
                ops.push({ op: 'pasteRuns', lines });
                apply(...ops);
              }}
              onContextMenu={(e) => menu.open(e, menuItems(commands, ['edit.undo', 'edit.redo', '-', 'format.bold', 'format.italic', 'format.underline', '-', 'edit.find']))}
              style={{
                // View → Zoom: the page scaled on its own, the window's chrome left alone.
                zoom: view.zoom && Math.abs(view.zoom - 1) > 0.001 ? view.zoom : undefined,
                // The document's own base font, from its stylesheet's defaults.
                fontFamily: model.resolvedStyles?.['*default*']?.fontName || undefined,
                fontSize: model.resolvedStyles?.['*default*']?.sizePx ? `${model.resolvedStyles['*default*'].sizePx}px` : undefined,
                width: section ? Math.round(section.widthPx) : 794,
                minHeight: paged ? pageTopOf(geo, pages.count - 1) + pageHeightOf(geo, pages.count - 1) : section ? Math.round(section.heightPx) : 1123,
                // An envelope in front starts the page at the envelope's own top margin.
                paddingTop: paged && geo?.first ? geo.first.top : section?.margins.top ?? 96,
                // In print layout the single flow is laid as wide as the
                // FIRST column only — see the note above `columnBoxes` — by
                // padding out the rest of the page on the right; the real
                // right margin is folded into that padding.
                paddingRight: paged && multiColumn
                  ? Math.round(section.widthPx - section.margins.left - columnBoxes[0].widthPx)
                  : section?.margins.right ?? 96,
                paddingBottom: section?.margins.bottom ?? 96,
                paddingLeft: section?.margins.left ?? 96,
                // The page colour, on the page itself when the flow is one sheet;
                // in print layout the sheets behind the flow carry it.
                background: !paged && section?.background ? section.background : undefined,
                // Outside print layout the flow is one box with no page
                // edges to break at, so the browser is asked to flow it into
                // columns itself, the way any web page's `column-count` does.
                ...(!paged && multiColumn ? {
                  columnCount: columnBoxes.length,
                  columnGap: (section.columns?.spacePx ?? 36),
                  columnRule: section.columns?.separator ? '0.5pt solid #808080' : 'none',
                } : {}),
                // The bands read the margins too, to line up with the body.
                '--wd-margin-left': `${section?.margins.left ?? 96}px`,
                '--wd-margin-right': `${section?.margins.right ?? 96}px`,
              }}
            >
              {/*
                The sheets, behind the flow — one per page the layout pass
                counted — and on each the watermark, the header and the footer
                that page calls for, its page number resolved.
              */}
              {paged
                ? Array.from({ length: pages.count }, (_, k) => (
                    <div key={`s${k}`} className={`wd-sheet${geo.first && k === 0 ? ' wd-envelope-sheet' : ''}`} contentEditable={false} aria-hidden="true" style={{ top: pageTopOf(geo, k), height: pageHeightOf(geo, k), ...(geo.first && k === 0 ? { left: geo.first.left, width: geo.first.W, right: 'auto' } : {}), background: section?.background || undefined }} />
                  ))
                : null}
              {/* Around a flow the watermark rides the one page; print layout puts it on every sheet below. */}
              {!paged && model.bands?.watermark ? (
                <div className="wd-watermark" contentEditable={false} aria-hidden="true" style={{ top: Math.round((section?.heightPx ?? 1123) / 3), color: model.bands.watermark.colour || 'silver', transform: `rotate(${model.bands.watermark.rotation ?? 315}deg)` }}>
                  {model.bands.watermark.text}
                </div>
              ) : null}
              {/* Line numbers, down the left margin, beside the lines the browser drew. */}
              {lineNos?.length ? (
                <div className="wd-linenos" contentEditable={false} aria-hidden="true" style={{ left: Math.max(0, (section?.margins.left ?? 96) - (section?.lineNumbers?.distancePx ?? 24) - 28) }}>
                  {lineNos.map((l) => <span key={`${l.n}-${l.top}`} style={{ top: l.top, height: l.height, lineHeight: `${l.height}px` }}>{l.n}</span>)}
                </div>
              ) : null}
              {/* The page borders: a frame on each sheet, or one around the flow. */}
              {section?.pageBorders
                ? (paged ? Array.from({ length: pages.count }, (_, k) => k) : [null]).map((k) => (
                    <div
                      key={`pb${k ?? 'flow'}`}
                      className="wd-pgborders"
                      contentEditable={false}
                      aria-hidden="true"
                      style={geo?.first && k === 0 ? { display: 'none' } : pageBordersStyle(section, k === null ? null : { top: pageTopOf(geo, k), height: pageHeightOf(geo, k) })}
                    />
                  ))
                : null}
              {paged
                ? Array.from({ length: pages.count }, (_, k) => (
                    <React.Fragment key={`b${k}`}>
                      {model.bands?.watermark && !(geo.first && k === 0) ? (
                        <div className="wd-watermark" contentEditable={false} aria-hidden="true" style={{ top: pageTopOf(geo, k) + Math.round(geo.H / 3), color: model.bands.watermark.colour || 'silver', transform: `rotate(${model.bands.watermark.rotation ?? 315}deg)` }}>
                          {model.bands.watermark.text}
                        </div>
                      ) : null}
                      {/* An envelope has no header or footer, and is not counted: the letter starts at page 1. */}
                      {geo.first && k === 0 ? null : <Band kind="header" bands={model.bands} section={section} page={k + 1 - (geo.first ? 1 : 0)} of={pages.count - (geo.first ? 1 : 0)} top={pageTopOf(geo, k)} height={pageHeightOf(geo, k)} onEdit={() => setDialog('header')} />}
                      {geo.first && k === 0 ? null : <Band kind="footer" bands={model.bands} section={section} page={k + 1 - (geo.first ? 1 : 0)} of={pages.count - (geo.first ? 1 : 0)} top={pageTopOf(geo, k)} height={pageHeightOf(geo, k)} onEdit={() => setDialog('footer')} />}
                      <PageNotes notes={model.footnotes} at={pages.notes} page={k} top={pageTopOf(geo, k) + geo.top} height={geo.H - geo.top - geo.bottom} styles={model.resolvedStyles} onEdit={(note) => act('editNote', { kind: 'footnote', id: note.id, initial: noteWords(note) })} />
                    </React.Fragment>
                  ))
                : null}
              {flowItems.slice(0, mounted).map((item) =>
                item.table ? (
                  <TableGroup key={`t${item.table.id}`} table={item.table} labels={model.listLabels} styles={model.resolvedStyles} tsplit={pages.tableSplits[item.table.id] || null} />
                ) : (
                  <Block key={item.index} block={item} labels={model.listLabels} styles={model.resolvedStyles} split={item.frame ? null : pages.splits[item.index] || null} pickedImage={picked?.block === item.index ? picked.image : null} inner={inner} markupMode={view.markupMode || 'simple'} place={placeOf(item, paged ? geo : null, pages.frames)} />
                )
              )}
              {mounted < flowItems.length ? <div className="wd-mounting" aria-hidden="true">{`Laying out… ${Math.round((mounted / flowItems.length) * 100)}%`}</div> : null}

              {picked ? <PictureHandles page={pageRef} picked={picked} model={model} pages={pages} onDrag={(on) => { pictureDrag.current = on; }} onResize={(size) => apply({ op: 'setImageSize', block: picked.block, image: picked.image, ...size })} /> : null}
              {tableAt ? <TableGrips page={pageRef} zoom={view.zoom ?? 1} model={model} pages={pages} tableId={tableAt.id} gridPx={tableAt.gridPx} onColumn={resizeColumn} onRow={resizeRow} onDrag={(on) => { pictureDrag.current = on; }} /> : null}
              {/*
                In print layout the footnotes are drawn on their pages (above);
                this copy is hidden and measured. In the other layouts it is
                the footnotes, under the body. Endnotes end the document.
              */}
              <Notes notes={model.footnotes} kind="footnotes" styles={model.resolvedStyles} measure={paged} onEdit={(note) => act('editNote', { kind: 'footnote', id: note.id, initial: noteWords(note) })} />
              <Notes notes={model.endnotes} kind="endnotes" styles={model.resolvedStyles} onEdit={(note) => act('editNote', { kind: 'endnote', id: note.id, initial: noteWords(note) })} />
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

      {dialog === 'print' && doc ? (
        <PrintDialog
          shell={shell}
          doc={doc}
          kind="doc"
          onClose={() => setDialog(null)}
          onSaveAs={(options) => exportAs('pdf', options)}
        />
      ) : null}

      {dialog === 'watermark' ? (
        <WatermarkDialog
          current={model?.bands?.watermark?.text || ''}
          onClose={() => setDialog(null)}
          onApply={async (text) => {
            await apply({ op: 'setWatermark', text });
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

      {noteDialog ? (
        <NoteDialog
          kind={noteDialog.kind}
          initial={noteDialog.initial || ''}
          onClose={() => setNoteDialog(null)}
          onSave={async (text) => {
            if (noteDialog.id) await apply({ op: 'setNoteText', kind: noteDialog.kind, id: noteDialog.id, text });
            else await apply({ op: 'insertNote', kind: noteDialog.kind, text });
            setNoteDialog(null);
            toast(noteDialog.id ? 'Note changed' : `${noteDialog.kind === 'endnote' ? 'Endnote' : 'Footnote'} inserted`, { tone: 'good' });
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

      {dialog === 'bookmark' ? (
        <BookmarkDialog
          bookmarks={model?.bookmarks || []}
          onClose={() => setDialog(null)}
          onAdd={(name) => apply({ op: 'addBookmark', name })}
          onDelete={(name) => apply({ op: 'removeBookmark', name })}
          onGoto={async (name) => {
            await apply({ op: 'gotoBookmark', name });
            setDialog(null);
          }}
        />
      ) : null}

      {dialog === 'crossReference' ? (
        <CrossReferenceDialog
          bookmarks={model?.bookmarks || []}
          onClose={() => setDialog(null)}
          onInsert={async (name) => {
            await apply({ op: 'insertCrossReference', name });
            setDialog(null);
          }}
        />
      ) : null}

      {dialog === 'caption' ? (
        <CaptionDialog
          fields={model?.fields || []}
          onClose={() => setDialog(null)}
          onInsert={async (label, text) => {
            await apply({ op: 'insertCaption', label, text });
            setDialog(null);
            toast('Caption inserted', { tone: 'good' });
          }}
        />
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

      {equation ? (
        <EquationDialog
          initial={equation.initial}
          display={equation.display}
          editing={equation.editing}
          onClose={() => { setEquation(null); requestAnimationFrame(() => pageRef.current?.focus({ preventScroll: true })); }}
          onInsert={async (linear, display) => {
            const target = equation;
            setEquation(null);
            const done = target.editing
              ? await apply({ op: 'replaceEquation', block: target.block, offset: target.offset, linear, display })
              : await apply({ op: 'insertEquation', linear, display });
            requestAnimationFrame(() => pageRef.current?.focus({ preventScroll: true }));
            if (done) toast(target.editing ? 'Equation updated' : 'Equation inserted', { tone: 'good' });
          }}
        />
      ) : null}

      {dialog === 'symbol' ? <SymbolDialog onClose={() => setDialog(null)} onInsert={(c) => apply({ op: 'insertText', text: c })} /> : null}

      {dialog === 'properties' ? <PropertiesDialog doc={doc} model={model} onClose={() => setDialog(null)} /> : null}

      {dialog === 'shortcuts' ? <ShortcutsDialog onClose={() => setDialog(null)} /> : null}

      {mailings.node}
      {envelopes.node}

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
/** The status bar's word on the mail merge: its kind and its recipients. */
function mergeChip(mm) {
  const kind = { formLetters: 'Letters', email: 'E-mail messages', envelopes: 'Envelopes', mailingLabels: 'Labels', catalog: 'Directory' }[mm.type] || 'Mail merge';
  if (!mm.source) return `${kind} — no recipients yet`;
  const who = `${mm.included} of ${mm.source.count} recipient${mm.source.count === 1 ? '' : 's'}`;
  return mm.preview ? `${kind} — record ${mm.record} of ${mm.included}` : `${kind} — ${who}`;
}

/** No pages laid yet: one sheet, nothing split. */
const NO_PAGES = { splits: {}, tableSplits: {}, notes: {}, frames: {}, count: 1, at: 1 };

/**
 * Where a block goes when the flow does not put it: a paragraph in a frame
 * placed on the page (an envelope's delivery address) at the frame's own
 * place on its sheet; the envelope's other words moved across to its own
 * margin. Null for everything else — the flow places it.
 */
function placeOf(block, geo, frames) {
  if (!geo) return null;
  if (block.frame && !block.container) {
    const k = frames?.[block.index] ?? (geo.first && block.index <= geo.first.endsAt ? 0 : 0);
    const left = (geo.first && k === 0 ? geo.first.left : 0) + block.frame.xPx;
    return { kind: 'frame', left: Math.round(left), top: Math.round(pageTopOf(geo, k) + block.frame.yPx), width: Math.round(block.frame.widthPx || 0) || undefined, height: block.frame.exact && block.frame.heightPx ? Math.round(block.frame.heightPx) : undefined };
  }
  if (geo.first && !block.container && block.index <= geo.first.endsAt && geo.first.dx) return { kind: 'shift', dx: geo.first.dx };
  return null;
}

/** Flow items mounted before the first paint, and per slice afterwards. */
const MOUNT_FIRST = 160;
const MOUNT_STEP = 240;

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
      // The file's grid and width ride every paragraph in the table; a sized
      // row and a merged cell say so on their own paragraphs.
      current = { id: at[1], rows: new Map(), gridPx: block.gridPx || null, width: block.tableWidth || null, rowHeights: new Map(), rowRules: new Map(), spans: new Map(), vAligns: new Map(), look: block.tableLook || null };
    }
    const row = Number(at[2]);
    const cell = Number(at[3]);
    if (!current.rows.has(row)) current.rows.set(row, new Map());
    const cells = current.rows.get(row);
    if (!cells.has(cell)) cells.set(cell, []);
    cells.get(cell).push(block);
    if (block.rowHeightPx && !current.rowHeights.has(row)) current.rowHeights.set(row, block.rowHeightPx);
    if (block.rowRule && !current.rowRules.has(row)) current.rowRules.set(row, block.rowRule);
    if (block.cellVAlign) current.vAligns.set(`${row}:${cell}`, block.cellVAlign);
    if (block.cellSpan > 1) current.spans.set(`${row}:${cell}`, block.cellSpan);
  }
  flush();
  return out;
}

function TableGroup({ table, labels, styles, tsplit }) {
  const rows = [...table.rows.entries()].sort((a, b) => a[0] - b[0]);
  // One table, or — split across pages at the rows the layout pass chose —
  // one per page, each knowing which row it starts at.
  const cuts = (tsplit || []).filter((r) => r > 0 && r < rows.length);
  const bounds = [0, ...cuts, rows.length];
  // The file's grid: each column its share of the table, and the table the
  // width the file gives it — a share of the text width, a fixed one, or
  // the grid's own sum. A file with no grid keeps the text width.
  const grid = table.gridPx;
  const sum = grid ? grid.reduce((a, b) => a + b, 0) : 0;
  const width = grid && sum > 0
    ? table.width?.type === 'pct' ? `${Math.min(100, Math.round(table.width.value / 50))}%`
      : table.width?.type === 'dxa' ? Math.round(table.width.value / 15)
      : Math.round(sum)
    : undefined;
  return (
    <>
      {bounds.slice(0, -1).map((from, j) => (
        <table key={j} className={`wd-table${table.look?.bare ? ' wd-table-bare' : ''}`} data-table={table.id} data-part={cuts.length ? j : undefined} data-row-from={from > 0 ? from : undefined} style={width || table.look?.fixed ? { width, ...(table.look?.fixed ? { tableLayout: 'fixed' } : {}) } : undefined}>
          {grid && sum > 0 ? <colgroup>{grid.map((w, i) => <col key={i} style={{ width: `${(w / sum) * 100}%` }} />)}</colgroup> : null}
          <tbody>
            {rows.slice(from, bounds[j + 1]).map(([r, cells]) => (
              <tr key={r} style={table.rowHeights?.has(r) ? { height: table.rowHeights.get(r) } : undefined}>
                {[...cells.entries()]
                  .sort((a, b) => a[0] - b[0])
                  .map(([c, paragraphs]) => {
                    // A row held to its height (a label's) is exactly that tall: its
                    // words sit in a box of that height, centred if the cell says so.
                    const exact = table.rowRules?.get(r) === 'exact' && table.rowHeights?.get(r);
                    const v = table.vAligns?.get(`${r}:${c}`);
                    const m = table.look?.cellMarginPx;
                    const blocks = paragraphs.map((block) => <Block key={block.index} block={block} labels={labels} styles={styles} />);
                    return (
                      <td key={c} colSpan={table.spans?.get(`${r}:${c}`) || undefined} style={m || v ? { ...(m ? { padding: `${m.top}px ${m.right}px ${m.bottom}px ${m.left}px` } : {}), ...(v ? { verticalAlign: v === 'center' ? 'middle' : v } : {}) } : undefined}>
                        {exact ? (
                          <div className="wd-cell-exact" style={{ height: table.rowHeights.get(r) - (m ? m.top + m.bottom : 0), justifyContent: v === 'center' ? 'center' : v === 'bottom' ? 'flex-end' : 'flex-start' }}>{blocks}</div>
                        ) : blocks}
                      </td>
                    );
                  })}
              </tr>
            ))}
          </tbody>
        </table>
      ))}
    </>
  );
}

/**
 * The header or footer of one page, drawn in the margin where Word draws it
 * and greyed the way Word greys it while the body has the caret. Which band
 * depends on the page: the first-page band on page 1 when the section asks
 * for one (and nothing at all when the file defines none, which is how a
 * cover page has no header), the even band on even pages when it asks for
 * that, the default otherwise. A PAGE field says this page's number; the
 * count is what the layout pass counted. Not editable in place — a
 * double-click opens the band's dialog, as in Word.
 */
function Band({ kind, bands, section, onEdit, page = 1, of = null, top = 0, height = null }) {
  const set = bands?.[kind === 'header' ? 'headers' : 'footers'];
  if (!set) return null;
  const band = page === 1 && section?.titlePage
    ? set.first ?? null
    : page % 2 === 0 && section?.evenAndOdd && set.even
      ? set.even
      : set.default || set.even || set.first;
  const paragraphs = band?.paragraphs || [];
  if (!paragraphs.length) return null;
  const count = of || section?.pageCount || null;
  // w:pgMar's header/footer distance: from the page edge to the band, 48 px by default.
  const distance = Math.round(section?.margins?.[kind] ?? 48);
  const place = kind === 'header'
    ? { top: top + distance }
    : height
      ? { top: top + height - distance, transform: 'translateY(-100%)' }
      : { bottom: distance };
  return (
    <div
      className={`wd-band wd-${kind}`}
      contentEditable={false}
      style={place}
      onDoubleClick={onEdit}
      title={`Double-click to edit the ${kind}`}
    >
      {paragraphs.map((p, i) => (
        <p key={i} className="wd-band-line" style={{ textAlign: p.align === 'both' ? 'justify' : p.align || undefined }}>
          {(p.runs || []).map((r, j) => {
            const text = r.field === 'PAGE' ? String(page) : r.field === 'NUMPAGES' ? (count ? String(count) : '') : r.text ?? '';
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
/*
 * Tab stops are measured, not computed: where a tab lands depends on the
 * width of the text before it in the font the browser actually used. The
 * measuring is the expensive part — every read of a position forces the
 * browser to lay the page out, and a page of a thousand paragraphs takes
 * tens of milliseconds to lay out. Measured one paragraph at a time, with a
 * write after each read, a tender with 274 tabs forced 274 layouts and the
 * window did not answer for seven seconds. So paragraphs REGISTER here
 * during the commit and are measured together once it is over: every width
 * reset (one write), every position and text extent read (one layout), every
 * width applied (one write). A microtask runs after React's commit and
 * before the browser paints, so nothing flashes.
 */
let pendingTabs = null;
function scheduleTabs(p, stops) {
  if (!pendingTabs) {
    pendingTabs = new Map();
    queueMicrotask(flushTabs);
  }
  pendingTabs.set(p, stops);
}

function flushTabs() {
  const batch = pendingTabs;
  pendingTabs = null;
  if (!batch) return;
  const items = [...batch].filter(([p]) => p.isConnected);
  for (const [p] of items) for (const span of p.querySelectorAll('.wd-tab')) span.style.width = '';
  const plans = items.map(([p, stops]) => planTabs(p, stops));
  for (const plan of plans) {
    for (const t of plan) {
      t.span.style.width = `${t.width}px`;
      t.span.dataset.leader = t.leader;
    }
  }
}

/** Reads only: what width each tab in `p` should get, with the widths reset. */
function planTabs(p, stops) {
  const tabs = p.querySelectorAll('.wd-tab');
  if (!tabs.length) return [];

  // Stops are measured from the LEFT MARGIN — the page's content edge, or the
  // text box's, or the header's — not from the paragraph's own indent. A TOC
  // entry indented 29 px with a right stop at the margin's far edge used to
  // put that stop 29 px past the edge, and its page number on the next line.
  const host = p.closest('.wd-textbox, .wd-notes, .wd-band, .wd-page') || p;
  const edge = host.classList.contains('wd-band') ? p : host;
  const left = rectOf(host).left + (parseFloat(getComputedStyle(edge).paddingLeft) || 0);
  const custom = (stops || []).filter((s) => s && s.posPx > 0);
  const image = p.querySelector('.wd-image');
  const plan = [];
  // With every width reset, a later tab's natural position is short by the
  // widths the earlier tabs ON ITS OWN LINE will be given; carry those
  // forward, and only those. A paragraph that wraps starts each line afresh —
  // carrying a first line's widths into the second pushed every tab on it
  // sideways by the sum of the ones above, which is how a wrapped contents
  // entry ended up with its page number off the page.
  let carried = 0;
  let lineTop = null;
  for (const span of tabs) {
    const rect = rectOf(span);
    if (lineTop === null || Math.abs(rect.top - lineTop) > 1) {
      carried = 0;
      lineTop = rect.top;
    }
    const x = rect.left - left + carried;

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
      let w = rectOf(range).width;
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
    const finalWidth = Math.max(2, Math.round(width));
    carried += finalWidth;
    plan.push({ span, width: finalWidth, leader: stop.leader || '' });
  }
  return plan;
}

/** Paragraph borders as CSS: one line per side, in the file's colour and weight. */
/**
 * Where the page borders sit: each side in from the page edge by its space
 * (Word's default), or out from the text by it — the margin less the space.
 * `sheet` places the frame on one sheet of print layout; null frames the flow.
 */
function pageBordersStyle(section, sheet) {
  const b = section.pageBorders;
  const m = section.margins;
  const inset = (side) => {
    const s = b[side];
    const spacePx = ((s && s.spacePt) || 0) * (96 / 72);
    if (b.offsetFrom !== 'text') return Math.round(spacePx);
    const margin = side === 'top' ? m.top : side === 'bottom' ? m.bottom : side === 'left' ? m.left + (m.gutter || 0) : m.right;
    return Math.round(Math.max(0, margin - spacePx));
  };
  const line = (side) => {
    const s = b[side];
    if (!s || s.style === 'none' || s.style === 'nil') return 'none';
    const css = BORDER_CSS[s.style] || 'solid';
    // A double line needs three pixels before the browser draws two.
    const width = css === 'double' ? Math.max(3, s.widthPx || 1) : Math.max(1, s.widthPx || 1);
    return `${width}px ${css} ${s.colour || '#000'}`;
  };
  const top = inset('top');
  const bottom = inset('bottom');
  return {
    left: inset('left'),
    right: inset('right'),
    top: sheet ? sheet.top + top : top,
    ...(sheet ? { height: Math.max(0, sheet.height - top - bottom) } : { bottom }),
    borderTop: line('top'),
    borderLeft: line('left'),
    borderBottom: line('bottom'),
    borderRight: line('right'),
  };
}

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
    marginTop: block.spaceBeforePx != null ? Math.round(block.spaceBeforePx) : named?.spaceBeforePx != null ? Math.round(named.spaceBeforePx) : heading ? '1.1em' : undefined,
    marginBottom: block.spaceAfterPx != null ? Math.round(block.spaceAfterPx) : named?.spaceAfterPx != null ? Math.round(named.spaceAfterPx) : undefined,
    // The paragraph's own left indent, to the pixel — including an explicit
    // zero, which is how a contract's "DATED" line sits at the margin while
    // its Normal style indents everything else. Only a paragraph that says
    // nothing takes the style's indent.
    marginLeft: block.indentPx != null ? Math.round(block.indentPx)
      : block.indentLevel ? block.indentLevel * 24
      : block.indent ? block.indent * 24
      : named?.indentPx ? Math.round(named.indentPx) : undefined,
    marginRight: block.rightPx ? Math.round(block.rightPx) : named?.rightPx ? Math.round(named.rightPx) : undefined,
    // A first-line indent pushes the first line in; a hanging one pulls it out
    // and the rest of the paragraph in by the same amount, the way a list does.
    textIndent: firstLine ? Math.round(firstLine) : hanging ? -Math.round(hanging) : undefined,
    paddingLeft: hanging ? Math.round(hanging) : undefined,
    // An exact line (pixels) beats a multiplier; the paragraph's own beats the
    // style's. Word's "single" for a Latin face is about 1.2 of the size.
    lineHeight: block.lineHeightPx ? `${Math.round(block.lineHeightPx)}px`
      : block.lineSpacing ? block.lineSpacing * 1.2
      : named?.lineExactPx ? `${Math.round(named.lineExactPx)}px`
      : named?.lineFactor ? named.lineFactor * 1.2 : undefined,
    // Shading and borders: the paragraph's own, or its style's, drawn edge to
    // edge like Word.
    backgroundColor: block.shading || named?.shading || undefined,
    ...borderStyle(block.borders ?? named?.borders),
  };
}

/** The tab stops that apply: the paragraph's own, else its style's. */
const tabStops = (block, styles) => block.tabs ?? (styles ? (styles[block.style] ?? styles['*default*'])?.tabs : null) ?? null;

/**
 * Word's Outline/Shadow/Glow text effects, in CSS. Outline hollows the
 * letters out — a stroke instead of a fill, the way Word draws it — and a
 * shadow and a glow are both `text-shadow`, so a run wearing both gets two
 * layers in the one property.
 */
function effectsStyle(run) {
  const shadows = [];
  if (run.shadow) shadows.push('1px 1px 0 rgba(0,0,0,.55)');
  if (run.glow?.colour) shadows.push(`0 0 ${Math.round(run.glow.radiusPt * 96 / 72)}px #${run.glow.colour}`);
  return {
    ...(run.outline ? { WebkitTextStroke: '0.6px currentColor', WebkitTextFillColor: 'transparent' } : {}),
    textShadow: shadows.length ? shadows.join(', ') : undefined,
  };
}

/** Author name -> a stable colour off a small palette, the way Word colours a reviewer. */
const TRACK_COLOURS = ['#C00000', '#2B5FD9', '#0F9D58', '#7B5CD6', '#E08B2B', '#0D8F6F', '#C0399F'];
function authorColour(author) {
  const s = String(author || '');
  let h = 0;
  for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) >>> 0;
  return TRACK_COLOURS[h % TRACK_COLOURS.length];
}

/** One run of text: its direct formatting, its link, its tabs, and — Review → Track Changes — an insertion or a deletion. */
function RunSpan({ run, markupMode = 'simple', at = null }) {
  // An equation: its one character in a host the MathML is drawn behind.
  if (run.math) return <MathRun run={run} at={at} />;
  // A footnote/endnote reference, or the mark at the head of the note: the
  // number is drawn by CSS from `data-n`, so the span holds no text and the
  // engine's caret offsets stay exactly right.
  if (run.noteRef || run.noteMark) {
    const n = run.noteRef ? run.noteRef.n : run.noteMark?.n;
    // The reference's one character (U+FFFC) stays in the DOM at size zero,
    // so the caret can step over it and Backspace can take it; the number is
    // drawn beside it by CSS.
    return (
      <span className={run.noteRef ? 'wd-noteref' : 'wd-notemark'} data-n={n ?? '?'} data-kind={run.noteRef?.kind} data-id={run.noteRef?.id} title={run.noteRef ? `${run.noteRef.kind} ${n}` : undefined}>
        {run.noteRef ? <span className="wd-noteref-char">{run.text}</span> : null}
      </span>
    );
  }
  // A deletion carries no words in the editable text at all (see
  // `flatDelRuns`/`trackedRemoveRange`) — `del.text` is a side channel for
  // display only, so this island is marked uneditable and never lengthens
  // what the engine's own offsets count. Shown in All Markup and Original;
  // Simple Markup and No Markup are the document AS IF it were accepted,
  // where a deletion is already gone.
  if (run.del) {
    if (markupMode !== 'all' && markupMode !== 'original') return null;
    const colour = authorColour(run.del.author);
    return (
      <span
        className="wd-del"
        contentEditable={false}
        suppressContentEditableWarning
        title={`Deleted by ${run.del.author || 'Someone'}${run.del.date ? ' · ' + formatWhen(run.del.date) : ''}`}
        style={{ color: colour, textDecoration: markupMode === 'all' ? 'line-through' : undefined, opacity: markupMode === 'original' ? 1 : undefined }}
      >
        {withTabs(run.del.text)}
      </span>
    );
  }
  // An insertion is real, live text — it is already part of the document,
  // so it keeps its place in the caret's own offsets. All Markup underlines
  // it in the author's colour; Original hides it (visually only: it stays
  // in the DOM, at zero width, so the offsets it belongs to still resolve).
  const insHidden = run.ins && markupMode === 'original';
  return (
    <span
      // A field's shading rides the same span as its formatting — Word
      // shades a field grey whatever else it carries — and `data-instr` and
      // `data-name` are what the page's Ctrl+click handler reads to follow a
      // REF to its bookmark (`gotoBookmark`), without the engine's frame
      // having to carry anything more than the run already does.
      className={[run.field ? 'wd-field' : null, run.field && MERGE_KINDS.has(run.field.kind) ? 'wd-mergefield' : null, run.link ? 'wd-link' : null, run.ins ? 'wd-ins' : null].filter(Boolean).join(' ') || undefined}
      data-kind={run.field && MERGE_KINDS.has(run.field.kind) ? run.field.kind : undefined}
      data-instr={run.field ? run.field.instr : undefined}
      data-name={run.field?.kind === 'ref' ? run.field.name : undefined}
      data-link={run.link || undefined}
      title={run.field
        ? (run.field.kind === 'ref' ? `REF ${run.field.name} — Ctrl+click to go to the bookmark` : run.field.instr.trim())
        : run.ins ? `Inserted by ${run.ins.author || 'Someone'}${run.ins.date ? ' · ' + formatWhen(run.ins.date) : ''}`
        : (run.link ? `Ctrl+click to go there` : undefined)}
      style={{
        fontWeight: run.bold ? 700 : undefined,
        fontStyle: run.italic ? 'italic' : undefined,
        textDecoration: [
          run.underline || (run.ins && markupMode === 'all') ? 'underline' : '',
          run.strike ? 'line-through' : '',
        ].filter(Boolean).join(' ') || undefined,
        // The engine reports `fontColour` as bare hex, the way the file
        // stores it; the size is in points, the way Word means it. The
        // painter used to read `colour` and paint the size in pixels, so
        // a colour never showed and 12 pt drew at two-thirds size.
        // A hyperlink is blue and underlined because its Hyperlink character
        // style says so — the engine folds that style into the run — and a
        // link without the style (a TOC entry) is as plain as Word draws it.
        color: run.ins && markupMode === 'all' ? authorColour(run.ins.author) : run.fontColour ? `#${run.fontColour}` : undefined,
        backgroundColor: run.highlight ? HIGHLIGHT_CSS[run.highlight] || run.highlight : undefined,
        fontFamily: run.fontName || undefined,
        fontSize: run.fontSize ? `${run.fontSize}pt` : undefined,
        // A footnote reference sits superscript, a chemical formula sub.
        ...(run.vertAlign === 'superscript' ? { verticalAlign: 'super', fontSize: '0.65em' } : run.vertAlign === 'subscript' ? { verticalAlign: 'sub', fontSize: '0.65em' } : {}),
        textTransform: run.caps ? 'uppercase' : undefined,
        fontVariant: run.smallCaps ? 'small-caps' : undefined,
        ...(insHidden ? { display: 'inline-block', width: 0, overflow: 'hidden' } : {}),
        ...effectsStyle(run),
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
      if (p.querySelector('.wd-tab')) scheduleTabs(p, p._tabs || null);
    }
  });
  const beside = floatsBeside(box);
  const d = box.dist || {};
  const style = {
    width: box.widthPx ? Math.min(box.widthPx, 720) : undefined,
    minHeight: box.heightPx ? Math.min(box.heightPx, 900) : undefined,
    backgroundColor: box.fill || undefined,
    border: box.line ? `1px solid ${box.line}` : undefined,
    ...(beside
      ? { float: floatSide(box), margin: `${Math.round(d.t || 0)}px ${floatSide(box) === 'left' ? Math.round(d.r || 12) : 0}px ${Math.round(d.b || 6)}px ${floatSide(box) === 'right' ? Math.round(d.l || 12) : 0}px` }
      : { margin: box.hAlign === 'center' ? '6px auto' : box.hAlign === 'right' ? '6px 0 6px auto' : '6px 0' }),
  };
  return (
    <div ref={ref} className={`wd-textbox${beside ? ' wd-float' : ''}`} contentEditable={false} style={style} title={box.name || undefined}>
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
/** A note's words, for the dialog that changes them: every paragraph, the mark left out. */
const noteWords = (note) => (note.paragraphs || []).map((p) => p.text || '').join('\n').trim();

function Notes({ notes, kind, styles, onEdit, measure = false }) {
  if (!notes?.length) return null;
  return (
    <div className={`wd-notes wd-${kind}${measure ? ' wd-notes-measure' : ''}`} contentEditable={false} aria-hidden={measure || undefined}>
      {notes.map((note) => (
        <div key={note.id} className="wd-note" id={measure ? undefined : `wd-${kind}-${note.n}`} data-note={note.id} title="Double-click to change the words" onDoubleClick={() => onEdit?.(note)}>
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

/**
 * A page's footnotes, at its foot: the notes whose references the layout
 * pass put on this page, drawn above the bottom margin under a short rule,
 * as Word draws them. The pass reserved the room, so the body stops above.
 */
function PageNotes({ notes, at, page, top, height, styles, onEdit }) {
  const mine = (notes || []).filter((note) => at?.[note.id] === page);
  if (!mine.length) return null;
  return (
    <div className="wd-pagenotes" contentEditable={false} style={{ top, height }}>
      <Notes notes={mine} kind="footnotes" styles={styles} onEdit={onEdit} />
    </div>
  );
}

/**
 * A paragraph — or, once the layout pass has split it at a line, its parts:
 * one element per page it spans, each carrying the block's index and the
 * character offset it starts at, so the caret's arithmetic is unchanged.
 * Memoised, and the split array keeps its identity while it is unchanged,
 * so a keystroke re-renders the one paragraph it touched.
 */
const Block = React.memo(function Block({ block, labels, styles, split, pickedImage = null, inner = null, markupMode = 'simple', place = null }) {
  if (!split || !split.length) return <Part block={block} labels={labels} styles={styles} from={0} to={Infinity} first last pickedImage={pickedImage} inner={inner} markupMode={markupMode} place={place} />;
  const bounds = [0, ...split, Infinity];
  return (
    <>
      {bounds.slice(0, -1).map((from, j) => (
        <Part key={j} block={block} labels={labels} styles={styles} from={from} to={bounds[j + 1]} first={j === 0} last={j === bounds.length - 2} pickedImage={pickedImage} inner={inner} markupMode={markupMode} />
      ))}
    </>
  );
});

/**
 * Does the drawing float beside the words — square, tight or through wrap,
 * at the left or the right? Then it goes into the paragraph before the words,
 * as a CSS float, and the lines run round it. Everything else — inline,
 * top-and-bottom, centred, behind, in front — is drawn after the words.
 */
function floatsBeside(d) {
  if (!d?.anchored) return false;
  if (!['square', 'tight', 'through'].includes(d.wrap)) return false;
  return d.hAlign !== 'center';
}

const floatSide = (d) => (d.hAlign === 'right' || d.hAlign === 'outside' ? 'right' : 'left');

/**
 * A picture's box from what the file says about it. The width is capped at
 * the column; a floating picture keeps the distances the file gives it from
 * the words, with Word's own quarter-inch-ish defaults where it gives none.
 */
function imageStyle(image, inner = null, inline = false) {
  let width = image.widthPx ? Math.min(image.widthPx, 640) : undefined;
  // A picture taller than the page's inside is drawn to fit it, its
  // proportions kept: a sheet cannot hold more, and a picture that ran
  // over the edge was drawn across two sheets (owner, 2026-09-20).
  if (width && inner && image.heightPx && image.widthPx) {
    const height = image.heightPx * (width / image.widthPx);
    if (height > inner) width = Math.max(16, Math.floor(width * (inner / height)));
  }
  const base = { width, height: 'auto', maxWidth: '100%' };
  const d = image.dist || {};
  // An inline picture in a paragraph with no words sits in the line as Word
  // draws it — several to a line while they fit, a scanner's four cards two
  // to a line — and the paragraph's alignment places them. Under words it
  // is a block beneath them (the engine keeps pictures apart from the runs).
  if (!image.anchored) return inline ? { ...base, display: 'inline-block', verticalAlign: 'baseline', margin: '6px 0' } : { ...base, display: 'block', margin: '6px 0' };
  const side = image.hAlign === 'center' ? 'center' : floatSide(image);
  if (floatsBeside(image)) {
    return {
      ...base,
      float: side,
      margin: `${Math.round(d.t || 0)}px ${side === 'left' ? Math.round(d.r || 12) : 0}px ${Math.round(d.b || 6)}px ${side === 'right' ? Math.round(d.l || 12) : 0}px`,
    };
  }
  if (image.wrap === 'topAndBottom' || image.wrap === 'square' || image.wrap === 'tight' || image.wrap === 'through') {
    return { ...base, display: 'block', margin: side === 'center' ? '6px auto' : side === 'right' ? '6px 0 6px auto' : '6px auto 6px 0' };
  }
  // No wrap: behind the words or in front of them, at the paragraph's edge.
  return {
    ...base,
    position: 'absolute',
    top: 0,
    ...(side === 'center' ? { left: '50%', transform: 'translateX(-50%)' } : { [side]: 0 }),
    zIndex: image.behind ? -1 : 2,
    pointerEvents: 'auto',
  };
}

/**
 * Four corner handles over the picked picture. Drag one and the picture
 * follows, keeping its proportions, the page re-wrapping round it as it
 * goes; on release the engine is told the size. The handles are measured
 * from the picture's box after every layout, so they stay on it when the
 * pages move.
 */
function PictureHandles({ page, picked, model, pages, onResize, onDrag }) {
  const [box, setBox] = React.useState(null);
  const find = () => {
    const part = page.current && partFor(page.current, picked.block, 0);
    return part ? part.querySelectorAll('.wd-image')[picked.image] || null : null;
  };
  const measure = React.useCallback(() => {
    const img = find();
    if (!img || !page.current) return setBox(null);
    // In the page's own pixels, whatever the zoom: the handles live inside the page.
    const r = rectOf(img);
    const p = rectOf(page.current);
    setBox({ left: r.left - p.left, top: r.top - p.top, width: r.width, height: r.height });
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [picked, page]);
  React.useLayoutEffect(() => {
    measure();
  }, [measure, model, pages]);
  React.useEffect(() => {
    const el = page.current;
    if (!el || typeof ResizeObserver === 'undefined') return undefined;
    const ro = new ResizeObserver(() => measure());
    ro.observe(el);
    return () => ro.disconnect();
  }, [measure, page]);
  if (!box) return null;
  const start = (e, handle) => {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const img = find();
    if (!img) return;
    const x0 = e.clientX;
    const w0 = box.width;
    const ratio = box.height / Math.max(1, box.width);
    let width = w0;
    onDrag?.(true);
    const move = (ev) => {
      const dx = (ev.clientX - x0) * (handle.includes('w') ? -1 : 1);
      width = Math.max(16, Math.round(w0 + dx));
      img.style.width = `${width}px`;
      measure();
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      // After the page's own mouseup has seen the flag.
      setTimeout(() => onDrag?.(false), 0);
      if (width !== Math.round(w0)) onResize({ widthPx: width, heightPx: Math.round(width * ratio) });
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };
  return (
    <div className="wd-handles" contentEditable={false} aria-hidden="true" style={{ left: box.left, top: box.top, width: box.width, height: box.height }}>
      {[['nw', 0, 0], ['ne', 1, 0], ['sw', 0, 1], ['se', 1, 1]].map(([name, fx, fy]) => (
        <div key={name} className="wd-handle" data-handle={name} style={{ left: `calc(${fx * 100}% - 5px)`, top: `calc(${fy * 100}% - 5px)`, cursor: fx === fy ? 'nwse-resize' : 'nesw-resize' }} onMouseDown={(e) => start(e, name)} />
      ))}
    </div>
  );
}

/** The paragraphs a paginator keeps with what follows, by convention as much as by w:keepNext. */
const KEEP_WITH_NEXT = /^(Heading[1-6]|Title|Subtitle)$/;

function Part({ block, labels, styles, from, to, first, last, pickedImage = null, inner = null, markupMode = 'simple', place = null }) {
  const ref = React.useRef(null);
  const whole = first && last;
  const pick = (e, i) => {
    e.stopPropagation();
    e.currentTarget.dispatchEvent(new CustomEvent('wd-pick', { bubbles: true, detail: { block: block.index, image: i } }));
  };
  const picture = (image, i) => (
    <img
      key={i}
      className={`wd-image${floatsBeside(image) ? ' wd-float' : ''}${image.anchored && image.wrap === 'none' ? (image.behind ? ' behind' : ' front') : ''}${pickedImage === i ? ' picked' : ''}`}
      contentEditable={false}
      data-image={i}
      src={image.href}
      alt={image.name || ''}
      draggable={false}
      onClick={(e) => pick(e, i)}
      style={imageStyle(image, inner, inlineRow)}
    />
  );
  const runs = whole ? block.runs || [] : sliceRuns(block.runs, from, to);
  const hasTabs = runs.some((r) => r.text && r.text.includes('\t'));
  React.useLayoutEffect(() => {
    if (hasTabs && ref.current) scheduleTabs(ref.current, tabStops(block, styles));
  });
  const style = paragraphCss(block, styles);
  // A drop cap's own paragraph floats beside the words that follow it — the
  // CSS class gives it its float and its close line-height, but the space
  // the engine put above and below the ORIGINAL paragraph would otherwise
  // land here too and push the body down a line for nothing, so it is
  // cleared for this half; the body paragraph keeps its own.
  if (block.dropCap) { style.marginTop = 0; style.marginBottom = 0; }

  // A list marker arrives as `{ label, indentPx, bullet }` — the text to draw,
  // how far in it sits, and whether it is a bullet or a number. Drawing the
  // object itself is a React crash, and it was one: pressing "Bulleted list"
  // took the whole page down.
  const mark = labels?.[block.index];
  const label = typeof mark === 'string' ? mark : mark?.label ?? null;
  const markerIndent = typeof mark === 'object' && mark?.indentPx ? mark.indentPx : null;
  // The marker hangs to the left of the text by the level's hanging indent —
  // Word's bullet at a quarter inch with the text at a half — so the text
  // starts at the indent and the bullet sits in the space before it.
  const markerHang = typeof mark === 'object' && mark?.hangingPx ? Math.round(mark.hangingPx) : null;
  const listStyle = markerIndent ? { ...style, marginLeft: markerIndent, ...(markerHang ? { textIndent: -markerHang } : {}) } : style;
  // A continuation starts flush, without the space before; a part that goes
  // on ends without the space after, its last line justified like the rest.
  const flowStyle = whole
    ? listStyle
    : {
        ...listStyle,
        ...(first ? {} : { textIndent: 0, marginTop: 0 }),
        ...(last ? {} : { marginBottom: 0, ...(style.textAlign === 'justify' ? { textAlignLast: 'justify' } : {}) }),
      };
  // A frame placed on the page stands where the frame says, out of the
  // flow; an envelope's own words move across to its margin.
  const partStyle = place?.kind === 'frame'
    ? { ...flowStyle, position: 'absolute', left: place.left, top: place.top, width: place.width, height: place.height, margin: 0, overflow: place.height ? 'hidden' : undefined, boxSizing: 'border-box', zIndex: 1 }
    : place?.kind === 'shift' ? { ...flowStyle, left: place.dx } : flowStyle;
  // The pictures drawn under the words continue the paragraph's address
  // space after its text — offset `length + i` is picture i, and the text
  // boxes come after every picture — so a split between two picture lines
  // (a sheet of scanned cards, two to a line) sends the pictures after it to
  // the next page and keeps the ones before it. The pass reads the two
  // lengths off the element.
  const length = (block.runs || []).reduce((n, r) => n + (r.text ?? '').length, 0);
  const images = block.images || [];
  const inlineRow = length === 0;
  const tail = length + images.length;
  const under = images.map((image, i) => !floatsBeside(image) && from <= length + i && length + i < to);
  const boxesHere = from <= tail && tail < to;
  const drawsUnder = under.some(Boolean) || (boxesHere && (block.textBoxes || []).some((box) => !floatsBeside(box)));
  return (
    <p
      ref={ref}
      className={`wd-block${place?.kind === 'frame' ? ' wd-frame' : ''}${block.dropCap ? ' wd-dropcap' : ''}${block.dropCap?.kind === 'margin' ? ' wd-dropcap-margin' : ''}${block.tracked && markupMode !== 'final' && markupMode !== 'original' ? ' wd-changebar' : ''}`}
      data-block={block.index}
      data-style={block.style || 'Normal'}
      data-from={from > 0 ? from : undefined}
      data-length={images.length || block.textBoxes?.length ? length : undefined}
      data-tail={images.length || block.textBoxes?.length ? tail : undefined}
      data-part={whole ? undefined : first ? 0 : 1}
      data-break={first && block.pageBreakBefore ? '1' : undefined}
      data-keep={block.keepNext || KEEP_WITH_NEXT.test(block.style || '') ? '1' : undefined}
      data-keeplines={block.keepLines ? '1' : undefined}
      style={partStyle}
    >
      {/* Floats first, so the lines that follow run round them. */}
      {first ? (block.images || []).map((image, i) => (floatsBeside(image) ? picture(image, i) : null)) : null}
      {first ? (block.textBoxes || []).map((box, i) => (floatsBeside(box) ? <TextBox key={`f${i}`} box={box} styles={styles} /> : null)) : null}
      {first && label ? <span className="wd-marker" contentEditable={false} style={markerHang ? { display: 'inline-block', minWidth: markerHang, whiteSpace: 'nowrap', textIndent: 0, marginRight: 0 } : undefined}>{label}</span> : null}

      {/*
        A paragraph with no words needs a line to have a height and to hold
        the caret — unless pictures under the words give it both: a line
        break there put an empty line above every inserted picture.
      */}
      {runs.length ? (() => {
        // Each run's offset in the block rides an equation's host, so a click
        // on it and the selection's highlight know which character it is.
        let at = whole ? 0 : from;
        return runs.map((run, i) => {
          const here = at;
          at += (run.text ?? '').length;
          return <RunSpan key={i} run={run} markupMode={markupMode} at={here} />;
        });
      })() : drawsUnder ? null : <br />}
      {/*
        Pictures, charts and shapes sit under the paragraph's text as blocks —
        the engine's own honest simplification of float layout. Not editable:
        the caret has no business inside a picture, and letting the browser
        put it there is how an image gets deleted by a stray Backspace. Each
        goes in the part whose range holds it.
      */}
      {images.map((image, i) => (under[i] ? picture(image, i) : null))}
      {boxesHere ? (block.textBoxes || []).map((box, i) => (floatsBeside(box) ? null : <TextBox key={i} box={box} styles={styles} />)) : null}
    </p>
  );
}

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
/* A column, so the ruler sits above the page and the page is as tall as its
   content: as a row's flex item the page was stretched to the viewport's
   height — a fixed height — and a long document ran out of the bottom of it. */
.wd-scroll { flex: 1; overflow: auto; padding: 26px 0 40px; display: flex; flex-direction: column; align-items: center; background: var(--window); }
.wd-page {
  background: #fff; color: #111; border-radius: 2px; flex: none;
  /* A sheet of paper on a desk: a close shadow for the edge, a wide soft one for the lift. */
  box-shadow: 0 0 0 1px rgba(15, 20, 30, 0.05), 0 2px 6px rgba(15, 20, 30, 0.07), 0 14px 36px rgba(15, 20, 30, 0.1);
  outline: none; font-family: Calibri, "Segoe UI", system-ui, sans-serif; font-size: 15px;
  line-height: 1.5; caret-color: var(--accent); position: relative;
}
/* In print layout the flow is transparent and the sheets are drawn behind it, one per page. */
.wd-page.paged { background: transparent; box-shadow: none; }
/* Review → Track Changes: a paragraph carrying a change wears a bar in the
   margin, All Markup and Simple Markup alike — Word's own change bar, one
   per PARAGRAPH here rather than per line (a stated simplification). */
/* Word draws a change bar out in the left margin, clear of the words. */
.wd-block.wd-changebar { position: relative; }
.wd-block.wd-changebar::before { content: ""; position: absolute; left: -14px; top: 0; bottom: 0; width: 2px; background: #2b5fd9; pointer-events: none; }
:root[data-theme='dark'] .wd-block.wd-changebar::before { background: #6f9bff; }
.wd-del { cursor: default; }
.wd-sheet {
  position: absolute; left: 0; right: 0; z-index: 0; background: #fff; border-radius: 2px; pointer-events: none;
  box-shadow: 0 0 0 1px rgba(15, 20, 30, 0.05), 0 2px 6px rgba(15, 20, 30, 0.07), 0 14px 36px rgba(15, 20, 30, 0.1);
}
:root[data-theme='dark'] .wd-sheet { background: #f7f7f5; }
/* The page borders: a frame in the margins, over the sheet and under nothing it could hide. */
.wd-pgborders { position: absolute; z-index: 1; box-sizing: border-box; pointer-events: none; }
/* Line numbers: a column down the left margin, each beside the line it counts. */
.wd-linenos { position: absolute; top: 0; width: 28px; z-index: 1; pointer-events: none; user-select: none; }
.wd-linenos span { position: absolute; left: 0; width: 28px; text-align: right; font: 11px Calibri, "Segoe UI", sans-serif; color: #8a8f98; }
/* Headers and footers sit in the margins, greyed while the body has the caret. */
.wd-band { position: absolute; left: 0; right: 0; z-index: 1; color: #777; font-size: 13px; line-height: 1.35; user-select: none; cursor: default; }
.wd-band .wd-band-line { margin: 0; padding: 0 var(--wd-margin-right, 96px) 0 var(--wd-margin-left, 96px); min-height: 1.2em; white-space: pre-wrap; }
.wd-band:hover { color: #333; }
.wd.mode-web .wd-band, .wd.mode-draft .wd-band, .wd.mode-read .wd-band, .wd.mode-outline .wd-band { display: none; }
:root[data-theme='dark'] .wd-page { background: #f7f7f5; }
/* pre-wrap: a run of spaces, a line break inside a paragraph (w:br) and a tab
   all mean what they meant in Word, instead of collapsing to one space. */
.wd-block { margin: 0 0 0.55em; min-height: 1.2em; white-space: pre-wrap; }
/* A drop cap's letter: its own tiny paragraph, floated so the body
   paragraph after it wraps round it exactly as the browser wraps text
   round any float. "In margin" then paints it shifted fully out of that
   reserved space by its own rendered width — a transform costs no layout
   pass, unlike measuring the glyph to write a pixel margin. */
.wd-block.wd-dropcap { float: left; margin: 0 6px 0 0; padding: 0; line-height: 0.8; }
.wd-block.wd-dropcap-margin { transform: translateX(-100%); }
.wd-tab { display: inline-block; white-space: pre; tab-size: 0; overflow: hidden; vertical-align: baseline; min-width: 2px; }
.wd-tab[data-leader="dot"] { background: radial-gradient(circle, currentColor 0.6px, transparent 0.9px) 0 calc(100% - 3px) / 4px 2px repeat-x; }
.wd-tab[data-leader="hyphen"] { background: linear-gradient(currentColor, currentColor) 0 calc(100% - 3px) / 3px 1px repeat-x; }
.wd-tab[data-leader="underscore"] { border-bottom: 1px solid currentColor; }
.wd-marker { color: #555; margin-right: 6px; user-select: none; }
/* A picture: click to pick it; a picked one wears the accent. A float sits
   beside the words; one with no wrap sits behind or in front of them. */
.wd-image { cursor: default; }
.wd-image.picked { outline: 2px solid var(--accent); outline-offset: 2px; }
.wd-handles { position: absolute; z-index: 3; pointer-events: none; }
.wd-handle { position: absolute; width: 10px; height: 10px; background: #fff; border: 1.5px solid var(--accent); border-radius: 2px; box-sizing: border-box; pointer-events: auto; }
.wd-image.behind { opacity: .92; }
.wd-mounting { margin: 12px 0 0; font-size: 12px; color: var(--ink-3); user-select: none; }
/* A wrapper only so the status bar's columns chip has its own selector; it
   must not become an extra flex item of its own beside the chip it holds. */
.wd-columns-chip { display: contents; }

/* A watermark: the header's WordArt, drawn behind the body as Word does. */
.wd-watermark {
  position: absolute; left: 0; right: 0; top: 380px; text-align: center; pointer-events: none; user-select: none;
  font-family: Calibri, "Segoe UI", sans-serif; font-size: 150px; font-weight: 400; opacity: .35; letter-spacing: .02em; z-index: 0;
}
.wd-page > .wd-block, .wd-page > .wd-table, .wd-page > .wd-notes { position: relative; z-index: 1; }
/* Footnote references and the notes themselves. */
.wd-noteref::after, .wd-notemark::after { content: attr(data-n); vertical-align: super; font-size: 0.65em; line-height: 0; }
.wd-noteref-char { font-size: 0; }
/* A field — a REF, a PAGE, anything cached by <w:fldSimple> — shaded grey the
   way Word shades every field, so its words read as computed rather than typed. */
.wd-field { background: rgba(0, 0, 0, 0.08); border-radius: 2px; }
/* A merge field reads as plain words, «chevrons» and all, the way Word draws
   one — until Mailings → Highlight Merge Fields shades every one grey. */
.wd-field.wd-mergefield { background: transparent; }
.wd-page.wd-mm-hl .wd-field.wd-mergefield { background: #d9d9d9; }
.wd-notes .wd-note { cursor: text; }
.wd-notemark::after { margin-right: 3px; }
.wd-notes { margin-top: 28px; padding-top: 6px; border-top: 1px solid #333; width: 33%; min-width: 220px; font-size: 0.85em; user-select: none; }
.wd-notes .wd-note { width: 300%; display: flow-root; }
.wd-notes .wd-box-p { margin: 0 0 3px; white-space: pre-wrap; }
/* Footnotes at the foot of their page: the copy in the flow is hidden and
   measured, and each page draws its own above the bottom margin. */
.wd-notes-measure { visibility: hidden; height: 0; overflow: hidden; margin: 0; padding: 0; border: 0; width: 100%; min-width: 0; pointer-events: none; }
.wd-notes-measure .wd-note { width: auto; }
.wd-pagenotes { position: absolute; left: var(--wd-margin-left, 96px); right: var(--wd-margin-right, 96px); display: flex; flex-direction: column; justify-content: flex-end; pointer-events: none; z-index: 1; }
.wd-pagenotes > .wd-notes { margin: 0; padding-top: 10px; border-top: 0; width: auto; min-width: 0; pointer-events: auto; }
.wd-pagenotes > .wd-notes::before { content: ''; display: block; width: 33%; min-width: 220px; border-top: 1px solid #333; margin-bottom: 6px; }
.wd-pagenotes .wd-note { width: auto; }
/* A text box: in flow under its paragraph, the caret kept out of it. */
.wd-textbox { display: block; box-sizing: border-box; padding: 4px 8px; max-width: 100%; overflow: hidden; user-select: none; }
.wd-textbox .wd-box-p { margin: 0; min-height: 1.2em; white-space: pre-wrap; }
.wd-table { border-collapse: collapse; width: 100%; margin: 0.6em 0; }
.wd-table td { border: 1px solid #bbb; padding: 4px 7px; vertical-align: top; }
/* A table with its borders off — a sheet of labels — has no lines, only
   Word's View Gridlines: faint dashes that take no room and never print. */
.wd-table.wd-table-bare { margin: 0; }
.wd-table.wd-table-bare td { border: 0; outline: 1px dashed rgba(70, 120, 200, 0.35); outline-offset: -1px; }
.wd-cell-exact { display: flex; flex-direction: column; overflow: hidden; }
.wd-cell-exact > .wd-block { margin-top: 0; margin-bottom: 0; }
/* An envelope in front of the letter: its sheet at the envelope's size. */
.wd-sheet.wd-envelope-sheet { box-shadow: 0 0 0 1px rgba(15, 20, 30, 0.08), 0 2px 6px rgba(15, 20, 30, 0.08), 0 10px 28px rgba(15, 20, 30, 0.1); }
.wd-find {
  position: absolute; top: 12px; right: 22px; display: flex; align-items: center; gap: 7px;
  background: var(--surface); border: 1px solid var(--line); border-radius: var(--r-3);
  padding: 7px 10px; box-shadow: var(--shadow-2); z-index: 20;
}
.wd-find .rw-input { width: 160px; }
`;
