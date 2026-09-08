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
import WordRibbon from './word/ribbon.js';
import {
  LinkDialog, TableDialog, BandDialog, CommentDialog, CommentsDialog, FindDialog, WordCountDialog,
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
    apply({ op: 'setSelection', anchor: pos.anchor || pos.focus, focus: pos.focus });
  }, [apply, currentPosition]);

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
        <div className="wd">
          <style>{CSS}</style>
          <div className="wd-scroll">
            <div
              className="wd-page"
              ref={pageRef}
              contentEditable
              suppressContentEditableWarning
              spellCheck
              onMouseUp={syncSelection}
              onKeyDown={(e) => {
                // Arrow keys and Home/End move the caret without an edit, so the
                // engine is told where it landed — after the browser has moved
                // it, which is why this waits a tick.
                if (/^(Arrow|Home|End|Page)/.test(e.key)) setTimeout(syncSelection, 0);
              }}
              onContextMenu={(e) => menu.open(e, menuItems(commands, ['edit.undo', 'edit.redo', '-', 'format.bold', 'format.italic', 'format.underline', '-', 'edit.find']))}
              style={{
                width: section ? Math.round(section.widthPx) : 794,
                minHeight: section ? Math.round(section.heightPx) : 1123,
                paddingTop: section?.margins.top ?? 96,
                paddingRight: section?.margins.right ?? 96,
                paddingBottom: section?.margins.bottom ?? 96,
                paddingLeft: section?.margins.left ?? 96,
              }}
            >
              {groupTables(model.blocks).map((item) =>
                item.table ? (
                  <TableGroup key={`t${item.table.id}`} table={item.table} labels={model.listLabels} />
                ) : (
                  <Block key={item.index} block={item} labels={model.listLabels} />
                )
              )}
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

function TableGroup({ table, labels }) {
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
                    <Block key={block.index} block={block} labels={labels} />
                  ))}
                </td>
              ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

const Block = React.memo(function Block({ block, labels }) {
  const style = {
    // 'both' is OOXML for justified; the other three are CSS already.
    textAlign: block.align === 'both' ? 'justify' : block.align || undefined,
    marginLeft: block.indentLevel ? block.indentLevel * 24 : block.indent ? block.indent * 24 : undefined,
    lineHeight: block.lineSpacing ? block.lineSpacing * 1.2 : undefined,
    fontSize: HEADING_SIZES[block.style] ? `${HEADING_SIZES[block.style]}px` : undefined,
    fontWeight: block.style && /Title|Heading/.test(block.style) ? 600 : undefined,
    marginTop: block.style && /Title|Heading/.test(block.style) ? '1.1em' : undefined,
  };

  // A list marker arrives as `{ label, indentPx, bullet }` — the text to draw,
  // how far in it sits, and whether it is a bullet or a number. Drawing the
  // object itself is a React crash, and it was one: pressing "Bulleted list"
  // took the whole page down.
  const mark = labels?.[block.index];
  const label = typeof mark === 'string' ? mark : mark?.label ?? null;
  const markerIndent = typeof mark === 'object' && mark?.indentPx ? mark.indentPx : null;
  return (
    <p className="wd-block" data-block={block.index} style={markerIndent ? { ...style, marginLeft: markerIndent } : style}>
      {label ? <span className="wd-marker" contentEditable={false}>{label}</span> : null}
      {(block.runs || []).length ? (
        block.runs.map((run, i) => (
          <span
            key={i}
            style={{
              fontWeight: run.bold ? 700 : undefined,
              fontStyle: run.italic ? 'italic' : undefined,
              textDecoration: [run.underline ? 'underline' : '', run.strike ? 'line-through' : ''].filter(Boolean).join(' ') || undefined,
              // The engine reports `fontColour` as bare hex, the way the file
              // stores it; the size is in points, the way Word means it. The
              // painter used to read `colour` and paint the size in pixels, so
              // a colour never showed and 12 pt drew at two-thirds size.
              color: run.fontColour ? `#${run.fontColour}` : undefined,
              backgroundColor: run.highlight ? HIGHLIGHT_CSS[run.highlight] || run.highlight : undefined,
              fontFamily: run.fontName || undefined,
              fontSize: run.fontSize ? `${run.fontSize}pt` : undefined,
            }}
          >
            {run.text}
          </span>
        ))
      ) : (
        <br />
      )}
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
  line-height: 1.5; caret-color: var(--accent);
}
:root[data-theme='dark'] .wd-page { background: #f7f7f5; }
.wd-block { margin: 0 0 0.55em; min-height: 1.2em; }
.wd-marker { color: #555; margin-right: 6px; user-select: none; }
.wd-table { border-collapse: collapse; width: 100%; margin: 0.6em 0; }
.wd-table td { border: 1px solid #bbb; padding: 4px 7px; vertical-align: top; }
.wd-find {
  position: absolute; top: 12px; right: 22px; display: flex; align-items: center; gap: 7px;
  background: var(--surface); border: 1px solid var(--line); border-radius: var(--r-3);
  padding: 7px 10px; box-shadow: var(--shadow-2); z-index: 20;
}
.wd-find .rw-input { width: 160px; }
`;
