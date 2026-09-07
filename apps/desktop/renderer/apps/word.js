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
import { Ribbon, Group, Button, Separator, Icon, Spacer, Chip, Empty, Spinner, Select, useToast, useMenu, useCommands, menuItems } from '@rutba/office-ui';
import { AppFrame, useAppMenu, pickOpen, pickSave, confirmDiscard, useFileDrop, openInApp , useDirtyGuard } from '../shell.js';

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

export default function Word({ app, shell, boot }) {
  const toast = useToast();
  const [doc, setDoc] = useState(null);
  const [model, setModel] = useState(null);
  const [busy, setBusy] = useState(true);
  const [error, setError] = useState(null);
  const [tab, setTab] = useState('home');
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
        <Ribbon
          tabs={[
            { id: 'home', label: 'Home' },
            { id: 'insert', label: 'Insert' },
            { id: 'layout', label: 'Layout' },
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
                <Button tall icon="new" label="New" onClick={() => shell.win.create({ app: 'word' })} />
                <Button tall icon="open" label="Open" onClick={openFile} />
                <Button tall icon="save" label="Save" onClick={() => save(false)} />
              </Group>
              <Group label="Font">
                <Button icon="bold" pressed={format.bold} title="Bold" onClick={() => commands['format.bold'].run()} />
                <Button icon="italic" pressed={format.italic} title="Italic" onClick={() => commands['format.italic'].run()} />
                <Button icon="underline" pressed={format.underline} title="Underline" onClick={() => commands['format.underline'].run()} />
                <Button icon="strike" pressed={format.strike} title="Strikethrough" onClick={() => apply({ op: 'toggleFormat', tag: 'strike' })} />
                <Separator />
                <Select
                  value={format.fontSize || ''}
                  onChange={(e) => apply({ op: 'setRunFormat', delta: { size: Number(e.target.value) || null } })}
                  style={{ width: 62 }}
                >
                  <option value="">Size</option>
                  {[8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48].map((s) => (
                    <option key={s} value={s}>{s}</option>
                  ))}
                </Select>
              </Group>
              <Group label="Paragraph">
                <Button icon="alignLeft" pressed={format.paragraphAlign === 'left'} title="Align left" onClick={() => align('left')} />
                <Button icon="alignCenter" pressed={format.paragraphAlign === 'center'} title="Centre" onClick={() => align('center')} />
                <Button icon="alignRight" pressed={format.paragraphAlign === 'right'} title="Align right" onClick={() => align('right')} />
                <Button icon="alignJustify" pressed={format.paragraphAlign === 'justify'} title="Justify" onClick={() => align('justify')} />
              </Group>
              <Group label="Styles">
                <Select value={format.paragraphStyle || ''} onChange={(e) => setStyle(e.target.value)} style={{ width: 128 }}>
                  <option value="">Body text</option>
                  {(model?.paragraphStyles || []).map((s) => (
                    <option key={s.id} value={s.id}>{s.name}</option>
                  ))}
                </Select>
              </Group>
              <Group label="Editing">
                <Button tall icon="find" label="Find" onClick={() => commands['edit.find'].run()} />
              </Group>
            </>
          ) : tab === 'insert' ? (
            <>
              <Group label="Tables">
                <Button tall icon="table" label="Table" onClick={() => commands['insert.table'].run()} />
              </Group>
              <Group label="Illustrations">
                <Button tall icon="picture" label="Picture" onClick={async () => {
                  const file = (await shell.dialog.open({ title: 'Insert picture', filters: [{ name: 'Images', extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp'] }] }))[0];
                  if (!file) return;
                  const { bytes, stat } = await shell.fs.read({ path: file });
                  apply({ op: 'insertImage', name: stat.name, contentType: stat.ext === '.png' ? 'image/png' : 'image/jpeg', data: bytes, widthPx: 420, heightPx: 280 });
                }} />
                <Button tall icon="chart" label="Chart" onClick={() => apply({ op: 'insertChart', kind: 'column' })} />
                <Button tall icon="shape" label="Shape" onClick={() => apply({ op: 'insertShape', preset: 'rect', widthPx: 200, heightPx: 120 })} />
              </Group>
              <Group label="Pages">
                <Button tall icon="file" label="Break" onClick={() => commands['insert.break'].run()} />
              </Group>
            </>
          ) : tab === 'layout' ? (
            <>
              <Group label="Page">
                <Button icon="file" label="Portrait" onClick={() => apply({ op: 'setPageSetup', spec: { orientation: 'portrait' } })} />
                <Button icon="file" label="Landscape" onClick={() => apply({ op: 'setPageSetup', spec: { orientation: 'landscape' } })} />
              </Group>
              <Group label="Export">
                <Button tall icon="pdf" label="PDF" onClick={() => exportAs('pdf')} />
                <Button tall icon="export" label="Markdown" onClick={() => exportAs('md')} />
                <Button tall icon="export" label="Text" onClick={() => exportAs('txt')} />
              </Group>
            </>
          ) : (
            <>
              <Group label="Zoom">
                <Button icon="zoomOut" label="Out" onClick={() => shell.win.zoom({ delta: -0.1 })} />
                <Button icon="zoomIn" label="In" onClick={() => shell.win.zoom({ delta: 0.1 })} />
                <Button icon="check" label="100%" onClick={() => shell.win.zoom({ reset: true })} />
              </Group>
              <Group label="Print">
                <Button tall icon="print" label="Print" onClick={() => shell.print.print({})} />
              </Group>
            </>
          )}
        </Ribbon>
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
              {model.blocks.map((block) => (
                <Block key={block.index} block={block} labels={model.listLabels} />
              ))}
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
    </AppFrame>
  );
}

// Memoised: with the patch above, a keystroke changes one block, and only
// that one should re-render. Without this the saving is thrown away in
// reconciliation.
const Block = React.memo(function Block({ block, labels }) {
  const style = {
    textAlign: block.align || undefined,
    marginLeft: block.indent ? block.indent * 24 : undefined,
    fontSize: HEADING_SIZES[block.style] ? `${HEADING_SIZES[block.style]}px` : undefined,
    fontWeight: block.style && /Title|Heading/.test(block.style) ? 600 : undefined,
    marginTop: block.style && /Title|Heading/.test(block.style) ? '1.1em' : undefined,
  };

  if (block.table) {
    return (
      <table className="wd-table" data-block={block.index}>
        <tbody>
          {block.table.rows.map((row, r) => (
            <tr key={r}>
              {row.cells.map((cell, c) => (
                <td key={c} colSpan={cell.colspan || 1} rowSpan={cell.rowspan || 1}>
                  {(cell.blocks || []).map((b) => (b.runs || []).map((run) => run.text).join('')).join('\n')}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    );
  }

  const label = labels?.[block.index];
  return (
    <p className="wd-block" data-block={block.index} style={style}>
      {label ? <span className="wd-marker" contentEditable={false}>{label}</span> : null}
      {(block.runs || []).length ? (
        block.runs.map((run, i) => (
          <span
            key={i}
            style={{
              fontWeight: run.bold ? 700 : undefined,
              fontStyle: run.italic ? 'italic' : undefined,
              textDecoration: [run.underline ? 'underline' : '', run.strike ? 'line-through' : ''].filter(Boolean).join(' ') || undefined,
              color: run.colour || undefined,
              backgroundColor: run.highlight || undefined,
              fontFamily: run.fontName || undefined,
              fontSize: run.fontSize ? `${run.fontSize}px` : undefined,
            }}
          >
            {run.text}
          </span>
        ))
      ) : (
        <br />
      )}
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
