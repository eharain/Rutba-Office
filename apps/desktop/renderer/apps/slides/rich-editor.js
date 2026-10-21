// The editor a text box (or a table cell) is typed in: its words as they
// look, bold where they are bold and the size they are, and the selected
// words formatted on their own (owner, 2026-10-21: "the entire text box has
// single formatting and does not allow formatting selected text
// separately"). The paragraphs are the truth: every keystroke, deletion,
// paste and format is made on them (rich-text.js) and the box drawn again,
// so what is shown is always what will be kept, and nothing the browser
// would add of its own gets into the file. Words typed through an input
// method are left to the browser until they are done, then read back.

import React, { forwardRef, useCallback, useEffect, useImperativeHandle, useRef } from 'react';
import { plainOf, runPlain, lookOf, replaceRange, formatRange, formatRun, formatParagraphsIn, formatState, runsIn, RUN_KEYS, PARAGRAPH_KEYS } from '@rutba/ooxml/rich-runs';

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const WORD = /[\p{L}\p{N}_'’-]/u;
/** Where focus may go without ending the editing: the ribbon, its menus and galleries, the Format pane, a dialog. */
export const KEEPS_EDITING = '.rw-ribbon, .rw-menu, .rw-popover, .rw-dialog, .rw-panel, [role="menu"], [role="listbox"], [role="dialog"], [data-keeps-editing]';

/** A run's look as CSS. */
function runCss(r) {
  const deco = [r.underline || r.link ? 'underline' : '', r.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  return [
    r.bold ? 'font-weight:700' : r.bold === false ? 'font-weight:400' : '',
    r.italic ? 'font-style:italic' : '',
    deco ? `text-decoration:${deco}` : '',
    r.color ? `color:${r.color}` : '',
    r.size ? `font-size:${r.size}px` : '',
    r.font ? `font-family:'${String(r.font).replace(/'/g, '')}', 'Segoe UI', sans-serif` : '',
    r.highlight ? `background:${r.highlight}` : '',
  ].filter(Boolean).join(';');
}

/** The box drawn from its paragraphs. Each paragraph a div, each run a span; a break, a field and an equation are kept whole. */
function draw(el, paragraphs) {
  let number = 0;
  el.innerHTML = paragraphs.map((p, i) => {
    const bullet = p.bullet?.type === 'char' ? p.bullet.char || '•' : p.bullet?.type === 'number' ? `${(number += 1) + (p.bullet.start || 1) - 1}.` : '';
    if (p.bullet?.type !== 'number') number = 0;
    const css = [
      p.align ? `text-align:${p.align === 'justify' ? 'justify' : p.align === 'center' ? 'center' : p.align === 'right' ? 'right' : 'left'}` : '',
      p.rtl ? 'direction:rtl' : '',
      p.level ? `padding-inline-start:${p.level * 1.5}em` : '',
      p.lineHeight ? `line-height:${p.lineHeight}` : '',
    ].filter(Boolean).join(';');
    const runs = (p.runs || []).map((r) => {
      if (r.break) return '<br data-k="break">';
      if (r.field || r.math || r.picture) return `<span contenteditable="false" data-k="atom" data-len="${runPlain(r).length}" style="${runCss(r)}">${esc(runPlain(r))}</span>`;
      return r.text ? `<span style="${runCss(r)}">${esc(r.text)}</span>` : '';
    }).join('');
    // A paragraph that is empty, or ends in a line break, needs a last <br> for the caret to stand on its line.
    const tail = !runs || p.runs[p.runs.length - 1]?.break ? '<br data-k="empty">' : '';
    return `<div data-p="${i}"${bullet ? ` data-bullet="${esc(bullet)}"` : ''}${css ? ` style="${css}"` : ''}>${runs}${tail}</div>`;
  }).join('');
}

/** The leaves of one paragraph's div, in order, each with how many characters it counts for. */
function leaves(div) {
  const out = [];
  const walk = (node) => {
    for (const n of node.childNodes) {
      if (n.nodeType === 3) out.push({ node: n, len: n.data.length });
      else if (n.nodeType === 1) {
        if (n.dataset?.k === 'atom') out.push({ node: n, len: Number(n.dataset.len) || 0, atom: true });
        else if (n.tagName === 'BR') out.push({ node: n, len: n.dataset?.k === 'break' ? 1 : 0, br: true });
        else walk(n);
      }
    }
  };
  walk(div);
  return out;
}

const paragraphsOf = (el) => [...el.children].filter((d) => d.tagName === 'DIV');

/** Where a leaf ends, as a DOM point. */
const endOf = (l) => (l.node.nodeType === 3 ? [l.node, l.node.data.length] : [l.node.parentNode, [...l.node.parentNode.childNodes].indexOf(l.node) + 1]);

/** The plain text the page shows now: what an input method typed is read from it. */
function plainOfPage(el) {
  return paragraphsOf(el).map((div) => leaves(div).map((l) => (l.node.nodeType === 3 ? l.node.data : l.br ? (l.len ? '\v' : '') : l.node.textContent)).join('')).join('\n');
}

/** A DOM point as an offset into the plain text. */
function offsetOf(el, node, offset) {
  const at = document.createRange();
  at.setStart(node, offset);
  const divs = paragraphsOf(el);
  let total = 0;
  for (let i = 0; i < divs.length; i++) {
    for (const l of leaves(divs[i])) {
      if (l.node === node && l.node.nodeType === 3) return total + offset;
      if (l.atom && l.node.contains(node)) return total;
      // The whole leaf lies before the point: counted.
      if (at.comparePoint(...endOf(l)) <= 0) total += l.len;
      else return total;
    }
    // The point is at the end of this paragraph unless it lies past the start of the next.
    const next = divs[i + 1];
    if (!next || at.comparePoint(next, 0) > 0) return total;
    total += 1;
  }
  return total;
}

/** An offset into the plain text as a DOM point. */
function positionOf(el, at) {
  const divs = paragraphsOf(el);
  let left = Math.max(0, at);
  for (let i = 0; i < divs.length; i++) {
    const ls = leaves(divs[i]);
    const len = ls.reduce((n, l) => n + l.len, 0);
    if (left > len && i < divs.length - 1) { left -= len + 1; continue; }
    let after = null; // the point just past the last leaf passed
    for (const l of ls) {
      if (!l.len) continue;
      if (l.node.nodeType === 3) {
        if (left <= l.len) return { node: l.node, offset: left };
      } else if (left === 0) {
        return { node: l.node.parentNode, offset: [...l.node.parentNode.childNodes].indexOf(l.node) };
      }
      left -= l.len;
      const [node, offset] = endOf(l);
      after = { node, offset };
    }
    return after || { node: divs[i], offset: 0 };
  }
  return { node: el, offset: 0 };
}

export const RichEditor = forwardRef(function RichEditor({ paragraphs: initial, base = {}, className = 'sl-rich', style, dir, anchor = 'top', caret = null, onCommit, onCancel, onCaret, onFormat, onKey }, ref) {
  const el = useRef(null);
  const model = useRef((initial || []).length ? initial : [{ runs: [{ text: '' }] }]);
  const sel = useRef({ start: 0, end: 0 });
  const pending = useRef(null); // a look chosen with nothing selected, for the words typed next
  const history = useRef({ list: [], at: -1, typedAt: 0 });
  const composing = useRef(false);
  const done = useRef(false);

  const report = useCallback(() => {
    const { start, end } = sel.current;
    onCaret?.({ start, end, text: plainOf(model.current) });
    onFormat?.(formatState(model.current, start, end, base));
  }, [onCaret, onFormat, base]);

  const place = useCallback((start, end = start) => {
    const node = el.current;
    if (!node) return;
    const a = positionOf(node, start);
    const b = positionOf(node, end);
    const range = document.createRange();
    range.setStart(a.node, a.offset);
    range.setEnd(b.node, b.offset);
    const s = window.getSelection();
    s.removeAllRanges();
    s.addRange(range);
    sel.current = { start, end };
  }, []);

  const remember = useCallback((typing = false) => {
    const h = history.current;
    const now = Date.now();
    const snap = { paragraphs: model.current, sel: { ...sel.current } };
    // Words typed close together are one step, as a word processor's undo counts them.
    if (typing && h.at >= 0 && now - h.typedAt < 800 && h.list[h.at]?.typing) h.list[h.at] = { ...snap, typing: true };
    else { h.list = h.list.slice(0, h.at + 1); h.list.push({ ...snap, typing }); h.at = h.list.length - 1; }
    h.typedAt = typing ? now : 0;
  }, []);

  const set = useCallback((paragraphs, start, end = start, { typing = false, focus = true } = {}) => {
    model.current = paragraphs;
    draw(el.current, paragraphs);
    if (focus) el.current.focus({ preventScroll: true });
    place(start, end);
    remember(typing);
    report();
  }, [place, remember, report]);

  /** The selection as offsets, read from the page. */
  const read = useCallback(() => {
    const s = window.getSelection();
    if (!s?.rangeCount || !el.current?.contains(s.anchorNode)) return sel.current;
    const r = s.getRangeAt(0);
    const a = offsetOf(el.current, r.startContainer, r.startOffset);
    const b = offsetOf(el.current, r.endContainer, r.endOffset);
    sel.current = { start: Math.min(a, b), end: Math.max(a, b) };
    return sel.current;
  }, []);

  useEffect(() => {
    draw(el.current, model.current);
    el.current.focus({ preventScroll: true });
    const len = plainOf(model.current).length;
    const at = caret == null ? len : Math.min(caret, len);
    place(at, at);
    history.current = { list: [{ paragraphs: model.current, sel: { start: at, end: at } }], at: 0, typedAt: 0 };
    report();
    // Only when the editor opens: later, the model is the truth.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Left with nothing changed, the box is let go without a step in Undo.
  const opening = useRef(JSON.stringify(model.current));
  const commit = useCallback(() => {
    if (done.current) return;
    done.current = true;
    if (JSON.stringify(model.current) === opening.current) onCancel?.();
    else onCommit?.(model.current);
  }, [onCommit, onCancel]);

  /** Format the words chosen: run keys to the words (the word at the caret, or the next typed), paragraph keys to their paragraphs. */
  const format = useCallback((delta) => {
    const { start, end } = sel.current;
    const runPart = Object.fromEntries(Object.entries(delta).filter(([k]) => RUN_KEYS.includes(k)));
    const paraPart = Object.fromEntries(Object.entries(delta).filter(([k]) => PARAGRAPH_KEYS.includes(k)));
    let next = model.current;
    let a = start;
    let b = end;
    if (Object.keys(runPart).length) {
      if (a === b) {
        // The word the caret is in, as PowerPoint takes it; between words, the look waits for what is typed.
        const plain = plainOf(next);
        let s = a;
        let e = a;
        while (s > 0 && WORD.test(plain[s - 1])) s--;
        while (e < plain.length && WORD.test(plain[e])) e++;
        if (s < a && e > a) { next = formatRange(next, s, e, runPart); } else {
          const look = lookOf(runsIn(next, a, a)[0] || {});
          const resolved = { ...runPart };
          for (const k of ['bold', 'italic', 'underline', 'strike']) if (resolved[k] === 'toggle') resolved[k] = !(pending.current?.[k] ?? look[k]);
          pending.current = formatRun({ ...look, ...(pending.current || {}) }, resolved);
          onFormat?.({ ...formatState(next, a, b, base), ...pending.current, font: pending.current.font || '' });
          el.current.focus({ preventScroll: true });
          place(a, b);
          return;
        }
      } else next = formatRange(next, a, b, runPart);
    }
    if (Object.keys(paraPart).length) next = formatParagraphsIn(next, a, b, paraPart);
    set(next, a, b);
  }, [set, place, onFormat, base]);

  useImperativeHandle(ref, () => ({
    format,
    commit,
    get paragraphs() { return model.current; },
    get selection() { return { ...sel.current }; },
    /** Words put in where the caret is (Insert Symbol), with the look there. */
    insert(text) {
      const { start, end } = sel.current;
      set(replaceRange(model.current, start, end, text), start + text.length);
    },
  }), [format, commit, set]);

  useEffect(() => {
    const onSelection = () => {
      if (composing.current || !el.current) return;
      const s = window.getSelection();
      if (!s?.rangeCount || !el.current.contains(s.anchorNode)) return;
      const before = sel.current;
      const now = read();
      if (before.start !== now.start || before.end !== now.end) pending.current = null;
      report();
    };
    document.addEventListener('selectionchange', onSelection);
    // A click anywhere but the box, the ribbon, a menu or a pane ends the editing.
    const onDown = (e) => {
      if (!el.current || el.current.contains(e.target) || e.target.closest?.(KEEPS_EDITING)) return;
      commit();
    };
    document.addEventListener('pointerdown', onDown, true);
    return () => {
      document.removeEventListener('selectionchange', onSelection);
      document.removeEventListener('pointerdown', onDown, true);
    };
  }, [read, report, commit]);

  // The page's own beforeinput, which says what the key or the paste would
  // do (React's onBeforeInput is a stand-in that does not).
  const onBeforeInput = (e) => {
    const type = String(e.inputType || '');
    if (!type || type === 'insertCompositionText') return;
    e.preventDefault();
    const target = e.getTargetRanges?.()[0];
    let { start, end } = read();
    if (target) {
      start = offsetOf(el.current, target.startContainer, target.startOffset);
      end = offsetOf(el.current, target.endContainer, target.endOffset);
    }
    const typed = (text) => {
      const look = pending.current ? pending.current : null;
      pending.current = null;
      set(replaceRange(model.current, start, end, text, look), start + text.length, start + text.length, { typing: true });
    };
    if (type === 'insertText' || type === 'insertReplacementText') return typed(e.data ?? e.dataTransfer?.getData('text/plain') ?? '');
    if (type === 'insertParagraph') return typed('\n');
    if (type === 'insertLineBreak') return typed('\v');
    if (type === 'insertFromPaste' || type === 'insertFromDrop') {
      const text = (e.dataTransfer?.getData('text/plain') || '').replace(/\r\n?/g, '\n');
      return set(replaceRange(model.current, start, end, text, pending.current), start + text.length);
    }
    if (type.startsWith('delete')) {
      if (start === end) {
        if (type.endsWith('Backward') && start > 0) start -= 1;
        else if (type.endsWith('Forward')) end += 1;
      }
      const len = plainOf(model.current).length;
      return set(replaceRange(model.current, Math.max(0, start), Math.min(len, end), ''), Math.max(0, start), Math.max(0, start), { typing: true });
    }
    if (type === 'historyUndo' || type === 'historyRedo') return step(type === 'historyUndo' ? -1 : 1);
    if (type === 'formatBold') return format({ bold: 'toggle' });
    if (type === 'formatItalic') return format({ italic: 'toggle' });
    if (type === 'formatUnderline') return format({ underline: 'toggle' });
  };

  /**
   * Words that reached the page without passing through here (an input
   * method's, a spelling suggestion, a script's execCommand) read back from
   * it, in the look of the run they went into, and the box drawn again.
   */
  const syncFromPage = () => {
    const text = plainOfPage(el.current);
    const was = plainOf(model.current);
    if (text === was) return;
    let a = 0;
    while (a < was.length && a < text.length && was[a] === text[a]) a++;
    let z = 0;
    while (z < was.length - a && z < text.length - a && was[was.length - 1 - z] === text[text.length - 1 - z]) z++;
    const inserted = text.slice(a, text.length - z);
    set(replaceRange(model.current, a, was.length - z, inserted), a + inserted.length, a + inserted.length, { typing: true });
  };

  const inputRef = useRef(null);
  inputRef.current = { before: onBeforeInput, after: () => { if (!composing.current) syncFromPage(); } };
  useEffect(() => {
    const node = el.current;
    const before = (e) => inputRef.current.before(e);
    const after = () => inputRef.current.after();
    node.addEventListener('beforeinput', before);
    node.addEventListener('input', after);
    return () => {
      node.removeEventListener('beforeinput', before);
      node.removeEventListener('input', after);
    };
  }, []);

  const step = (by) => {
    const h = history.current;
    const at = h.at + by;
    if (at < 0 || at >= h.list.length) return;
    h.at = at;
    h.typedAt = 0;
    const snap = h.list[at];
    model.current = snap.paragraphs;
    draw(el.current, snap.paragraphs);
    place(snap.sel.start, snap.sel.end);
    report();
  };

  const onKeyDown = (e) => {
    const mod = e.ctrlKey || e.metaKey;
    // Esc leaves the box with the words as typed, as PowerPoint does.
    if (e.key === 'Escape') { e.preventDefault(); commit(); return; }
    if (mod && !e.shiftKey && (e.key === 'z' || e.key === 'Z')) { e.preventDefault(); step(-1); return; }
    if (mod && (e.key === 'y' || e.key === 'Y' || (e.shiftKey && (e.key === 'z' || e.key === 'Z')))) { e.preventDefault(); step(1); return; }
    if (mod && !e.altKey && ['b', 'i', 'u'].includes(e.key.toLowerCase())) {
      e.preventDefault();
      e.stopPropagation();
      format({ [{ b: 'bold', i: 'italic', u: 'underline' }[e.key.toLowerCase()]]: 'toggle' });
      return;
    }
    onKey?.(e, { commit });
  };

  const onBlur = (e) => {
    if (done.current) return;
    if (e.relatedTarget && (e.relatedTarget.closest?.(KEEPS_EDITING) || el.current?.contains(e.relatedTarget))) return;
    // Focus gone to nothing in particular (another window): the editing stays open until a click elsewhere.
    if (!e.relatedTarget) return;
    commit();
  };

  return (
    <div
      ref={el}
      className={className}
      contentEditable
      suppressContentEditableWarning
      role="textbox"
      aria-multiline="true"
      spellCheck
      dir={dir}
      // The words stand where they stand in the box: at its top, middle or foot.
      style={{ display: 'flex', flexDirection: 'column', justifyContent: { middle: 'center', bottom: 'flex-end' }[anchor] || 'flex-start', whiteSpace: 'pre-wrap', fontSize: base.size ? `${base.size}px` : undefined, color: base.color || undefined, fontFamily: base.font ? `'${base.font}', 'Segoe UI', sans-serif` : undefined, ...style }}
      onKeyDown={onKeyDown}
      onBlur={onBlur}
      onCompositionStart={() => { composing.current = true; }}
      onCompositionEnd={() => {
        composing.current = false;
        syncFromPage();
      }}
    />
  );
});
