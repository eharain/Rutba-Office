// Rutba Word's side of Review → Check Accessibility and Spelling: where a
// finding or a misspelt word is on the page, the fixes that are ordinary
// edits (a style, a colour), and the right-click menu's suggestions.
// The panes, the pass and the dialogs are renderer/review.js.

import { useCallback, useMemo } from 'react';
import { useReview } from '../../review.js';

// The page and a pane side by side: the window is a column (ruler over the
// page) until a pane stands at its right.
if (typeof document !== 'undefined' && !document.getElementById('rutba-word-review-css')) {
  const style = document.createElement('style');
  style.id = 'rutba-word-review-css';
  style.textContent = '.wd:has(> .rw-panel.right) { flex-direction: row; } .wd > .rw-panel.right { height: 100%; }';
  document.head.appendChild(style);
}

const STORY_LABEL ={ footnote: 'Footnote', endnote: 'Endnote', header: 'Header', footer: 'Footer', textbox: 'Text box' };

/** The word under a point on the page: its text and its block, or null. */
function wordAtPoint(x, y) {
  const range = document.caretRangeFromPoint?.(x, y);
  const node = range?.startContainer;
  if (!node || node.nodeType !== Node.TEXT_NODE) return null;
  const text = node.nodeValue || '';
  let a = range.startOffset;
  let b = range.startOffset;
  const inWord = (c) => /[\p{L}\p{M}'’]/u.test(c || '');
  while (a > 0 && inWord(text[a - 1])) a -= 1;
  while (b < text.length && inWord(text[b])) b += 1;
  const word = text.slice(a, b).replace(/^['’]+|['’]+$/g, '');
  if (!word || !/\p{L}/u.test(word)) return null;
  const blockEl = node.parentElement?.closest?.('[data-block]');
  if (!blockEl) return null;
  // Roughly where in the paragraph: the text of the part before this node.
  let before = 0;
  const walker = document.createTreeWalker(blockEl, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n && n !== node; n = walker.nextNode()) before += n.nodeValue.length;
  return { word, block: Number(blockEl.dataset.block), near: before + a + Number(blockEl.dataset.from || 0) };
}

/**
 * Select characters `[from, to)` of a paragraph on the page, the way the
 * Editor pane shows the word it is asking about. The engine's selection is
 * set too (by the caller); this is the highlight a person sees, which a
 * selection op alone does not move.
 */
function selectInPage(page, block, from, to) {
  const parts = [...(page?.querySelectorAll(`[data-block="${block}"]`) || [])];
  const part = parts.filter((el) => Number(el.dataset.from || 0) <= from).pop() || parts[0];
  if (!part) return false;
  const base = Number(part.dataset.from || 0);
  const point = (offset) => {
    const walker = document.createTreeWalker(part, NodeFilter.SHOW_TEXT);
    let seen = 0;
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      if (seen + n.nodeValue.length >= offset - base) return { node: n, offset: Math.max(0, offset - base - seen) };
      seen += n.nodeValue.length;
    }
    return null;
  };
  const a = point(from);
  const b = point(to);
  if (!a || !b) return false;
  const range = document.createRange();
  range.setStart(a.node, a.offset);
  range.setEnd(b.node, b.offset);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  part.scrollIntoView({ block: 'center', behavior: 'smooth' });
  return true;
}

/** The offset of `word` in the paragraph's own text nearest to where it was seen. */
function offsetOf(text, word, near) {
  let best = -1;
  const re = new RegExp(`(?<![\\p{L}\\p{M}])${word.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?![\\p{L}\\p{M}])`, 'gu');
  for (const m of String(text).matchAll(re)) if (best < 0 || Math.abs(m.index - near) < Math.abs(best - near)) best = m.index;
  return best;
}

export function useWordReview({ shell, doc, model, apply, toast, pageRef, setPicked, view, patchView, menu }) {
  const scrollTo = useCallback((block) => {
    requestAnimationFrame(() => pageRef.current?.querySelector(`[data-block="${block}"]`)?.scrollIntoView({ block: 'center', behavior: 'smooth' }));
  }, [pageRef]);

  const adapter = useMemo(() => ({
    async goTo(where) {
      if (where?.block == null) return;
      await apply({ op: 'setSelection', anchor: { block: where.block, offset: 0 }, focus: { block: where.block, offset: 0 } });
      if (where.image != null) setPicked?.({ block: where.block, image: where.image });
      scrollTo(where.block);
    },
    async altText(target, props) {
      await apply({ op: 'setAltText', ...target, descr: props.descr, decorative: props.decorative });
    },
    async fix(issue, f) {
      const block = issue.where?.block;
      switch (f.kind) {
        case 'tableHeader':
          return apply({ op: 'setHeaderRow', table: f.target.table });
        case 'removeBlanks':
          return apply({ op: 'removeEmptyParagraphs', blocks: f.target.blocks });
        case 'headingLevel':
          return apply({ op: 'setSelection', anchor: { block, offset: 0 }, focus: { block, offset: 0 } }, { op: 'setParagraphFormat', delta: { styleId: `Heading${f.level}` } });
        case 'normalStyle':
          return apply({ op: 'setSelection', anchor: { block, offset: 0 }, focus: { block, offset: 0 } }, { op: 'setParagraphFormat', delta: { styleId: 'Normal' } });
        case 'textColour':
          return apply({ op: 'setSelection', anchor: { block, offset: 0 }, focus: { block, offset: f.target.length } }, { op: 'setRunFormat', delta: { fontColour: f.colour.replace('#', '') } });
        default:
          return null;
      }
    },
    async showWord(found) {
      const w = found.where || {};
      if (w.story === 'body') {
        await apply({ op: 'setSelection', anchor: { block: w.block, offset: found.offset }, focus: { block: w.block, offset: found.offset + found.length } });
        requestAnimationFrame(() => {
          if (!selectInPage(pageRef.current, w.block, found.offset, found.offset + found.length)) scrollTo(w.block);
        });
      } else if (w.block != null) {
        scrollTo(w.block);
      }
    },
    whereLabel(found) {
      const w = found.where || {};
      return w.story && w.story !== 'body' ? STORY_LABEL[w.story] || '' : '';
    },
    asYouType: { on: view?.spell !== false, set: (on) => patchView?.({ spell: on }) },
  }), [apply, setPicked, scrollTo, view?.spell, patchView]);

  const review = useReview({ shell, doc, model, apply, toast, adapter });

  /**
   * The page's right-click: on a picture, Edit Alt Text; on a word our
   * dictionary does not know, its suggestions, Ignore All and Add to
   * Dictionary above the usual verbs.
   */
  const contextMenu = useCallback(async (e, base) => {
    e.preventDefault();
    const at = { clientX: e.clientX, clientY: e.clientY, preventDefault() {}, stopPropagation() {} };
    const pic = e.target?.closest?.('[data-image]');
    const blockEl = e.target?.closest?.('[data-block]');
    if (pic && blockEl) {
      const target = { block: Number(blockEl.dataset.block), image: Number(pic.dataset.image) };
      menu.open(at, [{ label: 'Edit Alt Text…', icon: 'textbox', run: () => review.openAltText(target) }, '-', ...base]);
      return;
    }
    const hit = window.getSelection()?.isCollapsed === false ? null : wordAtPoint(e.clientX, e.clientY);
    const text = hit ? model?.blocks?.[hit.block]?.text : null;
    const offset = hit && text != null ? offsetOf(text, hit.word, hit.near) : -1;
    if (!hit || offset < 0) {
      menu.open(at, base);
      return;
    }
    const answer = await Promise.race([review.checkWord(hit.word), new Promise((r) => setTimeout(() => r(null), 2500))]);
    if (!answer?.misspelt) {
      menu.open(at, base);
      return;
    }
    const key = `b:${hit.block}`;
    const replace = (s) => apply({ op: 'spellReplace', edits: [{ key, from: offset, to: offset + hit.word.length, text: s }] });
    const suggestions = (answer.suggestions || []).slice(0, 5);
    menu.open(at, [
      ...(suggestions.length ? suggestions.map((s) => ({ label: s, run: () => replace(s) })) : [{ label: '(No spelling suggestions)', disabled: true, run: () => {} }]),
      '-',
      { label: 'Ignore All', run: async () => { await review.ignoreWord(hit.word); } },
      { label: 'Add to Dictionary', icon: 'plus', run: () => review.addWord(hit.word) },
      '-',
      ...base,
    ]);
  }, [menu, model, apply, review]);

  return { ...review, contextMenu };
}
