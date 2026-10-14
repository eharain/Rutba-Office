// Documents: forms. The Developer tab's content controls (rich text, plain
// text, picture, check box, combo box, drop-down list, date picker) and
// legacy form fields, Design Mode and Properties; each field drawn in its
// paragraph as Word draws it, its words or its placeholder, a box Word's
// way round it when the pointer or the keyboard is on it, a list or a
// calendar dropping from it, and filled in place. Restrict Editing's
// Filling in forms lets nothing else take input, and Tab walks the fields.
//
// A field is one character of its paragraph (the engine's forms.js): its
// host holds that character for the caret's arithmetic and draws everything
// else in a shadow root, as an equation does, so nothing the field shows is
// counted as the paragraph's words.

import React, { createContext, useCallback, useContext, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { Button, Group, Rows, Dialog, Field, Input, Select, Icon, t } from '@rutba/office-ui';

export const FormsContext = createContext(null);

/** The words a kind of field shows until it is filled, in the window's language, written into the file as Word writes its own. */
export const placeholderWords = (kind) => (kind === 'comboBox' || kind === 'dropDownList' ? t('Choose an item.') : kind === 'date' ? t('Click or tap to enter a date.') : t('Click or tap here to enter text.'));

const KIND_NAMES = {
  richText: () => t('Rich Text Content Control'),
  plainText: () => t('Plain Text Content Control'),
  picture: () => t('Picture Content Control'),
  checkBox: () => t('Check Box Content Control'),
  comboBox: () => t('Combo Box Content Control'),
  dropDownList: () => t('Drop-Down List Content Control'),
  date: () => t('Date Picker Content Control'),
  formText: () => t('Text Form Field'),
  formCheckBox: () => t('Check Box Form Field'),
  formDropDown: () => t('Drop-Down Form Field'),
  other: () => t('Content Control'),
};
export const kindName = (kind) => (KIND_NAMES[kind] || KIND_NAMES.other)();

const TEXT_KINDS = new Set(['richText', 'plainText', 'comboBox', 'formText']);
const LIST_KINDS = new Set(['comboBox', 'dropDownList', 'formDropDown']);

/** Word's date formats, its default first (as forms.js lists them). */
const DATE_FORMATS = ['dd/MM/yyyy', 'd/M/yyyy', 'dd-MM-yy', 'yyyy-MM-dd', 'd MMMM yyyy', 'dddd, d MMMM yyyy', 'MMMM d, yyyy', 'M/d/yyyy', 'd-MMM-yy', 'MMMM yy'];

/* ── a field on the page ─────────────────────────────────────────────────── */

const SHADOW_CSS = `
:host { display: inline; }
.box { position: relative; display: inline; border-radius: 2px; padding: 0 1px; outline: 1px solid transparent; outline-offset: 1px; transition: outline-color 80ms ease; }
:host(:hover) .box, :host(.on) .box, :host(:focus) .box { outline-color: #a6b6cc; background: rgba(166, 182, 204, 0.10); }
:host(.on) .box, :host(:focus) .box { outline-color: #7f97b8; }
.ph { color: #808080; }
.title { position: absolute; left: -2px; bottom: calc(100% + 1px); display: none; white-space: nowrap; font: 600 10.5px/16px "Segoe UI", system-ui, sans-serif; color: #2e4a72; background: #e4ebf5; border: 1px solid #a6b6cc; border-bottom: 0; border-radius: 3px 3px 0 0; padding: 0 6px; pointer-events: none; }
:host(.on) .title.has, :host(:focus) .title.has { display: inline-block; }
.drop { display: none; margin-left: 3px; padding: 0 3px; border-left: 1px solid #a6b6cc; color: #2e4a72; font: 10px/1 "Segoe UI", system-ui, sans-serif; vertical-align: middle; cursor: pointer; }
:host(.on) .drop, :host(:focus) .drop, :host(:hover) .drop { display: inline-block; }
.glyph { font-family: "MS Gothic", "Segoe UI Symbol", "Segoe UI Emoji", sans-serif; cursor: pointer; }
.legacy { background: #d9d9d9; outline: 0; }
:host(.noshade) .legacy { background: transparent; }
.lbox { display: inline-block; width: 0.72em; height: 0.72em; border: 1px solid currentColor; vertical-align: -0.06em; position: relative; cursor: pointer; background: #fff; }
.lbox.on::before, .lbox.on::after { content: ''; position: absolute; left: 50%; top: -12%; width: 1px; height: 124%; background: currentColor; transform: rotate(45deg); }
.lbox.on::after { transform: rotate(-45deg); }
.edit { outline: none; white-space: pre-wrap; caret-color: #1f4e8c; min-width: 1ch; display: inline-block; }
.tag { display: none; font: 600 9.5px/14px "Segoe UI", system-ui, sans-serif; color: #fff; background: #5a7bb0; padding: 0 5px; vertical-align: 1px; white-space: nowrap; min-width: 6px; min-height: 14px; box-sizing: border-box; }
.tag.start { border-radius: 3px 0 0 3px; margin-right: 2px; }
.tag.end { border-radius: 0 3px 3px 0; margin-left: 2px; }
:host(.design) .tag { display: inline-block; }
:host(.design) .title { display: none !important; }
img { vertical-align: bottom; display: inline-block; }
.picture { display: inline-block; position: relative; cursor: pointer; }
.picture .hint { position: absolute; left: 0; right: 0; bottom: 6px; text-align: center; font: 11px/1.3 "Segoe UI", system-ui, sans-serif; color: #555; }
.locked { cursor: default; }
`;

/** The page's look for a field's words, from its run. */
function lookCss(look = {}) {
  return [
    look.bold ? 'font-weight:700' : '', look.italic ? 'font-style:italic' : '', look.underline ? 'text-decoration:underline' : '',
    look.colour ? `color:#${look.colour}` : '', look.sizePt ? `font-size:${look.sizePt}pt` : '', look.font ? `font-family:"${look.font}"` : '',
  ].filter(Boolean).join(';');
}

const escHtml = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** What a field draws, as markup for its shadow root. */
function fieldMarkup(c, { design }) {
  const title = c.title || '';
  // Design Mode's tags, as Word draws them: the title at both ends, empty tags for a control with none; a legacy field has none.
  const tags = (inner) => (design && !c.legacy ? `<span class="tag start">${escHtml(title)}</span>${inner}<span class="tag end">${escHtml(title)}</span>` : inner);
  const head = `<span class="title${title && !c.legacy ? ' has' : ''}">${escHtml(title)}</span>`;
  if (c.kind === 'checkBox') {
    const glyph = String.fromCodePoint(parseInt((c.checked ? c.checkedState?.val : c.uncheckedState?.val) || (c.checked ? '2612' : '2610'), 16));
    return `<span class="box" part="box">${head}${tags(`<span class="glyph" style="${lookCss(c.look)}">${escHtml(glyph)}</span>`)}</span>`;
  }
  if (c.kind === 'formCheckBox') return `<span class="box legacy">${tags(`<span class="lbox${c.checked ? ' on' : ''}"></span>`)}</span>`;
  if (c.kind === 'picture') {
    const w = Math.round(c.widthPx || 120);
    const h = Math.round(c.heightPx || 120);
    return `<span class="box">${head}${tags(`<span class="picture">${c.href ? `<img src="${escHtml(c.href)}" width="${w}" height="${h}" alt="">` : ''}${c.value ? '' : `<span class="hint">${escHtml(t('Click or tap to insert a picture.'))}</span>`}</span>`)}</span>`;
  }
  const empty = c.legacy ? false : c.showingPlaceholder;
  const words = c.legacy && c.kind === 'formText' ? (c.text || '     ') : empty ? (c.placeholder || c.text || placeholderWords(c.kind)) : (c.text || '');
  const drop = LIST_KINDS.has(c.kind) || c.kind === 'date' ? `<span class="drop" data-drop="1">▼</span>` : '';
  return `<span class="box${c.legacy ? ' legacy' : ''}">${head}${tags(`<span class="val${empty ? ' ph' : ''}" style="${lookCss(c.look)}">${escHtml(words).replace(/\n/g, '<br>')}</span>${drop}`)}</span>`;
}

/**
 * A form field in its paragraph: the one character it is in the text,
 * hidden, and what it shows drawn in a shadow root. Pressed, it does what
 * Word's does, a check box ticks, a list or a calendar drops, a picture
 * control asks for a picture, words are typed in place.
 */
export function ControlRun({ run, at }) {
  const ctx = useContext(FormsContext);
  const ref = useRef(null);
  const c = run.control || {};
  const [editing, setEditing] = useState(false);
  const design = Boolean(ctx?.design);
  const shading = ctx?.shading !== false;
  const markup = useMemo(() => fieldMarkup(c, { design }), [c, design]);
  const where = () => {
    const host = ref.current;
    const block = Number(host?.closest('[data-block]')?.dataset.block);
    return Number.isFinite(block) ? { block, offset: at } : null;
  };

  useLayoutEffect(() => {
    const host = ref.current;
    if (!host || editing) return;
    const root = host.shadowRoot || host.attachShadow({ mode: 'open' });
    const html = `<style>${SHADOW_CSS}</style>${markup}`;
    if (root.__html !== html) {
      root.innerHTML = html;
      root.__html = html;
    }
  }, [markup, editing]);

  // Typing in place: the words, or nothing for a placeholder, made editable and selected.
  const startEditing = useCallback(() => {
    const host = ref.current;
    if (!host || c.noEdit || !ctx?.canFill?.()) return;
    setEditing(true);
    const root = host.shadowRoot || host.attachShadow({ mode: 'open' });
    const value = c.legacy ? (c.value || '') : c.showingPlaceholder ? '' : c.text || '';
    root.innerHTML = `<style>${SHADOW_CSS}</style><span class="box"><span class="title${c.title && !c.legacy ? ' has' : ''}">${escHtml(c.title || '')}</span><span class="edit" contenteditable="plaintext-only" spellcheck="false" style="${lookCss(c.look)}">${escHtml(value)}</span></span>`;
    root.__html = null;
    host.classList.add('on');
    const edit = root.querySelector('.edit');
    let done = false;
    const finish = (commit, then = null) => {
      if (done) return;
      done = true;
      host.classList.remove('on');
      const text = (edit.innerText || '').replace(/\n$/, '');
      setEditing(false);
      if (commit && text !== value) ctx.fill(where(), c, text);
      if (then) then();
    };
    edit.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Escape') { e.preventDefault(); finish(false, () => ctx.toPage(host)); }
      else if (e.key === 'Tab') { e.preventDefault(); finish(true, () => ctx.next(where(), e.shiftKey ? -1 : 1)); }
      else if (e.key === 'Enter' && !(c.kind === 'plainText' && c.multiLine) && c.kind !== 'richText') { e.preventDefault(); finish(true, () => ctx.toPage(host)); }
    });
    edit.addEventListener('blur', () => finish(true));
    for (const type of ['keypress', 'keyup', 'input', 'beforeinput', 'paste', 'cut', 'copy']) edit.addEventListener(type, (e) => e.stopPropagation());
    edit.focus();
    const range = document.createRange();
    range.selectNodeContents(edit);
    const sel = root.getSelection ? root.getSelection() : window.getSelection();
    sel?.removeAllRanges?.();
    sel?.addRange?.(range);
  }, [c, ctx]);

  // What a press does, by kind.
  const activate = useCallback((e, how = 'press') => {
    const host = ref.current;
    if (!ctx || !host) return;
    const spot = where();
    if (!spot) return;
    ctx.focused(spot, c);
    if (design && how === 'press') { host.focus({ preventScroll: true }); return; }
    if (c.kind === 'checkBox' || c.kind === 'formCheckBox') { ctx.fill(spot, c, !c.checked); host.focus({ preventScroll: true }); return; }
    const onDrop = how === 'press' && e?.composedPath?.().some((n) => n?.dataset?.drop);
    if (c.kind === 'comboBox' && !onDrop && how === 'press') { startEditing(); return; }
    if (LIST_KINDS.has(c.kind) || c.kind === 'date') { host.focus({ preventScroll: true }); ctx.drop(host, spot, c); return; }
    if (c.kind === 'picture') { host.focus({ preventScroll: true }); ctx.pickPicture(spot, c); return; }
    if (TEXT_KINDS.has(c.kind)) { startEditing(); return; }
    host.focus({ preventScroll: true });
  }, [c, ctx, design, startEditing]);

  // Tab and Shift+Tab, and Restrict Editing's walk, arrive here.
  useEffect(() => {
    const host = ref.current;
    if (!host) return undefined;
    const on = () => activate(null, 'tab');
    host.addEventListener('rutba-field', on);
    return () => host.removeEventListener('rutba-field', on);
  }, [activate]);

  return (
    <span
      ref={ref}
      className={`wd-cc${design ? ' design' : ''}${shading ? '' : ' noshade'}`}
      data-at={at ?? undefined}
      data-kind={c.kind}
      data-cc-title={c.title || undefined}
      data-cc-tag={c.tag || undefined}
      data-value={c.kind === 'checkBox' || c.kind === 'formCheckBox' ? String(Boolean(c.checked)) : c.showingPlaceholder ? '' : String(c.value ?? '')}
      data-placeholder={c.showingPlaceholder ? '1' : undefined}
      data-tip={c.legacy ? kindName(c.kind) + (c.title ? ': ' + c.title : '') : c.title ? c.title + ': ' + kindName(c.kind) : kindName(c.kind)}
      contentEditable={false}
      suppressContentEditableWarning
      tabIndex={-1}
      role={c.kind === 'checkBox' || c.kind === 'formCheckBox' ? 'checkbox' : LIST_KINDS.has(c.kind) ? 'combobox' : 'textbox'}
      aria-checked={c.kind === 'checkBox' || c.kind === 'formCheckBox' ? Boolean(c.checked) : undefined}
      aria-label={c.title || kindName(c.kind)}
      onMouseDown={(e) => { if (e.button !== 0) return; e.preventDefault(); e.stopPropagation(); }}
      onMouseUp={(e) => { e.stopPropagation(); }}
      onClick={(e) => { if (e.button !== 0 || editing) return; e.preventDefault(); e.stopPropagation(); activate(e.nativeEvent, 'press'); }}
      onDoubleClick={(e) => { if (design) { e.stopPropagation(); ctx?.properties(where(), c); } }}
      onKeyDown={(e) => {
        if (editing) return;
        e.stopPropagation();
        const spot = where();
        if (e.key === 'Tab') { e.preventDefault(); ctx?.next(spot, e.shiftKey ? -1 : 1); return; }
        if (e.key === 'Escape') { e.preventDefault(); ctx?.toPage(ref.current); return; }
        if (e.key === ' ' && (c.kind === 'checkBox' || c.kind === 'formCheckBox')) { e.preventDefault(); ctx?.fill(spot, c, !c.checked); return; }
        if (e.key === 'Enter' || e.key === ' ' || (e.altKey && e.key === 'ArrowDown') || e.key === 'F4') { e.preventDefault(); activate(null, 'tab'); return; }
        if ((e.key === 'Delete' || e.key === 'Backspace') && ctx?.canFill?.() && !c.legacy && c.kind !== 'checkBox' && c.kind !== 'picture') { e.preventDefault(); ctx?.fill(spot, c, ''); }
      }}
    >
      {run.text}
    </span>
  );
}

/* ── the list and the calendar that drop from a field ────────────────────── */

const POPUP_CSS = `
.wd-cc-pop { position: fixed; z-index: 60; background: var(--surface, #fff); border: 1px solid var(--line, #c8ccd2); border-radius: 6px; box-shadow: 0 8px 24px rgba(15, 23, 42, 0.18), 0 1px 3px rgba(15, 23, 42, 0.12); font: 13px/1.35 var(--font, "Segoe UI", system-ui, sans-serif); color: var(--ink, #1b1f24); overflow: hidden; }
.wd-cc-list { min-width: 160px; max-height: 260px; overflow: auto; padding: 4px; }
.wd-cc-item { display: block; width: 100%; text-align: start; border: 0; background: transparent; padding: 6px 10px; border-radius: 4px; font: inherit; color: inherit; cursor: pointer; white-space: nowrap; }
.wd-cc-item:hover, .wd-cc-item.hot { background: var(--hover, #eef2f7); }
.wd-cc-item.on { background: var(--accent-soft, #dbe7fb); color: var(--accent-ink, #12407d); font-weight: 600; }
.wd-cc-cal { width: 252px; padding: 10px 10px 8px; }
.wd-cc-cal-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 6px; }
.wd-cc-cal-head strong { font-size: 13px; }
.wd-cc-cal-head button { border: 0; background: transparent; width: 28px; height: 26px; border-radius: 5px; cursor: pointer; color: var(--ink-2, #4a5361); display: grid; place-items: center; }
.wd-cc-cal-head button:hover { background: var(--hover, #eef2f7); }
.wd-cc-cal-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 2px; }
.wd-cc-cal-dow { text-align: center; font-size: 11px; font-weight: 600; color: var(--ink-3, #7a8392); padding: 2px 0 4px; }
.wd-cc-day { height: 30px; border: 0; border-radius: 5px; background: transparent; font: inherit; color: inherit; cursor: pointer; }
.wd-cc-day:hover { background: var(--hover, #eef2f7); }
.wd-cc-day.out { color: var(--ink-3, #9aa1ac); }
.wd-cc-day.today { box-shadow: inset 0 0 0 1px var(--accent, #2b5fd9); }
.wd-cc-day.on { background: var(--accent, #2b5fd9); color: #fff; font-weight: 600; }
.wd-cc-cal-foot { display: flex; justify-content: center; margin-top: 8px; border-top: 1px solid var(--line-soft, #e3e6ea); padding-top: 6px; }
.wd-cc-cal-foot button { border: 0; background: transparent; color: var(--accent, #2b5fd9); font: inherit; font-weight: 600; cursor: pointer; padding: 4px 12px; border-radius: 5px; }
.wd-cc-cal-foot button:hover { background: var(--hover, #eef2f7); }
.wd-cc-props-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px 14px; }
.wd-cc-props-section { margin-top: 14px; padding-top: 10px; border-top: 1px solid var(--line-soft, #e3e6ea); }
.wd-cc-props-section h4 { margin: 0 0 8px; font-size: 12px; font-weight: 600; color: var(--ink-2, #4a5361); }
.wd-cc-check { display: flex; align-items: center; gap: 8px; font-size: 12.5px; margin: 5px 0; cursor: pointer; }
.wd-cc-items { width: 100%; border: 1px solid var(--line, #c8ccd2); border-radius: 6px; max-height: 150px; overflow: auto; }
.wd-cc-items-row { display: grid; grid-template-columns: 1fr 1fr 28px; gap: 6px; padding: 4px 6px; border-bottom: 1px solid var(--line-soft, #e3e6ea); align-items: center; }
.wd-cc-items-row:last-child { border-bottom: 0; }
.wd-cc-items-row.head { font-size: 11px; font-weight: 600; color: var(--ink-3, #7a8392); background: var(--surface-2, #f6f7f9); }
.wd-cc-items-row input { width: 100%; box-sizing: border-box; }
.wd-cc-items-add { margin-top: 6px; }
.wd-cc-kind { display: inline-flex; align-items: center; gap: 6px; font-size: 12px; color: var(--ink-2, #4a5361); margin-bottom: 10px; padding: 3px 9px; border-radius: 999px; background: var(--surface-2, #f1f3f6); }
`;

let popupCssInstalled = false;
function installPopupCss() {
  if (popupCssInstalled || typeof document === 'undefined') return;
  popupCssInstalled = true;
  const style = document.createElement('style');
  style.id = 'rutba-word-forms-css';
  style.textContent = POPUP_CSS;
  document.head.appendChild(style);
}

/** Close on a press elsewhere, Escape or a scroll of the page. */
function useDismiss(ref, onClose) {
  useEffect(() => {
    const off = (e) => { if (!ref.current?.contains(e.target)) onClose(); };
    const key = (e) => { if (e.key === 'Escape') { e.stopPropagation(); onClose(true); } };
    const scroll = (e) => { if (!ref.current?.contains(e.target)) onClose(); };
    document.addEventListener('mousedown', off, true);
    window.addEventListener('keydown', key, true);
    window.addEventListener('scroll', scroll, true);
    return () => {
      document.removeEventListener('mousedown', off, true);
      window.removeEventListener('keydown', key, true);
      window.removeEventListener('scroll', scroll, true);
    };
  }, [ref, onClose]);
}

/** Where a popup goes: under the field, inside the window. */
function placeUnder(rect, width, height) {
  const left = Math.max(8, Math.min(rect.left, window.innerWidth - width - 8));
  const below = rect.bottom + 3;
  const top = below + height > window.innerHeight - 8 ? Math.max(8, rect.top - height - 3) : below;
  return { left, top };
}

function ListPopup({ rect, items, current, onPick, onClose }) {
  const ref = useRef(null);
  const [hot, setHot] = useState(Math.max(0, items.findIndex((it) => it.text === current)));
  useDismiss(ref, onClose);
  useEffect(() => {
    const key = (e) => {
      if (e.key === 'ArrowDown') { e.preventDefault(); e.stopPropagation(); setHot((h) => Math.min(items.length - 1, h + 1)); }
      else if (e.key === 'ArrowUp') { e.preventDefault(); e.stopPropagation(); setHot((h) => Math.max(0, h - 1)); }
      else if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); if (items[hot]) onPick(items[hot]); }
    };
    window.addEventListener('keydown', key, true);
    return () => window.removeEventListener('keydown', key, true);
  }, [items, hot, onPick]);
  const width = Math.max(160, Math.round(rect.width) + 24);
  const at = placeUnder(rect, width, Math.min(260, items.length * 31 + 10));
  return createPortal(
    <div className="wd-cc-pop wd-cc-list" ref={ref} role="listbox" style={{ left: at.left, top: at.top, minWidth: width }}>
      {items.map((it, i) => (
        <button key={i} type="button" role="option" aria-selected={it.text === current} className={`wd-cc-item${it.text === current ? ' on' : ''}${i === hot ? ' hot' : ''}`} data-item={it.text} onMouseEnter={() => setHot(i)} onMouseDown={(e) => e.preventDefault()} onClick={() => onPick(it)}>
          {it.text || '\u00A0'}
        </button>
      ))}
    </div>,
    document.body,
  );
}

const pad = (n) => String(n).padStart(2, '0');
const isoOf = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

function CalendarPopup({ rect, value, onPick, onClose }) {
  const ref = useRef(null);
  const start = value && /^\d{4}-\d{2}-\d{2}$/.test(value) ? new Date(Number(value.slice(0, 4)), Number(value.slice(5, 7)) - 1, 1) : new Date(new Date().getFullYear(), new Date().getMonth(), 1);
  const [month, setMonth] = useState(start);
  useDismiss(ref, onClose);
  const today = isoOf(new Date());
  const first = new Date(month.getFullYear(), month.getMonth(), 1);
  // Weeks start on the window's language's own first day.
  let firstDay = 1;
  try { firstDay = (new Intl.Locale(document.documentElement.lang || 'en-GB').weekInfo?.firstDay ?? 1) % 7; } catch { /* an older engine */ }
  const lead = (first.getDay() - firstDay + 7) % 7;
  const days = Array.from({ length: 42 }, (_, i) => new Date(month.getFullYear(), month.getMonth(), 1 - lead + i));
  const lang = document.documentElement.lang || undefined;
  const title = new Intl.DateTimeFormat(lang, { month: 'long', year: 'numeric' }).format(month);
  const dow = Array.from({ length: 7 }, (_, i) => new Intl.DateTimeFormat(lang, { weekday: 'narrow' }).format(new Date(2024, 0, 7 + ((firstDay + i) % 7))));
  const at = placeUnder(rect, 252, 300);
  return createPortal(
    <div className="wd-cc-pop wd-cc-cal" ref={ref} role="dialog" aria-label={t('Choose a date')} style={{ left: at.left, top: at.top }}>
      <div className="wd-cc-cal-head">
        <button type="button" className="wd-cc-cal-prev" aria-label={t('Previous month')} onMouseDown={(e) => e.preventDefault()} onClick={() => setMonth(new Date(month.getFullYear(), month.getMonth() - 1, 1))}><Icon name="chevronLeft" size={14} /></button>
        <strong className="wd-cc-cal-title">{title}</strong>
        <button type="button" className="wd-cc-cal-next" aria-label={t('Next month')} onMouseDown={(e) => e.preventDefault()} onClick={() => setMonth(new Date(month.getFullYear(), month.getMonth() + 1, 1))}><Icon name="chevronRight" size={14} /></button>
      </div>
      <div className="wd-cc-cal-grid">
        {dow.map((d, i) => <div key={'h' + i} className="wd-cc-cal-dow">{d}</div>)}
        {days.map((d) => {
          const iso = isoOf(d);
          return (
            <button key={iso} type="button" data-date={iso} className={`wd-cc-day${d.getMonth() !== month.getMonth() ? ' out' : ''}${iso === today ? ' today' : ''}${iso === value ? ' on' : ''}`} onMouseDown={(e) => e.preventDefault()} onClick={() => onPick(iso)}>
              {new Intl.NumberFormat(lang).format(d.getDate())}
            </button>
          );
        })}
      </div>
      <div className="wd-cc-cal-foot"><button type="button" className="wd-cc-cal-today" onMouseDown={(e) => e.preventDefault()} onClick={() => onPick(today)}>{t('Today')}</button></div>
    </div>,
    document.body,
  );
}

/* ── Properties ──────────────────────────────────────────────────────────── */

/** Developer → Properties, as Word's Content Control Properties box, with the placeholder's words beside its title. */
export function ControlPropertiesDialog({ control, onClose, onSave }) {
  const c = control;
  const [title, setTitle] = useState(c.title || '');
  const [tag, setTag] = useState(c.tag || '');
  const [placeholder, setPlaceholder] = useState(c.placeholder || '');
  const [noDelete, setNoDelete] = useState(Boolean(c.noDelete));
  const [noEdit, setNoEdit] = useState(Boolean(c.noEdit));
  const [multiLine, setMultiLine] = useState(Boolean(c.multiLine));
  // An item with no value of its own (Word's "Choose an item.") keeps none: picking it empties the control.
  const [items, setItems] = useState((c.items || []).map((it) => ({ ...it, keepEmpty: it.value === '' })));
  const [dateFormat, setDateFormat] = useState(c.dateFormat || DATE_FORMATS[0]);
  const list = LIST_KINDS.has(c.kind);
  const save = () => {
    const props = {};
    if (title !== (c.title || '')) props.title = title;
    if (!c.legacy) {
      if (tag !== (c.tag || '')) props.tag = tag;
      if (noDelete !== Boolean(c.noDelete) || noEdit !== Boolean(c.noEdit)) Object.assign(props, { noDelete, noEdit });
      if (c.kind !== 'checkBox' && c.kind !== 'picture' && placeholder && placeholder !== (c.placeholder || '')) props.placeholder = placeholder;
      if (c.kind === 'plainText' && multiLine !== Boolean(c.multiLine)) props.multiLine = multiLine;
      if (c.kind === 'date' && dateFormat !== c.dateFormat) props.dateFormat = dateFormat;
    }
    if (list) {
      const clean = items.filter((it) => (it.text || '').trim()).map((it) => ({ text: it.text.trim(), value: c.legacy ? it.text.trim() : it.keepEmpty && !it.value ? '' : (it.value || it.text).trim() }));
      if (JSON.stringify(clean) !== JSON.stringify((c.items || []).map((it) => ({ text: it.text, value: it.value })))) props.items = clean;
    }
    onSave(props);
  };
  return (
    <Dialog
      title={c.legacy ? t('Form Field Options') : t('Content Control Properties')}
      width={520}
      onClose={onClose}
      className="wd-cc-props"
      actions={<><Button label={t('Cancel')} onClick={onClose} /><Button primary className="wd-cc-props-ok" label={t('OK')} onClick={save} /></>}
    >
      <div className="wd-cc-kind"><Icon name={c.kind === 'date' ? 'calendar' : c.kind === 'picture' ? 'picture' : c.kind === 'checkBox' || c.kind === 'formCheckBox' ? 'check' : list ? 'list' : 'textbox'} size={13} />{kindName(c.kind)}</div>
      <div className="wd-cc-props-grid">
        <Field label={c.legacy ? t('Bookmark') : t('Title')}><Input className="wd-cc-title" value={title} onChange={(e) => setTitle(e.target.value)} autoFocus /></Field>
        {c.legacy ? <div /> : <Field label={t('Tag')}><Input className="wd-cc-tag" value={tag} onChange={(e) => setTag(e.target.value)} /></Field>}
      </div>
      {!c.legacy && c.kind !== 'checkBox' && c.kind !== 'picture' ? (
        <Field label={t('Placeholder text')} hint={t('Shown until the control is filled in.')}><Input className="wd-cc-placeholder" value={placeholder} onChange={(e) => setPlaceholder(e.target.value)} /></Field>
      ) : null}
      {!c.legacy ? (
        <div className="wd-cc-props-section">
          <h4>{t('Locking')}</h4>
          <label className="wd-cc-check"><input type="checkbox" className="wd-cc-nodelete" checked={noDelete} onChange={(e) => setNoDelete(e.target.checked)} />{t('Content control cannot be deleted')}</label>
          <label className="wd-cc-check"><input type="checkbox" className="wd-cc-noedit" checked={noEdit} onChange={(e) => setNoEdit(e.target.checked)} />{t('Contents cannot be edited')}</label>
        </div>
      ) : null}
      {c.kind === 'plainText' ? (
        <div className="wd-cc-props-section">
          <h4>{t('Plain Text Properties')}</h4>
          <label className="wd-cc-check"><input type="checkbox" className="wd-cc-multiline" checked={multiLine} onChange={(e) => setMultiLine(e.target.checked)} />{t('Allow carriage returns (multiple paragraphs)')}</label>
        </div>
      ) : null}
      {c.kind === 'date' ? (
        <div className="wd-cc-props-section">
          <h4>{t('Date Picker Properties')}</h4>
          <Field label={t('Display the date like this')}>
            <Select className="wd-cc-dateformat" value={dateFormat} onChange={(e) => setDateFormat(e.target.value)}>
              {[...new Set([dateFormat, ...DATE_FORMATS])].map((f) => <option key={f} value={f}>{f}</option>)}
            </Select>
          </Field>
        </div>
      ) : null}
      {list ? (
        <div className="wd-cc-props-section">
          <h4>{c.kind === 'comboBox' ? t('Combo Box Properties') : t('Drop-Down List Properties')}</h4>
          <div className="wd-cc-items">
            <div className="wd-cc-items-row head"><span>{t('Display Name')}</span><span>{c.legacy ? '' : t('Value')}</span><span /></div>
            {items.map((it, i) => (
              <div className="wd-cc-items-row" key={i}>
                <Input className="wd-cc-item-text" value={it.text} onChange={(e) => setItems(items.map((x, j) => (j === i ? { ...x, text: e.target.value } : x)))} />
                {c.legacy ? <span /> : <Input className="wd-cc-item-value" value={it.value ?? ''} onChange={(e) => setItems(items.map((x, j) => (j === i ? { ...x, value: e.target.value } : x)))} />}
                <Button icon="close" title={t('Remove')} onClick={() => setItems(items.filter((_, j) => j !== i))} />
              </div>
            ))}
          </div>
          <Button className="wd-cc-items-add wd-cc-add" icon="plus" label={t('Add')} onClick={() => setItems([...items, { text: '', value: '' }])} />
        </div>
      ) : null}
    </Dialog>
  );
}

/* ── the hook the Word window uses ───────────────────────────────────────── */

/**
 * Forms for one Word window: what Developer's buttons do, the context every
 * field on the page reads, the popups and the Properties box, and Tab in a
 * document protected for filling in forms.
 */
export function useWordForms({ shell, model, apply, toast, pageRef, view, patchView }) {
  installPopupCss();
  const modelRef = useRef(model);
  modelRef.current = model;
  const [popup, setPopup] = useState(null);
  const [props, setProps] = useState(null);
  const current = useRef(null);
  const protection = model?.protection || null;
  const formsOnly = Boolean(protection?.enforced && protection.edit === 'forms');
  const design = Boolean(view.designMode);

  /** Every field on the page, in document order, as the engine's blocks list them. */
  const fields = useCallback(() => {
    const out = [];
    (modelRef.current?.blocks || []).forEach((b, block) => {
      let at = 0;
      for (const r of b.runs || []) {
        if (r.control) out.push({ block, offset: at, control: r.control });
        at += (r.text || '').length;
      }
    });
    return out;
  }, []);

  const hostOf = useCallback((spot) => {
    const page = pageRef.current;
    if (!page || !spot) return null;
    for (const part of page.querySelectorAll(`[data-block="${spot.block}"]`)) {
      const host = part.querySelector(`.wd-cc[data-at="${spot.offset}"]`);
      if (host) return host;
    }
    return null;
  }, [pageRef]);

  const fill = useCallback(async (spot, control, value) => {
    if (!spot) return;
    if (control.noEdit) { toast(t('This content control cannot be edited: its contents are locked.'), { ms: 4000 }); return; }
    await apply({ op: 'setControlValue', block: spot.block, offset: spot.offset, value });
  }, [apply, toast]);

  const goTo = useCallback((spot) => {
    current.current = spot;
    const host = hostOf(spot);
    if (!host) return false;
    host.scrollIntoView({ block: 'nearest' });
    host.dispatchEvent(new CustomEvent('rutba-field'));
    return true;
  }, [hostOf]);

  const next = useCallback((spot, dir = 1) => {
    const all = fields().filter((f) => !f.control.noEdit || f.control.kind === 'checkBox');
    if (!all.length) return;
    const here = spot ? all.findIndex((f) => f.block === spot.block && f.offset === spot.offset) : -1;
    const at = here < 0 ? (dir > 0 ? 0 : all.length - 1) : (here + dir + all.length) % all.length;
    // The window draws the new value a beat after the engine answers: wait for the host.
    setTimeout(() => goTo(all[at]), 30);
  }, [fields, goTo]);

  const ctx = useMemo(() => ({
    design,
    shading: view.formShading !== false,
    canFill: () => true,
    fill,
    next,
    focused: (spot) => { current.current = spot; },
    toPage: (host) => {
      host?.blur?.();
      pageRef.current?.focus({ preventScroll: true });
    },
    drop: (host, spot, control) => {
      const rect = host.getBoundingClientRect();
      setPopup({ kind: control.kind === 'date' ? 'date' : 'list', rect: { left: rect.left, top: rect.top, bottom: rect.bottom, width: rect.width }, spot, control });
    },
    pickPicture: async (spot, control) => {
      if (control.noEdit) { toast(t('This content control cannot be edited: its contents are locked.'), { ms: 4000 }); return; }
      const [file] = await shell.dialog.open({ title: t('Insert picture'), filters: [{ name: t('Pictures'), extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp'] }] });
      if (!file) return;
      const { bytes } = await shell.fs.read({ path: file });
      const ext = file.split('.').pop().toLowerCase();
      const contentType = { png: 'image/png', gif: 'image/gif', webp: 'image/webp', bmp: 'image/bmp' }[ext] || 'image/jpeg';
      // The picture is fitted to the control's box, its proportions kept, as Word fits it.
      let w = control.widthPx || 160;
      let h = control.heightPx || 160;
      try {
        const url = URL.createObjectURL(new Blob([bytes]));
        const img = new Image();
        img.src = url;
        await img.decode();
        const k = Math.min(w / img.naturalWidth, h / img.naturalHeight);
        w = Math.max(1, Math.round(img.naturalWidth * k));
        h = Math.max(1, Math.round(img.naturalHeight * k));
        URL.revokeObjectURL(url);
      } catch { /* the box's own size, then */ }
      await apply({ op: 'setControlPicture', block: spot.block, offset: spot.offset, picture: { data: bytes, contentType, widthPx: w, heightPx: h } });
    },
    properties: (spot, control) => { current.current = spot; setProps({ spot, control }); },
  }), [design, view.formShading, fill, next, pageRef, shell, apply, toast]);

  /** Developer's buttons. */
  const act = useCallback(async (name, arg) => {
    switch (name) {
      case 'insert': {
        const legacy = arg.legacy || false;
        const props = legacy ? (arg.kind === 'formDropDown' ? { items: [] } : {}) : {
          placeholder: arg.kind === 'checkBox' || arg.kind === 'picture' ? undefined : placeholderWords(arg.kind),
          ...(arg.kind === 'comboBox' || arg.kind === 'dropDownList' ? { items: [{ text: placeholderWords(arg.kind), value: '' }] } : {}),
        };
        // Word's own default placeholders are written as its own glossary names them; the words go with them.
        if (!legacy && props.placeholder) props.defaultPlaceholder = true;
        await apply({ op: 'insertControl', kind: arg.kind, legacy, props: stripDefault(props) });
        return;
      }
      case 'designMode':
        patchView((v) => ({ designMode: !v.designMode }));
        return;
      case 'shading':
        patchView((v) => ({ formShading: v.formShading === false }));
        return;
      case 'properties': {
        const all = fields();
        const sel = modelRef.current?.selection;
        const near = current.current && all.find((f) => f.block === current.current.block && f.offset === current.current.offset)
          || (sel ? all.find((f) => f.block === sel.focus?.block && (f.offset === sel.focus.offset || f.offset + 1 === sel.focus.offset)) : null);
        if (!near) { toast(t('Select a content control first: click it, or put the caret beside it.'), { ms: 4000 }); return; }
        setProps({ spot: { block: near.block, offset: near.offset }, control: near.control });
        return;
      }
      case 'resetFields': {
        const legacy = fields().filter((f) => f.control.legacy);
        for (const f of legacy) await apply({ op: 'setControlValue', block: f.block, offset: f.offset, value: f.control.kind === 'formCheckBox' ? false : f.control.kind === 'formDropDown' ? f.control.items?.[0]?.text ?? '' : '' });
        return;
      }
      default:
    }
  }, [apply, fields, patchView, toast]);

  /** Tab in a document protected for filling in forms goes to the next field, as in Word. */
  const pageKey = useCallback((e) => {
    if (e.key !== 'Tab' || !formsOnly) return false;
    e.preventDefault();
    const sel = modelRef.current?.selection?.focus || { block: 0, offset: 0 };
    const all = fields();
    if (!all.length) return true;
    // From the caret: the first field after it, or with Shift the last before it, round the end.
    const after = (f) => f.block > sel.block || (f.block === sel.block && f.offset >= sel.offset);
    const target = e.shiftKey ? [...all].reverse().find((f) => !after(f)) || all[all.length - 1] : all.find(after) || all[0];
    setTimeout(() => goTo({ block: target.block, offset: target.offset }), 30);
    return true;
  }, [formsOnly, fields, goTo]);

  // A document opened protected for forms starts at its first field, as Word does.
  const started = useRef(false);
  useEffect(() => {
    if (!formsOnly) { started.current = false; return; }
    if (started.current || !fields().length) return;
    started.current = true;
    const first = fields()[0];
    current.current = { block: first.block, offset: first.offset };
  }, [formsOnly, fields, model]);

  const overlay = (
    <>
      {popup?.kind === 'list' ? (
        <ListPopup
          rect={popup.rect}
          items={popup.control.items || []}
          current={popup.control.showingPlaceholder ? null : popup.control.value}
          onClose={(viaKey) => { setPopup(null); if (viaKey) hostOf(popup.spot)?.focus({ preventScroll: true }); }}
          onPick={(it) => { const p = popup; setPopup(null); fill(p.spot, p.control, it.value === '' && p.control.kind !== 'formDropDown' ? '' : p.control.legacy ? it.text : it.value); }}
        />
      ) : null}
      {popup?.kind === 'date' ? (
        <CalendarPopup
          rect={popup.rect}
          value={popup.control.date || null}
          onClose={() => setPopup(null)}
          onPick={(iso) => { const p = popup; setPopup(null); fill(p.spot, p.control, iso); }}
        />
      ) : null}
      {props ? (
        <ControlPropertiesDialog
          control={props.control}
          onClose={() => setProps(null)}
          onSave={async (change) => {
            const p = props;
            setProps(null);
            if (Object.keys(change).length) await apply({ op: 'setControlProps', block: p.spot.block, offset: p.spot.offset, props: change });
          }}
        />
      ) : null}
    </>
  );

  return { ctx, act, overlay, pageKey, design, formsOnly };
}

/** Word's own default placeholder is not a placeholder of one's own: the engine writes its glossary name. */
function stripDefault(props) {
  const { defaultPlaceholder, ...rest } = props;
  return defaultPlaceholder ? { ...rest, placeholderDefault: rest.placeholder, placeholder: undefined } : rest;
}

/* ── the Developer tab ───────────────────────────────────────────────────── */

/** Developer: Controls (the seven content controls, Design Mode, Properties, Legacy Tools) and Protect. */
export function DeveloperTab({ forms, menu, act, view }) {
  const insert = (kind) => () => forms.act('insert', { kind });
  const legacy = (kind) => () => forms.act('insert', { kind, legacy: true });
  return (
    <>
      <Group label={t('Controls')}>
        <Rows>
          <>
            <Button icon="styles" title={t('Rich Text Content Control: words that can be formatted')} onClick={insert('richText')} />
            <Button icon="textbox" title={t('Plain Text Content Control: words in one look')} onClick={insert('plainText')} />
            <Button icon="picture" title={t('Picture Content Control: a picture chosen from a file')} onClick={insert('picture')} />
            <Button icon="check" title={t('Check Box Content Control: ticked or not')} onClick={insert('checkBox')} />
          </>
          <>
            <Button icon="list" title={t('Combo Box Content Control: one of a list, or words of one\'s own')} onClick={insert('comboBox')} />
            <Button icon="chevronDown" title={t('Drop-Down List Content Control: one of a list')} onClick={insert('dropDownList')} />
            <Button icon="calendar" title={t('Date Picker Content Control: a date from a calendar')} onClick={insert('date')} />
            <Button dropdown icon="grid" title={t('Legacy Tools: the form fields of Word 97 to 2003')} onClick={(e) => menu.open(e, [
              { heading: true, label: t('Legacy Forms') },
              { label: t('Text Form Field'), icon: 'textbox', run: legacy('formText') },
              { label: t('Check Box Form Field'), icon: 'check', run: legacy('formCheckBox') },
              { label: t('Drop-Down Form Field'), icon: 'list', run: legacy('formDropDown') },
              '-',
              { label: t('Form Field Shading'), icon: view.formShading !== false ? 'check' : undefined, run: () => forms.act('shading') },
              { label: t('Reset Form Fields'), icon: 'refresh', run: () => forms.act('resetFields') },
            ])} />
          </>
        </Rows>
        <Rows>
          <Button icon="pen" label={t('Design Mode')} pressed={forms.design} title={t('Design Mode: show each control\'s tags, and open its Properties with a double-click')} onClick={() => forms.act('designMode')} />
          <Button icon="settings" label={t('Properties')} title={t('Properties: the selected control\'s title, tag, placeholder, locks and list')} onClick={() => forms.act('properties')} />
        </Rows>
      </Group>
      <Group label={t('Protect')}>
        <Button tall icon="shield" label={t('Restrict Editing')} title={t('Restrict Editing: allow only filling in forms, comments or tracked changes')} onClick={() => act('restrictPane', true)} />
      </Group>
    </>
  );
}
