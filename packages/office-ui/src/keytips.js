// The ribbon from the keyboard, as Office's: KeyTips and the arrow keys.
//
// Alt pressed and let go alone (or F10) puts a letter on every tab and a
// number on every quick-access button; typing a tab's letter chooses the
// tab and puts letters on its controls; typing a control's letters presses
// it. A folded group's letters drop the group and letter its controls in
// turn. Escape goes back a step and, from the tabs, puts the letters away;
// so does a press of the mouse or Alt again. The letters are drawn over the
// ribbon, never in the buttons, so no button's name changes while they show.
//
// The arrow keys move through the ribbon whenever the keyboard is in it:
// along the tabs (choosing each, as Office does), down into the controls,
// and among the controls to the nearest one that way. Escape from a control
// goes back to the page the keyboard came from.

import React, { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';

/** The tabs' letters Office uses, where the tab is one Office has. */
const TAB_KEYS = {
  home: 'H', insert: 'N', draw: 'D', design: 'G', layout: 'P', pageLayout: 'P', references: 'S', mailings: 'M',
  review: 'R', view: 'W', help: 'Y', formulas: 'M', data: 'A', transitions: 'K', animations: 'A', slideShow: 'S',
  record: 'C', automate: 'U',
};

const CONTROLS = 'button, select, input:not([type="hidden"]), textarea, [role="button"]';
const letterOf = (ch) => (/[\p{L}\p{N}]/u.test(ch) ? ch.toLocaleUpperCase() : '');
const lettersIn = (name) => [...String(name || '').normalize('NFC')].map(letterOf).filter(Boolean);
const shown = (n) => {
  const r = n.getBoundingClientRect();
  return r.width > 0 && r.height > 0 && getComputedStyle(n).visibility !== 'hidden';
};
/** What a control is called: its name as shown, its label for a screen reader, or its tip. */
export const nameOf = (n) => (n.querySelector(':scope > span')?.textContent || n.getAttribute('data-label') || n.getAttribute('aria-label') || n.getAttribute('data-tip') || n.getAttribute('title') || n.textContent || '').trim();

/**
 * Letters for each of `names`, none of them the beginning of another's: a
 * name whose first letter is its own alone gets that letter; names that
 * share a first letter get two, the second taken from further along each
 * name (or a digit), so typing the first narrows to them and the second
 * picks one. `taken` are codes already spoken for, preferred letters kept.
 */
export function keyTipCodes(names, preferred = []) {
  const codes = new Array(names.length).fill('');
  const used = new Set();
  // Office's own letters first, where they are free.
  preferred.forEach((p, i) => {
    if (p && !used.has(p) && ![...used].some((u) => u.startsWith(p) || p.startsWith(u))) {
      codes[i] = p;
      used.add(p);
    }
  });
  const firsts = new Map();
  names.forEach((name, i) => {
    if (codes[i]) return;
    const letters = lettersIn(name);
    let first = letters.find((ch) => !used.has(ch)) || letters[0] || '';
    if (!first) first = String((i % 9) + 1);
    if (!firsts.has(first)) firsts.set(first, []);
    firsts.get(first).push(i);
  });
  for (const [first, list] of firsts) {
    if (list.length === 1 && !used.has(first) && ![...used].some((u) => u.startsWith(first))) {
      codes[list[0]] = first;
      used.add(first);
      continue;
    }
    const seconds = new Set();
    for (const i of list) {
      const letters = lettersIn(names[i]).slice(1);
      let second = letters.find((ch) => !seconds.has(ch));
      for (let d = 1; !second && d <= 9; d++) if (!seconds.has(String(d))) second = String(d);
      for (let c = 65; !second && c <= 90; c++) if (!seconds.has(String.fromCharCode(c))) second = String.fromCharCode(c);
      seconds.add(second);
      codes[i] = first + second;
      used.add(codes[i]);
    }
  }
  return codes;
}

/** The control nearest `from` in a direction, among `all`, or null. */
function nearest(from, all, dir) {
  const a = from.getBoundingClientRect();
  const ax = a.left + a.width / 2;
  const ay = a.top + a.height / 2;
  let best = null;
  let score = Infinity;
  for (const n of all) {
    if (n === from) continue;
    const b = n.getBoundingClientRect();
    const bx = b.left + b.width / 2;
    const by = b.top + b.height / 2;
    const dx = bx - ax;
    const dy = by - ay;
    let along;
    let across;
    if (dir === 'ArrowRight') { if (b.left < a.right - 2) continue; along = dx; across = Math.max(0, Math.abs(dy) - a.height / 2); }
    else if (dir === 'ArrowLeft') { if (b.right > a.left + 2) continue; along = -dx; across = Math.max(0, Math.abs(dy) - a.height / 2); }
    else if (dir === 'ArrowDown') { if (b.top < a.bottom - 2) continue; along = dy; across = Math.abs(dx); }
    else { if (b.bottom > a.top + 2) continue; along = -dy; across = Math.abs(dx); }
    const s = along + across * 3;
    if (s < score) { score = s; best = n; }
  }
  return best;
}

const usable = (root) => [...root.querySelectorAll(CONTROLS)].filter((n) => !n.disabled && shown(n) && !n.closest('[aria-hidden="true"]'));

/**
 * KeyTips and the arrow keys for one ribbon. `ref` is the ribbon's node;
 * `peek(true)` drops a folded ribbon's groups. Returns the overlay to draw
 * and the ribbon's own keydown handler.
 */
export function useRibbonKeys({ ref, tabs, active, onTab, collapsed, peek }) {
  // null, or { level: 'tabs' | 'controls' | 'group', typed, group }
  const [mode, setMode] = useState(null);
  const [tips, setTips] = useState([]);
  const back = useRef(null);
  const armed = useRef(false);
  const live = useRef({ tabs, active, onTab, collapsed, peek });
  live.current = { tabs, active, onTab, collapsed, peek };

  const stop = useCallback(() => setMode(null), []);

  // Alt alone, or F10: the letters on, or off again.
  useEffect(() => {
    const down = (e) => {
      if (e.key === 'Alt' && !e.repeat && !e.ctrlKey && !e.shiftKey && !e.metaKey) { armed.current = true; return; }
      armed.current = false;
      if (e.key === 'F10' && !e.ctrlKey && !e.altKey && !e.metaKey && !e.shiftKey) {
        e.preventDefault();
        setMode((m) => (m ? null : { level: 'tabs', typed: '' }));
      }
    };
    const up = (e) => {
      if (e.key !== 'Alt') return;
      const was = armed.current;
      armed.current = false;
      if (!was) return;
      e.preventDefault();
      if (!ref.current || !ref.current.isConnected || !ref.current.offsetParent) return;
      if (document.querySelector('.rw-dialog, [role="dialog"][aria-modal="true"]')) return;
      setMode((m) => (m ? null : { level: 'tabs', typed: '' }));
    };
    const off = () => { armed.current = false; };
    window.addEventListener('keydown', down, true);
    window.addEventListener('keyup', up, true);
    window.addEventListener('mousedown', off, true);
    return () => {
      window.removeEventListener('keydown', down, true);
      window.removeEventListener('keyup', up, true);
      window.removeEventListener('mousedown', off, true);
    };
  }, [ref]);

  // Where the keyboard was, to give it back.
  useEffect(() => {
    if (mode && !back.current) {
      const a = document.activeElement;
      back.current = a && a !== document.body && !ref.current?.contains(a) ? a : null;
    }
  }, [mode, ref]);

  // The letters, placed over what they press. Measured after the tab's
  // controls have drawn (a tab chosen by its letter draws its groups first).
  useLayoutEffect(() => {
    if (!mode || !ref.current) { setTips([]); return undefined; }
    const root = ref.current;
    let frame = requestAnimationFrame(() => {
      frame = 0;
      const rtl = getComputedStyle(root).direction === 'rtl';
      const place = (n, how) => {
        const r = n.getBoundingClientRect();
        if (how === 'centre') return { x: r.left + r.width / 2, y: r.bottom - 7, centre: true };
        // A small control's tag sits on its picture's lower corner, clear of its name.
        const pic = n.querySelector(':scope > svg:not(.rw-caret)')?.getBoundingClientRect();
        if (pic && pic.width) return { x: rtl ? pic.right - 2 : pic.left + 2, y: pic.bottom - 1, centre: true };
        return { x: rtl ? r.right - 10 : r.left + 10, y: r.bottom - 5, centre: true };
      };
      let list = [];
      if (mode.level === 'tabs') {
        const { tabs: all } = live.current;
        const nodes = [...root.querySelectorAll('.rw-tab')];
        const codes = keyTipCodes(nodes.map((n) => n.textContent), all.map((t) => TAB_KEYS[t.id] || ''));
        list = nodes.map((n, i) => ({ node: n, code: codes[i], tab: all[i]?.id, ...place(n, 'centre') }));
        [...root.querySelectorAll('.rw-quick button')].filter(shown).slice(0, 9).forEach((n, i) => list.push({ node: n, code: String(i + 1), ...place(n, 'centre') }));
      } else {
        const scope = mode.level === 'group' ? mode.group?.querySelector(':scope > .rw-group-body') : root.querySelector('.rw-groups');
        if (scope) {
          const nodes = usable(scope).filter((n) => mode.level === 'group' || !n.closest('.rw-group-body') || !n.closest('.rw-group.folded'));
          const codes = keyTipCodes(nodes.map(nameOf));
          list = nodes.map((n, i) => ({ node: n, code: codes[i], ...place(n, n.matches('.rw-fold') || (n.matches('.rw-btn') && getComputedStyle(n).flexDirection === 'column') ? 'centre' : 'start') }));
        }
      }
      setTips(list);
    });
    return () => { if (frame) cancelAnimationFrame(frame); };
  }, [mode, active, ref]);

  // While they show, the keys are theirs.
  useEffect(() => {
    if (!mode) return undefined;
    const press = (tip) => {
      const { onTab: choose, collapsed: folded, peek: drop } = live.current;
      if (tip.tab) {
        choose(tip.tab);
        if (folded) drop(true);
        setMode({ level: 'controls', typed: '' });
        return;
      }
      const n = tip.node;
      if (n.matches('.rw-fold')) {
        const group = n.closest('.rw-group');
        if (group) group.dataset.keyOpened = '1';
        n.click();
        setMode({ level: 'group', typed: '', group });
        return;
      }
      setMode(null);
      if (n.matches('select, input, textarea')) n.focus();
      else n.click();
    };
    const key = (e) => {
      if (e.key === 'Alt' || e.key === 'Shift' || e.key === 'Control' || e.key === 'Meta') return;
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        setMode((m) => (!m || m.level === 'tabs' ? null : m.level === 'group' ? { level: 'controls', typed: '' } : { level: 'tabs', typed: '' }));
        if (mode.level === 'group') mode.group?.querySelector(':scope > .rw-fold')?.click();
        return;
      }
      if (e.key.startsWith('Arrow') || e.key === 'Tab') {
        // The arrows take the keyboard into the ribbon, and the letters go.
        e.preventDefault();
        e.stopPropagation();
        const root = ref.current;
        setMode(null);
        if (!root) return;
        if (mode.level === 'tabs' && e.key !== 'ArrowDown') root.querySelector('.rw-tab[aria-selected="true"]')?.focus();
        else usable(root.querySelector('.rw-groups') || root)[0]?.focus();
        return;
      }
      if (e.key.length !== 1 || e.ctrlKey || e.metaKey) return;
      e.preventDefault();
      e.stopPropagation();
      const typed = (mode.typed || '') + letterOf(e.key);
      const matches = tips.filter((t) => t.code.startsWith(typed));
      if (!matches.length) return;
      const exact = matches.find((t) => t.code === typed);
      if (exact && matches.length === 1) press(exact);
      else setMode({ ...mode, typed });
    };
    const away = (e) => { if (!e.target.closest?.('.rw-keytips')) stop(); };
    window.addEventListener('keydown', key, true);
    window.addEventListener('mousedown', away, true);
    window.addEventListener('blur', stop);
    window.addEventListener('resize', stop);
    return () => {
      window.removeEventListener('keydown', key, true);
      window.removeEventListener('mousedown', away, true);
      window.removeEventListener('blur', stop);
      window.removeEventListener('resize', stop);
    };
  }, [mode, tips, ref, stop]);

  // The arrows, whenever the keyboard is in the ribbon.
  const onKeyDown = useCallback((e) => {
    if (e.defaultPrevented || e.altKey || e.ctrlKey || e.metaKey) return;
    const root = ref.current;
    const at = e.target;
    if (!root || !(at instanceof Element)) return;
    if (e.key === 'Escape' && !at.closest('.rw-group[data-open]')) {
      // From the ribbon back to the page.
      const to = back.current;
      e.preventDefault();
      if (to && to.isConnected) to.focus({ preventScroll: true });
      else at.blur();
      back.current = null;
      return;
    }
    if (!e.key.startsWith('Arrow')) return;
    if (at.matches('select, input, textarea')) return;
    const { tabs: all, onTab: choose } = live.current;
    if (at.matches('.rw-tab')) {
      const nodes = [...root.querySelectorAll('.rw-tab')];
      const i = nodes.indexOf(at);
      const rtl = getComputedStyle(root).direction === 'rtl';
      const step = e.key === 'ArrowRight' ? (rtl ? -1 : 1) : e.key === 'ArrowLeft' ? (rtl ? 1 : -1) : 0;
      if (step) {
        e.preventDefault();
        const next = nodes[(i + step + nodes.length) % nodes.length];
        next.focus();
        const id = all[nodes.indexOf(next)]?.id;
        if (id) choose(id);
      } else if (e.key === 'ArrowDown') {
        e.preventDefault();
        usable(root.querySelector('.rw-groups') || root)[0]?.focus();
      }
      return;
    }
    const pop = at.closest('.rw-group[data-open] > .rw-group-body');
    const scope = pop || root.querySelector('.rw-groups');
    if (!scope || !scope.contains(at)) return;
    const all2 = usable(scope).filter((n) => pop || !n.closest('.rw-group.folded > .rw-group-body'));
    const to = nearest(at, all2, e.key);
    e.preventDefault();
    if (to) to.focus();
    else if (e.key === 'ArrowUp' && !pop) root.querySelector('.rw-tab[aria-selected="true"]')?.focus();
  }, [ref]);

  const typed = mode?.typed || '';
  const overlay = mode && tips.length ? (
    <div className="rw-keytips" aria-hidden="true">
      {tips.map((tip, i) => (
        <span
          key={i}
          className={`rw-keytip${typed && !tip.code.startsWith(typed) ? ' off' : ''}`}
          data-code={tip.code}
          data-tab={tip.tab || undefined}
          style={{ left: Math.round(tip.x), top: Math.round(tip.y) }}
        >
          {tip.code}
        </span>
      ))}
    </div>
  ) : null;
  return { overlay, onKeyDown, keyTips: Boolean(mode), stopKeyTips: stop };
}
