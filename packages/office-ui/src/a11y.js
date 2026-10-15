// What a screen reader is told about the window's controls.
//
// The suite draws its own tips (data-tip), and a picture button's name is in
// its tip alone: to a screen reader such a button had no name at all. So each
// window names every control nothing else names, from its tip (the part
// before its long dash), and gives the rest of a tip to the control as its
// description, as Office's ScreenTips are read: drawn once when the window
// opens and again for whatever is added or re-tipped later. A picture inside
// a control that has a name is the control's, not a picture of its own.

import { useEffect } from 'react';

const CONTROLS = 'button, [role="button"], [role="tab"], [role="menuitem"], [role="menuitemcheckbox"], [role="option"], [role="slider"], [role="checkbox"], [role="switch"], a[href], select, textarea, input:not([type="hidden"])';
const FIELDS = new Set(['INPUT', 'SELECT', 'TEXTAREA']);
const DASHES = String.fromCharCode(0x2013, 0x2014);
const SPLIT = new RegExp(`\\s+[${DASHES}]\\s+`);
const LEAD = new RegExp(`^[\\s${DASHES}:(),.-]+`);

/** The tip a control carries, if any: the suite's own, or a title the tips layer has not taken yet. */
const tipOf = (el) => (el.getAttribute('data-tip') || el.getAttribute('title') || '').trim();

/** The words in `el` a screen reader is given: none from inside a drawing or anything hidden. */
function spokenText(el) {
  const walk = el.ownerDocument.createTreeWalker(el, 4);
  for (let n = walk.nextNode(); n; n = walk.nextNode()) {
    if (!n.data.trim()) continue;
    if (n.parentElement?.closest('svg, [aria-hidden="true"]') && el.contains(n.parentElement.closest('svg, [aria-hidden="true"]'))) continue;
    return n.data.trim();
  }
  return '';
}

/** Whether something other than this pass already names `el`. */
function namedElsewhere(el) {
  if (el.dataset.nameFrom === 'tip') return false;
  if ((el.getAttribute('aria-label') || '').trim()) return true;
  if (el.getAttribute('aria-labelledby')) return true;
  if (FIELDS.has(el.tagName)) {
    if (el.labels && [...el.labels].some((l) => l.textContent.trim())) return true;
    if (el.tagName === 'INPUT' && /^(button|submit|reset)$/i.test(el.type) && el.value) return true;
    return false;
  }
  // A control named by what it holds: its words (not a picture's, which is the control's and hidden), or a picture with a title.
  if (spokenText(el)) return true;
  if (el.querySelector('img[alt]:not([alt=""]), svg > title')) return true;
  return false;
}

/** The name a tip gives, and what it says besides. */
export function splitTip(tip) {
  const text = String(tip || '').trim();
  if (!text) return { name: '', more: '' };
  const parts = text.split(SPLIT);
  const name = parts[0].trim();
  const more = parts.slice(1).join(` ${DASHES.slice(1)} `).trim();
  return { name, more };
}

/** Name `el` from its tip where nothing else does; describe it by the rest. */
export function nameControl(el) {
  const tip = tipOf(el);
  if (!tip) return;
  const { name, more } = splitTip(tip);
  if (!namedElsewhere(el)) {
    if (el.getAttribute('aria-label') !== name) el.setAttribute('aria-label', name);
    el.dataset.nameFrom = 'tip';
  }
  // The rest of the tip, unless the control's own words already say it all.
  if (!el.hasAttribute('aria-describedby')) {
    const own = (el.getAttribute('aria-label') || spokenText(el) || '').trim();
    const said = el.dataset.nameFrom === 'tip' ? more : tip.startsWith(own) ? tip.slice(own.length).replace(LEAD, '').replace(/\)$/, '').trim() : own.includes(tip) ? '' : tip;
    if (said && said !== own) {
      if (el.getAttribute('aria-description') !== said) el.setAttribute('aria-description', said);
    } else if (el.hasAttribute('aria-description') && el.dataset.describedFrom === 'tip') el.removeAttribute('aria-description');
    if (said) el.dataset.describedFrom = 'tip';
  }
}

const LABELLED = 'input:not([type="hidden"]), select, textarea, [role="combobox"], [role="spinbutton"], [role="slider"]';

/**
 * A row's label, given to the first control in it that has no name of its
 * own (and the row's hint as its description): a label drawn beside its
 * field rather than around it is still the field's name.
 */
export function labelFirstControl(box, labelId, hintId = null) {
  const control = box?.querySelector(LABELLED);
  if (!control || !labelId || !box.ownerDocument.getElementById(labelId)) return;
  const by = control.getAttribute('aria-labelledby');
  const named = (control.getAttribute('aria-label') || '').trim() || (by && by !== labelId) || [...(control.labels || [])].some((l) => l.textContent.trim());
  if (!named) control.setAttribute('aria-labelledby', labelId);
  if (hintId && !control.getAttribute('aria-describedby')) control.setAttribute('aria-describedby', hintId);
}

/** A picture inside a named control is part of it. */
function quietPictures(root) {
  for (const svg of root.querySelectorAll('svg[role="img"]:not([aria-hidden])')) {
    if (svg.parentElement?.closest(CONTROLS)) svg.setAttribute('aria-hidden', 'true');
  }
}

/** Every control under `root` (and `root` itself), named. */
export function nameControls(root) {
  if (!root || root.nodeType !== 1) return;
  if (root.matches(CONTROLS)) nameControl(root);
  for (const el of root.querySelectorAll(CONTROLS)) nameControl(el);
  quietPictures(root);
}

/* ── the keyboard ─────────────────────────────────────────────────────────── */

/** How the last thing was done: by key or by pointer. A menu opened from the keyboard takes the keyboard. */
export const modality = { last: 'pointer' };

const TABBABLE = 'a[href], button:not([disabled]), input:not([disabled]):not([type="hidden"]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"]), [contenteditable="true"]';
const shownEl = (el) => Boolean(el && el.isConnected && el.getClientRects().length && getComputedStyle(el).visibility !== 'hidden');
const firstTabbable = (root, skip = null) => [...root.querySelectorAll(TABBABLE)].find((el) => shownEl(el) && !(skip && skip(el)));

/** Where F6 last left the keyboard in each part of the window. */
const lastIn = new WeakMap();

/**
 * The parts of the window F6 goes round, in the order they stand: the
 * ribbon, the panes, the document and the status bar, as Office's F6 does.
 * The document is the window's work outside its panes: what the app marks
 * `data-f6-home`, or where the keyboard last was there, or the first thing
 * there that takes it.
 */
export function f6Regions(doc = document) {
  const out = [];
  const ribbon = doc.querySelector('.rw-ribbon');
  if (shownEl(ribbon)) out.push({ kind: 'ribbon', el: ribbon });
  const body = doc.querySelector('.rw-body');
  const panels = body ? [...body.querySelectorAll('.rw-panel')].filter((p) => shownEl(p) && !p.parentElement.closest('.rw-panel')) : [];
  const marked = body ? [...body.querySelectorAll('[data-f6-home]')].find(shownEl) : null;
  // The app's own mark says where its document is, even in a pane (Mail's message list).
  for (const p of panels) out.push(marked && p.contains(marked) ? { kind: 'document', el: p, home: marked, inPane: () => false } : { kind: 'pane', el: p });
  if (body) {
    const inPane = (el) => panels.some((p) => p.contains(el));
    const home = marked && !inPane(marked) ? marked : firstTabbable(body, inPane);
    if (home) out.push({ kind: marked && !inPane(marked) || !marked ? 'document' : 'pane', el: body, home, inPane });
  }
  const status = doc.querySelector('.rw-status');
  if (shownEl(status)) out.push({ kind: 'status', el: status });
  // In the order they stand on the page: a pane on the left before the document, one on the right after it.
  const at = (r) => (r.kind === 'document' ? r.home : r.el);
  return out.sort((a, b) => (at(a) === at(b) ? 0 : at(a).compareDocumentPosition(at(b)) & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1));
}

/** The part of the window the keyboard is in now. */
function regionOf(list, el) {
  if (!el || el === document.body) return -1;
  // A pane before the window's work that holds it.
  const pane = list.findIndex((r) => r.el !== r.el.ownerDocument.querySelector('.rw-body') && r.el.contains(el));
  if (pane >= 0) return pane;
  return list.findIndex((r) => r.inPane && r.el.contains(el) && !r.inPane(el));
}

/** Put the keyboard in a part of the window: back where it was there, or at the part's own start. */
export function enterRegion(r) {
  const kept = lastIn.get(r.el);
  const ok = (el) => el && shownEl(el) && r.el.contains(el) && !el.matches?.('[disabled]') && (!r.inPane || !r.inPane(el));
  let to = null;
  if (r.kind === 'ribbon') to = r.el.querySelector('.rw-tab[aria-selected="true"]') || firstTabbable(r.el);
  else if (r.home) to = ok(kept) ? kept : r.home;
  else to = ok(kept) ? kept : firstTabbable(r.el);
  if (!to) {
    // A part with nothing that takes the keyboard (a status bar of words) is read as a whole.
    if (!r.el.hasAttribute('tabindex')) r.el.setAttribute('tabindex', '-1');
    to = r.el;
  }
  to.focus({ preventScroll: r.kind === 'ribbon' });
  return to;
}

/** F6 forward or Shift+F6 back round the window's parts; false when there is nowhere to go. */
export function cycleRegions(back = false) {
  const list = f6Regions();
  if (!list.length) return false;
  const now = regionOf(list, document.activeElement);
  const next = now < 0 ? (back ? list.length - 1 : 0) : (now + (back ? -1 : 1) + list.length) % list.length;
  enterRegion(list[next]);
  return true;
}

/** The document's own place for the keyboard, when what held it has gone. */
export function focusDocument() {
  const doc = f6Regions().find((r) => r.kind === 'document');
  if (doc) enterRegion(doc);
  return Boolean(doc);
}

/** Once per window: name what is there and whatever comes, F6 round the parts, and the keyboard never left on nothing. */
export function useControlNames() {
  useEffect(() => {
    const onKey = (e) => {
      modality.last = 'key';
      if (e.key !== 'F6' || e.ctrlKey || e.altKey || e.metaKey || e.defaultPrevented) return;
      // A dialog keeps the keyboard until it closes; a menu, until it is put away.
      if (document.querySelector('.rw-dialog, .rw-menu')) return;
      if (cycleRegions(e.shiftKey)) e.preventDefault();
    };
    const onPointer = () => { modality.last = 'pointer'; };
    let held = null;
    const onFocus = (e) => {
      held = e.target;
      const list = f6Regions();
      const r = list[regionOf(list, e.target)];
      if (r) lastIn.set(r.el, e.target);
    };
    window.addEventListener('keydown', onKey);
    window.addEventListener('pointerdown', onPointer, true);
    document.addEventListener('focusin', onFocus);
    // What held the keyboard was taken away (a pane closed, a list redrawn) and
    // nothing took it after: it goes to the document rather than to nothing.
    // Late enough that a dialog or a menu gives it back to what opened it first.
    let fixing = 0;
    const lost = () => {
      if (fixing || !held || held.isConnected) return;
      fixing = setTimeout(() => {
        fixing = 0;
        if (held && !held.isConnected && (!document.activeElement || document.activeElement === document.body)) {
          if (!document.querySelector('.rw-dialog, .rw-menu')) focusDocument();
        }
      }, 80);
    };
    if (typeof MutationObserver !== 'function') return undefined;
    nameControls(document.body);
    const watch = new MutationObserver((list) => {
      lost();
      const roots = new Set();
      for (const m of list) {
        if (m.type === 'attributes') { if (m.target.nodeType === 1) roots.add(m.target); continue; }
        for (const n of m.addedNodes) if (n.nodeType === 1) roots.add(n);
        // A control whose words changed may have lost or gained its name.
        if (m.target.nodeType === 1) { const c = m.target.closest(CONTROLS); if (c) roots.add(c); }
      }
      for (const r of roots) if (r.isConnected) nameControls(r);
    });
    watch.observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ['data-tip', 'title'] });
    return () => {
      watch.disconnect();
      clearTimeout(fixing);
      window.removeEventListener('keydown', onKey);
      window.removeEventListener('pointerdown', onPointer, true);
      document.removeEventListener('focusin', onFocus);
    };
  }, []);
}
