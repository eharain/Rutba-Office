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

/** Once per window: name what is there and whatever comes. */
export function useControlNames() {
  useEffect(() => {
    if (typeof MutationObserver !== 'function') return undefined;
    nameControls(document.body);
    const watch = new MutationObserver((list) => {
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
    return () => watch.disconnect();
  }, []);
}
