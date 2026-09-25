/**
 * Equations on the Word page.
 *
 * An equation is ONE character of its paragraph to the engine (U+FFFC, like
 * a footnote reference) and a piece of MathML to the eye. The two are kept
 * apart with a shadow root: the host span holds the one character as its
 * only light-DOM child — so every offset the page computes from text nodes
 * and `textContent` counts the equation as exactly one — and the MathML is
 * drawn inside the shadow, where no tree walker, selection or text count of
 * the page's can see its letters. Chromium lays it out with its own MathML
 * Core: a fraction's bar, a radical's overbar, stacked limits, italic
 * variables and upright function names, in Cambria Math where Windows has
 * it, the way Word draws them.
 *
 * The host is not editable: the caret steps over it as a unit, Backspace
 * and Delete take it whole, and a click selects it.
 */
import React from 'react';

/** Inside the shadow: the math face and the spacing Word gives a display equation. */
const SHADOW_CSS = `
math { font-family: 'Cambria Math', 'STIX Two Math', 'Latin Modern Math', math; font-size: 1.04em; }
math[display="block"] { margin: 0.3em 0; }
:host([data-jc="left"]) math, :host([data-jc="right"]) math { display: inline math; math-style: normal; }
`;

/** On the page, around the shadow host. */
export const EQUATION_CSS = `
.wd-math { display: inline-block; border-radius: 3px; padding: 0 1px; cursor: default; }
.wd-math-display { display: block; padding: 2px 4px; margin: 2px 0; }
.wd-math[data-jc="left"] { text-align: left; }
.wd-math[data-jc="right"] { text-align: right; }
.wd-math:hover { background: rgba(60, 64, 67, 0.06); }
.wd-math.sel { background: rgba(38, 110, 214, 0.18); box-shadow: 0 0 0 1px rgba(38, 110, 214, 0.35); }
`;

/** Word's jc names as the three sides the page can put a display equation on. */
const SIDE = { left: 'left', right: 'right', center: 'center', centerGroup: 'center' };

/**
 * One equation in the text. `at` is its character offset in the block, which
 * the selection highlight and a click read off the element.
 */
export function MathRun({ run, at }) {
  const ref = React.useRef(null);
  const math = run.math || {};
  const markup = math.mathml || '';
  React.useLayoutEffect(() => {
    const host = ref.current;
    if (!host) return;
    const root = host.shadowRoot || host.attachShadow({ mode: 'open' });
    const html = `<style>${SHADOW_CSS}</style>${markup}`;
    // Only when it changed: a shadow tree rewritten on every render would
    // make the page lay the equation out again for every keystroke nearby.
    if (root.__html !== html) {
      root.innerHTML = html;
      root.__html = html;
    }
  }, [markup]);
  const side = math.display ? SIDE[math.jc] || 'center' : null;
  return (
    <span
      ref={ref}
      className={`wd-math ${math.display ? 'wd-math-display' : 'wd-math-inline'}`}
      contentEditable={false}
      suppressContentEditableWarning
      data-at={at ?? undefined}
      data-jc={side && side !== 'center' ? side : undefined}
      data-linear={math.linear || ''}
      role="math"
      aria-label={math.linear || 'Equation'}
      title="Equation — double-click to edit"
    >
      {run.text}
    </span>
  );
}

/** The equation host a DOM node sits in (its one light-DOM character, or the host itself), or null. */
export function mathHostOf(node) {
  const el = node?.nodeType === 3 ? node.parentElement : node;
  return el?.closest?.('.wd-math') || null;
}
