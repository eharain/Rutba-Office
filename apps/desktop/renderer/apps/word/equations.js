/**
 * Equations on the Word page, and the editor that makes them.
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
 * and Delete take it whole, a click selects it and a double-click (or Enter
 * on a selected one) opens it in the editor below.
 */
import React from 'react';
import { Button, Dialog } from '@rutba/office-ui';
import { ommlToMathml } from '@rutba/ooxml/math';
import { linearToOmml, EQUATION_GALLERY } from '@rutba/ooxml/math-linear';

/** Inside the shadow: the math face and the spacing Word gives a display equation. */
const SHADOW_CSS = `
math { font-family: 'Cambria Math', 'STIX Two Math', 'Latin Modern Math', math; font-size: 1.04em; }
math[display="block"] { margin: 0.3em 0; }
/* The page marks a selected equation with its own box; the browser's
   highlight over pieces of the MathML is left out. */
::selection { background: transparent; }
:host([data-jc="left"]) math, :host([data-jc="right"]) math { display: inline math; math-style: normal; }
`;

/** On the page, around the shadow host, and the editor's own look. */
export const EQUATION_CSS = `
.wd-math { display: inline-block; border-radius: 3px; padding: 0 1px; cursor: default; }
.wd-math-display { display: block; padding: 2px 4px; margin: 2px 0; }
.wd-math[data-jc="left"] { text-align: left; }
.wd-math[data-jc="right"] { text-align: right; }
.wd-math:hover { background: rgba(60, 64, 67, 0.06); }
.wd-math.sel { background: rgba(38, 110, 214, 0.18); box-shadow: 0 0 0 1px rgba(38, 110, 214, 0.35); }

.wd-eq { display: flex; flex-direction: column; gap: 8px; }
.wd-eq math { font-family: 'Cambria Math', 'STIX Two Math', 'Latin Modern Math', math; }
.wd-eq-structs { display: grid; grid-template-columns: repeat(8, minmax(0, 1fr)); gap: 4px; }
/* Word's Structures group: the shape on top, its name under it. */
.wd-eq-struct {
  display: flex; flex-direction: column; align-items: center; justify-content: flex-end; gap: 3px; height: 62px; padding: 4px 2px 5px;
  border-radius: var(--r-1); border: 1px solid transparent; background: transparent; color: var(--ink); font: inherit;
  font-size: 11.5px; line-height: 1.15; cursor: pointer; white-space: nowrap;
}
.wd-eq-struct > span:first-child { flex: 1; display: flex; align-items: center; }
.wd-eq-struct:hover { background: var(--hover); border-color: var(--accent-line); }
.wd-eq-struct.on { background: var(--selected); border-color: var(--accent-line); font-weight: 600; }
.wd-eq-struct math { font-size: 14px; }
.wd-eq-templates {
  display: grid; grid-template-columns: repeat(auto-fill, minmax(92px, 1fr)); gap: 6px; padding: 8px;
  background: var(--sunken); border-radius: var(--r-2);
}
.wd-eq-template {
  display: flex; align-items: center; justify-content: center; min-height: 52px; padding: 4px 6px; border-radius: var(--r-1);
  border: 1px solid var(--line); background: var(--surface); cursor: pointer; font-size: 17px; color: var(--ink);
}
.wd-eq-template:hover { border-color: var(--accent-line); background: var(--hover); }
.wd-eq-row { display: flex; align-items: center; gap: 8px; }
.wd-eq-label { font-size: 11.5px; font-weight: 600; color: var(--ink-2); text-transform: uppercase; letter-spacing: 0.04em; }
.wd-eq-sets { display: flex; gap: 4px; }
.wd-eq-set {
  padding: 2px 10px; border-radius: 20px; border: 1px solid transparent; background: transparent; color: var(--ink-2);
  font: inherit; font-size: 11.5px; cursor: pointer;
}
.wd-eq-set:hover { background: var(--hover); }
.wd-eq-set.on { background: var(--selected); border-color: var(--accent-line); color: var(--ink); font-weight: 600; }
.wd-eq-symbols { display: grid; grid-template-columns: repeat(auto-fill, minmax(28px, 1fr)); gap: 3px; }
.wd-eq-symbol {
  height: 28px; border: 1px solid var(--line); background: var(--surface); color: var(--ink); border-radius: var(--r-1);
  font-size: 16px; cursor: pointer; font-family: 'Cambria Math', 'Segoe UI Symbol', serif; padding: 0;
}
.wd-eq-symbol:hover { border-color: var(--accent-line); background: var(--hover); }
.wd-eq-input {
  width: 100%; box-sizing: border-box; min-height: 46px; resize: vertical; padding: 8px 10px; border-radius: var(--r-1);
  border: 1px solid var(--line-strong); background: var(--surface); color: var(--ink);
  font-family: 'Cambria Math', Cambria, serif; font-size: 17px; line-height: 1.4;
}
.wd-eq-input:focus { outline: 2px solid var(--accent-line); outline-offset: 0; border-color: var(--accent); }
.wd-eq-preview {
  min-height: 84px; display: flex; align-items: center; justify-content: center; padding: 10px; border-radius: var(--r-2);
  border: 1px solid var(--line); background: #ffffff; color: #000000; font-size: 22px; overflow-x: auto;
}
.wd-eq-preview.bad { justify-content: flex-start; color: var(--bad); font-size: 13px; background: var(--surface); }
.wd-eq-gallery { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 6px; max-height: 172px; overflow: auto; padding-right: 2px; }
.wd-eq-tile {
  display: flex; flex-direction: column; align-items: stretch; gap: 4px; padding: 6px 8px 8px; border-radius: var(--r-1);
  border: 1px solid var(--line); background: #ffffff; color: #000000; cursor: pointer; text-align: left; font: inherit;
}
.wd-eq-tile:hover { border-color: var(--accent-line); box-shadow: 0 0 0 2px var(--selected); }
.wd-eq-tile-name { font-size: 11px; color: var(--ink-2); font-weight: 600; }
.wd-eq-tile-math { display: flex; justify-content: safe center; align-items: center; min-height: 40px; font-size: 10.5px; overflow: hidden; }
.wd-eq-check { display: inline-flex; align-items: center; gap: 6px; font-size: 12.5px; color: var(--ink); cursor: pointer; }
.wd-eq-hint { font-size: 11.5px; color: var(--ink-3); margin: 0; }
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

/* ── the editor ────────────────────────────────────────────────────────── */

/** The linear form drawn as MathML, or the reason it cannot be. */
export function linearPreview(linear, display = true) {
  const built = linearToOmml(linear, { display });
  if (!built.ok) return { ok: false, error: built.error, at: built.at };
  return { ok: true, mathml: ommlToMathml(built.xml), xml: built.xml };
}

/** A few characters of MathML markup, set inline (a button's face, a gallery tile). */
function MathFace({ linear, display = false }) {
  const html = React.useMemo(() => {
    const p = linearPreview(linear, display);
    return p.ok ? p.mathml : '';
  }, [linear, display]);
  return <span dangerouslySetInnerHTML={{ __html: html }} />;
}

/**
 * Word's Structures, each a gallery of templates in the linear form. `caret`
 * is where the typing goes after one is put in: the first empty slot. `〖〗`
 * is an invisible group, so an empty one draws as Word's dotted placeholder.
 */
export const STRUCTURES = [
  { name: 'Fraction', face: 'x/y', templates: [
    { label: 'Stacked fraction', text: '()/()', caret: 1 },
    { label: 'Skewed fraction', text: '()∕()', caret: 1 },
    { label: 'Linear fraction', text: '()⊘()', caret: 1 },
    { label: 'Fraction with no bar', text: '(¦)', caret: 1 },
    { label: 'dy/dx', text: '(dy)/(dx)', caret: 9 },
  ] },
  { name: 'Script', face: 'e^x', templates: [
    { label: 'Superscript', text: '〖〗^()', caret: 1 },
    { label: 'Subscript', text: '〖〗_()', caret: 1 },
    { label: 'Subscript and superscript', text: '〖〗_()^()', caret: 1 },
    { label: 'Left subscript and superscript', text: '_()^() 〖〗', caret: 2 },
    { label: 'x squared', text: 'x^2', caret: 3 },
  ] },
  { name: 'Radical', face: '√x', templates: [
    { label: 'Square root', text: '√()', caret: 2 },
    { label: 'Radical with degree', text: '√(&)', caret: 2 },
    { label: 'Cube root', text: '∛()', caret: 2 },
    { label: 'Square root of a² + b²', text: '√(a^2+b^2)', caret: 10 },
  ] },
  { name: 'Integral', face: '∫_0^1▒〖〗', templates: [
    { label: 'Integral', text: '∫▒〖〗', caret: 3 },
    { label: 'Integral with limits', text: '∫_()^()▒〖〗', caret: 3 },
    { label: 'Double integral', text: '∬▒〖〗', caret: 3 },
    { label: 'Contour integral', text: '∮▒〖〗', caret: 3 },
    { label: 'Differential dx', text: 'dx', caret: 2 },
  ] },
  { name: 'Large Operator', face: '∑_(i=1)^n▒〖〗', templates: [
    { label: 'Summation', text: '∑▒〖〗', caret: 3 },
    { label: 'Summation with limits', text: '∑_()^()▒〖〗', caret: 3 },
    { label: 'Summation below', text: '∑_()▒〖〗', caret: 3 },
    { label: 'Product with limits', text: '∏_()^()▒〖〗', caret: 3 },
    { label: 'Union with limits', text: '⋃_()^()▒〖〗', caret: 3 },
  ] },
  { name: 'Bracket', face: '(x)', templates: [
    { label: 'Parentheses', text: '()', caret: 1 },
    { label: 'Brackets', text: '[]', caret: 1 },
    { label: 'Braces', text: '{}', caret: 1 },
    { label: 'Absolute value', text: '||', caret: 1 },
    { label: 'Norm', text: '‖‖', caret: 1 },
    { label: 'Angle brackets', text: '⟨⟩', caret: 1 },
    { label: 'Floor', text: '⌊⌋', caret: 1 },
    { label: 'Ceiling', text: '⌈⌉', caret: 1 },
  ] },
  { name: 'Function', face: 'sin θ', templates: [
    { label: 'Sine', text: 'sin〖〗', caret: 4 },
    { label: 'Cosine', text: 'cos〖〗', caret: 4 },
    { label: 'Tangent', text: 'tan〖〗', caret: 4 },
    { label: 'Logarithm', text: 'log〖〗', caret: 4 },
    { label: 'Natural logarithm', text: 'ln〖〗', caret: 3 },
    { label: 'Limit', text: 'lim_(n→∞) 〖〗', caret: 11 },
  ] },
  { name: 'Matrix', face: '(■(1&0@0&1))', templates: [
    { label: '2 × 2 empty matrix', text: '■(&@&)', caret: 2 },
    { label: '2 × 2 matrix in parentheses', text: '(■(&@&))', caret: 3 },
    { label: '2 × 2 matrix in brackets', text: '[■(&@&)]', caret: 3 },
    { label: '3 × 3 empty matrix', text: '■(&&@&&@&&)', caret: 2 },
    { label: 'Identity matrix', text: '(■(1&0@0&1))', caret: 12 },
  ] },
];

/** The Symbols palette — Greek letters, operators and arrows. */
export const SYMBOL_SETS = [
  ['Greek', 'α β γ δ ε ζ η θ ι κ λ μ ν ξ π ρ σ τ υ φ χ ψ ω Γ Δ Θ Λ Ξ Π Σ Φ Ψ Ω'],
  ['Operators', '± ∓ × ÷ ⋅ ∘ = ≠ ≈ ≡ ∼ ∝ < > ≤ ≥ ≪ ≫ ∞ ∂ ∇ ∈ ∉ ⊂ ⊃ ⊆ ⊇ ∪ ∩ ∀ ∃ ¬ ∧ ∨ ′ ° … ⋯'],
  ['Arrows', '→ ← ↔ ⇒ ⇐ ⇔ ↑ ↓ ↦ ⟶ ⟹ ⟺ ↗ ↘ ↙ ↖'],
];

/**
 * The equation editor — Insert → Equation (Alt+=), and an equation opened
 * again with a double-click. A line of Word's linear format, a live preview
 * of what it builds, the structures and symbols at the input's caret, and
 * Word's built-in equations. OK hands back the linear form and whether the
 * equation stands on its own line.
 */
export function EquationDialog({ initial = '', display: initialDisplay = true, editing = false, onClose, onInsert }) {
  const [text, setText] = React.useState(initial);
  const [display, setDisplay] = React.useState(Boolean(initialDisplay));
  const [open, setOpen] = React.useState(null);
  const [set, setSet] = React.useState(0);
  const input = React.useRef(null);
  const preview = React.useMemo(() => (text.trim() ? linearPreview(text, display) : null), [text, display]);

  React.useEffect(() => {
    const el = input.current;
    if (!el) return;
    el.focus();
    el.setSelectionRange(el.value.length, el.value.length);
  }, []);

  /** Put characters in at the input's caret, and the caret `caret` into them. */
  const put = (chars, caret = chars.length) => {
    const el = input.current;
    const from = el ? el.selectionStart : text.length;
    const to = el ? el.selectionEnd : text.length;
    const next = text.slice(0, from) + chars + text.slice(to);
    setText(next);
    requestAnimationFrame(() => {
      if (!input.current) return;
      input.current.focus();
      input.current.setSelectionRange(from + caret, from + caret);
    });
  };

  const ok = () => {
    if (!preview?.ok) return;
    onInsert(text, display);
  };

  const struct = STRUCTURES.find((s) => s.name === open) || null;
  const chars = SYMBOL_SETS[set][1].split(/\s+/).filter(Boolean);

  return (
    <Dialog
      title={editing ? 'Edit Equation' : 'Equation'}
      width={720}
      onClose={onClose}
      actions={
        <>
          <label className="wd-eq-check" style={{ marginRight: 'auto' }}>
            <input type="checkbox" className="wd-eq-display" checked={display} onChange={(e) => setDisplay(e.target.checked)} />
            Display — on a line of its own, centred
          </label>
          <Button label="Cancel" onClick={onClose} />
          <Button primary className="wd-eq-ok" label={editing ? 'Update' : 'Insert'} disabled={!preview?.ok} onClick={ok} />
        </>
      }
    >
      <div className="wd-eq">
        <div className="wd-eq-structs" role="toolbar" aria-label="Structures">
          {STRUCTURES.map((s) => (
            <button
              key={s.name}
              type="button"
              className={`wd-eq-struct${open === s.name ? ' on' : ''}`}
              data-struct={s.name}
              data-tip={`${s.name} — pick a template to put in at the caret`}
              onClick={() => setOpen(open === s.name ? null : s.name)}
            >
              <MathFace linear={s.face} />
              <span className="wd-eq-struct-name">{s.name}</span>
            </button>
          ))}
        </div>
        {struct ? (
          <div className="wd-eq-templates">
            {struct.templates.map((t) => (
              <button key={t.label} type="button" className="wd-eq-template" data-tip={t.label} onClick={() => { put(t.text, t.caret); setOpen(null); }}>
                <MathFace linear={t.text} display />
              </button>
            ))}
          </div>
        ) : null}

        <textarea
          ref={input}
          className="wd-eq-input"
          value={text}
          spellCheck={false}
          placeholder="Type an equation — x=(-b±√(b^2-4ac))/2a, \sum_(i=1)^n i, \int_0^1 f(x)dx, a_i, \alpha"
          onChange={(e) => setText(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); ok(); }
          }}
        />
        <div className={`wd-eq-preview${preview && !preview.ok ? ' bad' : ''}`} aria-live="polite">
          {!preview ? <span className="wd-eq-hint">The equation appears here as you type.</span>
            : preview.ok ? <span className="wd-eq-built" dangerouslySetInnerHTML={{ __html: preview.mathml }} />
            : <span className="wd-eq-error">{`${preview.error} (at character ${preview.at + 1})`}</span>}
        </div>

        <div className="wd-eq-row">
          <span className="wd-eq-label">Symbols</span>
          <div className="wd-eq-sets">
            {SYMBOL_SETS.map(([label], i) => (
              <button key={label} type="button" className={`wd-eq-set${set === i ? ' on' : ''}`} onClick={() => setSet(i)}>{label}</button>
            ))}
          </div>
        </div>
        <div className="wd-eq-symbols">
          {chars.map((c) => (
            <button key={c} type="button" className="wd-eq-symbol" data-tip={c} onClick={() => put(c)}>{c}</button>
          ))}
        </div>

        <span className="wd-eq-label">Built-in</span>
        <div className="wd-eq-gallery">
          {EQUATION_GALLERY.map((g) => (
            <button key={g.name} type="button" className="wd-eq-tile" data-gallery={g.name} data-tip={`${g.name} — put it in the editor`} onClick={() => setText(g.linear)}>
              <span className="wd-eq-tile-name">{g.name}</span>
              <span className="wd-eq-tile-math"><MathFace linear={g.linear} /></span>
            </button>
          ))}
        </div>
      </div>
    </Dialog>
  );
}

/* ── the clipboard ─────────────────────────────────────────────────────── */

/** The clipboard type a copy within the suite carries its equations on. */
export const CLIP_TYPE = 'application/x-rutba-word-runs';

/**
 * What a copy of [from, to) holds, from the blocks the window draws: plain
 * text for anything else (each equation as its linear form) and, when an
 * equation is in it, the pieces a paste in the suite puts back as they were.
 */
export function clipOf(blocks, from, to) {
  const lines = [];
  let hasMath = false;
  for (let b = from.block; b <= to.block; b++) {
    const block = blocks[b];
    if (!block) continue;
    const start = b === from.block ? from.offset : 0;
    const end = b === to.block ? to.offset : Infinity;
    const pieces = [];
    let at = 0;
    for (const run of block.runs || []) {
      const text = run.text ?? '';
      const s = at;
      const e = at + text.length;
      at = e;
      if (e <= start || s >= end || run.del) continue;
      if (run.math) {
        if (start <= s && end >= e) {
          pieces.push({ math: { xml: run.math.xml, display: Boolean(run.math.display) }, linear: run.math.linear || '' });
          hasMath = true;
        }
        continue;
      }
      const slice = text.slice(Math.max(0, start - s), Math.min(text.length, end - s));
      if (slice) pieces.push({ text: slice });
    }
    lines.push(pieces);
  }
  const text = lines.map((pieces) => pieces.map((p) => (p.math ? p.linear : p.text)).join('')).join('\n');
  return { text, lines: lines.map((pieces) => pieces.map((p) => (p.math ? { math: p.math } : p))), hasMath };
}
