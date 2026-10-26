// The Symbol dialog: a grid of the characters people insert most, by set,
// opened by Documents, Worksheets and Presentations.

import React, { useState } from 'react';
import { Button, Dialog, t } from '@rutba/office-ui';

const SYMBOLS = [
  // The two dashes are written by their code points, U+2014 and U+2013.
  [t('Common'), '\u2014 \u2013 … • · © ® ™ § ¶ † ‡ ° ± × ÷ ≠ ≤ ≥ ≈ ∞ √ ∑ ∏ ∆ µ « » “ ” ‘ ’ ‹ › ¡ ¿'],
  [t('Currency'), '€ £ $ ¥ ₹ ₽ ₩ ₺ ₪ ¢ ₫ ₴ ₦ ₱'],
  [t('Arrows'), '← → ↑ ↓ ↔ ↕ ⇐ ⇒ ⇑ ⇓ ⇔ ↩ ↪ ➜ ➔ ➤'],
  [t('Marks'), '✓ ✔ ✗ ✘ ☐ ☑ ☒ ★ ☆ ♥ ♦ ♣ ♠ ♪ ♫ ☎ ✉ ⚠ ⚡ ☀ ☁ ☂'],
  [t('Greek'), 'α β γ δ ε ζ η θ ι κ λ μ ν ξ ο π ρ σ τ υ φ χ ψ ω Γ Δ Θ Λ Ξ Π Σ Φ Ψ Ω'],
  [t('Fractions and numbers'), '½ ⅓ ⅔ ¼ ¾ ⅛ ⅜ ⅝ ⅞ ¹ ² ³ ⁴ ⁿ ₀ ₁ ₂ ₃ ①  ② ③ ④ ⑤'],
  [t('Letters'), 'à á â ä å æ ç è é ê ë ì í î ï ñ ò ó ô ö ø ù ú û ü ý ÿ ß À Á Â Ä Å Æ Ç È É Ê Ë Ñ Ö Ø Ü'],
];

/**
 * The symbol grid's look travels with the dialog: Worksheets and
 * Presentations open it too, and in their windows the rules that lived in
 * Documents' stylesheet were never loaded, so the grid fell to a column of
 * plain buttons.
 */
const SYMBOL_CSS = `
.wd-symbols { display: grid; grid-template-columns: repeat(12, 1fr); gap: 3px; }
.wd-symbol {
  border: 1px solid var(--line); background: var(--surface); color: var(--ink); border-radius: var(--r-1);
  font-size: 18px; height: 34px; cursor: pointer; font-family: "Segoe UI Symbol", "Segoe UI", system-ui, sans-serif;
}
.wd-symbol:hover { border-color: var(--accent-line); background: var(--hover); }
`;

export function SymbolDialog({ onClose, onInsert }) {
  const [set, setSet] = useState(0);
  const chars = SYMBOLS[set][1].split(/\s+/).filter(Boolean);
  return (
    <Dialog title={t('Symbol')} width={520} onClose={onClose} actions={<Button label={t('Close')} onClick={onClose} />}>
      <style>{SYMBOL_CSS}</style>
      <div className="ml-filters" style={{ flexWrap: 'wrap', marginBottom: 10 }}>
        {SYMBOLS.map(([label], i) => (
          <button key={label} type="button" className={`ml-filter${set === i ? ' on' : ''}`} onClick={() => setSet(i)}>{label}</button>
        ))}
      </div>
      <div className="wd-symbols">
        {chars.map((c, i) => (
          <button key={`${c}${i}`} type="button" className="wd-symbol" title={`U+${c.codePointAt(0).toString(16).toUpperCase().padStart(4, '0')}`} onClick={() => onInsert(c)}>
            {c}
          </button>
        ))}
      </div>
      <p className="rw-hint">{t('Click a symbol to insert it at the caret. The dialog stays open for the next one.')}</p>
    </Dialog>
  );
}
