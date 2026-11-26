// Conditional formatting and data validation: what each dialog's choices
// ask the engine for.
//
// The dialogs speak in their own lists (a condition, its values, how to show
// the cells); `addConditionalRule` and `addValidationRule` speak in
// `{ kind, ... }`. The two met nowhere before, so every rule was refused.
// Kept apart from the dialogs so a test can prove each choice a dialog
// offers is one the engine writes.

/** Each condition the dialog offers: how many values it needs, and whether it shows a look. */
export const CONDITION_RULES = [
  { id: 'greaterThan', operands: 1, look: true },
  { id: 'lessThan', operands: 1, look: true },
  { id: 'between', operands: 2, look: true },
  { id: 'equal', operands: 1, look: true },
  { id: 'notEqual', operands: 1, look: true },
  { id: 'containsText', operands: 1, look: true, text: true },
  { id: 'duplicateValues', operands: 0, look: true },
  { id: 'uniqueValues', operands: 0, look: true },
  // Top and bottom: `count` rules take how many (or what share), 10 when left empty.
  { id: 'top10', operands: 1, look: true, count: true },
  { id: 'top10Percent', operands: 1, look: true, count: true },
  { id: 'bottom10', operands: 1, look: true, count: true },
  { id: 'bottom10Percent', operands: 1, look: true, count: true },
  { id: 'aboveAverage', operands: 0, look: true },
  { id: 'belowAverage', operands: 0, look: true },
  // Excel's "Use a formula to decide which cells to format".
  { id: 'formula', operands: 1, look: true, formula: true },
  { id: 'colorScale', operands: 0, look: false },
  { id: 'colorScaleReverse', operands: 0, look: false },
  { id: 'dataBarBlue', operands: 0, look: false },
  { id: 'dataBarGreen', operands: 0, look: false },
  { id: 'dataBarRed', operands: 0, look: false },
];

/** The dialog's "Show them as" list, in its order, as the engine names each look. */
export const LOOKS = ['red', 'yellow', 'green', 'bold'];

/** The engine's spec for one rule as the dialog holds it. */
export function ruleSpec({ rule, a = '', b = '', style = 0 }) {
  const look = LOOKS[style] ?? 'red';
  switch (rule) {
    case 'greaterThan': case 'lessThan': case 'equal': case 'notEqual':
      return { kind: 'cellIs', operator: rule, values: [a], look };
    case 'between':
      return { kind: 'cellIs', operator: 'between', values: [a, b], look };
    case 'containsText':
      return { kind: 'containsText', text: a, look };
    case 'duplicateValues':
      return { kind: 'duplicates', look };
    case 'uniqueValues':
      return { kind: 'unique', look };
    case 'top10': case 'top10Percent': case 'bottom10': case 'bottom10Percent':
      return { kind: 'top10', rank: String(a).trim() === '' ? 10 : Number(a), bottom: rule.startsWith('bottom'), percent: rule.endsWith('Percent'), look };
    case 'aboveAverage':
      return { kind: 'aboveAverage', look };
    case 'belowAverage':
      return { kind: 'aboveAverage', below: true, look };
    case 'formula':
      return { kind: 'expression', formula: a, look };
    case 'colorScale':
      return { kind: 'colorScale' };
    case 'colorScaleReverse':
      return { kind: 'colorScale', reverse: true };
    case 'dataBarBlue':
      return { kind: 'dataBar', colour: 'blue' };
    case 'dataBarGreen':
      return { kind: 'dataBar', colour: 'green' };
    case 'dataBarRed':
      return { kind: 'dataBar', colour: 'red' };
    default:
      return { kind: rule };
  }
}

/**
 * Whether the dialog's fields are filled for a condition: a count may be
 * left empty (it is then 10) but must be a whole number in its range, a
 * formula and the other values must be there.
 */
export function ruleReady({ rule, a = '', b = '' }) {
  const chosen = CONDITION_RULES.find((c) => c.id === rule);
  if (!chosen) return false;
  if (chosen.count) {
    if (String(a).trim() === '') return true;
    const n = Number(a);
    return Number.isInteger(n) && n >= 1 && n <= (rule.endsWith('Percent') ? 100 : 1000);
  }
  if (chosen.operands > 0 && !String(a).trim()) return false;
  if (chosen.operands > 1 && !String(b).trim()) return false;
  return true;
}

/** The kinds of value the validation dialog's Allow list offers. */
export const VALIDATION_KINDS = ['list', 'whole', 'decimal', 'date', 'textLength'];

/** A range, on this sheet or another one: A1:B2, $H$1:$H$9, Products!A2:A11, 'Price list'!A2:A11. */
const LIST_RANGE = /^=?(?:(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?\$?[A-Z]+\$?\d+(:\$?[A-Z]+\$?\d+)?$/i;

/**
 * The engine's spec for the validation dialog: a list of choices (or the
 * range they come from, on this sheet or another, or after an equals sign a
 * name the workbook defines), or bounds, either or both. One bound alone is a
 * floor or a ceiling, as Excel's "greater than or equal to" and "less than
 * or equal to".
 */
export function validationSpec({ kind, list = '', min = '', max = '', prompt = '' }) {
  const extra = prompt.trim() ? { prompt: prompt.trim() } : {};
  if (kind === 'list') {
    const text = list.trim();
    if (LIST_RANGE.test(text) || /^=[A-Za-z_][\w.]*$/.test(text)) return { kind, source: text.replace(/^=/, ''), ...extra };
    return { kind, items: text.split(',').map((s) => s.trim()).filter(Boolean), ...extra };
  }
  const lo = min.trim();
  const hi = max.trim();
  if (lo && hi) return { kind, operator: 'between', values: [lo, hi], ...extra };
  if (lo) return { kind, operator: 'greaterThanOrEqual', values: [lo], ...extra };
  return { kind, operator: 'lessThanOrEqual', values: [hi], ...extra };
}
