/**
 * A filter by a rule: a number greater than a figure, a text that contains a
 * word, a date in a span, the top ten. Excel keeps such a filter in the
 * file as `customFilters` or `top10`; this module turns the rule a person
 * picks into that form, turns it back, and says whether a cell passes.
 * Pure: no workbook, no view.
 *
 * A rule as a person picks it:
 *   { kind: 'number', op, value, value2 }   op from FILTER_OPS.number
 *   { kind: 'date',   op, value, value2 }   value a date's number, op from FILTER_OPS.date
 *   { kind: 'text',   op, value }           op from FILTER_OPS.text
 *   { kind: 'top',    bottom, count, percent }
 * A rule as the file holds it:
 *   { custom: { and, tests: [{ operator, val }] } }
 *   { top10: { top, percent, val, filterVal } }
 *   { other: true }                         a kind kept as it is and not read
 */

/** The comparisons offered for each kind of column, in the order Excel lists them. */
export const FILTER_OPS = {
  number: ['equals', 'notEquals', 'greater', 'greaterOrEqual', 'less', 'lessOrEqual', 'between'],
  text: ['equals', 'notEquals', 'begins', 'ends', 'contains', 'notContains'],
  date: ['equals', 'before', 'after', 'between'],
};

const NUMBER_OPERATOR = {
  equals: 'equal', notEquals: 'notEqual', greater: 'greaterThan', greaterOrEqual: 'greaterThanOrEqual',
  less: 'lessThan', lessOrEqual: 'lessThanOrEqual',
};
const NUMBER_OP = Object.fromEntries(Object.entries(NUMBER_OPERATOR).map(([op, operator]) => [operator, op]));

/** A number as the file writes one: no exponent for the figures people type. */
const written = (n) => String(Number(n));

/** The file's form of a rule a person picked. Throws a plain sentence for a rule that cannot be one. */
export function ruleToFile(rule) {
  const kind = rule?.kind;
  if (kind === 'top') {
    const count = Math.floor(Number(rule.count));
    if (!Number.isFinite(count) || count < 1) throw new Error('Type how many to keep: 1 or more.');
    if (rule.percent && count > 100) throw new Error('A percentage is 100 or less.');
    return { top10: { top: !rule.bottom, percent: Boolean(rule.percent), val: count, filterVal: null } };
  }
  if (kind === 'text') {
    const text = String(rule.value ?? '');
    if (!FILTER_OPS.text.includes(rule.op)) throw new Error('"' + rule.op + '" is not a text filter');
    if (!text && rule.op !== 'equals' && rule.op !== 'notEquals') throw new Error('Type the text to look for.');
    const test = {
      equals: { operator: 'equal', val: text },
      notEquals: { operator: 'notEqual', val: text },
      begins: { operator: 'equal', val: text + '*' },
      ends: { operator: 'equal', val: '*' + text },
      contains: { operator: 'equal', val: '*' + text + '*' },
      notContains: { operator: 'notEqual', val: '*' + text + '*' },
    }[rule.op];
    return { custom: { and: false, tests: [test] } };
  }
  if (kind === 'number' || kind === 'date') {
    if (!FILTER_OPS[kind].includes(rule.op)) throw new Error('"' + rule.op + '" is not a ' + kind + ' filter');
    const a = Number(rule.value);
    if (rule.value === '' || rule.value == null || !Number.isFinite(a)) {
      throw new Error(kind === 'date' ? 'Type a date to compare with.' : 'Type a number to compare with.');
    }
    if (rule.op === 'between') {
      const b = Number(rule.value2);
      if (rule.value2 === '' || rule.value2 == null || !Number.isFinite(b)) {
        throw new Error(kind === 'date' ? 'Type both dates.' : 'Type both numbers.');
      }
      const [low, high] = a <= b ? [a, b] : [b, a];
      // A date's last day counts to its end: a time on that day is still inside.
      return {
        custom: {
          and: true,
          tests: kind === 'date'
            ? [{ operator: 'greaterThanOrEqual', val: written(Math.floor(low)) }, { operator: 'lessThan', val: written(Math.floor(high) + 1) }]
            : [{ operator: 'greaterThanOrEqual', val: written(low) }, { operator: 'lessThanOrEqual', val: written(high) }],
        },
      };
    }
    if (kind === 'date') {
      const day = Math.floor(a);
      if (rule.op === 'equals') {
        return { custom: { and: true, tests: [{ operator: 'greaterThanOrEqual', val: written(day) }, { operator: 'lessThan', val: written(day + 1) }] } };
      }
      // After a day is after its end; before a day is before its start.
      return { custom: { and: false, tests: [rule.op === 'before' ? { operator: 'lessThan', val: written(day) } : { operator: 'greaterThanOrEqual', val: written(day + 1) }] } };
    }
    return { custom: { and: false, tests: [{ operator: NUMBER_OPERATOR[rule.op], val: written(a) }] } };
  }
  throw new Error('not a filter rule');
}

/**
 * The rule a filled form asks for, as `applyFilterRule` takes one, or null while
 * the form is not filled. A number and a date go as they were typed: the
 * view reads them as a cell would. `type` is what the column holds; `op`
 * may also be `top` or `bottom`.
 */
export function ruleFromForm(type, { op, a = '', b = '', count = '', percent = false }) {
  if (op === 'top' || op === 'bottom') {
    const n = Math.floor(Number(count));
    if (!Number.isFinite(n) || n < 1 || (percent && n > 100)) return null;
    return { kind: 'top', bottom: op === 'bottom', count: n, percent: Boolean(percent) };
  }
  if (type === 'text') {
    // Equals nothing is a filter for the blanks; the other comparisons need a word.
    if (a === '' && op !== 'equals' && op !== 'notEquals') return null;
    return { kind: 'text', op, value: a };
  }
  if (a.trim() === '' || (op === 'between' && b.trim() === '')) return null;
  return op === 'between' ? { kind: type, op, value: a.trim(), value2: b.trim() } : { kind: type, op, value: a.trim() };
}

const numberOf = (val) => {
  const raw = String(val ?? '').trim();
  return raw !== '' && Number.isFinite(Number(raw)) ? Number(raw) : null;
};

/**
 * The rule a person would have picked to get what the file holds, for a
 * column of `type` (`number`, `text` or `date`); null for a filter the menu
 * has no form for, which is then only shown as being there.
 */
export function fileToRule(file, type = 'text') {
  if (file?.top10) {
    return { kind: 'top', bottom: !file.top10.top, count: Number(file.top10.val) || 10, percent: Boolean(file.top10.percent) };
  }
  const tests = file?.custom?.tests ?? [];
  if (tests.length === 1) {
    const { operator = 'equal', val } = tests[0];
    const n = numberOf(val);
    if (n !== null && type !== 'text') {
      if (type === 'date') {
        if (operator === 'lessThan') return { kind: 'date', op: 'before', value: n };
        if (operator === 'greaterThanOrEqual') return { kind: 'date', op: 'after', value: n - 1 };
        if (operator === 'greaterThan') return { kind: 'date', op: 'after', value: n };
        return null;
      }
      return NUMBER_OP[operator] ? { kind: 'number', op: NUMBER_OP[operator], value: n } : null;
    }
    const text = String(val ?? '');
    if (operator !== 'equal' && operator !== 'notEqual') return null;
    const not = operator === 'notEqual';
    const inner = /^\*([^*?~]+)\*$/.exec(text);
    if (inner) return { kind: 'text', op: not ? 'notContains' : 'contains', value: inner[1] };
    if (/[*?~]/.test(text) && !not) {
      const begins = /^([^*?~]+)\*$/.exec(text);
      if (begins) return { kind: 'text', op: 'begins', value: begins[1] };
      const ends = /^\*([^*?~]+)$/.exec(text);
      if (ends) return { kind: 'text', op: 'ends', value: ends[1] };
    }
    return { kind: 'text', op: not ? 'notEquals' : 'equals', value: text };
  }
  if (tests.length === 2 && file.custom.and) {
    const low = tests.find((t) => t.operator === 'greaterThanOrEqual');
    const high = tests.find((t) => t.operator === 'lessThanOrEqual' || t.operator === 'lessThan');
    const a = numberOf(low?.val);
    const b = numberOf(high?.val);
    if (a === null || b === null || type === 'text') return null;
    if (type === 'date') {
      const last = high.operator === 'lessThan' ? b - 1 : b;
      return last === a ? { kind: 'date', op: 'equals', value: a } : { kind: 'date', op: 'between', value: a, value2: last };
    }
    return high.operator === 'lessThanOrEqual' ? { kind: 'number', op: 'between', value: a, value2: b } : null;
  }
  return null;
}

/** A filter's pattern, with Excel's `*`, `?` and `~`, as a test of a whole text without regard to case. */
function patternTest(pattern) {
  let source = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length) source += pattern[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    else if (ch === '*') source += '[\\s\\S]*';
    else if (ch === '?') source += '[\\s\\S]';
    else source += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  const re = new RegExp('^' + source + '$', 'i');
  return (text) => re.test(text);
}

/** Whether one cell, by its value and the text it shows, passes one test of a custom filter. */
function testPasses(test, value, text) {
  const operator = test.operator || 'equal';
  const n = numberOf(test.val);
  if (n !== null && typeof value === 'number') {
    switch (operator) {
      case 'equal': return value === n;
      case 'notEqual': return value !== n;
      case 'greaterThan': return value > n;
      case 'greaterThanOrEqual': return value >= n;
      case 'lessThan': return value < n;
      case 'lessThanOrEqual': return value <= n;
      default: return false;
    }
  }
  const shown = String(text ?? '');
  const pattern = String(test.val ?? '');
  if (operator === 'equal' || operator === 'notEqual') {
    // A blank asked for (Excel writes a space) is a cell that shows nothing.
    const hit = pattern.trim() === '' ? shown.trim() === '' : patternTest(pattern)(shown);
    return operator === 'equal' ? hit : !hit;
  }
  // A number asked for and a cell that holds none: no order between them.
  if (n !== null || shown === '') return false;
  const order = shown.localeCompare(pattern, undefined, { sensitivity: 'accent' });
  switch (operator) {
    case 'greaterThan': return order > 0;
    case 'greaterThanOrEqual': return order >= 0;
    case 'lessThan': return order < 0;
    case 'lessThanOrEqual': return order <= 0;
    default: return false;
  }
}

/**
 * The figure a top or bottom filter cuts at, for the column's numbers: the
 * smallest of the top `val` (or `val` percent), or the largest of the bottom.
 * Ties with it stay in, as Excel keeps them. Null when the column has no number.
 */
export function topThreshold(numbers, top10) {
  const sorted = numbers.filter((n) => typeof n === 'number' && Number.isFinite(n)).sort((a, b) => a - b);
  if (!sorted.length) return null;
  const asked = Math.max(1, Math.floor(Number(top10.val) || 1));
  const n = Math.min(sorted.length, top10.percent ? Math.max(1, Math.floor((sorted.length * asked) / 100)) : asked);
  return top10.top === false ? sorted[n - 1] : sorted[sorted.length - n];
}

/**
 * Whether a cell passes a file rule. `threshold` is `topThreshold` of the
 * column, for a top or bottom rule. A rule that is not read (`other`) lets
 * every row pass: its rows stay as the file hid them.
 */
export function rulePasses(file, value, text, threshold = null) {
  if (file?.top10) {
    if (typeof value !== 'number' || threshold === null) return false;
    return file.top10.top === false ? value <= threshold : value >= threshold;
  }
  const tests = file?.custom?.tests ?? [];
  if (!tests.length) return true;
  return file.custom.and
    ? tests.every((t) => testPasses(t, value, text))
    : tests.some((t) => testPasses(t, value, text));
}
