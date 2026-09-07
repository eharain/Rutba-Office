/**
 * The function catalog: what each implemented function is called, what it
 * takes, and where it belongs in a "Formulas" menu.
 *
 * This file exists for TOOLBARS, not for the engine — the engine's truth is
 * `FUNCTIONS` in functions.js. The two are held in lock-step by a test that
 * fails when a function exists without a catalog entry or a catalog entry
 * without a function, which is what lets a ribbon render this list and honour
 * the house rule that a control which cannot act is not shown.
 *
 * `args` is display text for a person, not a parseable grammar. Optional
 * arguments are in brackets, repeatable ones end with an ellipsis.
 */

export const CATEGORIES = [
  'Math & Trig',
  'Statistical',
  'Logical',
  'Text',
  'Lookup & Reference',
  'Date & Time',
  'Financial',
  'Information',
];

const entry = (name, args, help, category) => ({ name, args, help, category });
const math = (name, args, help) => entry(name, args, help, 'Math & Trig');
const stat = (name, args, help) => entry(name, args, help, 'Statistical');
const logic = (name, args, help) => entry(name, args, help, 'Logical');
const text = (name, args, help) => entry(name, args, help, 'Text');
const lookup = (name, args, help) => entry(name, args, help, 'Lookup & Reference');
const date = (name, args, help) => entry(name, args, help, 'Date & Time');
const fin = (name, args, help) => entry(name, args, help, 'Financial');
const info = (name, args, help) => entry(name, args, help, 'Information');

export const FUNCTION_CATALOG = [
  // ---- Math & Trig --------------------------------------------------------
  math('SUM', '(number1, …)', 'Add numbers and ranges'),
  math('PRODUCT', '(number1, …)', 'Multiply numbers and ranges'),
  math('ABS', '(number)', 'Absolute value'),
  math('SIGN', '(number)', '-1, 0 or 1 by the sign'),
  math('SQRT', '(number)', 'Square root'),
  math('POWER', '(number, power)', 'A number raised to a power'),
  math('EXP', '(number)', 'e raised to a power'),
  math('LN', '(number)', 'Natural logarithm'),
  math('LOG', '(number, [base])', 'Logarithm, base 10 unless given'),
  math('LOG10', '(number)', 'Base-10 logarithm'),
  math('PI', '()', 'The constant π'),
  math('MOD', '(number, divisor)', 'Remainder, with the divisor’s sign'),
  math('INT', '(number)', 'Round down to an integer'),
  math('TRUNC', '(number, [digits])', 'Cut off decimals without rounding'),
  math('ROUND', '(number, [digits])', 'Round half away from zero'),
  math('ROUNDUP', '(number, [digits])', 'Round away from zero'),
  math('ROUNDDOWN', '(number, [digits])', 'Round toward zero'),
  math('CEILING', '(number, [significance])', 'Round up to a multiple'),
  math('FLOOR', '(number, [significance])', 'Round down to a multiple'),
  math('MROUND', '(number, multiple)', 'Round to the nearest multiple'),
  math('EVEN', '(number)', 'Round away from zero to an even integer'),
  math('ODD', '(number)', 'Round away from zero to an odd integer'),
  math('SUMSQ', '(number1, …)', 'Sum of squares'),
  math('SUMPRODUCT', '(array1, array2, …)', 'Multiply arrays position by position and add'),
  math('SUMIF', '(range, criteria, [sum_range])', 'Add the cells a criterion selects'),
  math('SUMIFS', '(sum_range, range1, criteria1, …)', 'Add cells matching every criterion'),
  math('RAND', '()', 'Random number between 0 and 1'),
  math('RANDBETWEEN', '(bottom, top)', 'Random whole number in a range'),
  math('SEQUENCE', '(rows, [cols], [start], [step])', 'A spilled run of numbers'),

  // ---- Statistical --------------------------------------------------------
  stat('AVERAGE', '(number1, …)', 'Arithmetic mean'),
  stat('AVERAGEIF', '(range, criteria, [average_range])', 'Mean of the cells a criterion selects'),
  stat('AVERAGEIFS', '(average_range, range1, criteria1, …)', 'Mean of cells matching every criterion'),
  stat('MIN', '(number1, …)', 'Smallest number'),
  stat('MAX', '(number1, …)', 'Largest number'),
  stat('MAXIFS', '(range, range1, criteria1, …)', 'Largest value among matching cells'),
  stat('MINIFS', '(range, range1, criteria1, …)', 'Smallest value among matching cells'),
  stat('COUNT', '(value1, …)', 'How many numbers'),
  stat('COUNTA', '(value1, …)', 'How many non-empty cells'),
  stat('COUNTBLANK', '(range)', 'How many empty cells'),
  stat('COUNTIF', '(range, criteria)', 'How many cells match a criterion'),
  stat('COUNTIFS', '(range1, criteria1, …)', 'How many rows match every criterion'),
  stat('MEDIAN', '(number1, …)', 'Middle value'),
  stat('MODE', '(number1, …)', 'Most frequent value'),
  stat('MODE.SNGL', '(number1, …)', 'Most frequent value'),
  stat('STDEV', '(number1, …)', 'Sample standard deviation'),
  stat('STDEV.S', '(number1, …)', 'Sample standard deviation'),
  stat('STDEVP', '(number1, …)', 'Population standard deviation'),
  stat('STDEV.P', '(number1, …)', 'Population standard deviation'),
  stat('VAR', '(number1, …)', 'Sample variance'),
  stat('VAR.S', '(number1, …)', 'Sample variance'),
  stat('VARP', '(number1, …)', 'Population variance'),
  stat('VAR.P', '(number1, …)', 'Population variance'),
  stat('LARGE', '(range, k)', 'k-th largest value'),
  stat('SMALL', '(range, k)', 'k-th smallest value'),
  stat('RANK', '(number, range, [order])', 'Rank within the data'),
  stat('RANK.EQ', '(number, range, [order])', 'Rank within the data'),

  // ---- Logical ------------------------------------------------------------
  logic('IF', '(condition, when_true, [when_false])', 'One value or the other'),
  logic('IFS', '(condition1, value1, …)', 'The first value whose condition holds'),
  logic('SWITCH', '(value, case1, result1, …, [default])', 'Match a value against cases'),
  logic('AND', '(logical1, …)', 'TRUE when everything is true'),
  logic('OR', '(logical1, …)', 'TRUE when anything is true'),
  logic('XOR', '(logical1, …)', 'TRUE when an odd number are true'),
  logic('NOT', '(logical)', 'The opposite'),
  logic('TRUE', '()', 'The value TRUE'),
  logic('FALSE', '()', 'The value FALSE'),
  logic('IFERROR', '(value, [fallback])', 'A fallback when the value is an error'),
  logic('IFNA', '(value, [fallback])', 'A fallback when the value is #N/A'),

  // ---- Text ---------------------------------------------------------------
  text('CONCAT', '(text1, …)', 'Join text'),
  text('CONCATENATE', '(text1, …)', 'Join text'),
  text('TEXTJOIN', '(delimiter, ignore_empty, text1, …)', 'Join text with a separator'),
  text('LEN', '(text)', 'Length of the text'),
  text('LEFT', '(text, [count])', 'Leading characters'),
  text('RIGHT', '(text, [count])', 'Trailing characters'),
  text('MID', '(text, start, count)', 'Characters from the middle'),
  text('UPPER', '(text)', 'To upper case'),
  text('LOWER', '(text)', 'To lower case'),
  text('PROPER', '(text)', 'Capitalise Each Word'),
  text('TRIM', '(text)', 'Collapse runs of spaces'),
  text('SUBSTITUTE', '(text, old, new)', 'Replace every occurrence of a text'),
  text('REPLACE', '(text, start, count, new)', 'Replace characters by position'),
  text('REPT', '(text, times)', 'Repeat text'),
  text('FIND', '(find, within, [start])', 'Position of text, case-sensitive'),
  text('SEARCH', '(find, within, [start])', 'Position of text, case-insensitive, wildcards'),
  text('EXACT', '(text1, text2)', 'Case-sensitive equality'),
  text('CHAR', '(number)', 'The character for a code'),
  text('CODE', '(text)', 'The code of the first character'),
  text('VALUE', '(text)', 'Text to a number'),
  text('TEXT', '(value, format)', 'A number as formatted text'),

  // ---- Lookup & Reference -------------------------------------------------
  lookup('VLOOKUP', '(value, table, column, [approximate])', 'Find a row by its first column'),
  lookup('HLOOKUP', '(value, table, row, [approximate])', 'Find a column by its first row'),
  lookup('XLOOKUP', '(value, lookup_range, return_range, [if_not_found], [match_mode], [search_mode])', 'Modern lookup in any direction'),
  lookup('INDEX', '(range, row, [column])', 'The value at a position'),
  lookup('MATCH', '(value, range, [match_type])', 'The position of a value'),
  lookup('CHOOSE', '(index, value1, …)', 'Pick a value by number'),
  lookup('ROW', '([reference])', 'The row number, of a reference or of this cell'),
  lookup('COLUMN', '([reference])', 'The column number, of a reference or of this cell'),
  lookup('ROWS', '(range)', 'How many rows a range spans'),
  lookup('COLUMNS', '(range)', 'How many columns a range spans'),
  lookup('OFFSET', '(anchor, rows, cols, [height], [width])', 'A range moved and resized from an anchor'),
  lookup('INDIRECT', '(text, [a1])', 'The reference a text string names'),
  lookup('ADDRESS', '(row, column, [abs], [a1], [sheet])', 'A cell address, as text'),
  lookup('TRANSPOSE', '(array)', 'Rows become columns'),
  lookup('UNIQUE', '(array, [by_col], [exactly_once])', 'The distinct rows, spilled'),
  lookup('SORT', '(array, [index], [order], [by_col])', 'The rows sorted, spilled'),
  lookup('FILTER', '(array, include, [if_empty])', 'The rows a condition keeps, spilled'),

  // ---- Date & Time --------------------------------------------------------
  date('TODAY', '()', 'Today’s date'),
  date('NOW', '()', 'The current date and time'),
  date('DATE', '(year, month, day)', 'A date from its parts'),
  date('YEAR', '(date)', 'The year of a date'),
  date('MONTH', '(date)', 'The month of a date'),
  date('DAY', '(date)', 'The day of a date'),
  date('EDATE', '(start, months)', 'The same day, months away'),
  date('EOMONTH', '(start, months)', 'The last day of a month, months away'),
  date('DAYS', '(end, start)', 'Days between two dates'),
  date('WEEKDAY', '(date, [type])', 'Day of the week as a number'),
  date('HOUR', '(time)', 'The hour of a time'),
  date('MINUTE', '(time)', 'The minute of a time'),
  date('SECOND', '(time)', 'The second of a time'),
  date('TIME', '(hour, minute, second)', 'A time from its parts'),

  // ---- Financial ----------------------------------------------------------
  fin('PMT', '(rate, nper, pv, [fv], [type])', 'Payment for a loan'),
  fin('FV', '(rate, nper, pmt, [pv], [type])', 'Future value of an investment'),
  fin('PV', '(rate, nper, pmt, [fv], [type])', 'Present value of an investment'),
  fin('NPER', '(rate, pmt, pv, [fv], [type])', 'Number of payment periods'),
  fin('NPV', '(rate, value1, …)', 'Net present value of cash flows'),

  // ---- Information --------------------------------------------------------
  info('ISBLANK', '(value)', 'Is the cell empty'),
  info('ISNUMBER', '(value)', 'Is it a number'),
  info('ISTEXT', '(value)', 'Is it text'),
  info('ISLOGICAL', '(value)', 'Is it TRUE or FALSE'),
  info('ISERROR', '(value)', 'Is it any error'),
  info('ISNA', '(value)', 'Is it #N/A'),
  info('ISEVEN', '(number)', 'Is the integer part even'),
  info('ISODD', '(number)', 'Is the integer part odd'),
  info('NA', '()', 'The error #N/A'),
  info('N', '(value)', 'A value as a number, text as 0'),
];

/** The catalog grouped by category, in CATEGORIES order, for a menu to render. */
export function catalogByCategory() {
  const groups = new Map(CATEGORIES.map((c) => [c, []]));
  for (const fn of FUNCTION_CATALOG) groups.get(fn.category).push(fn);
  return CATEGORIES.map((category) => ({ category, functions: groups.get(category) }));
}
