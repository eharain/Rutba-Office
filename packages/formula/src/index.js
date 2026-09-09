/**
 * @rutba/formula — the calculation engine.
 *
 * Written from the published, observable behaviour of spreadsheets: A1
 * references, operator precedence, error values, coercion rules and the 1900
 * date serial. None of that is anyone's intellectual property — it is the
 * shared convention every spreadsheet has implemented for forty years, and the
 * file format that carries it is an ISO standard.
 *
 * Layers, each usable alone:
 *   values.js     coercion, comparison, error values, date serials
 *   parser.js     tokenizer, precedence-correct parser, reference parsing
 *   functions.js  the function library
 *   evaluator.js  AST + resolver -> value
 *   engine.js     sheet model, dependency graph, ordered recalculation
 */
export {
  FormulaError, ERR, ERROR_TYPES, isError, isBlank, firstError,
  toNumber, toText, toBoolean, compareValues, formatNumber,
  roundHalfAwayFromZero, dateToSerial, serialToDate, serialToParts, partsToSerial, numericText,
} from './values.js';

export {
  parse, tokenize, dependencies, parseReference, colToIndex, indexToCol, TOKEN,
  shiftFormula,
} from './parser.js';

export { FUNCTIONS, FUNCTION_NAMES, isVolatile } from './functions.js';
export { FUNCTION_CATALOG, CATEGORIES, catalogByCategory } from './catalog.js';
export { evaluate, calculate } from './evaluator.js';
export { Spreadsheet, prettyKey } from './engine.js';
