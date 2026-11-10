/**
 * The `@` in a formula.
 *
 * A formula is held the way it is shown: `=@A1:A3` answers with the one cell in
 * line with it, `=A1:A3` spills. A file says the same thing another way. A
 * formula stored with no array mark has always meant the first (and carries no
 * `@`), and one with the mark means the second. These turn one into the other,
 * when a file is opened and when it is saved, so the rest of the program deals
 * only in what a person reads.
 */
import { parse } from './parser.js';
import { ARGUMENT_KINDS, ANSWERS_SEVERAL, SEVERAL_WHEN_GIVEN, ANSWERS_A_REFERENCE } from './argument-kinds.js';

const bare = (text) => String(text).replace(/^=/, '');

function kindAt(kinds, i) {
  if (i < kinds.length) return kinds[i];
  if (kinds.length < 2) return kinds[kinds.length - 1];
  return kinds[kinds.length - 2 + ((i - kinds.length) % 2)];
}

const ROWS_AND_COLUMNS = new Set(['ROW', 'COLUMN']);

// What Excel answered: INDEX keeps the @ off where it is given a reference and each place is a whole number written out
// other than 0, or worked out by a function that answers with one value. INDEX(A1:A3,3) and INDEX(B1:B3,MATCH(10,B1:B3,0))
// have none; INDEX(A1:A3,2.7), INDEX(A:A,ROW()), INDEX(A1:A3,0) and INDEX(A1:A3*2,3) have one.
const REFERENCE_NODES = new Set(['cell', 'range', 'name', 'structref', 'spillref']);
function indexMaySpread(node, names) {
  const from = node.args[0];
  if (!from || !(REFERENCE_NODES.has(from.type) || (from.type === 'call' && ANSWERS_A_REFERENCE.has(from.name)))) return true;
  const fixed = (a) => !a || a.type === 'empty'
    || (a.type === 'literal' && typeof a.value === 'number' && Number.isInteger(a.value) && a.value !== 0)
    || (a.type === 'call' && !ROWS_AND_COLUMNS.has(a.name) && !callGivesSeveral(a, names));
  return !fixed(node.args[1]) || !fixed(node.args[2]);
}

function offsetSpreads(node, names) {
  if (node.args[0] && several(node.args[0], names)) return true;
  return [node.args[3], node.args[4]].some((a) => a && !(a.type === 'literal' && a.value === 1));
}

function callGivesSeveral(node, names) {
  if (node.name === 'INDEX') return indexMaySpread(node, names);
  // One value looked up gives one answer, whatever the arrays it looks in: XLOOKUP(20,A1:A3,B1:B3) has no @ in Excel.
  if (node.name === 'XLOOKUP') return Boolean(node.args[0]) && several(node.args[0], names);
  if (node.name === 'OFFSET') return offsetSpreads(node, names);
  if (ANSWERS_SEVERAL.has(node.name)) return true;
  if (!SEVERAL_WHEN_GIVEN.has(node.name)) return false;
  const kinds = ARGUMENT_KINDS.get(node.name);
  return node.args.some((a, i) => (!kinds || kindAt(kinds, i) !== 'V') && several(a, names));
}

function several(node, names) {
  switch (node.type) {
    case 'range':
    case 'array':
    case 'spillref':
      return true;
    case 'name':
      return names(node.name);
    case 'structref':
      return !(node.thisRow && (node.endCol === null || node.endCol === node.startCol));
    case 'call':
      return callGivesSeveral(node, names);
    default:
      return false;
  }
}

const within = (where) => (where === 'array' || where === 'value' ? where : 'operand');

function walk(node, where, names, found) {
  switch (node.type) {
    case 'range':
    case 'spillref':
    case 'name':
    case 'structref':
      if ((where === 'value' || where === 'operand') && several(node, names)) found(node);
      return;
    case 'array':
      if (where === 'value') found(node);
      return;
    case 'unary':
      if (node.op !== '@') walk(node.operand, within(where), names, found);
      return;
    case 'binary':
      walk(node.left, within(where), names, found);
      walk(node.right, within(where), names, found);
      return;
    case 'call': {
      // A LAMBDA's own names may be given ranges, so in its body they are narrowed as a range would be: LAMBDA(a,b,@a+@b).
      if (node.name === 'LAMBDA' && node.args.length) {
        const own = new Set(node.args.slice(0, -1).filter((a) => a.type === 'name').map((a) => a.name.toUpperCase()));
        const inner = (name) => own.has(String(name).toUpperCase()) || names(name);
        walk(node.args[node.args.length - 1], where === 'array' ? 'array' : 'ref', inner, found);
        return;
      }
      const kinds = ARGUMENT_KINDS.get(node.name);
      if (!kinds) return;
      if (callGivesSeveral(node, names)) {
        if (where === 'value') found(node);
        else if (where === 'operand' && ANSWERS_A_REFERENCE.has(node.name)) found(node);
        else if (where === 'ref' && ROWS_AND_COLUMNS.has(node.name)) found(node);
      }
      node.args.forEach((arg, i) => {
        const kind = kindAt(kinds, i);
        let inner = where === 'array' ? 'array' : kind === 'V' ? 'value' : kind === 'R' ? 'ref' : 'array';
        if (kind === 'A') inner = 'array';
        // IF has never worked on a whole array without the array mark, even inside SUMPRODUCT.
        if (node.name === 'IF' && i === 0) inner = 'value';
        walk(arg, inner, names, found);
      });
      return;
    }
    default:
  }
}

const noNames = () => false;

/** Whether a formula read with no array mark would need an `@` anywhere. */
export function needsAt(ast, names = noNames) {
  let any = false;
  walk(ast, 'value', names, () => { any = true; });
  return any;
}

/** The formula, with the leading = as it came, as a formula with no array mark is shown. */
export function addAt(text, names = noNames) {
  const source = String(text);
  const body = bare(source);
  const at = [];
  walk(parse(body, { spans: true }), 'value', names, (node) => at.push(node.span[0]));
  let out = body;
  for (const place of at.sort((a, b) => b - a)) out = out.slice(0, place) + '@' + out.slice(place);
  return (source.startsWith('=') ? '=' : '') + out;
}

function eachAt(node, visit) {
  if (!node || typeof node !== 'object') return;
  if (node.type === 'unary' && node.op === '@') visit(node);
  if (node.type === 'unary') eachAt(node.operand, visit);
  if (node.type === 'binary') { eachAt(node.left, visit); eachAt(node.right, visit); }
  if (node.type === 'call' || node.type === 'invoke') (node.args ?? []).forEach((a) => eachAt(a, visit));
  if (node.type === 'invoke') eachAt(node.callee, visit);
  if (node.type === 'array') node.rows.forEach((row) => row.forEach((a) => eachAt(a, visit)));
}

function withoutAt(body) {
  const at = [];
  eachAt(parse(body, { spans: true }), (node) => at.push(node.span[0]));
  let out = body;
  for (const place of at.sort((a, b) => b - a)) out = out.slice(0, place) + out.slice(place + 1);
  return out;
}

/** `@x` as the file's function: every `@` becomes SINGLE(x). */
function atAsCall(body) {
  const pieces = [];
  eachAt(parse(body, { spans: true }), (node) => {
    pieces.push({ at: node.span[0], drop: 1, put: 'SINGLE(' });
    pieces.push({ at: node.span[1], drop: 0, put: ')' });
  });
  let out = body;
  for (const p of pieces.sort((a, b) => b.at - a.at || a.drop - b.drop)) out = out.slice(0, p.at) + p.put + out.slice(p.at + p.drop);
  return out;
}

/** A file's SINGLE(x) as the `@x` a person reads. */
export function singleAsAt(text) {
  const source = String(text);
  if (!/\bSINGLE\s*\(/i.test(source)) return source;
  let body = bare(source);
  for (let guard = 0; guard < 64; guard++) {
    let hit = null;
    const look = (node) => {
      if (!node || typeof node !== 'object' || hit) return;
      if (node.type === 'call' && node.name === 'SINGLE' && node.args.length === 1) { hit = node; return; }
      if (node.type === 'unary') look(node.operand);
      if (node.type === 'binary') { look(node.left); look(node.right); }
      if (node.type === 'call' || node.type === 'invoke') (node.args ?? []).forEach(look);
      if (node.type === 'array') node.rows.forEach((row) => row.forEach(look));
    };
    try { look(parse(body, { spans: true })); } catch { return source; }
    if (!hit) break;
    const arg = hit.args[0];
    const inner = body.slice(...(arg.outer ?? arg.span));
    const simple = ['cell', 'range', 'name', 'call', 'array', 'structref', 'spillref'].includes(arg.type) || arg.outer;
    body = body.slice(0, hit.span[0]) + '@' + (simple ? inner : '(' + inner + ')') + body.slice(hit.span[1]);
  }
  return (source.startsWith('=') ? '=' : '') + body;
}

/**
 * How a formula goes in a file. `plain`: stored with no array mark it reads back
 * the same, and `text` has no `@` to carry. Otherwise it needs the mark, and any
 * `@` it has is written as the function a file can hold.
 */
export function formulaForFile(text, ast, names = noNames) {
  const body = bare(text);
  if (!ast) return { plain: true, text: '=' + body };
  if (!body.includes('@')) return { plain: !needsAt(ast, names), text: '=' + body };
  let stripped;
  let spans;
  try { stripped = withoutAt(body); spans = addAt(stripped, names); } catch { return { plain: true, text: '=' + body }; }
  if (spans === body) return { plain: true, text: '=' + stripped };
  if (stripped === body) return { plain: !needsAt(ast, names), text: '=' + body };
  return { plain: false, text: '=' + atAsCall(body) };
}

/** A range or array down to the one value in line with the cell asking; the first when it has no place of its own. */
export function oneInLine(raw, row, col, onMissing) {
  if (!Array.isArray(raw)) return raw;
  const grid = Array.isArray(raw[0]) ? raw : [raw];
  const at = raw.origin;
  if (!at || row === undefined || col === undefined) return grid[0]?.length ? grid[0][0] : '';
  const rows = raw.full?.rows ?? grid.length;
  const cols = raw.full?.cols ?? (grid[0]?.length ?? 0);
  const y = rows === 1 ? 0 : row - at.row;
  const x = cols === 1 ? 0 : col - at.col;
  if (y < 0 || y >= rows || x < 0 || x >= cols) return onMissing();
  return grid[y + at.row - at.top]?.[x + at.col - at.left] ?? null;
}
