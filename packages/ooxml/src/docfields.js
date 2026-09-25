/**
 * Insert → Quick Parts → Field: the document fields a person puts in the
 * body by hand — PAGE, NUMPAGES, DATE, TIME, FILENAME, AUTHOR and TITLE —
 * their codes as Word's Field dialog writes them, and their results worked
 * out the way Word works them out. Nothing here touches a package.
 */

/** The fields the dialog lists, in Word's order within its categories, with each one's formats. */
export const DOC_FIELDS = [
  { name: 'AUTHOR', category: 'Document Information', description: 'The name of the document\'s author, from its properties', formats: 'text' },
  { name: 'DATE', category: 'Date and Time', description: 'Today\'s date', formats: 'date' },
  { name: 'FILENAME', category: 'Document Information', description: 'The document\'s file name', formats: 'text', path: true },
  { name: 'NUMPAGES', category: 'Document Information', description: 'The number of pages in the document', formats: 'number' },
  { name: 'PAGE', category: 'Numbering', description: 'The number of the page it is on', formats: 'number' },
  { name: 'TIME', category: 'Date and Time', description: 'The time now', formats: 'time' },
  { name: 'TITLE', category: 'Document Information', description: 'The document\'s title, from its properties', formats: 'text' },
];

/** Word's date pictures in the Field dialog, and the time ones. */
export const DATE_PICTURES = [
  'dd/MM/yyyy', 'dddd, d MMMM yyyy', 'd MMMM yyyy', 'dd/MM/yy', 'yyyy-MM-dd', 'd-MMM-yy', 'dd.MM.yyyy', 'd MMM. yy',
  'M/d/yyyy', 'dddd, MMMM d, yyyy', 'MMMM d, yyyy', 'MMMM yy', 'MMM-yy',
];
export const TIME_PICTURES = ['HH:mm', 'HH:mm:ss', 'h:mm am/pm', 'h:mm:ss am/pm', 'dd/MM/yyyy HH:mm'];
/** Number formats, as `\*` switches, with what each shows for 1. */
export const NUMBER_FORMATS = [
  ['', '1, 2, 3, …'], ['roman', 'i, ii, iii, …'], ['ROMAN', 'I, II, III, …'], ['alphabetic', 'a, b, c, …'], ['ALPHABETIC', 'A, B, C, …'],
];
/** Text formats, as `\*` switches. */
export const TEXT_FORMATS = [
  ['', '(none)'], ['Upper', 'Uppercase'], ['Lower', 'Lowercase'], ['FirstCap', 'First capital'], ['Caps', 'Title case'],
];

/** A field's code: ` DATE \@ "d MMMM yyyy" \* MERGEFORMAT `, the shape Word's dialog writes. */
export function docFieldInstr({ name, picture = '', format = '', path = false } = {}) {
  const n = String(name || '').toUpperCase();
  let instr = ' ' + n + ' ';
  if ((n === 'DATE' || n === 'TIME') && picture) instr += '\\@ "' + picture + '" ';
  if (n === 'FILENAME' && path) instr += '\\p ';
  if (format) instr += '\\* ' + format + ' ';
  return instr + '\\* MERGEFORMAT ';
}

/** A field's code read back: `{ name, picture, format, path }`. */
export function parseDocFieldInstr(instr) {
  const s = String(instr);
  const name = (/^\s*(\S+)/.exec(s)?.[1] || '').toUpperCase();
  const picture = /\\@\s*"([^"]*)"/.exec(s)?.[1] ?? /\\@\s*(\S+)/.exec(s)?.[1] ?? '';
  const formats = [...s.matchAll(/\\\*\s*(\S+)/g)].map((m) => m[1]).filter((f) => f.toUpperCase() !== 'MERGEFORMAT' && f.toUpperCase() !== 'CHARFORMAT');
  return { name, picture, format: formats[0] || '', path: /\\p\b/i.test(s) };
}

const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

/** A date in one of Word's `\@` pictures: d dd ddd dddd M MM MMM MMMM yy yyyy h hh H HH m mm s ss am/pm, 'quoted text'. */
export function formatDatePicture(date, picture) {
  const d = date instanceof Date ? date : new Date(date);
  const pic = String(picture || 'dd/MM/yyyy');
  const ampm = /am\/pm|AM\/PM/.test(pic);
  let out = '';
  let i = 0;
  const run = (ch) => { let j = i; while (pic[j] === ch) j++; const n = j - i; i = j; return n; };
  while (i < pic.length) {
    const c = pic[i];
    if (c === "'") {
      const j = pic.indexOf("'", i + 1);
      out += pic.slice(i + 1, j < 0 ? pic.length : j);
      i = j < 0 ? pic.length : j + 1;
      continue;
    }
    if (/^am\/pm/i.test(pic.slice(i))) {
      const pm = d.getHours() >= 12;
      out += pic.slice(i, i + 5) === 'am/pm' ? (pm ? 'pm' : 'am') : (pm ? 'PM' : 'AM');
      i += 5;
      continue;
    }
    if (c === 'd') {
      const n = run('d');
      out += n === 1 ? d.getDate() : n === 2 ? String(d.getDate()).padStart(2, '0') : n === 3 ? DAYS[d.getDay()].slice(0, 3) : DAYS[d.getDay()];
    } else if (c === 'M') {
      const n = run('M');
      out += n === 1 ? d.getMonth() + 1 : n === 2 ? String(d.getMonth() + 1).padStart(2, '0') : n === 3 ? MONTHS[d.getMonth()].slice(0, 3) : MONTHS[d.getMonth()];
    } else if (c === 'y') {
      const n = run('y');
      out += n <= 2 ? String(d.getFullYear()).slice(-2) : String(d.getFullYear());
    } else if (c === 'h') {
      const n = run('h');
      const h = d.getHours() % 12 || 12;
      out += n === 1 ? h : String(h).padStart(2, '0');
    } else if (c === 'H') {
      const n = run('H');
      out += n === 1 ? (ampm ? d.getHours() % 12 || 12 : d.getHours()) : String(d.getHours()).padStart(2, '0');
    } else if (c === 'm') {
      const n = run('m');
      out += n === 1 ? d.getMinutes() : String(d.getMinutes()).padStart(2, '0');
    } else if (c === 's') {
      const n = run('s');
      out += n === 1 ? d.getSeconds() : String(d.getSeconds()).padStart(2, '0');
    } else {
      out += c;
      i++;
    }
  }
  return out;
}

const ROMAN = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
/** A number in a `\*` numbering format: Arabic, roman, ROMAN, alphabetic, ALPHABETIC. */
export function formatNumber(n, format = '') {
  const v = Math.max(0, Math.floor(Number(n) || 0));
  if (format === 'roman' || format === 'ROMAN') {
    let x = v;
    let s = '';
    for (const [k, r] of ROMAN) while (x >= k) { s += r; x -= k; }
    return format === 'ROMAN' ? s.toUpperCase() : s;
  }
  if (format === 'alphabetic' || format === 'ALPHABETIC') {
    if (!v) return '';
    // Word repeats the letter: 27 is aa, 28 bb.
    const s = String.fromCharCode(97 + ((v - 1) % 26)).repeat(Math.floor((v - 1) / 26) + 1);
    return format === 'ALPHABETIC' ? s.toUpperCase() : s;
  }
  return String(v);
}

/** Text in a `\*` text format. */
export function formatText(text, format = '') {
  const t = String(text ?? '');
  switch (String(format).toLowerCase()) {
    case 'upper': return t.toUpperCase();
    case 'lower': return t.toLowerCase();
    case 'firstcap': return t.replace(/^\s*\p{L}/u, (m) => m.toUpperCase());
    case 'caps': return t.replace(/(^|\s)(\p{L})/gu, (m, a, b) => a + b.toUpperCase());
    default: return t;
  }
}

/**
 * A field's result from what the document knows: `page` (the page it is on),
 * `pages` (how many), `now`, `fileName`, `filePath`, `author`, `title`.
 * Undefined when the field cannot be worked out here — its result then stays.
 */
export function evaluateDocField(instr, ctx = {}) {
  const { name, picture, format, path } = parseDocFieldInstr(instr);
  switch (name) {
    case 'PAGE': return ctx.page == null ? undefined : formatNumber(ctx.page, format);
    case 'NUMPAGES': return ctx.pages == null ? undefined : formatNumber(ctx.pages, format);
    case 'DATE': return formatDatePicture(ctx.now ?? new Date(), picture || 'dd/MM/yyyy');
    case 'TIME': return formatDatePicture(ctx.now ?? new Date(), picture || 'HH:mm');
    case 'FILENAME': {
      const file = path ? ctx.filePath || ctx.fileName : ctx.fileName;
      return file ? formatText(file, format) : undefined;
    }
    case 'AUTHOR': return ctx.author == null ? undefined : formatText(ctx.author, format);
    case 'TITLE': return ctx.title == null ? undefined : formatText(ctx.title, format);
    default: return undefined;
  }
}
