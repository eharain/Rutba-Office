// Many mail accounts at once, from rows of text.
//
// Someone setting up an office, or moving every address of a family, has a
// list: an address and a password a line, written however it was written -
// separated by a space, a tab, a comma, a semicolon, a colon or a bar, pasted
// from a spreadsheet, kept in a .csv or .tsv. This reads all of those.
//
// Two rules make it unambiguous:
//   - A line without a header is an address, then the separator that follows
//     it, then the password: everything after that separator, so a password
//     may hold any character at all, spaces and commas included.
//   - A first line that names its columns (email, password, name, username,
//     the incoming and outgoing servers and ports) makes the rest a table in
//     that line's separator, a quoted field keeping its separators, as a
//     spreadsheet writes CSV.
// A line starting # is a comment. Each row says what is wrong with it, if
// anything, and a password is never put in a message. A line that does not
// start with an address is not shown back at all, only its number: it may be
// a password on a line of its own.

// The part before the @: anything but spaces, separators, double quotes,
// brackets and slashes (an apostrophe is fine: o'brien@). The domain: a host name - labels of letters (any script),
// digits and inner hyphens, dots between them, ending in a name of letters
// (or an xn-- one) - so nothing in it can make a lookup's address another one.
const LOCAL = /^[^\s@,;:|"<>()[\]\\/]{1,64}$/u;
const LABEL = /^[\p{L}\p{N}](?:[\p{L}\p{N}-]{0,61}[\p{L}\p{N}])?$/u;
const TOP = /^(?:\p{L}{2,63}|xn--[a-z0-9-]{1,59})$/iu;

/** Is this a host name a mail domain can be: two labels or more, each a valid one, the last a top-level name. */
export function isDomain(domain) {
  const d = String(domain || '');
  if (d.length > 253) return false;
  const labels = d.split('.');
  return labels.length >= 2 && labels.every((l) => LABEL.test(l)) && TOP.test(labels[labels.length - 1]);
}

/** Is this text an email address. */
export function isAddress(s) {
  const text = String(s || '').trim();
  const at = text.lastIndexOf('@');
  if (at < 1 || text.length > 254) return false;
  return LOCAL.test(text.slice(0, at)) && isDomain(text.slice(at + 1));
}

const HEADERS = {
  email: /^(e-?mail|e-?mail address|address|account|login|user ?name|username|user)$/i,
  password: /^(pass(word)?|pwd|app password|secret)$/i,
  name: /^(name|display name|full name)$/i,
  user: /^(imap user|login name|sign-?in name)$/i,
  imapHost: /^(imap( server| host)?|incoming( server| host)?|imap_server|imap_host|incoming_server)$/i,
  imapPort: /^(imap port|incoming port|imap_port)$/i,
  smtpHost: /^(smtp( server| host)?|outgoing( server| host)?|smtp_server|smtp_host|outgoing_server)$/i,
  smtpPort: /^(smtp port|outgoing port|smtp_port)$/i,
};

/** Split one line of a table by `sep`, quotes keeping their separators ("" is a quote). */
export function splitRow(line, sep) {
  const out = [];
  let field = '';
  let quoted = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (quoted) {
      if (c === '"' && line[i + 1] === '"') {
        field += '"';
        i++;
      } else if (c === '"') quoted = false;
      else field += c;
    } else if (c === '"' && field.trim() === '') {
      quoted = true;
      field = '';
    } else if (sep === ' ' ? /\s/.test(c) : c === sep) {
      out.push(field.trim());
      field = '';
      if (sep === ' ') while (/\s/.test(line[i + 1] || '')) i++;
    } else field += c;
  }
  out.push(field.trim());
  return out;
}

/** The separator a header line uses: the first of tab, comma, semicolon, bar, colon found. */
function headerSeparator(line) {
  for (const sep of ['\t', ',', ';', '|', ':']) if (line.includes(sep)) return sep;
  return ' ';
}

/** The column a header names, or null. */
function columnOf(cell) {
  const s = String(cell || '').trim();
  for (const [key, re] of Object.entries(HEADERS)) if (re.test(s)) return key;
  return null;
}

/** A row's problem, given the addresses above it; and the address noted as seen. */
function problemOf(row, seen) {
  const key = String(row.email || '').toLowerCase();
  let problem = null;
  if (row.unreadable) problem = 'unreadable';
  else if (!isAddress(row.email)) problem = 'not-an-address';
  else if (seen.has(key)) problem = 'repeated';
  else if (!row.password) problem = 'no-password';
  if (!row.unreadable && isAddress(row.email)) seen.add(key);
  return problem;
}

/**
 * Say what is wrong with each of rows read at different times (the box read
 * more than once): a line that could not be read, an address that is not
 * one, one already above, no password.
 * @param {Array<{ email: string, password: string, unreadable?: boolean }>} rows
 */
export function checkRows(rows) {
  const seen = new Set();
  return rows.map((row) => {
    const out = { ...row };
    delete out.problem;
    const problem = problemOf(row, seen);
    if (problem) out.problem = problem;
    return out;
  });
}

/**
 * Read rows of accounts.
 * @param {string} text
 * @returns {{ rows: Array<{ line: number, email: string, password: string, name?: string, user?: string, imap?: object, smtp?: object, problem?: string }>, header: string[]|null, separator: string|null }}
 *   `problem` one of 'not-an-address', 'repeated', 'no-password'
 */
export function parseAccountRows(text) {
  const lines = String(text || '').replace(/^\uFEFF/, '').split(/\r?\n/);
  const rows = [];
  let header = null;
  let sep = null;
  let first = true;
  const seen = new Set();
  lines.forEach((raw, i) => {
    const line = raw.replace(/\s+$/, '');
    if (!line.trim() || /^\s*#/.test(line)) return;
    if (first) {
      first = false;
      const s = headerSeparator(line);
      const cells = splitRow(line, s).map(columnOf);
      if (cells.includes('email') && !isAddress(splitRow(line, s)[cells.indexOf('email')])) {
        header = cells;
        sep = s;
        return;
      }
    }
    let row;
    if (header) {
      const cells = splitRow(line, sep);
      const get = (key) => {
        const at = header.indexOf(key);
        return at >= 0 ? cells[at] || '' : '';
      };
      const email = get('email').trim();
      // A cell that is not address-like (columns out of step, a password where the address goes) is not shown back.
      if (!isAddress(email)) row = { line: i + 1, email: '', password: '', unreadable: true };
      else {
        row = { line: i + 1, email, password: get('password') };
        if (get('name')) row.name = get('name');
        if (get('user')) row.user = get('user');
        if (get('imapHost')) row.imap = { host: get('imapHost'), port: Number(get('imapPort')) || 993 };
        if (get('smtpHost')) row.smtp = { host: get('smtpHost'), port: Number(get('smtpPort')) || 465 };
      }
    } else {
      // An address, its separator, and the rest of the line as the password.
      const trimmed = line.trim();
      const quoted = /^"([^"]*)"\s*([\t,;:| ])\s*(.*)$/.exec(trimmed);
      const m = quoted || /^([^\s,;:|\t]+@[^\s,;:|\t]+?)(?:\s*([\t,;:|])\s*|\s+)(.*)$/.exec(trimmed);
      if (m && isAddress(m[1])) {
        let password = m[3];
        // A password a spreadsheet quoted ("p,w") is the text inside its quotes.
        const q = /^"((?:[^"]|"")*)"$/.exec(password);
        if (q) password = q[1].replace(/""/g, '"');
        row = { line: i + 1, email: m[1].trim(), password };
      } else if (isAddress(trimmed)) row = { line: i + 1, email: trimmed, password: '' };
      // Not an address and a password: its text is never shown back, only that line N could not be read.
      else row = { line: i + 1, email: '', password: '', unreadable: true };
    }
    const problem = problemOf(row, seen);
    if (problem) row.problem = problem;
    rows.push(row);
  });
  return { rows, header: header ? header.filter(Boolean) : null, separator: sep };
}
