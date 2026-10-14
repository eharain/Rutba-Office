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
// anything, and a password is never put in a message.

const EMAIL = /^[^\s@,;:|"'<>()]+@[^\s@,;:|"'<>()]+\.[^\s@,;:|"'<>()]+$/;

/** Is this text an email address. */
export const isAddress = (s) => EMAIL.test(String(s || '').trim());

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

/**
 * Say what is wrong with each of rows read at different times (the box read
 * more than once): an address that is not one, one already above, no password.
 * @param {Array<{ email: string, password: string }>} rows
 */
export function checkRows(rows) {
  const seen = new Set();
  return rows.map((row) => {
    const out = { ...row };
    delete out.problem;
    const key = String(row.email || '').toLowerCase();
    if (!isAddress(row.email)) out.problem = 'not-an-address';
    else if (seen.has(key)) out.problem = 'repeated';
    else if (!row.password) out.problem = 'no-password';
    if (isAddress(row.email)) seen.add(key);
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
      row = { line: i + 1, email: get('email').trim(), password: get('password') };
      if (get('name')) row.name = get('name');
      if (get('user')) row.user = get('user');
      if (get('imapHost')) row.imap = { host: get('imapHost'), port: Number(get('imapPort')) || 993 };
      if (get('smtpHost')) row.smtp = { host: get('smtpHost'), port: Number(get('smtpPort')) || 465 };
    } else {
      // An address, its separator, and the rest of the line as the password.
      const trimmed = line.trim();
      const quoted = /^"([^"]*)"\s*([\t,;:| ])\s*(.*)$/.exec(trimmed);
      const m = quoted || /^([^\s,;:|\t]+@[^\s,;:|\t]+?)(?:\s*([\t,;:|])\s*|\s+)(.*)$/.exec(trimmed);
      if (m) {
        let password = m[3];
        // A password a spreadsheet quoted ("p,w") is the text inside its quotes.
        const q = /^"((?:[^"]|"")*)"$/.exec(password);
        if (q) password = q[1].replace(/""/g, '"');
        row = { line: i + 1, email: m[1].trim(), password };
      } else row = { line: i + 1, email: trimmed, password: '' };
    }
    const key = row.email.toLowerCase();
    if (!isAddress(row.email)) row.problem = 'not-an-address';
    else if (seen.has(key)) row.problem = 'repeated';
    else if (!row.password) row.problem = 'no-password';
    if (isAddress(row.email)) seen.add(key);
    rows.push(row);
  });
  return { rows, header: header ? header.filter(Boolean) : null, separator: sep };
}
