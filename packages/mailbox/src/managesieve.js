// The mail server's own automatic reply (owner, 2026-10-22, D-2026-10-22-01:
// "the out office the mail should have a way to activate and deactivate that
// servers function with custom message"), set over ManageSieve (RFC 5804)
// as a Sieve script with the vacation action (RFC 5230), so replies go out
// while this computer is off.
//
// A server keeps one active script. Someone's own filters may be that
// script, so Office's never replaces it: Office's script answers first and
// then runs theirs (`include :personal`), and turning the reply off makes
// theirs active again. Which one that was is written in Office's script, so
// another computer can turn it off too.

import net from 'node:net';
import tls from 'node:tls';

/** The name Office's script is kept under on the server. */
export const SCRIPT_NAME = 'rutba-out-of-office';
/** How often one sender is answered, in days (Gmail's is four). */
export const REPLY_DAYS = 4;
const MARK = '# rutba-office: ';

/* ── the script ─────────────────────────────────────────────────── */

/** A Sieve quoted string: backslash and double quote escaped, line breaks as CRLF. */
export const sieveString = (text) => `"${String(text ?? '').replace(/\r?\n/g, '\r\n').replace(/[\\"]/g, (c) => `\\${c}`)}"`;

/** An instant as Sieve's currentdate compares it in UTC: "yyyy-mm-ddThh:mm:ss". */
const utc = (when) => new Date(when).toISOString().slice(0, 19);

/**
 * The script for an automatic reply: `settings` { subject, message, start,
 * end }, `addresses` the account's own, `previous` the script that was
 * active before, run after the reply.
 */
export function vacationScript({ subject = '', message = '', start = null, end = null }, { addresses = [], previous = null, days = REPLY_DAYS } = {}) {
  const tests = [];
  if (start) tests.push(`currentdate :zone "+0000" :value "ge" "iso8601" ${sieveString(utc(start))}`);
  if (end) tests.push(`currentdate :zone "+0000" :value "le" "iso8601" ${sieveString(utc(end))}`);
  const need = ['vacation', ...(tests.length ? ['date', 'relational'] : []), ...(previous ? ['include'] : [])];
  const args = [`:days ${days}`];
  if (subject.trim()) args.push(`:subject ${sieveString(subject.trim())}`);
  if (addresses.length) args.push(`:addresses [${addresses.map(sieveString).join(', ')}]`);
  const action = `vacation ${args.join(' ')} ${sieveString(message || '')};`;
  const lines = [
    '# Rutba Office: out of office, turned on and off from Mail.',
    `${MARK}${JSON.stringify({ subject, message, start, end, previous })}`,
    `require [${need.map(sieveString).join(', ')}];`,
    tests.length === 0 ? action : tests.length === 1 ? `if ${tests[0]} {\r\n  ${action}\r\n}` : `if allof(${tests.join(', ')}) {\r\n  ${action}\r\n}`,
  ];
  if (previous) lines.push(`include :personal ${sieveString(previous)};`);
  return `${lines.join('\r\n')}\r\n`;
}

/** What Office's script says it holds, or null for one it did not write. */
export function readScript(text) {
  const line = String(text || '').split(/\r?\n/).find((l) => l.startsWith(MARK));
  if (!line) return null;
  try { return JSON.parse(line.slice(MARK.length)); } catch { return null; }
}

/* ── the protocol ───────────────────────────────────────────────── */

/**
 * One ManageSieve connection. `host`, `port` (4190), signed in with `user`
 * and `pass` (PLAIN, after STARTTLS: a password never crosses in the clear,
 * unless `plain` is set for a check's own stand-in on this computer).
 */
export async function openSieve({ host, port = 4190, user, pass, plain = false, timeoutMs = 15000, servername = host }) {
  let socket = net.connect({ host, port });
  let buf = Buffer.alloc(0);
  let waiting = null;
  let failed = null;
  const feed = (chunk) => { buf = Buffer.concat([buf, chunk]); waiting?.(); };
  const attach = (s) => {
    s.on('data', feed);
    s.on('error', (err) => { failed = err; waiting?.(); });
    s.on('close', () => { failed ||= new Error('The mail server closed the connection.'); waiting?.(); });
    s.setTimeout?.(timeoutMs, () => { failed = new Error('The mail server did not answer.'); s.destroy(); });
  };
  attach(socket);
  // A port that is filtered answers nothing at all: given up on, not waited for.
  await new Promise((resolve, reject) => {
    const timer = setTimeout(() => { socket.destroy(); reject(Object.assign(new Error('The mail server did not answer.'), { code: 'unreachable' })); }, timeoutMs);
    socket.once('connect', () => { clearTimeout(timer); resolve(); });
    socket.once('error', (err) => { clearTimeout(timer); reject(Object.assign(err, { code: 'unreachable' })); });
  });

  /** Bytes as they come: a line, or a literal of `n` bytes. */
  const more = () => new Promise((resolve) => { waiting = () => { waiting = null; resolve(); }; });
  const line = async () => {
    for (;;) {
      const i = buf.indexOf('\r\n');
      if (i >= 0) { const l = buf.subarray(0, i).toString('utf8'); buf = buf.subarray(i + 2); return l; }
      if (failed) throw failed;
      await more();
    }
  };
  const bytes = async (n) => {
    while (buf.length < n) { if (failed) throw failed; await more(); }
    const b = buf.subarray(0, n); buf = buf.subarray(n); return b.toString('utf8');
  };
  /** The lines of an answer up to OK, NO or BYE: { ok, code, text, lines } (a literal read whole into its line). */
  const answer = async () => {
    const lines = [];
    for (;;) {
      let l = await line();
      const lit = /\{(\d+)\+?\}$/.exec(l);
      if (lit) { const body = await bytes(Number(lit[1])); l = l.slice(0, lit.index) + JSON.stringify(body); await line().then((rest) => { l += rest; }); }
      const end = /^(OK|NO|BYE)(?:\s+\(([^)]*)\))?(?:\s+(.*))?$/i.exec(l);
      if (end) return { ok: end[1].toUpperCase() === 'OK', code: end[2] || '', text: unquote(end[3] || ''), lines };
      lines.push(l);
    }
  };
  const send = (text) => socket.write(text);
  const command = async (text) => { send(`${text}\r\n`); const a = await answer(); if (!a.ok) throw Object.assign(new Error(a.text || 'The mail server refused that.'), { code: a.code }); return a; };

  let caps = capabilities((await answer()).lines);
  if (!plain) {
    if (!caps.STARTTLS) { socket.destroy(); throw Object.assign(new Error('This mail server does not offer a secure connection for its filters.'), { code: 'no-tls' }); }
    await command('STARTTLS');
    socket.removeAllListeners('data');
    socket = tls.connect({ socket, servername });
    buf = Buffer.alloc(0);
    attach(socket);
    await new Promise((resolve, reject) => { socket.once('secureConnect', resolve); socket.once('error', reject); });
    // After STARTTLS the server says what it offers again.
    caps = capabilities((await answer()).lines);
  }
  if (!String(caps.SIEVE || '').split(/\s+/).includes('vacation')) { socket.destroy(); throw Object.assign(new Error('This mail server cannot send automatic replies itself.'), { code: 'no-vacation' }); }
  const auth = Buffer.from(`\0${user}\0${pass}`, 'utf8').toString('base64');
  await command(`AUTHENTICATE "PLAIN" "${auth}"`).catch((err) => { socket.destroy(); throw Object.assign(new Error('The mail server did not accept the sign-in for its filters.'), { code: 'auth', cause: err }); });

  return {
    capabilities: caps,
    /** [{ name, active }] */
    async list() {
      const a = await command('LISTSCRIPTS');
      return a.lines.map((l) => { const m = /^("(?:[^"\\]|\\.)*")(\s+ACTIVE)?$/i.exec(l.trim()); return m ? { name: unquote(m[1]), active: Boolean(m[2]) } : null; }).filter(Boolean);
    },
    async get(name) {
      const a = await command(`GETSCRIPT ${quote(name)}`);
      return a.lines.length ? unquote(a.lines.join('\r\n')) : '';
    },
    async put(name, text) {
      const body = Buffer.from(text, 'utf8');
      send(`PUTSCRIPT ${quote(name)} {${body.length}+}\r\n`);
      send(body);
      send('\r\n');
      const a = await answer();
      if (!a.ok) throw Object.assign(new Error(a.text || 'The mail server refused the script.'), { code: a.code || 'script' });
    },
    setActive: (name) => command(`SETACTIVE ${quote(name || '')}`),
    remove: (name) => command(`DELETESCRIPT ${quote(name)}`),
    async close() { try { await command('LOGOUT'); } catch { /* gone */ } socket.destroy(); },
  };
}

const quote = (s) => `"${String(s).replace(/[\\"]/g, (c) => `\\${c}`)}"`;
function unquote(s) {
  const t = String(s).trim();
  if (t.startsWith('"') && t.endsWith('"')) {
    try { return JSON.parse(t); } catch { return t.slice(1, -1).replace(/\\(.)/g, '$1'); }
  }
  return t;
}
function capabilities(lines) {
  const caps = {};
  for (const l of lines) {
    const m = /^"([^"]+)"(?:\s+"((?:[^"\\]|\\.)*)")?$/.exec(l.trim());
    if (m) caps[m[1].toUpperCase()] = m[2] ?? true;
  }
  return caps;
}

/* ── turning it on and off ─────────────────────────────────────── */

/**
 * The server's automatic reply on, with these settings: Office's script
 * written and made active, whatever was active before run after it.
 */
export async function turnOn(session, settings, { addresses = [] } = {}) {
  const scripts = await session.list();
  const active = scripts.find((s) => s.active)?.name || null;
  let previous = active && active !== SCRIPT_NAME ? active : null;
  // On already: keep what it said was before.
  if (active === SCRIPT_NAME) previous = readScript(await session.get(SCRIPT_NAME))?.previous || null;
  if (previous && !scripts.some((s) => s.name === previous)) previous = null;
  await session.put(SCRIPT_NAME, vacationScript(settings, { addresses, previous }));
  await session.setActive(SCRIPT_NAME);
  return { on: true, previous };
}

/** The server's automatic reply off: the script before it active again, or none. */
export async function turnOff(session) {
  const scripts = await session.list();
  const ours = scripts.find((s) => s.name === SCRIPT_NAME);
  if (!ours) return { on: false, previous: null };
  const said = readScript(await session.get(SCRIPT_NAME)) || {};
  const previous = said.previous && scripts.some((s) => s.name === said.previous) ? said.previous : null;
  if (ours.active) await session.setActive(previous || '');
  return { on: false, previous };
}

/** What the server says now: { on, settings } (settings as Office wrote them, or null). */
export async function readState(session) {
  const scripts = await session.list();
  const ours = scripts.find((s) => s.name === SCRIPT_NAME);
  if (!ours) return { on: false, settings: null };
  return { on: ours.active, settings: readScript(await session.get(SCRIPT_NAME)) };
}
