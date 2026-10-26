// An IMAP server for the tests and the window checks, run in the same process.
//
// Small, and written for them: IMAP4rev1 as ImapFlow speaks it, with UIDPLUS,
// MOVE, CONDSTORE, SPECIAL-USE and IDLE, each of which can be left out to see
// how Mail does without it. Implicit TLS on 127.0.0.1 with a certificate made
// as it starts (Mail never sends a password in the clear, and a check's
// account says not to verify this one). Mailboxes live in memory.
//
// Besides answering Mail, it is "the server elsewhere": a phone marking a
// message read, another client deleting or filing one, mail arriving, the
// server renumbering a folder (a new UIDVALIDITY), going away for a while, or
// refusing one change. Every command it hears is kept in `heard`.

import crypto from 'node:crypto';
import tls from 'node:tls';
import { seq, set, oid, int, utf8, bits, explicit, time, algorithm } from '../src/smime/der.js';
import { OID } from '../src/smime/oids.js';
import { pem } from '../src/smime/cert.js';

/** A self-signed certificate for 127.0.0.1, made now: { key, cert } in PEM. */
export function certificate() {
  const { publicKey, privateKey } = crypto.generateKeyPairSync('ec', { namedCurve: 'P-256' });
  const name = seq(set(seq(oid(OID.commonName), utf8('127.0.0.1'))));
  const sigAlg = algorithm(OID.ecdsaWithSHA256, false);
  const day = 86_400_000;
  const tbs = seq(explicit(0, int(2)), int(crypto.randomBytes(8)), sigAlg, name, seq(time(new Date(Date.now() - day)), time(new Date(Date.now() + 30 * day))), name, Buffer.from(publicKey.export({ type: 'spki', format: 'der' })));
  const signature = crypto.sign('sha256', tbs, { key: privateKey, dsaEncoding: 'der' });
  return { key: privateKey.export({ type: 'pkcs8', format: 'pem' }), cert: pem(seq(tbs, sigAlg, bits(signature))) };
}

/** A header's value in a raw message, or ''. */
function headerOf(raw, name) {
  const head = String(raw).split(/\r?\n\r?\n/)[0];
  const m = new RegExp(`^${name}:[ \\t]*((?:.*)(?:\\r?\\n[ \\t].*)*)`, 'im').exec(head);
  return m ? m[1].replace(/\r?\n[ \t]+/g, ' ').trim() : '';
}

const quote = (s) => `"${String(s).replace(/[\\"]/g, (c) => `\\${c}`)}"`;
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const imapDate = (d) => `"${String(d.getUTCDate()).padStart(2, '0')}-${MONTHS[d.getUTCMonth()]}-${d.getUTCFullYear()} ${d.toISOString().slice(11, 19)} +0000"`;

/** The words of a command: atoms, quoted strings, (lists) and literals (Buffers). */
function tokenize(text, out = []) {
  const stack = [out];
  let i = 0;
  while (i < text.length) {
    const c = text[i];
    if (c === ' ') { i++; continue; }
    if (c === '(') { const list = []; stack.at(-1).push(list); stack.push(list); i++; continue; }
    if (c === ')') { stack.pop(); i++; continue; }
    if (c === '"') {
      let s = '';
      i++;
      while (i < text.length && text[i] !== '"') { if (text[i] === '\\') i++; s += text[i++]; }
      i++;
      stack.at(-1).push(s);
      continue;
    }
    let s = '';
    let depth = 0;
    while (i < text.length && (depth > 0 || !/[ ()]/.test(text[i]))) {
      if (text[i] === '[') depth++;
      if (text[i] === ']') depth--;
      s += text[i++];
    }
    stack.at(-1).push(s);
  }
  return out;
}

/**
 * Start one. `users` { address: password }; `capabilities` turns extensions
 * off ({ move: false, uidplus: false, condstore: false, idle: false,
 * specialUse: false }) or marks a Gmail-like server ({ gmail: true }, whose
 * folders are then named as Gmail names them). Resolves once it listens.
 */
export async function imapServer({ users = { 'sam@checks.example': 'check' }, capabilities: want = {}, folders: given = null } = {}) {
  const caps = { move: true, uidplus: true, condstore: true, idle: true, specialUse: true, gmail: false, ...want };
  const capability = ['IMAP4rev1', 'LITERAL+', 'ID', 'ENABLE', 'NAMESPACE', ...(caps.uidplus ? ['UIDPLUS'] : []), ...(caps.move ? ['MOVE'] : []), ...(caps.condstore ? ['CONDSTORE'] : []), ...(caps.specialUse ? ['SPECIAL-USE'] : []), ...(caps.idle ? ['IDLE'] : []), ...(caps.gmail ? ['X-GM-EXT-1'] : [])].join(' ');
  const { key, cert } = certificate();

  /** name -> { name, use, uidValidity, uidNext, modseq, messages: [{ uid, flags: Set, raw: Buffer, date, modseq }] } */
  const boxes = new Map();
  let validity = 1000 + Math.floor(Math.random() * 1000);
  const makeBox = (name, use = null) => {
    const box = { name, use, uidValidity: validity++, uidNext: 1, modseq: 1, messages: [], subscribed: true };
    boxes.set(name, box);
    return box;
  };
  const prefix = caps.gmail ? '[Gmail]/' : '';
  for (const [name, use] of given || [['INBOX', null], [`${prefix}${caps.gmail ? 'Sent Mail' : 'Sent'}`, '\\Sent'], [`${prefix}Drafts`, '\\Drafts'], [`${prefix}${caps.gmail ? 'Bin' : 'Trash'}`, '\\Trash'], [`${prefix}${caps.gmail ? 'Spam' : 'Junk'}`, '\\Junk'], ...(caps.gmail ? [] : [['Archive', '\\Archive']])]) makeBox(name, use);
  const box = (name) => boxes.get(/^inbox$/i.test(name) ? 'INBOX' : name) || null;

  const heard = [];
  const refusals = [];
  const sessions = new Set();
  let server = null;
  let port = 0;
  let peak = 0;

  /** Every other session looking at this mailbox hears what changed (at once while it idles). */
  const tell = (name, line, except = null) => {
    for (const s of sessions) if (s !== except && s.selected === name) { s.pending.push(line); if (s.idling) s.flush(); }
  };
  const bump = (b, m) => { b.modseq++; m.modseq = b.modseq; };
  const flagsText = (m) => `(${[...m.flags].join(' ')})`;

  function addMessage(name, raw, { flags = [], date = new Date() } = {}, except = null) {
    const b = box(name);
    if (!b) return null;
    const m = { uid: b.uidNext++, flags: new Set(flags), raw: Buffer.from(raw), date, modseq: 0 };
    bump(b, m);
    b.messages.push(m);
    tell(b.name, `* ${b.messages.length} EXISTS`, except);
    return m.uid;
  }
  function removeAt(b, index, except = null) {
    b.messages.splice(index, 1);
    b.modseq++;
    tell(b.name, `* ${index + 1} EXPUNGE`, except);
  }

  /** "1:4,7,9:*" over the values given (UIDs or sequence numbers), * the largest. */
  function inSet(text, values) {
    const max = values.length ? Math.max(...values) : 0;
    const wanted = new Set();
    for (const part of String(text).split(',')) {
      const [a, b = a] = part.split(':').map((v) => (v === '*' ? max : Number(v)));
      const lo = Math.min(a, b);
      const hi = Math.max(a, b);
      for (const v of values) if (v >= lo && v <= hi) wanted.add(v);
    }
    return wanted;
  }

  function session(socket) {
    const s = { socket, user: null, selected: null, readOnly: false, idling: null, pending: [], condstore: false };
    sessions.add(s);
    peak = Math.max(peak, sessions.size);
    const write = (data) => { if (!socket.destroyed) socket.write(data); };
    s.flush = () => { if (s.pending.length) write(`${s.pending.splice(0).join('\r\n')}\r\n`); };

    let buf = Buffer.alloc(0);
    let parts = [];
    let need = 0;
    let chain = Promise.resolve();
    socket.on('data', (chunk) => {
      buf = Buffer.concat([buf, chunk]);
      for (;;) {
        if (need) {
          if (buf.length < need) return;
          parts.push(buf.subarray(0, need));
          buf = buf.subarray(need);
          need = 0;
          continue;
        }
        const i = buf.indexOf('\r\n');
        if (i < 0) return;
        const line = buf.subarray(0, i).toString('utf8');
        buf = buf.subarray(i + 2);
        const lit = /\{(\d+)(\+?)\}$/.exec(line);
        if (lit) {
          parts.push(line.slice(0, lit.index));
          need = Number(lit[1]);
          if (!lit[2]) write('+ go ahead\r\n');
          continue;
        }
        parts.push(line);
        const command = parts;
        parts = [];
        chain = chain.then(() => run(command)).catch((err) => write(`* BAD ${err.message}\r\n`));
      }
    });
    socket.on('error', () => {});
    socket.on('close', () => sessions.delete(s));
    write(`* OK [CAPABILITY ${capability}] Rutba test IMAP ready\r\n`);

    async function run(command) {
      // Words, with each literal standing where it came.
      const words = [];
      for (const p of command) {
        if (Buffer.isBuffer(p)) words.push(p);
        else tokenize(p, words);
      }
      const text = command.map((p) => (Buffer.isBuffer(p) ? `{${p.length}}` : p)).join('');
      if (s.idling) {
        if (/^DONE$/i.test(text.trim())) { const tag = s.idling; s.idling = null; write(`${tag} OK IDLE terminated\r\n`); }
        return;
      }
      const [tag, ...rest] = words;
      let name = String(rest[0] || '').toUpperCase();
      let args = rest.slice(1);
      let byUid = false;
      if (name === 'UID') { byUid = true; name = String(args[0] || '').toUpperCase(); args = args.slice(1); }
      heard.push(text.replace(/^(\S+ LOGIN \S+) .*/i, '$1 ****'));
      const ok = (msg = 'done', code = '') => { s.flush(); write(`${tag} OK ${code ? `[${code}] ` : ''}${msg}\r\n`); };
      const no = (msg, code = '') => { s.flush(); write(`${tag} NO ${code ? `[${code}] ` : ''}${msg}\r\n`); };
      if (refusals.some((r) => r.test(text))) return no('Refused by the test server', 'CANNOT');
      if (name !== 'CAPABILITY' && name !== 'LOGIN' && name !== 'LOGOUT' && name !== 'ID' && name !== 'NOOP' && !s.user) return write(`${tag} BAD Sign in first\r\n`);
      const current = () => (s.selected ? box(s.selected) : null);
      const str = (v) => (Buffer.isBuffer(v) ? v.toString('utf8') : String(v ?? ''));

      switch (name) {
        case 'CAPABILITY': write(`* CAPABILITY ${capability}\r\n`); return ok();
        case 'NOOP': return ok();
        case 'ID': write('* ID ("name" "Rutba test IMAP")\r\n'); return ok();
        case 'LOGIN': {
          const [u, p] = args.map(str);
          if (!(u in users) || users[u] !== p) return no('Invalid credentials', 'AUTHENTICATIONFAILED');
          s.user = u;
          return ok('signed in', `CAPABILITY ${capability}`);
        }
        case 'LOGOUT': write('* BYE logging out\r\n'); ok(); socket.end(); return undefined;
        case 'ENABLE': {
          const on = args.map((a) => String(a).toUpperCase()).filter((a) => a === 'CONDSTORE' && caps.condstore);
          if (on.length) s.condstore = true;
          write(`* ENABLED${on.map((a) => ` ${a}`).join('')}\r\n`);
          return ok();
        }
        case 'NAMESPACE': write('* NAMESPACE (("" "/")) NIL NIL\r\n'); return ok();
        case 'LIST':
        case 'LSUB': {
          const pattern = str(args[1]);
          const re = new RegExp(`^${pattern.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/%/g, '[^/]*')}$`, 'i');
          for (const b of [...boxes.values()].sort((x, y) => (x.name === 'INBOX' ? -1 : y.name === 'INBOX' ? 1 : x.name.localeCompare(y.name)))) {
            if (!re.test(b.name) || (name === 'LSUB' && !b.subscribed)) continue;
            const children = [...boxes.keys()].some((n) => n.startsWith(`${b.name}/`));
            const attrs = [children ? '\\HasChildren' : '\\HasNoChildren', ...(caps.specialUse && b.use && name === 'LIST' ? [b.use] : [])];
            write(`* ${name} (${attrs.join(' ')}) "/" ${quote(b.name)}\r\n`);
          }
          return ok();
        }
        case 'STATUS': {
          const b = box(str(args[0]));
          if (!b) return no('No such mailbox', 'NONEXISTENT');
          const items = (Array.isArray(args[1]) ? args[1] : [args[1]]).map((i) => String(i).toUpperCase());
          const value = { MESSAGES: b.messages.length, UIDNEXT: b.uidNext, UIDVALIDITY: b.uidValidity, UNSEEN: b.messages.filter((m) => !m.flags.has('\\Seen')).length, RECENT: 0, HIGHESTMODSEQ: b.modseq };
          const got = items.filter((i) => i in value && (i !== 'HIGHESTMODSEQ' || caps.condstore));
          write(`* STATUS ${quote(b.name)} (${got.map((i) => `${i} ${value[i]}`).join(' ')})\r\n`);
          return ok();
        }
        case 'SELECT':
        case 'EXAMINE': {
          const b = box(str(args[0]));
          s.selected = null;
          s.pending = [];
          if (!b) return no('No such mailbox', 'NONEXISTENT');
          s.selected = b.name;
          s.readOnly = name === 'EXAMINE';
          if (Array.isArray(args[1]) && args[1].some((a) => /^condstore$/i.test(a)) && caps.condstore) s.condstore = true;
          write(`* ${b.messages.length} EXISTS\r\n* 0 RECENT\r\n* FLAGS (\\Answered \\Flagged \\Deleted \\Seen \\Draft $Junk $NotJunk)\r\n* OK [PERMANENTFLAGS (\\Answered \\Flagged \\Deleted \\Seen \\Draft $Junk $NotJunk \\*)] Flags\r\n* OK [UIDVALIDITY ${b.uidValidity}] UIDs\r\n* OK [UIDNEXT ${b.uidNext}] Next\r\n`);
          if (caps.condstore) write(`* OK [HIGHESTMODSEQ ${b.modseq}] Modseq\r\n`);
          return ok('selected', s.readOnly ? 'READ-ONLY' : 'READ-WRITE');
        }
        case 'CLOSE':
        case 'UNSELECT': {
          const b = current();
          if (b && name === 'CLOSE' && !s.readOnly) b.messages = b.messages.filter((m) => !m.flags.has('\\Deleted'));
          s.selected = null;
          return ok();
        }
        case 'IDLE':
          if (!caps.idle) return write(`${tag} BAD No IDLE here\r\n`);
          s.idling = tag;
          write('+ idling\r\n');
          s.flush();
          return undefined;
        case 'CREATE': {
          const n = str(args[0]).replace(/\/$/, '');
          if (box(n)) return no('Mailbox exists', 'ALREADYEXISTS');
          makeBox(n);
          return ok();
        }
        case 'DELETE': {
          const b = box(str(args[0]));
          if (!b) return no('No such mailbox', 'NONEXISTENT');
          if (b.name === 'INBOX') return no('The Inbox stays', 'CANNOT');
          boxes.delete(b.name);
          return ok();
        }
        case 'RENAME': {
          const from = box(str(args[0]));
          const to = str(args[1]);
          if (!from) return no('No such mailbox', 'NONEXISTENT');
          if (box(to)) return no('Mailbox exists', 'ALREADYEXISTS');
          const old = from.name;
          for (const [n, b] of [...boxes]) {
            if (n !== old && !n.startsWith(`${old}/`)) continue;
            boxes.delete(n);
            b.name = to + n.slice(old.length);
            boxes.set(b.name, b);
          }
          return ok();
        }
        case 'SUBSCRIBE':
        case 'UNSUBSCRIBE': {
          const b = box(str(args[0]));
          if (b) b.subscribed = name === 'SUBSCRIBE';
          return ok();
        }
        case 'APPEND': {
          const b = box(str(args[0]));
          const literal = args.find((a) => Buffer.isBuffer(a));
          if (!b) return no('No such mailbox', 'TRYCREATE');
          const flags = Array.isArray(args[1]) ? args[1] : [];
          const uid = addMessage(b.name, literal || Buffer.alloc(0), { flags }, s);
          return ok('appended', caps.uidplus ? `APPENDUID ${b.uidValidity} ${uid}` : '');
        }
      }

      // The rest work on the mailbox selected.
      const b = current();
      if (!b) return write(`${tag} BAD Select a mailbox first\r\n`);
      const values = () => (byUid ? b.messages.map((m) => m.uid) : b.messages.map((_, i) => i + 1));
      const chosen = (setText) => {
        const keep = inSet(setText, values());
        return b.messages.map((m, i) => ({ m, seq: i + 1 })).filter(({ m, seq: n }) => keep.has(byUid ? m.uid : n));
      };

      switch (name) {
        case 'FETCH': {
          const list = chosen(str(args[0]));
          let items = Array.isArray(args[1]) ? args[1] : [args[1]];
          items = items.map((i) => String(i));
          const changedSince = Array.isArray(args[2]) && /^changedsince$/i.test(args[2][0]) ? Number(args[2][1]) : null;
          if (changedSince != null) s.condstore = true;
          const parts = [];
          for (const { m, seq: n } of list) {
            if (changedSince != null && m.modseq <= changedSince) continue;
            const out = [];
            const want = new Set(items.map((i) => i.toUpperCase()));
            if (byUid) want.add('UID');
            if (s.condstore && caps.condstore) want.add('MODSEQ');
            let seen = false;
            const bodies = [];
            for (const item of items) {
              const u = item.toUpperCase();
              const section = /^BODY(\.PEEK)?\[(.*)\]$/i.exec(item);
              if (section) {
                if (!section[1]) seen = true;
                const what = section[2].toUpperCase();
                const raw = m.raw;
                const split = raw.indexOf('\r\n\r\n');
                const head = split >= 0 ? raw.subarray(0, split + 4) : raw;
                let bytes = raw;
                if (what === 'HEADER') bytes = head;
                else if (what === 'TEXT') bytes = split >= 0 ? raw.subarray(split + 4) : Buffer.alloc(0);
                else if (what.startsWith('HEADER.FIELDS')) {
                  const fields = (/\((.*)\)/.exec(section[2])?.[1] || '').split(/\s+/).map((f) => f.toLowerCase());
                  const lines = head.toString('utf8').split(/\r\n(?![ \t])/).filter((l) => fields.includes(l.split(':')[0].trim().toLowerCase()));
                  bytes = Buffer.from(`${lines.join('\r\n')}${lines.length ? '\r\n' : ''}\r\n`, 'utf8');
                }
                bodies.push([`BODY[${section[2]}]`, bytes]);
                continue;
              }
              if (u === 'RFC822.SIZE') out.push(`RFC822.SIZE ${m.raw.length}`);
              else if (u === 'INTERNALDATE') out.push(`INTERNALDATE ${imapDate(m.date)}`);
            }
            if (seen && !s.readOnly && !m.flags.has('\\Seen')) { m.flags.add('\\Seen'); bump(b, m); want.add('FLAGS'); }
            const head = [];
            if (want.has('UID')) head.push(`UID ${m.uid}`);
            if (want.has('FLAGS')) head.push(`FLAGS ${flagsText(m)}`);
            if (want.has('MODSEQ')) head.push(`MODSEQ (${m.modseq})`);
            const chunks = [Buffer.from(`* ${n} FETCH (${[...head, ...out].join(' ')}`)];
            for (const [label, bytes] of bodies) chunks.push(Buffer.from(` ${label} {${bytes.length}}\r\n`), bytes);
            chunks.push(Buffer.from(')\r\n'));
            parts.push(Buffer.concat(chunks));
          }
          s.flush();
          if (parts.length) write(Buffer.concat(parts));
          return ok();
        }
        case 'STORE': {
          if (s.readOnly) return no('Read only');
          const list = chosen(str(args[0]));
          let at = 1;
          let unchangedSince = null;
          if (Array.isArray(args[1])) { unchangedSince = Number(args[1][1]); at = 2; }
          const action = String(args[at]).toUpperCase();
          const flags = Array.isArray(args[at + 1]) ? args[at + 1] : [args[at + 1]];
          const silent = action.endsWith('.SILENT');
          const modified = [];
          for (const { m, seq: n } of list) {
            if (unchangedSince != null && m.modseq > unchangedSince) { modified.push(byUid ? m.uid : n); continue; }
            const before = flagsText(m);
            if (action.startsWith('+')) for (const f of flags) m.flags.add(f);
            else if (action.startsWith('-')) for (const f of flags) m.flags.delete(f);
            else m.flags = new Set(flags);
            if (flagsText(m) !== before) {
              bump(b, m);
              tell(b.name, `* ${n} FETCH (FLAGS ${flagsText(m)})`, s);
            }
            if (!silent || s.condstore) write(`* ${n} FETCH (${byUid ? `UID ${m.uid} ` : ''}FLAGS ${flagsText(m)}${s.condstore && caps.condstore ? ` MODSEQ (${m.modseq})` : ''})\r\n`);
          }
          return ok('stored', modified.length ? `MODIFIED ${modified.join(',')}` : '');
        }
        case 'COPY':
        case 'MOVE': {
          if (name === 'MOVE' && !caps.move) return write(`${tag} BAD No MOVE here\r\n`);
          const list = chosen(str(args[0]));
          const target = box(str(args[1]));
          if (!target) return no('No such mailbox', 'TRYCREATE');
          const from = [];
          const to = [];
          for (const { m } of list) {
            from.push(m.uid);
            const flags = [...m.flags].filter((f) => f !== '\\Deleted' || name === 'COPY');
            to.push(addMessage(target.name, m.raw, { flags, date: m.date }, null));
          }
          const code = caps.uidplus && from.length ? `COPYUID ${target.uidValidity} ${from.join(',')} ${to.join(',')}` : '';
          if (name === 'COPY') return ok('copied', code);
          if (code) write(`* OK [${code}] Moved\r\n`);
          for (const uid of from) {
            const index = b.messages.findIndex((m) => m.uid === uid);
            if (index < 0) continue;
            write(`* ${index + 1} EXPUNGE\r\n`);
            removeAt(b, index, s);
          }
          return ok('moved');
        }
        case 'EXPUNGE': {
          if (s.readOnly) return no('Read only');
          const keep = byUid ? inSet(str(args[0]), b.messages.map((m) => m.uid)) : null;
          for (let i = b.messages.length - 1; i >= 0; i--) {
            const m = b.messages[i];
            if (!m.flags.has('\\Deleted') || (keep && !keep.has(m.uid))) continue;
            write(`* ${i + 1} EXPUNGE\r\n`);
            removeAt(b, i, s);
          }
          return ok('expunged');
        }
        case 'SEARCH': {
          const found = search(b, args, byUid);
          write(`* SEARCH${found.map((v) => ` ${v}`).join('')}\r\n`);
          return ok();
        }
        default:
          return write(`${tag} BAD Unknown command ${name}\r\n`);
      }
    }
  }

  /** SEARCH with the keys Mail uses: ALL, UID set, HEADER, SEEN, UNSEEN, FLAGGED, DELETED, UNDELETED, a sequence set. */
  function search(b, args, byUid) {
    const keys = args.map((a) => (Buffer.isBuffer(a) ? a.toString('utf8') : a));
    let rows = b.messages.map((m, i) => ({ m, n: i + 1 }));
    for (let i = 0; i < keys.length; i++) {
      const k = String(keys[i]).toUpperCase();
      if (k === 'ALL' || k === 'CHARSET') { if (k === 'CHARSET') i++; continue; }
      if (k === 'UID') { const keep = inSet(keys[++i], b.messages.map((m) => m.uid)); rows = rows.filter(({ m }) => keep.has(m.uid)); continue; }
      if (k === 'HEADER') {
        const field = String(keys[++i]);
        const value = String(keys[++i]).toLowerCase();
        rows = rows.filter(({ m }) => headerOf(m.raw.toString('utf8'), field).toLowerCase().includes(value));
        continue;
      }
      if (k === 'SEEN') rows = rows.filter(({ m }) => m.flags.has('\\Seen'));
      else if (k === 'UNSEEN') rows = rows.filter(({ m }) => !m.flags.has('\\Seen'));
      else if (k === 'FLAGGED') rows = rows.filter(({ m }) => m.flags.has('\\Flagged'));
      else if (k === 'DELETED') rows = rows.filter(({ m }) => m.flags.has('\\Deleted'));
      else if (k === 'UNDELETED') rows = rows.filter(({ m }) => !m.flags.has('\\Deleted'));
      else if (/^[\d*:,]+$/.test(k)) { const keep = inSet(k, b.messages.map((_, j) => j + 1)); rows = rows.filter(({ n }) => keep.has(n)); }
    }
    return rows.map(({ m, n }) => (byUid ? m.uid : n));
  }

  const listen = (at = 0) => new Promise((resolve, reject) => {
    server = tls.createServer({ key, cert }, session);
    server.once('error', reject);
    server.listen(at, '127.0.0.1', () => { port = server.address().port; resolve(port); });
  });
  await listen();

  const find = (name, uid) => box(name)?.messages.find((m) => m.uid === uid) || null;
  return {
    host: '127.0.0.1',
    get port() { return port; },
    heard,
    /** The account a test or a check adds for this server. */
    account: (email = Object.keys(users)[0]) => ({ email, name: email, imap: { host: '127.0.0.1', port, secure: true, rejectUnauthorized: false, user: email }, smtp: { host: '127.0.0.1', port: 1, secure: false } }),
    /** Mail arriving: its UID in that folder. */
    deliver: (name, raw, options = {}) => addMessage(name, String(raw).replace(/\r?\n/g, '\r\n'), options),
    /** What a folder holds: [{ uid, flags, messageId, subject }] in order. */
    messages: (name) => (box(name)?.messages || []).map((m) => ({ uid: m.uid, flags: [...m.flags], messageId: headerOf(m.raw.toString('utf8'), 'message-id'), subject: headerOf(m.raw.toString('utf8'), 'subject'), raw: m.raw.toString('utf8') })),
    folders: () => [...boxes.keys()],
    uidValidity: (name) => box(name)?.uidValidity ?? null,
    /** Another client changed a message's flags. */
    setFlags(name, uid, flags) {
      const b = box(name);
      const m = find(name, uid);
      if (!m) return false;
      m.flags = new Set(flags);
      bump(b, m);
      tell(b.name, `* ${b.messages.indexOf(m) + 1} FETCH (FLAGS ${flagsText(m)})`);
      return true;
    },
    /** Another client deleted one, for good. */
    expunge(name, uid) {
      const b = box(name);
      const i = b ? b.messages.findIndex((m) => m.uid === uid) : -1;
      if (i < 0) return false;
      removeAt(b, i);
      return true;
    },
    /** Another client filed one: its UID in the folder it went to. */
    move(name, uid, to) {
      const m = find(name, uid);
      if (!m || !box(to)) return null;
      const now = addMessage(to, m.raw, { flags: [...m.flags], date: m.date });
      this.expunge(name, uid);
      return now;
    },
    /** The server renumbers a folder: a new UIDVALIDITY and every UID new. */
    renumber(name) {
      const b = box(name);
      b.uidValidity = validity++;
      b.uidNext = 1 + Math.floor(Math.random() * 50) + 100;
      for (const m of b.messages) m.uid = b.uidNext++;
      b.modseq++;
      return b.uidValidity;
    },
    /** Another client made a folder. */
    makeFolder: (name) => (box(name) ? false : Boolean(makeBox(name))),
    /** Another client renamed a folder, with the folders inside it. */
    renameFolder(from, to) {
      const b = box(from);
      if (!b || box(to)) return false;
      const old = b.name;
      for (const [n, inner] of [...boxes]) {
        if (n !== old && !n.startsWith(`${old}/`)) continue;
        boxes.delete(n);
        inner.name = to + n.slice(old.length);
        boxes.set(inner.name, inner);
      }
      return true;
    },
    /** Another client deleted a folder. */
    deleteFolder: (name) => boxes.delete(box(name)?.name ?? ''),
    /** Commands matching this are answered NO until `allow()`. */
    refuse: (pattern) => { refusals.push(pattern); },
    allow: () => { refusals.length = 0; },
    /** Gone away: every connection dropped and nothing listening, until `start()`. */
    async stop() {
      for (const s of sessions) s.socket.destroy();
      sessions.clear();
      await new Promise((resolve) => server.close(() => resolve()));
    },
    /** Back, on the same port. */
    start: () => listen(port),
    connections: () => sessions.size,
    /** The most connections open at once since the last `resetPeak()`. */
    peak: () => peak,
    resetPeak: () => { peak = sessions.size; },
    idling: () => [...sessions].filter((s) => s.idling).length,
    async close() {
      for (const s of sessions) s.socket.destroy();
      sessions.clear();
      await new Promise((resolve) => (server.listening ? server.close(() => resolve()) : resolve()));
    },
  };
}
