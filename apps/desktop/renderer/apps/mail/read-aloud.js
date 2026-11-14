// Read Aloud in Mail (D-2026-11-13-34), on the reader the apps share
// (read-aloud/reader.js, C7). Three ways to play: the message open; every
// message in the list, in the list's order; and the unread ones in the list.
// The list is the list as drawn when the play starts, so a filter (Unread,
// Starred, Attachments) or a search decides what Play all reads.
//
// Each message is read as who sent it, its subject, then its words: the
// plain text, or the page's words when there is none, with the quoted
// message it answers, the signature and the long web addresses left out,
// and what a page hides (its title, a hidden preheader, Outlook's own copy
// for itself) not read. One message is read at a time; the next is asked
// for only when its turn comes, so a list of five hundred costs nothing
// until it is played. A message is marked read only once it has been read
// to its end (the PM, on review 1); one skipped or stopped stays as it was.

import { t } from '@rutba/office-ui/messages';
import { stripTags, displayName } from './parts.js';

/* ── the words of a page ─────────────────────────────────────────────────── */

// Every pass over a page here is one walk forward through it, whatever it
// holds: a page is mail from anyone, and Play all reaches one by itself. A
// lazy pattern that starts a fresh scan to the end at every "<div" or "<!--"
// took over a minute on 200 KB of them (review 2), with the window frozen.

const ATTRS = {};
/**
 * An attribute's value, or null. One pattern whose parts never run over the
 * same characters (the spaces either side of "=" are parted by it, and a
 * value is quoted or has no spaces), so a tag of a megabyte of spaces costs
 * one pass; the value is then searched plainly (review 3).
 */
const attrOf = (attrs, name) => {
  const re = (ATTRS[name] ||= new RegExp(`(?:^|[\\s"'/])${name}[ \\t\\r\\n\\f]*=[ \\t\\r\\n\\f]*(?:"([^"]*)"|'([^']*)'|([^\\s"'=<>\`]+))`, 'i'));
  const m = re.exec(attrs);
  return m ? (m[1] ?? m[2] ?? m[3]) : null;
};
const attrIs = (attrs, name, ...values) => values.includes(String(attrOf(attrs, name) ?? '').trim().toLowerCase());
const QUOTE_CLASSES = ['gmail_quote', 'moz-cite-prefix', 'yahoo_quoted'];

/** Where a reply's quoted message starts in a page, in the mail programs in wide use: the tag, and what its attributes say. */
const QUOTE_START = {
  // Gmail, Thunderbird, Yahoo; Outlook, Outlook on the web.
  div: (attrs) => String(attrOf(attrs, 'class') ?? '').toLowerCase().split(/[ \t\r\n\f]+/).some((c) => QUOTE_CLASSES.includes(c))
    || attrIs(attrs, 'id', 'divrplyfwdmsg', 'appendonsend'),
  hr: (attrs) => attrIs(attrs, 'id', 'stopspelling'), // Outlook.com, older
  blockquote: (attrs) => attrIs(attrs, 'type', 'cite'), // Apple Mail, Thunderbird
};

/** A forward's own line, in the languages mail programs write it in. */
const FORWARD_LINE = /^(?:-{2,}\s*(?:Forwarded message|Weitergeleitete Nachricht|Message transf[ée]r[ée]|Mensaje reenviado|Messaggio inoltrato|Doorgestuurd bericht|Mensagem encaminhada)\s*-{2,}|Begin forwarded message:?|D[ée]but du message r[ée]exp[ée]di[ée]\s*:?)$/i;
const FORWARD_IN_PAGE = /Forwarded message|Begin forwarded message|Weitergeleitete Nachricht|Message transf[ée]r[ée]|Mensaje reenviado|Messaggio inoltrato|Doorgestuurd bericht/i;

/** A style that keeps any element out of sight: display:none, visibility:hidden, Outlook's mso-hide:all. */
const HIDING = /(?:^|[;\s])(?:display\s*:\s*none|visibility\s*:\s*hidden|mso-hide\s*:\s*all)/i;
const zeroOf = (prop) => new RegExp(`(?:^|[;\\s])${prop}\\s*:\\s*0(?:\\.0*)?(?:px|pt|em|rem|%)?\\s*(?:!important\\s*)?(?:;|$)`, 'i');
const MAX_HEIGHT_0 = zeroOf('max-height');
const OVERFLOW_HIDDEN = /(?:^|[;\s])overflow\s*:\s*hidden/i;
const FONT_SIZE_0 = zeroOf('font-size');
const OPACITY_0 = zeroOf('opacity');
/** Elements that hold words, not layout: a zero size or opacity on one hides its words. */
const INLINE = new Set(['span', 'a', 'font', 'b', 'i', 'u', 'em', 'strong', 'small', 'sup', 'sub', 'br', 'img', 'wbr']);
/**
 * Does this element keep its words out of sight? display:none and the like
 * anywhere; no height with its overflow cut; and a zero font size or opacity
 * on an inline element only, since MJML and hybrid layouts put font-size:0
 * on the cells that hold the whole message (review 3).
 */
const hides = (name, attrs) => {
  const style = attrOf(attrs, 'style') || '';
  if (!style) return false;
  if (HIDING.test(style)) return true;
  if (MAX_HEIGHT_0.test(style) && OVERFLOW_HIDDEN.test(style)) return true;
  return INLINE.has(name) && (FONT_SIZE_0.test(style) || OPACITY_0.test(style));
};

const TABLE_PARTS = ['tbody', 'thead', 'tfoot'];
/**
 * Does this tag end, by HTML's own rule, an element being left out that was
 * written without its end tag: a p, li, td, th, tr, dd, dt or option? Asked
 * only at the element's own level, never inside an element it holds.
 */
const IMPLIED_END = {
  p: (close, name) => (name === 'p' ? !close : !INLINE.has(name)),
  li: (close, name) => (close ? ['ul', 'ol', 'menu'].includes(name) : name === 'li'),
  td: (close, name) => (close ? ['tr', 'table', ...TABLE_PARTS].includes(name) : ['td', 'th', 'tr', ...TABLE_PARTS].includes(name)),
  tr: (close, name) => (close ? ['table', ...TABLE_PARTS].includes(name) : ['tr', ...TABLE_PARTS].includes(name)),
  dd: (close, name) => (close ? name === 'dl' : name === 'dd' || name === 'dt'),
  option: (close, name) => (close ? name === 'select' || name === 'optgroup' : name === 'option' || name === 'optgroup'),
};
IMPLIED_END.th = IMPLIED_END.td;
IMPLIED_END.dt = IMPLIED_END.dd;

/** A tag: "<", its name, anything that is neither "<" nor ">", then ">". A "<" with no ">" costs only the way to the next "<". */
const TAG = /<(\/?)([a-zA-Z][a-zA-Z0-9]*)\b([^<>]*)>/g;
const VOID = new Set(['area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input', 'link', 'meta', 'source', 'track', 'wbr']);
const selfClosed = (attrs) => attrs.endsWith('/') || /\/\s*$/.test(attrs.slice(-64));

/**
 * The comments (Outlook's <!--[if mso]> copies with them), the head and any
 * title taken off, in one walk forward. A comment with no end hides the rest,
 * as it does in a browser; a head with no end ends at the body.
 */
function withoutUnseen(html) {
  const s = String(html || '');
  const open = /<!--|<(head|title)\b/gi;
  const failedFrom = {}; // a search for an end that failed from here fails from anywhere after
  const endOf = (kind, re, from) => {
    if (failedFrom[kind] != null && from >= failedFrom[kind]) return -1;
    re.lastIndex = from;
    const m = re.exec(s);
    if (!m) {
      failedFrom[kind] = from;
      return -1;
    }
    return kind === 'body' ? m.index : m.index + m[0].length;
  };
  // words-ok: regular expressions for the page's tags, not words
  const ends = { head: /<\/head\s*>/gi, title: /<\/title\s*>/gi, body: /<body\b/gi, gt: />/g };
  let out = '';
  let at = 0;
  for (let m = open.exec(s); m; m = open.exec(s)) {
    out += s.slice(at, m.index);
    const kind = m[1] ? m[1].toLowerCase() : 'comment';
    let end;
    if (kind === 'comment') {
      end = s.indexOf('-->', m.index + 4);
      if (end < 0) return out;
      end += 3;
    } else {
      end = endOf(kind, ends[kind], m.index);
      if (end < 0 && kind === 'head') end = endOf('body', ends.body, m.index);
      if (end < 0) end = endOf('gt', ends.gt, m.index); // its own tag alone
      if (end < 0) return out;
    }
    at = end;
    open.lastIndex = end;
  }
  return out + s.slice(at);
}

/**
 * A page with what nobody sees taken off (comments, Outlook's [if mso]
 * copies, the head and its title, hidden preheaders) and the message it
 * answers too: from the first quote's start, and every blockquote with what
 * it holds. A forward's quote is kept, since what it forwards is the
 * message. An element left out with no end leaves out the rest.
 */
export function pageWithoutQuotes(html) {
  return walkPage(withoutUnseen(html), true);
}

/** The walk: what is hidden left out, and with `dropQuotes` the quoted message too. */
function walkPage(s, dropQuotes) {
  let out = '';
  let at = 0;
  // An element being left out: `depth` counts those of its name inside it, `nest` the other elements open inside it.
  let skip = null;
  let forwardChecks = 0;
  TAG.lastIndex = 0;
  for (let m = TAG.exec(s); m; m = TAG.exec(s)) {
    const close = m[1] === '/';
    const name = m[2].toLowerCase();
    const attrs = m[3];
    if (skip) {
      const endsHere = skip.depth === 1 && skip.nest === 0 && Boolean(IMPLIED_END[skip.name]?.(close, name));
      if (!endsHere) {
        const opens = !close && !VOID.has(name) && !selfClosed(attrs);
        if (name === skip.name) skip.depth += close ? -1 : opens ? 1 : 0;
        else if (close) skip.nest = Math.max(0, skip.nest - 1);
        else if (opens) skip.nest += 1;
        if (skip.depth === 0) {
          skip = null;
          at = TAG.lastIndex;
        }
        continue;
      }
      // Ended by this tag, which is the page's again.
      skip = null;
      at = m.index;
    }
    if (close) continue;
    const quote = dropQuotes && Boolean(QUOTE_START[name]?.(attrs));
    if (quote) {
      // A forward is read: what it forwards is the message. Past a few dozen, a page is not a forward.
      const forward = forwardChecks++ < 50 && FORWARD_IN_PAGE.test(stripTags(s.slice(m.index, m.index + 600)));
      if (!forward) return out + s.slice(at, m.index);
      continue;
    }
    if (VOID.has(name) || selfClosed(attrs)) continue;
    if ((dropQuotes && name === 'blockquote') || hides(name, attrs)) {
      out += s.slice(at, m.index);
      skip = { name, depth: 1, nest: 0 };
    }
  }
  return skip ? out : out + s.slice(at);
}

const NAMED = {
  nbsp: ' ', amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", rsquo: '’', lsquo: '‘', rdquo: '”', ldquo: '“',
  sbquo: '‚', bdquo: '„', ndash: '\u2013', mdash: '\u2014', hellip: '…', bull: '•', middot: '·', copy: '©',
  reg: '®', trade: '™', euro: '€', pound: '£', yen: '¥', cent: '¢', deg: '°', times: '×',
  divide: '÷', laquo: '«', raquo: '»', shy: '', zwnj: '', zwj: '', lrm: '', rlm: '', zwsp: '', szlig: 'ß',
  aelig: 'æ', AElig: 'Æ', oslash: 'ø', Oslash: 'Ø', eth: 'ð', ETH: 'Ð', thorn: 'þ', THORN: 'Þ',
  iexcl: '¡', iquest: '¿', sect: '§', para: '¶', ordm: 'º', ordf: 'ª', frac12: '½', frac14: '¼', frac34: '¾',
};
const MARKS = { acute: '́', grave: '̀', circ: '̂', uml: '̈', tilde: '̃', ring: '̊', cedil: '̧', caron: '̌' };

/** HTML's character references, named and numbered, as the characters they stand for; one not known is left as written. */
export function decodeEntities(text) {
  return String(text || '').replace(/&(#x[0-9a-f]{1,6}|#[0-9]{1,7}|[a-z][a-z0-9]{1,10});/gi, (all, ref) => {
    if (ref[0] === '#') {
      const code = ref[1] === 'x' || ref[1] === 'X' ? parseInt(ref.slice(2), 16) : parseInt(ref.slice(1), 10);
      if (!Number.isFinite(code) || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff)) return all;
      if (code === 0xa0) return ' ';
      return code === 0 ? '' : String.fromCodePoint(code);
    }
    if (Object.hasOwn(NAMED, ref)) return NAMED[ref];
    const m = /^([a-zA-Z])(acute|grave|circ|uml|tilde|ring|cedil|caron)$/.exec(ref);
    return m ? (m[1] + MARKS[m[2]]).normalize('NFC') : all;
  });
}

/* ── the words of a message's text ───────────────────────────────────────── */

const WROTE_OPEN = /^(?:On|Am|Le|El|Il giorno|Il|Op|Em|W dniu)\b/i;
// German and Dutch name the writer after the verb: "schrieb Amina Yusuf <amina@...>:".
const WROTE_END = /(?:\b(?:wrote|a [ée]crit|escribi[óo]|ha scritto|escreveu|napisa[łl](?:\(a\)|a)?)|\b(?:schrieb|schreef)(?:\s+[^:\n]{1,160})?)\s*:$/i;
const FROM_LINE = /^(?:From|Von|De|Da|Van|Od|Från|Fra|Lähettäjä|От|من|发件人|發件人|差出人)\s*:\s*\S/i;
const SENT_LINE = /^(?:Sent|Date|Gesendet|Datum|Envoy[ée]|Enviado|Fecha|Inviato|Data|Verzonden|Wys[łl]ano|Skickat|Sendt|L[äa]hetetty|Отправлено|Дата|تاريخ|التاريخ|已发送|发送时间|日期|送信日時)\s*:\s*\S/i;
const ORIGINAL = /^-{2,}\s*(?:Original Message|Urspr[üu]ngliche Nachricht|Message d'origine|Mensaje original|Messaggio originale|Oorspronkelijk bericht|Mensagem original|Ursprungligt meddelande)\s*-{2,}$/i;
/** A line of the header a forward carries: From, Date, Subject, To and Cc in the languages above. */
const FORWARD_HEADER = /^(?:From|Date|Sent|Subject|To|Cc|Von|Datum|Gesendet|Betreff|An|De|Date|Envoy[ée]|Objet|[ÀA]|Para|Fecha|Asunto|Da|Oggetto|Inviato|Van|Onderwerp|Aan|Verzonden)\s*:/i;

/** Is this line where a quoted message, or the signature, starts? */
function quoteStartsAt(lines, i) {
  const line = lines[i].trim();
  const next = (lines[i + 1] || '').trim();
  if (/^-- ?$/.test(lines[i])) return true; // the signature's own line
  if (ORIGINAL.test(line)) return true;
  if (WROTE_OPEN.test(line) && line.length <= 240 && WROTE_END.test(line)) return true;
  if (WROTE_OPEN.test(line) && line.length <= 240 && next.length <= 60 && WROTE_END.test(next)) return true; // the verb wrapped onto a line of its own
  if (/^_{10,}$/.test(line) && FROM_LINE.test(next)) return true; // Outlook's rule above the header
  if (FROM_LINE.test(line)) {
    // Outlook's header of the message answered: From, then Sent or Date within a few lines.
    for (let j = i + 1; j < Math.min(lines.length, i + 5); j++) if (SENT_LINE.test(lines[j].trim())) return true;
  }
  return false;
}

/**
 * Plain words with the quoted message, the signature and the "> " lines
 * taken off. A forward's own line and the header under it are left out and
 * what it forwards is read, as the message it is.
 */
export function textWithoutQuotes(text) {
  const lines = String(text || '').replace(/\r\n?/g, '\n').split('\n');
  const kept = [];
  for (let i = 0; i < lines.length; i++) {
    // A blank line starts no quote, signature or forward.
    if (!lines[i].trim()) {
      kept.push('');
      continue;
    }
    if (FORWARD_LINE.test(lines[i].trim())) {
      // The header under the forward's line (Apple Mail leaves a blank line first): its From, Date, Subject and To, to the blank line.
      let j = i + 1;
      while (j < lines.length && !lines[j].trim()) j++;
      if (j < lines.length && FORWARD_HEADER.test(lines[j].trim())) {
        i = j;
        while (i + 1 < lines.length && FORWARD_HEADER.test(lines[i + 1].trim())) i++;
      }
      kept.push('');
      continue;
    }
    if (quoteStartsAt(lines, i)) break;
    if (/^\s*>/.test(lines[i])) continue;
    kept.push(lines[i]);
  }
  return kept.join('\n').trim();
}

/** A web address as it is said: its site alone, not every character of it. */
const siteOf = (url) => {
  try {
    return new URL(url).hostname.replace(/^www\./i, '');
  } catch {
    return '';
  }
};

/**
 * The words of a message to read: its text, or its page's, with what it
 * quotes left out. A message that is all quote is read whole, less the
 * "> " marks, rather than not at all. Paragraphs come back one each, each a
 * piece of its own for the reader.
 * @returns {string[]}
 */
export function wordsToRead(message) {
  const fromText = message?.text ? String(message.text) : '';
  const fromPage = !fromText && message?.html ? String(message.html) : '';
  const pageWords = (html) => decodeEntities(stripTags(html));
  let words = fromText ? textWithoutQuotes(fromText) : textWithoutQuotes(pageWords(pageWithoutQuotes(fromPage)));
  // Line by line: under /m, a "\s" at a line's start runs on over the newlines, a scan to the end from each.
  const unquoted = (s) => s.split('\n').map((line) => line.replace(/^[ \t]*>+ ?/, '')).join('\n').trim();
  // All quote: read whole, still without what nobody sees (review 3); all hidden: read less its head and comments.
  if (!words) words = unquoted(fromText || pageWords(walkPage(withoutUnseen(fromPage), false)));
  if (!words && fromPage) words = unquoted(pageWords(withoutUnseen(fromPage)));
  return words
    // words-ok: a regular expression for a link, not words
    .replace(/<?\bhttps?:\/\/[^\s<>()"]+>?/gi, (url) => siteOf(url.replace(/^<|>$/g, '')))
    .replace(/\u00a0/g, ' ')
    .split(/\n[ \t\r\f\v]*\n/)
    .map((p) => p.split('\n').map((line) => line.trim()).filter(Boolean).join(' ').replace(/[ \t]{2,}/g, ' ').trim())
    .filter((p) => /[\p{L}\p{N}]/u.test(p));
}

/**
 * The language of a paragraph, from its script, for the voice that reads it:
 * Urdu or Arabic (Urdu by the letters Arabic does not use), Hindi, or null
 * for the default voice. Latin script is left to the default.
 */
export function languageOf(text) {
  const s = String(text || '');
  const letters = (s.match(/\p{L}/gu) || []).length;
  if (!letters) return null;
  const arabic = (s.match(/[\u0600-\u06ff\u0750-\u077f\ufb50-\ufdff\ufe70-\ufeff]/g) || []).length;
  const devanagari = (s.match(/[ऀ-ॿ]/g) || []).length;
  if (arabic * 2 > letters) return /[ٹڈڑںھہے]/.test(s) ? 'ur' : 'ar';
  if (devanagari * 2 > letters) return 'hi';
  return null;
}

/** Who sent a message, as it is said: the name, else the address. */
const senderOf = (message) => displayName(Array.isArray(message?.from) ? message.from[0] : message?.from);

/**
 * A message as the reader takes it: who sent it, its subject, then its words.
 * Each piece says which message and part it is, for the window to follow.
 * @param {object} message  as mail.message() answers it
 * @param {string} key      the row's key, to find it in the list
 * @returns {Array<{ text, source: { key, part }, language }>}
 */
export function piecesOf(message, key) {
  const subject = String(message?.subject || '').trim();
  return [
    { text: t('mail.read-aloud.from-name', { name: senderOf(message) }), source: { key, part: 'from' }, language: null },
    { text: subject ? t('mail.read-aloud.subject-subject', { subject }) : t('mail.no-subject'), source: { key, part: 'subject' }, language: null },
    ...wordsToRead(message).map((text) => ({ text, source: { key, part: 'body' }, language: languageOf(text) })),
  ];
}

/* ── what a play reads, and the window around it ─────────────────────────── */

/**
 * The messages a play reads, in the list's order: a thread's messages one
 * after another where the list is in threads.
 * @param {object} o
 * @param {Array} o.threads   the list as drawn
 * @param {(thread) => Array} o.rowsIn  a thread's messages (the thread alone when not threaded)
 * @param {'all'|'unread'} o.mode
 */
export function queueOf({ threads = [], rowsIn = (x) => x.messages || [x], mode = 'all' }) {
  const out = [];
  for (const thread of threads) for (const row of rowsIn(thread)) if (mode !== 'unread' || row.unread) out.push(row);
  return out;
}

/**
 * Where the list being read came from. Reading stops when this changes: an
 * account, a folder, a folder of All accounts (its role), another view
 * (Attachments, People) or another search scope.
 */
export const placeOfList = ({ accountId = null, folder = null, role = null, view = null, scope = null, unified = false } = {}) =>
  JSON.stringify([accountId, unified ? null : folder, unified ? role : null, view, scope]);

/**
 * What opening a message does, as the selection asks: a draft goes to the
 * form, the form held in the reading pane is left for what was clicked, and
 * an unread message is marked read. A message shown by Read Aloud
 * (`quiet`) is shown and nothing else: no draft opened, no form left, and
 * it is marked read only once it has been read to its end.
 */
export function openingDoes({ quiet = false, readOnly = false, unread = false, inDrafts = false, moved = false, formUntouched = false } = {}) {
  if (quiet) return { editDraft: false, leaveForm: false, markRead: false };
  const editDraft = !readOnly && inDrafts;
  return { editDraft, leaveForm: !editDraft && moved && formUntouched, markRead: unread && !readOnly };
}

/**
 * The Read aloud group's buttons. A play's button stays pressed and live
 * while its run goes on, whatever the list does meanwhile (a message read
 * and marked leaves the Unread filter). Play all and Play unread are not
 * offered in a message's own window, which has no list to see.
 */
export function readButtons({ reading = null, rows = [], alone = false, hasMessage = false, openKey = null } = {}) {
  const on = (mode) => reading?.mode === mode;
  return {
    open: { pressed: on('open') && openKey != null && reading.key === openKey, disabled: !hasMessage },
    all: alone ? null : { pressed: on('all'), disabled: !on('all') && !rows.length },
    unread: alone ? null : { pressed: on('unread'), disabled: !on('unread') && !rows.some((r) => r.unread) },
  };
}

/* ── the player ──────────────────────────────────────────────────────────── */

/**
 * The player of a list of messages: each asked for when its turn comes,
 * read through the shared reader, then the next.
 * @param {object} o
 * @param {object} o.reader     createReader()'s answer
 * @param {(row) => Promise<object>} o.load   the message for a row
 * @param {(row) => string} o.keyOf
 * @param {(row, n, count) => void} [o.onMessage]  each message as its reading starts
 * @param {(row, message) => void} [o.onRead]  each message read to its end
 * @param {(state) => void} [o.onState]   { playing, n, count, row }
 * @param {() => void} [o.onEnd]   the last one read
 * @param {() => void} [o.onStall] the reader stopped by itself (a voice that failed, no speech): the run is ended
 */
export function createMailReading({ reader, load, keyOf, onMessage = () => {}, onRead = () => {}, onState = () => {}, onEnd = () => {}, onStall = () => {} }) {
  let rows = [];
  let at = -1;
  let run = 0;
  let playing = false;
  let waiting = null; // { run, done } while a message is being read
  const state = () => ({ playing, n: at + 1, count: rows.length, row: rows[at] || null });
  const tell = () => onState(state());

  async function playAt(i, mine) {
    for (; i < rows.length; i++) {
      if (mine !== run) return;
      at = i;
      tell();
      let message = null;
      try {
        message = await load(rows[i]);
      } catch {
        message = null; // gone since the list was drawn: the next one
      }
      if (mine !== run) return;
      if (!message) continue;
      onMessage(rows[i], i + 1, rows.length);
      const ended = new Promise((resolve) => {
        waiting = { run: mine, done: resolve };
        reader.read(piecesOf(message, keyOf(rows[i])), { onEnd: () => resolve(true) }).then((started) => { if (!started) resolve(true); }, () => resolve(false));
      });
      const toTheEnd = await ended;
      if (waiting?.run === mine) waiting = null;
      if (mine !== run) return;
      if (!toTheEnd) return; // stalled: the run was ended there
      onRead(rows[i], message);
    }
    if (mine !== run) return;
    playing = false;
    tell();
    onEnd();
  }

  return {
    state,
    /** Plays these rows from the first (or from `from`). Answers false when there is nothing to play. */
    play(list, from = 0) {
      // The run before is over first, so stopping the reader is not taken for a stall.
      const mine = (run += 1);
      waiting = null;
      reader.stop();
      rows = [...(list || [])];
      if (!rows.length) {
        playing = false;
        at = -1;
        tell();
        return false;
      }
      playing = true;
      playAt(Math.max(0, Math.min(rows.length - 1, from)), mine);
      return true;
    },
    /** The next message, now. The one left is not marked read. */
    next() {
      if (!playing) return;
      const mine = (run += 1);
      reader.stop();
      playAt(at + 1, mine);
    },
    /** The message before, from its start. */
    previous() {
      if (!playing) return;
      const mine = (run += 1);
      reader.stop();
      playAt(Math.max(0, at - 1), mine);
    },
    stop() {
      run += 1;
      waiting = null;
      reader.stop();
      if (!playing) return;
      playing = false;
      tell();
    },
    /**
     * The reader stopped by itself, mid-message, and will not say it ended
     * (the system voice failed, or there is no speech): the run ends here,
     * the message stays as it was. A stop the player made itself is not one.
     */
    stalled() {
      if (!playing || !waiting || waiting.run !== run) return false;
      const { done } = waiting;
      waiting = null;
      run += 1;
      playing = false;
      done(false);
      tell();
      onStall();
      return true;
    },
  };
}

/**
 * Has the shared reader stopped by itself mid-message? Its state says so:
 * not playing, not paused, and not past its last sentence (the end of a
 * message, which it does say).
 */
export const readerStalled = (s) => Boolean(s) && !s.playing && !s.paused && s.count > 0 && s.index < s.count;
/** Which runs mark a message read once it has been read to its end: Play all and Play unread. Read aloud on the open message changes nothing (the PM). */
export const marksReadAtEnd = (mode) => mode === 'all' || mode === 'unread';

/**
 * Whether a message read to its end is marked read: in a run that marks
 * (Play all, Play unread), unread, not read-only, not a local archive or a
 * file opened, and not marked read or unread by hand since the run began
 * (`touched`): what a person sets wins (review 2).
 */
export const marksReadAfterReading = ({ mode, unread, readOnly, local, touched }) =>
  marksReadAtEnd(mode) && Boolean(unread) && !readOnly && !local && !touched;

/**
 * The message Read Aloud shows, kept apart from the selection's own work on
 * a message opened (draft to the form, the form left, marked read). Read
 * Aloud marks the row it is about to show; the selection takes the mark when
 * it shows that row. A re-run for the same message because the certificates
 * changed (`tick`) is quiet too; a re-run for any other reason is not. Any
 * other selection, nothing selected, and the reading's end, stop or stall
 * drop the mark, so it never reaches a later click (reviews 2 and 3).
 */
export function createQuietOpen() {
  let pending = null;
  let shown = null; // { key, tick }: shown quietly, as of that tick
  return {
    /** Read Aloud is about to show this row. */
    mark(key) { pending = key; },
    /** The row was open already: the selection does not change, so nothing takes the mark. */
    unmark(key) { if (pending === key) pending = null; },
    /** The selection shows `key` now (null: nothing), as of `tick`. True when Read Aloud showed it. */
    take(key, tick = null) {
      const quiet = key != null && (pending === key || (shown?.key === key && shown.tick !== tick));
      pending = null;
      shown = quiet ? { key, tick } : null;
      return quiet;
    },
    /** The reading ended, stopped or stalled: nothing it showed stays quiet. */
    clear() {
      pending = null;
      shown = null;
    },
  };
}
