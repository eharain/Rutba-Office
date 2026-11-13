// Read Aloud in Mail (D-2026-11-13-34), on the reader the apps share
// (read-aloud/reader.js, C7). Three ways to play: the message open; every
// message in the list, in the list's order; and the unread ones in the list.
// The list is the list as drawn, so a filter (Unread, Starred, Attachments)
// or a search decides what Play all reads.
//
// Each message is read as who sent it, its subject, then its words: the
// plain text, or the page's words when there is none, with the quoted
// message it answers, the signature and the long web addresses left out.
// One message is read at a time; the next is asked for only when its turn
// comes, so a list of five hundred costs nothing until it is played.

import { t } from '@rutba/office-ui/messages';
import { stripTags, displayName } from './parts.js';

/** Where a reply's quoted message starts in a page, in the mail programs in wide use. */
const QUOTE_STARTS = [
  /<div[^>]*class=["'][^"']*\bgmail_quote\b/i, // Gmail
  /<div[^>]*id=["']divRplyFwdMsg["']/i, // Outlook
  /<div[^>]*id=["']appendonsend["']/i, // Outlook on the web
  /<hr[^>]*id=["']stopSpelling["']/i, // Outlook.com, older
  /<div[^>]*class=["'][^"']*\bmoz-cite-prefix\b/i, // Thunderbird
  /<blockquote[^>]*type=["']cite["']/i, // Apple Mail, Thunderbird
  /<div[^>]*class=["'][^"']*\byahoo_quoted\b/i, // Yahoo
];

/** A page with the message it answers taken off: from the first quote's start, and every blockquote. */
export function pageWithoutQuotes(html) {
  let s = String(html || '');
  let cut = s.length;
  for (const re of QUOTE_STARTS) {
    const m = re.exec(s);
    if (m && m.index < cut) cut = m.index;
  }
  s = s.slice(0, cut);
  // Innermost first, so a quote inside a quote goes with it.
  for (let i = 0; i < 20 && /<blockquote/i.test(s); i++) {
    const next = s.replace(/<blockquote\b[^>]*>(?:(?!<blockquote\b)[\s\S])*?<\/blockquote>/gi, '');
    if (next === s) break;
    s = next;
  }
  return s;
}

/** Is this line where a quoted message, or the signature, starts? */
function quoteStartsAt(lines, i) {
  const line = lines[i].trim();
  const next = (lines[i + 1] || '').trim();
  if (/^-- ?$/.test(lines[i])) return true; // the signature's own line
  if (/^-{2,}\s*Original Message\s*-{2,}$/i.test(line)) return true;
  if (/^On\b.{4,200}\bwrote:$/i.test(line)) return true;
  if (/^On\b.{4,200}$/i.test(line) && /^wrote:$/i.test(next)) return true; // "On ... <x@y>" / "wrote:", wrapped
  if (/^_{10,}$/.test(line) && /^From:/i.test(next)) return true; // Outlook's rule above the header
  if (/^From:\s*\S/i.test(line)) {
    // Outlook's header of the message answered: From, then Sent or Date within a few lines.
    for (let j = i + 1; j < Math.min(lines.length, i + 5); j++) if (/^(Sent|Date):\s*\S/i.test(lines[j].trim())) return true;
  }
  return false;
}

/** Plain words with the quoted message, the signature and the "> " lines taken off. */
export function textWithoutQuotes(text) {
  const lines = String(text || '').replace(/\r\n?/g, '\n').split('\n');
  const kept = [];
  for (let i = 0; i < lines.length; i++) {
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
 * quotes left out. A message that is all quote (a forward with nothing
 * added, say) is read whole, less the "> " marks, rather than not at all.
 * Paragraphs come back one each, each a piece of its own for the reader.
 * @returns {string[]}
 */
export function wordsToRead(message) {
  const fromText = message?.text ? String(message.text) : '';
  const fromPage = !fromText && message?.html ? String(message.html) : '';
  let words = fromText ? textWithoutQuotes(fromText) : textWithoutQuotes(stripTags(pageWithoutQuotes(fromPage)));
  if (!words) words = (fromText || stripTags(fromPage)).replace(/^\s*>+ ?/gm, '').trim();
  return words
    .replace(/<?\bhttps?:\/\/[^\s<>()"]+>?/gi, (url) => siteOf(url.replace(/^<|>$/g, '')))
    .split(/\n\s*\n/)
    .map((p) => p.replace(/\s*\n\s*/g, ' ').replace(/[ \t]{2,}/g, ' ').trim())
    .filter((p) => /[\p{L}\p{N}]/u.test(p));
}

/** Who sent a message, as it is said: the name, else the address. */
const senderOf = (message) => displayName(Array.isArray(message?.from) ? message.from[0] : message?.from);

/**
 * A message as the reader takes it: who sent it, its subject, then its words.
 * Each piece says which message and part it is, for the window to follow.
 * @param {object} message  as mail.message() answers it
 * @param {string} key      the row's key, to find it in the list
 * @returns {Array<{ text, source: { key, part }, language: null }>}
 */
export function piecesOf(message, key) {
  const subject = String(message?.subject || '').trim();
  return [
    { text: t('From {name}', { name: senderOf(message) }), source: { key, part: 'from' }, language: null },
    { text: subject ? t('Subject: {subject}', { subject }) : t('(no subject)'), source: { key, part: 'subject' }, language: null },
    ...wordsToRead(message).map((text) => ({ text, source: { key, part: 'body' }, language: null })),
  ];
}

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
 * The player of a list of messages: each asked for when its turn comes,
 * read through the shared reader, then the next.
 * @param {object} o
 * @param {object} o.reader     createReader()'s answer
 * @param {(row) => Promise<object>} o.load   the message for a row
 * @param {(row) => string} o.keyOf
 * @param {(row, n, count) => void} [o.onMessage]  each message as its reading starts
 * @param {(state) => void} [o.onState]   { playing, n, count, row }
 * @param {() => void} [o.onEnd]
 */
export function createMailReading({ reader, load, keyOf, onMessage = () => {}, onState = () => {}, onEnd = () => {} }) {
  let rows = [];
  let at = -1;
  let run = 0;
  let playing = false;
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
        reader.read(piecesOf(message, keyOf(rows[i])), { onEnd: resolve }).then((started) => { if (!started) resolve(); });
      });
      await ended;
      if (mine !== run) return;
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
      reader.stop();
      rows = [...(list || [])];
      const mine = (run += 1);
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
    /** The next message, now. */
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
      reader.stop();
      if (!playing) return;
      playing = false;
      tell();
    },
  };
}
