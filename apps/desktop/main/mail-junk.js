// Junk.
//
// Outlook's junk filter in the parts people use: senders always let through
// and senders always sent to Junk, the server's own verdict where it gives
// one, and a filter that learns from what a person marks Junk and Not Junk.
// Nothing leaves the machine to decide: what the filter has learned is a
// count of words kept in settings, the way Thunderbird's has worked for
// twenty years.
//
// The filter is the one SpamBayes settled on: each word's chance of being
// junk from how often it turned up in each kind (Robinson's smoothing, so a
// word seen once is not taken as proof), the most telling words combined
// with Fisher's chi-square method. It says nothing until it has seen a few
// of each kind: a filter that has learned from two messages files mail by
// accident.
//
// Pure: settings and messages in, verdicts and new settings out.

import { withoutBlocks } from '@rutba/mailbox/mime';
import { authenticationReport } from './mail-insight.js';

/** How hard the filter looks: Outlook's four settings. */
export const JUNK_LEVELS = ['off', 'low', 'high', 'safeOnly'];

/** The score at or above which a level files a message as junk. */
const CUTOFF = { low: 0.9, high: 0.7 };

/** What the filter needs to have seen, of each kind, before it says anything. */
export const MIN_LEARNED = 5;

/** The most telling words a verdict weighs, and how far from even a word must be to count. */
const CLUES = 150;
const MIN_STRENGTH = 0.1;

/** The words kept, at most: past this the ones seen once go first. */
const MAX_TOKENS = 40000;
/** The messages remembered as learned, so a second mark does not count twice. */
const MAX_TRAINED = 5000;

/**
 * Outlook's Blocked Encodings list: a language's character sets, by the
 * name the dialog shows. Unicode and Western European are left off, as
 * blocking them would block nearly everything.
 */
export const JUNK_ENCODINGS = {
  arabic: { label: 'Arabic', charsets: ['windows-1256', 'iso-8859-6', 'asmo-708'] },
  baltic: { label: 'Baltic', charsets: ['windows-1257', 'iso-8859-4', 'iso-8859-13'] },
  centralEuropean: { label: 'Central European', charsets: ['windows-1250', 'iso-8859-2'] },
  chineseSimplified: { label: 'Chinese Simplified', charsets: ['gb2312', 'gbk', 'gb18030', 'hz-gb-2312', 'x-gbk'] },
  chineseTraditional: { label: 'Chinese Traditional', charsets: ['big5', 'big5-hkscs', 'x-x-big5'] },
  cyrillic: { label: 'Cyrillic', charsets: ['koi8-r', 'koi8-u', 'windows-1251', 'iso-8859-5', 'ibm866', 'cp866'] },
  greek: { label: 'Greek', charsets: ['windows-1253', 'iso-8859-7'] },
  hebrew: { label: 'Hebrew', charsets: ['windows-1255', 'iso-8859-8', 'iso-8859-8-i'] },
  japanese: { label: 'Japanese', charsets: ['iso-2022-jp', 'shift_jis', 'shift-jis', 'sjis', 'euc-jp', 'csiso2022jp'] },
  korean: { label: 'Korean', charsets: ['euc-kr', 'ks_c_5601-1987', 'iso-2022-kr', 'cp949'] },
  thai: { label: 'Thai', charsets: ['windows-874', 'tis-620', 'iso-8859-11'] },
  turkish: { label: 'Turkish', charsets: ['windows-1254', 'iso-8859-9'] },
  vietnamese: { label: 'Vietnamese', charsets: ['windows-1258'] },
};

export function defaultJunk() {
  return {
    level: 'low', safe: [], blocked: [], trustContacts: true,
    // Outlook's Safe Recipients, and its International tab: the top-level
    // domains and the encodings whose mail goes to Junk.
    safeRecipients: [], blockedTlds: [], blockedEncodings: [],
    model: { junk: 0, good: 0, tokens: {}, trained: {} },
  };
}

/** Settings as stored, with anything missing filled in — an older settings file has none of this. */
export function junkSettings(stored) {
  const base = defaultJunk();
  const s = stored && typeof stored === 'object' ? stored : {};
  const model = s.model && typeof s.model === 'object' ? s.model : {};
  return {
    level: JUNK_LEVELS.includes(s.level) ? s.level : base.level,
    safe: Array.isArray(s.safe) ? s.safe : [],
    blocked: Array.isArray(s.blocked) ? s.blocked : [],
    trustContacts: s.trustContacts !== false,
    safeRecipients: Array.isArray(s.safeRecipients) ? s.safeRecipients : [],
    blockedTlds: Array.isArray(s.blockedTlds) ? s.blockedTlds.map(tldEntry).filter(Boolean) : [],
    blockedEncodings: Array.isArray(s.blockedEncodings) ? s.blockedEncodings.filter((k) => k in JUNK_ENCODINGS) : [],
    model: {
      junk: Number(model.junk) || 0,
      good: Number(model.good) || 0,
      tokens: model.tokens && typeof model.tokens === 'object' ? model.tokens : {},
      trained: model.trained && typeof model.trained === 'object' ? model.trained : {},
    },
  };
}

/** A list entry tidied: an address, or a domain written "example.com" or "@example.com". */
export function listEntry(value) {
  const v = String(value || '').trim().toLowerCase().replace(/^mailto:/, '');
  if (!v) return null;
  if (v.includes('@') && !v.startsWith('@')) return /^[^\s@]+@[^\s@]+$/.test(v) ? v : null;
  const domain = v.replace(/^@/, '');
  return /^[a-z0-9-]+(\.[a-z0-9-]+)+$/.test(domain) ? '@' + domain : null;
}

/** A Blocked Top-Level Domain entry tidied: "ru", ".ru" or "*.ru" is "ru"; anything else, null. */
export function tldEntry(value) {
  const v = String(value || '').trim().toLowerCase().replace(/^\*?\./, '');
  return /^[a-z]{2,24}$/.test(v) || /^xn--[a-z0-9-]{1,59}$/.test(v) ? v : null;
}

/** Whether an address is on a list: the address itself, or its domain or one above it. */
export function onList(list, address) {
  const a = String(address || '').trim().toLowerCase();
  if (!a.includes('@')) return false;
  const domain = a.split('@').pop();
  return (list || []).some((raw) => {
    const entry = listEntry(raw);
    if (!entry) return false;
    if (!entry.startsWith('@')) return entry === a;
    const d = entry.slice(1);
    return domain === d || domain.endsWith('.' + d);
  });
}

const senderOf = (message) => {
  const from = Array.isArray(message?.from) ? message.from[0] : message?.from;
  return { address: String(from?.address || '').toLowerCase(), name: String(from?.name || '') };
};

/** A mailing list's posting address, from its List-Post header (RFC 2369), lower case; null for any other mail. */
export function listPostOf(message) {
  for (const h of message?.headers || []) {
    if (String(h.key || h.name || '').toLowerCase() !== 'list-post') continue;
    const m = /<mailto:([^>?]+)/i.exec(String(h.value || ''));
    if (m) {
      try { return decodeURIComponent(m[1]).trim().toLowerCase(); } catch { return m[1].trim().toLowerCase(); }
    }
  }
  return null;
}

/** The addresses a message was sent to: its To and Cc, and a mailing list's own posting address. */
export function recipientsOf(message) {
  const out = [...(message?.to || []), ...(message?.cc || [])].map((a) => String(a?.address || '').trim().toLowerCase()).filter((a) => a.includes('@'));
  const post = listPostOf(message);
  if (post) out.push(post);
  return [...new Set(out)];
}

/**
 * The address "Never block this group or mailing list" keeps: the list's
 * posting address when it has one, else the one address the message went
 * to when that is not one of `own` — the group's. Null for mail sent to a
 * person rather than a group.
 */
export function groupAddressOf(message, own = []) {
  const mine = new Set(own.map((a) => String(a || '').trim().toLowerCase()));
  const post = listPostOf(message);
  if (post && !mine.has(post)) return post;
  const to = (message?.to || []).map((a) => String(a?.address || '').trim().toLowerCase()).filter((a) => a.includes('@'));
  return to.length === 1 && !mine.has(to[0]) ? to[0] : null;
}

/** The character sets a message is written in: as its parse kept them, else from its headers. */
export function charsetsOf(message) {
  const out = new Set((message?.charsets || []).map((c) => String(c).toLowerCase()));
  for (const h of message?.headers || []) {
    const key = String(h.key || h.name || '').toLowerCase();
    const value = String(h.value || '');
    if (key === 'content-type') {
      const m = /charset\s*=\s*"?([^";\s]+)/i.exec(value);
      if (m) out.add(m[1].toLowerCase());
    }
    if (key === 'subject' || key === 'from') for (const m of value.matchAll(/=\?([^?*]+)(?:\*[^?]*)?\?[bqBQ]\?/g)) out.add(m[1].toLowerCase());
  }
  return [...out];
}

const stripHtml = (html) => withoutBlocks(html)
  .replace(/<[^<>]*>/g, ' ')
  .replace(/&nbsp;/g, ' ')
  .replace(/&amp;/g, '&');

/**
 * The words a message is judged by: its subject's (twice — once marked as
 * the subject's), the sender's address, domain and name, the body's, the
 * sites its links go to, and whether it is pictures and markup with no
 * plain words or carries an attachment.
 */
export function tokensOf(message) {
  const out = new Set();
  const add = (t, max = 24) => { if (t.length >= 2 && t.length <= max) out.add(t); };
  const words = (text, prefix = '') => {
    for (const w of String(text || '').toLowerCase().match(/[\p{L}\p{N}][\p{L}\p{N}'$%-]*/gu) || []) add(prefix + w, prefix ? 24 + prefix.length : 24);
  };
  words(message?.subject, 'subject:');
  words(message?.subject);
  const sender = senderOf(message);
  if (sender.address) {
    add('from:' + sender.address, 80);
    add('domain:' + sender.address.split('@').pop(), 80);
  }
  words(sender.name, 'name:');
  const body = message?.text || stripHtml(message?.html) || message?.preview || '';
  words(body.slice(0, 20000));
  for (const m of String(message?.html || body).matchAll(/https?:\/\/([a-z0-9.-]+)/gi)) add('url:' + m[1].toLowerCase(), 80);
  if (message?.html && !message?.text) add('meta:markup-only');
  if (message?.hasAttachments || message?.attachments?.length) add('meta:attachment');
  return [...out];
}

/** One word's chance of being junk, smoothed towards even when it has been seen little. */
function wordProbability(model, token) {
  const seen = model.tokens[token];
  if (!seen) return null;
  const [j, g] = seen;
  const junkRate = j / Math.max(1, model.junk);
  const goodRate = g / Math.max(1, model.good);
  if (junkRate + goodRate === 0) return null;
  const p = junkRate / (junkRate + goodRate);
  const n = j + g;
  // Robinson: a word seen n times moves from 0.5 towards what it showed, by n parts in n + 1.
  return (0.5 + n * p) / (1 + n);
}

/** Fisher's inverse chi-square: the chance of a value this large by luck, for an even number of degrees. */
function chi2Q(x2, degrees) {
  const m = x2 / 2;
  let term = Math.exp(-m);
  let sum = term;
  for (let i = 1; i < degrees / 2; i++) {
    term *= m / i;
    sum += term;
  }
  return Math.min(sum, 1);
}

/** Whether the filter has seen enough of each kind to say anything. */
export const ready = (model) => model.junk >= MIN_LEARNED && model.good >= MIN_LEARNED;

/**
 * How junk-like a message is, 0 to 1 — null while the filter has learned
 * too little. Near 1 is junk, near 0 good, near 0.5 unsure.
 */
export function junkScore(model, message) {
  if (!ready(model)) return null;
  const clues = tokensOf(message)
    .map((t) => wordProbability(model, t))
    .filter((p) => p !== null && Math.abs(p - 0.5) >= MIN_STRENGTH)
    .sort((a, b) => Math.abs(b - 0.5) - Math.abs(a - 0.5))
    .slice(0, CLUES)
    // Never quite certain, either way: one word must not settle it.
    .map((p) => Math.min(0.99, Math.max(0.01, p)));
  if (!clues.length) return 0.5;
  let junkLog = 0;
  let goodLog = 0;
  for (const p of clues) {
    junkLog += Math.log(1 - p);
    goodLog += Math.log(p);
  }
  const junk = 1 - chi2Q(-2 * junkLog, 2 * clues.length);
  const good = 1 - chi2Q(-2 * goodLog, 2 * clues.length);
  return (1 + junk - good) / 2;
}

/** Whether the server that delivered it already called it junk (SpamAssassin's headers, and Exchange's). */
export function serverSaysJunk(message) {
  for (const h of message?.headers || []) {
    const key = String(h.key || h.name || '').toLowerCase();
    const value = String(h.value || '').trim();
    if (key === 'x-spam-flag' && /^yes\b/i.test(value)) return true;
    if (key === 'x-spam-status' && /^yes\b/i.test(value)) return true;
    // Exchange's spam confidence level: 5 and over is junk to Outlook itself.
    if (key === 'x-ms-exchange-organization-scl' && Number(value) >= 5) return true;
  }
  return false;
}

/**
 * Where a message belongs, and why.
 *
 * @param {object} message as the store keeps it
 * @param {object} settings `junkSettings(...)`
 * @param {{ isContact?: (address: string) => boolean }} [o]
 * @returns {{ junk: boolean, why: string|null, score: number|null }}
 *   `why` is 'blocked', 'safe', 'safeRecipient', 'contact', 'blockedTld',
 *   'blockedEncoding', 'server', 'notSafe' or 'filter'.
 */
export function junkVerdict(message, settings, { isContact = null } = {}) {
  const { address } = senderOf(message);
  if (onList(settings.blocked, address)) return { junk: true, why: 'blocked', score: null };
  // A From or a To anyone can write: a message the delivering server called
  // junk is let in on a Safe Sender's, a Safe Recipient's or a contact's
  // name only when the server also found it really came from there.
  const trusted = !serverSaysJunk(message) || authenticationReport(message).ok === true;
  if (trusted && onList(settings.safe, address)) return { junk: false, why: 'safe', score: null };
  // Safe Recipients: mail sent to a group or a mailing list kept safe.
  if (trusted && (settings.safeRecipients || []).length && recipientsOf(message).some((a) => onList(settings.safeRecipients, a))) return { junk: false, why: 'safeRecipient', score: null };
  if (trusted && settings.trustContacts && address && isContact?.(address)) return { junk: false, why: 'contact', score: null };
  // The International lists, which work as Blocked Senders do, whatever the level.
  const tld = address.includes('@') ? address.split('@').pop().split('.').pop() : '';
  if (tld && (settings.blockedTlds || []).includes(tld)) return { junk: true, why: 'blockedTld', score: null };
  if ((settings.blockedEncodings || []).length) {
    const charsets = charsetsOf(message);
    if (settings.blockedEncodings.some((k) => JUNK_ENCODINGS[k]?.charsets.some((c) => charsets.includes(c)))) return { junk: true, why: 'blockedEncoding', score: null };
  }
  if (settings.level === 'off') return { junk: false, why: null, score: null };
  if (serverSaysJunk(message)) return { junk: true, why: 'server', score: null };
  if (settings.level === 'safeOnly') return { junk: true, why: 'notSafe', score: null };
  const score = junkScore(settings.model, message);
  if (score === null) return { junk: false, why: null, score: null };
  return { junk: score >= CUTOFF[settings.level], why: score >= CUTOFF[settings.level] ? 'filter' : null, score };
}

/** Words for the reading pane's note on a message in Junk. */
export const JUNK_REASONS = {
  blocked: 'its sender is on your Blocked Senders list',
  blockedTld: "its sender's address ends in a country or region on your Blocked Top-Level Domain list",
  blockedEncoding: 'it is written in an encoding on your Blocked Encodings list',
  server: 'the server that delivered it marked it as junk',
  notSafe: 'Junk Email Options lets in only Safe Senders and contacts',
  filter: 'the junk filter judged it junk from what you have marked before',
};

/** The key a message is remembered by once learned: its Message-ID, else its sender, subject and date. */
const trainedKey = (message) => message?.messageId || `${senderOf(message).address}|${message?.subject || ''}|${message?.date || ''}`;

/**
 * Learn from a message marked junk (true) or good (false). A message learned
 * the other way before is unlearned first; one learned this way already is
 * left as it is. Returns new settings; the old are not touched.
 */
export function learn(settings, message, junk) {
  const model = {
    junk: settings.model.junk,
    good: settings.model.good,
    tokens: { ...settings.model.tokens },
    trained: { ...settings.model.trained },
  };
  const key = trainedKey(message);
  const was = model.trained[key];
  const now = junk ? 'junk' : 'good';
  if (was === now) return settings;
  const tokens = tokensOf(message);
  const count = (kind, by) => {
    model[kind] = Math.max(0, model[kind] + by);
    const slot = kind === 'junk' ? 0 : 1;
    for (const t of tokens) {
      const pair = model.tokens[t] ? [...model.tokens[t]] : [0, 0];
      pair[slot] = Math.max(0, pair[slot] + by);
      if (pair[0] + pair[1] === 0) delete model.tokens[t];
      else model.tokens[t] = pair;
    }
  };
  if (was) count(was, -1);
  count(now, 1);
  delete model.trained[key];
  model.trained[key] = now;
  // The oldest remembered go first.
  const keys = Object.keys(model.trained);
  for (let i = 0; i < keys.length - MAX_TRAINED; i++) delete model.trained[keys[i]];
  prune(model);
  return { ...settings, model };
}

/** Keep the word count bounded: the words seen once go first, then twice. */
function prune(model) {
  let entries = Object.keys(model.tokens);
  for (let floor = 1; entries.length > MAX_TOKENS && floor < 1000; floor++) {
    for (const t of entries) {
      const [j, g] = model.tokens[t];
      if (j + g <= floor) delete model.tokens[t];
    }
    entries = Object.keys(model.tokens);
  }
}

/** Never Block this Group or Mailing List: the address put on the Safe Recipients list. Returns new settings. */
export function listRecipient(settings, address) {
  const entry = listEntry(address);
  if (!entry) throw new Error(`"${address}" is not an address or a domain.`);
  return { ...settings, safeRecipients: [...new Set([...(settings.safeRecipients || []).map(listEntry).filter(Boolean), entry])] };
}

/**
 * A sender put on one list and taken off the other: Block Sender, Never
 * Block Sender. Returns new settings.
 */
export function listSender(settings, address, list) {
  const entry = listEntry(address);
  if (!entry) throw new Error(`"${address}" is not an address or a domain.`);
  const other = list === 'blocked' ? 'safe' : 'blocked';
  return {
    ...settings,
    [list]: [...new Set([...settings[list].map(listEntry).filter(Boolean), entry])],
    [other]: settings[other].map(listEntry).filter((e) => e && e !== entry),
  };
}
