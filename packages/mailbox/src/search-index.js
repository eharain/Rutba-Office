// One search over every mailbox and every archive.
//
// The index is made of segments, one per folder of each account or opened
// archive, and is held in memory only. A segment is rebuilt when its folder
// changes and thrown away when the folder goes; nothing is written to disk,
// so there is no index file that can disagree with the mail. (An account's
// folders are the truth; an archive is never written to at all.)
//
// A segment is compact: its words are numbers in one vocabulary shared by
// every segment, and its postings - which messages hold which word - are one
// sorted run per word in a single typed array (compressed sparse rows). So a
// large archive costs a few bytes per word per message rather than a set of
// strings per word, and a search walks arrays rather than objects.
//
// Message bodies are indexed as words. A quoted phrase is first narrowed by
// its words in the index, then confirmed against the subject, the people and
// the attachment names held in memory, and only then against the body, read
// again from its source for those messages alone.

import { parseSearch, words } from './query.js';

/**
 * Words worth indexing.
 *
 * Split on everything that is not a letter, a digit or one of the three
 * characters that hold an address together - a search for `ali@example.com`
 * has to find it, and a tokeniser that breaks on `@` and `.` turns one term
 * into three that match half the mailbox.
 */
export function tokenise(text, out = new Set(), seq = null) {
  // A scanner rather than a split and a regular expression per word: bodies
  // are most of what is indexed, and this was the costliest step of it.
  // Letters include their marks, so a Hindi word is not cut at its vowels.
  const s = String(text || '').toLowerCase();
  const n = s.length;
  let start = -1;
  let joined = false; // the word holds @ or . and is also its parts
  for (let i = 0; i <= n; i++) {
    const kind = i < n ? charKind(s.charCodeAt(i), s, i) : 0;
    if (kind) {
      if (start < 0) {
        start = i;
        joined = false;
      }
      if (kind === 3) joined = true;
      continue;
    }
    if (start < 0) continue;
    let a = start;
    let b = i;
    start = -1;
    while (a < b && isTrim(s.charCodeAt(a))) a++;
    while (b > a && isTrim(s.charCodeAt(b - 1))) b--;
    const len = b - a;
    if (len < 2 || len > 64) continue;
    const token = s.slice(a, b);
    out.add(token);
    seq?.push(token);
    // An address is also its parts, so `ali@example.com` is found by `ali` and
    // by `example`.
    if (joined) for (const part of token.split(/[@.]/)) if (part.length > 1) out.add(part);
  }
  return out;
}

const isTrim = (c) => c === 46 || c === 95 || c === 45; // . _ -
const LETTER = /[\p{L}\p{N}\p{M}]/u;
const wide = new Map();

/** 0: a break; 1: part of a word; 3: @ or ., which also split it into parts. */
function charKind(c, s, i) {
  if (c < 128) {
    if ((c >= 97 && c <= 122) || (c >= 48 && c <= 57) || (c >= 65 && c <= 90)) return 1;
    if (c === 64 || c === 46) return 3;
    if (c === 95 || c === 43 || c === 45) return 1;
    return 0;
  }
  // A surrogate pair is one character; both halves take its kind.
  let ch = s[i];
  if (c >= 0xd800 && c <= 0xdbff) ch = s.slice(i, i + 2);
  else if (c >= 0xdc00 && c <= 0xdfff) ch = s.slice(i - 1, i + 1);
  let k = wide.get(ch);
  if (k === undefined) {
    k = LETTER.test(ch) ? 1 : 0;
    if (wide.size < 65536) wide.set(ch, k);
  }
  return k;
}

const lower = (s) => String(s || '').toLowerCase();
const norm = (s) => lower(s).replace(/\s+/g, ' ');
const person = (p) => (p ? `${p.name || ''} ${p.address || ''}` : '');
const people = (list) => (list || []).map(person).join(' ');

/** Everything about a message, but its body, that a phrase might be in. */
export function headlineText(doc) {
  return [doc.subject, person(doc.from), people(doc.to), people(doc.cc), people(doc.bcc), (doc.attachments || []).join(' '), doc.preview].join(' ');
}

/** A message as the index keeps it: what a result row shows, and nothing heavy. */
export function docOf(m, id) {
  const iso = m.date ? new Date(m.date) : null;
  const t = iso && !Number.isNaN(iso.getTime()) ? iso.getTime() : 0;
  const names = (m.attachments || []).map((a) => (typeof a === 'string' ? a : a.filename)).filter(Boolean);
  const from = Array.isArray(m.from) ? m.from[0] : m.from;
  return {
    id: String(id),
    subject: m.subject || '',
    from: from ? personOf(from) : null,
    to: peopleOf(m.to),
    cc: peopleOf(m.cc),
    bcc: peopleOf(m.bcc),
    t,
    size: Number(m.size || 0),
    unread: Boolean(m.unread),
    flagged: Boolean(m.flagged),
    pinned: Boolean(m.pinned),
    hasAttachments: Boolean(m.hasAttachments || names.length),
    attachmentCount: names.length || Number(m.attachmentCount || 0),
    attachments: names.length ? names : NONE,
    messageId: m.messageId ? String(m.messageId).replace(/^<|>$/g, '') : null,
    inReplyTo: m.inReplyTo ? String(m.inReplyTo).replace(/^<|>$/g, '') : null,
    references: m.references?.length ? m.references.slice(-6) : NONE,
    preview: own(String(m.preview || '').slice(0, 400).replace(/\s+/g, ' ').slice(0, 140)),
  };
}

// A hundred thousand messages name a few thousand people: each person is
// kept once and shared, which took the index of a 5 GB archive from 229 MB
// of records to a fraction of it.
const NONE = Object.freeze([]);

/**
 * A string of its own. A slice of a longer string can keep the whole of it
 * alive (a word cut from a body, the body), which in an index of a hundred
 * thousand messages is most of the memory it takes.
 */
function own(s) {
  return s.length < 13 ? s : Buffer.from(s, 'utf8').toString('utf8');
}
const PEOPLE = new Map();
function personOf(p) {
  const name = p.name || null;
  const address = p.address || '';
  const key = `${name}\u0001${address}`;
  let held = PEOPLE.get(key);
  if (!held) {
    held = Object.freeze({ name, address });
    if (PEOPLE.size > 500_000) PEOPLE.clear();
    PEOPLE.set(key, held);
  }
  return held;
}
function peopleOf(list) {
  if (!list?.length) return NONE;
  return list.slice(0, 50).map(personOf);
}

/**
 * Bytes of a message's pair signature: a bit set for each pair of words that
 * stand together in it, three bits a pair. A phrase whose pairs are not all
 * there is not in the message, so most messages that merely hold a phrase's
 * words are set aside without reading their bodies again. (A pair's bits can
 * be set by other pairs, so a match is still read and confirmed.)
 */
export const SIGNATURE = 256;
const pairBits = (a, b) => [(Math.imul(a, 0x9e3779b1) ^ Math.imul(b, 0x85ebca6b)) >>> 0, (Math.imul(a, 0xc2b2ae35) + Math.imul(b, 0x27d4eb2f)) >>> 0, (Math.imul(a ^ 0x5bd1e995, 0x01000193) ^ Math.imul(b + 0x6b43a9b5, 0x2545f491)) >>> 0].map((h) => (h ^ (h >>> 15)) & (SIGNATURE * 8 - 1));

/** Does the run of words `want` appear in `seq`, one after another? */
function inSequence(seq, want) {
  if (!want.length) return true;
  outer: for (let i = 0; i + want.length <= seq.length; i++) {
    for (let k = 0; k < want.length; k++) if (seq[i + k] !== want[k]) continue outer;
    return true;
  }
  return false;
}

const sequenceOf = (text) => {
  const seq = [];
  tokenise(text, new Set(), seq);
  return seq;
};

/** Binary search for `value` in a sorted Int32Array; its index or -1. */
function find(sorted, value) {
  let lo = 0;
  let hi = sorted.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    const v = sorted[mid];
    if (v === value) return mid;
    if (v < value) lo = mid + 1;
    else hi = mid - 1;
  }
  return -1;
}

export class SearchIndex {
  #vocab = new Map();
  #tokens = [];
  /** Token ids in the order of their text, for prefix lookups. */
  #sorted = [];
  #sortedCount = 0;
  #scratch = new Int32Array(1024);

  /** @param {{ now?: () => number }} [opts] */
  constructor({ now = () => Date.now() } = {}) {
    this.now = now;
    /** @type {Map<string, object>} */
    this.segments = new Map();
  }

  get vocabularySize() {
    return this.#tokens.length;
  }

  /** The id of a word, made when it is new. */
  idOf(token) {
    let id = this.#vocab.get(token);
    if (id === undefined) {
      token = own(token);
      id = this.#tokens.length;
      this.#vocab.set(token, id);
      this.#tokens.push(token);
    }
    return id;
  }

  /**
   * A message's words and its pair signature, from its texts (the headline,
   * the body): what a builder's `add` takes.
   */
  analyse(...texts) {
    const words = new Set();
    const sig = new Uint8Array(SIGNATURE);
    for (const text of texts) {
      const seq = [];
      tokenise(text, words, seq);
      let prev = -1;
      for (const token of seq) {
        const id = this.idOf(token);
        if (prev >= 0) for (const bit of pairBits(prev, id)) sig[bit >> 3] |= 1 << (bit & 7);
        prev = id;
      }
    }
    return { words, sig };
  }

  /** A set of words as sorted ids. */
  ids(tokens) {
    const out = new Int32Array(tokens.size ?? tokens.length);
    let i = 0;
    for (const token of tokens) out[i++] = this.idOf(token);
    return out.sort();
  }

  /**
   * A segment in the making. `add` each message with its words (a Set) or
   * its word ids (as `ids` returned them, kept by a caller that rebuilds the
   * same folder often); `finish` makes it searchable.
   */
  builder({ key, stamp = null, source, folder, textOf = null }) {
    const index = this;
    const docs = [];
    const lists = [];
    const sigs = [];
    return {
      add(doc, tokens, sig = null) {
        const ids = tokens instanceof Int32Array ? tokens : index.ids(tokens);
        docs.push(doc);
        lists.push(ids);
        sigs.push(sig);
        return ids;
      },
      get size() {
        return docs.length;
      },
      finish() {
        const segment = index.#compile({ key, stamp, source, folder, textOf }, docs, lists);
        // Pair signatures side by side; a message without one is all ones (always read).
        segment.sigs = new Uint8Array(docs.length * SIGNATURE);
        sigs.forEach((sig, d) => (sig ? segment.sigs.set(sig, d * SIGNATURE) : segment.sigs.fill(255, d * SIGNATURE, (d + 1) * SIGNATURE)));
        index.segments.set(key, segment);
        return segment;
      },
    };
  }

  /** Postings as compressed sparse rows. */
  #compile(meta, docs, lists) {
    const vocab = this.#tokens.length;
    if (this.#scratch.length < vocab) this.#scratch = new Int32Array(Math.max(vocab, this.#scratch.length * 2));
    const count = this.#scratch;
    const touched = [];
    let total = 0;
    for (const ids of lists) {
      for (let k = 0; k < ids.length; k++) {
        const id = ids[k];
        if (count[id]++ === 0) touched.push(id);
      }
      total += ids.length;
    }
    const tokens = Int32Array.from(touched).sort();
    const offsets = new Uint32Array(tokens.length + 1);
    let at = 0;
    for (let i = 0; i < tokens.length; i++) {
      offsets[i] = at;
      const n = count[tokens[i]];
      count[tokens[i]] = at; // now the write cursor
      at += n;
    }
    offsets[tokens.length] = at;
    const postings = new Uint32Array(total);
    lists.forEach((ids, d) => {
      for (let k = 0; k < ids.length; k++) postings[count[ids[k]]++] = d;
    });
    for (const id of touched) count[id] = 0;
    return { ...meta, docs, tokens, offsets, postings };
  }

  drop(key) {
    this.segments.delete(key);
  }

  /** Bytes the postings take, for the memory a large archive costs. */
  postingBytes() {
    let n = 0;
    for (const s of this.segments.values()) n += s.tokens.byteLength + s.offsets.byteLength + s.postings.byteLength;
    return n;
  }

  /** Every word id that starts with `prefix`. */
  #prefixIds(prefix, cap = 20000) {
    if (this.#sortedCount !== this.#tokens.length) {
      const fresh = [];
      for (let i = this.#sortedCount; i < this.#tokens.length; i++) fresh.push(i);
      const t = this.#tokens;
      fresh.sort((a, b) => (t[a] < t[b] ? -1 : t[a] > t[b] ? 1 : 0));
      // Merge the new words into the sorted list rather than sort it all again.
      const merged = new Array(this.#sorted.length + fresh.length);
      let i = 0;
      let j = 0;
      let k = 0;
      const old = this.#sorted;
      while (i < old.length && j < fresh.length) merged[k++] = t[old[i]] <= t[fresh[j]] ? old[i++] : fresh[j++];
      while (i < old.length) merged[k++] = old[i++];
      while (j < fresh.length) merged[k++] = fresh[j++];
      this.#sorted = merged;
      this.#sortedCount = this.#tokens.length;
    }
    const t = this.#tokens;
    const s = this.#sorted;
    let lo = 0;
    let hi = s.length;
    while (lo < hi) {
      const mid = (lo + hi) >>> 1;
      if (t[s[mid]] < prefix) lo = mid + 1;
      else hi = mid;
    }
    const out = [];
    for (let i = lo; i < s.length && out.length < cap; i++) {
      if (!t[s[i]].startsWith(prefix)) break;
      out.push(s[i]);
    }
    return Int32Array.from(out).sort();
  }

  /** Mark in `bits` every message of `segment` holding any of `ids`. */
  static #mark(segment, ids, bits) {
    const { tokens, offsets, postings } = segment;
    // Walk whichever list is shorter.
    if (ids.length < tokens.length / 8) {
      for (let k = 0; k < ids.length; k++) {
        const at = find(tokens, ids[k]);
        if (at < 0) continue;
        for (let p = offsets[at]; p < offsets[at + 1]; p++) bits[postings[p]] = 1;
      }
    } else {
      let a = 0;
      let b = 0;
      while (a < ids.length && b < tokens.length) {
        if (ids[a] === tokens[b]) {
          for (let p = offsets[b]; p < offsets[b + 1]; p++) bits[postings[p]] = 1;
          a++;
          b++;
        } else if (ids[a] < tokens[b]) a++;
        else b++;
      }
    }
    return bits;
  }

  /**
   * Search.
   * @param {string} query
   * @param {{ limit?: number, offset?: number, sources?: Iterable<string>|null, keys?: Iterable<string>|null, verifyCap?: number }} [opts]
   * @returns {{ rows: object[], total: number, exact: boolean, searched: number }}
   */
  search(query, { limit = 200, offset = 0, sources = null, keys = null, verifyCap = 600 } = {}) {
    const { clauses, empty } = parseSearch(query, { now: this.now() });
    if (empty) return { rows: [], total: 0, exact: true, searched: 0 };
    const wantSources = sources ? new Set(sources) : null;
    const wantKeys = keys ? new Set(keys) : null;

    // Word ids, once for every segment.
    const prepared = clauses.map((c) => ({
      negate: c.negate,
      any: c.any.map((atom) => {
        if (atom.kind === 'term') return { ...atom, ids: this.#prefixIds(atom.value) };
        if (atom.kind === 'phrase') {
          const ids = atom.words.map((w) => this.#vocab.get(w));
          const known = !ids.some((x) => x === undefined);
          const bits = [];
          if (known) for (let k = 1; k < ids.length; k++) bits.push(...pairBits(ids[k - 1], ids[k]));
          return { ...atom, wordIds: known ? ids.map((x) => Int32Array.of(x)) : null, bits };
        }
        return atom;
      }),
    }));
    const tokenish = (atom) => atom.kind === 'term' || atom.kind === 'phrase';
    const lead = prepared.filter((c) => !c.negate && c.any.every(tokenish));
    const rest = prepared.filter((c) => !lead.includes(c));

    const hits = [];
    let searched = 0;
    for (const segment of this.segments.values()) {
      if (wantSources && !wantSources.has(segment.source.id)) continue;
      if (wantKeys && !wantKeys.has(segment.key)) continue;
      const n = segment.docs.length;
      if (!n) continue;
      searched += n;
      const cache = new Map();
      const bitsOf = (atom) => {
        if (cache.has(atom)) return cache.get(atom);
        let bits;
        if (atom.kind === 'term') bits = SearchIndex.#mark(segment, atom.ids, new Uint8Array(n));
        else if (!atom.wordIds) bits = new Uint8Array(n);
        else {
          bits = SearchIndex.#mark(segment, atom.wordIds[0], new Uint8Array(n));
          for (let w = 1; w < atom.wordIds.length; w++) {
            const next = SearchIndex.#mark(segment, atom.wordIds[w], new Uint8Array(n));
            for (let i = 0; i < n; i++) bits[i] &= next[i];
          }
        }
        cache.set(atom, bits);
        return bits;
      };

      let cand = null;
      for (const clause of lead) {
        const bits = new Uint8Array(n);
        for (const atom of clause.any) {
          const b = bitsOf(atom);
          for (let i = 0; i < n; i++) bits[i] |= b[i];
        }
        if (cand) for (let i = 0; i < n; i++) cand[i] &= bits[i];
        else cand = bits;
      }

      const folderText = lower(`${segment.folder.path} ${segment.folder.name || ''}`);
      const sourceText = lower(`${segment.source.name || ''} ${segment.source.id}`);
      const test = (atom, doc, i) => {
        switch (atom.kind) {
          case 'term':
          case 'phrase':
            return bitsOf(atom)[i] === 1;
          case 'field': {
            const v = atom.value;
            if (atom.field === 'from') return lower(person(doc.from)).includes(v);
            if (atom.field === 'to') return lower(people(doc.to)).includes(v);
            if (atom.field === 'cc') return lower(people(doc.cc)).includes(v);
            if (atom.field === 'bcc') return lower(people(doc.bcc)).includes(v);
            if (atom.field === 'subject') return norm(doc.subject).includes(v);
            if (atom.field === 'filename') return doc.attachments.some((a) => lower(a).includes(v));
            if (atom.field === 'folder') return folderText.includes(v);
            if (atom.field === 'source') return sourceText.includes(v);
            return false;
          }
          case 'has':
            return doc.hasAttachments;
          case 'is':
            if (atom.value === 'unread') return doc.unread;
            if (atom.value === 'read') return !doc.unread;
            if (atom.value === 'starred') return doc.flagged;
            if (atom.value === 'unstarred') return !doc.flagged;
            return true;
          case 'date':
            if (!doc.t) return false;
            return (atom.from == null || doc.t >= atom.from) && (atom.to == null || doc.t < atom.to);
          case 'size':
            return (atom.min == null || doc.size >= atom.min) && (atom.max == null || doc.size <= atom.max);
          default:
            return true;
        }
      };

      for (let i = 0; i < n; i++) {
        if (cand && !cand[i]) continue;
        const doc = segment.docs[i];
        let ok = true;
        for (const clause of rest) {
          let holds = false;
          for (const atom of clause.any) {
            if (test(atom, doc, i)) {
              holds = true;
              break;
            }
          }
          if (clause.negate && holds && clause.any.length === 1 && clause.any[0].kind === 'phrase') {
            // Its words are all here; the phrase itself may not be.
            holds = this.#contains(segment, i, clause.any[0]);
          }
          if (holds === clause.negate) {
            ok = false;
            break;
          }
        }
        if (!ok) continue;
        // Phrases that must hold, still to be confirmed against the text.
        let pending = null;
        for (const clause of lead) {
          if (clause.any.length !== 1 || clause.any[0].kind !== 'phrase' || clause.any[0].loose) continue;
          const phrase = clause.any[0];
          // One word is its posting, already matched.
          if (phrase.words.length < 2) continue;
          if (inSequence(sequenceOf(headlineText(doc)), phrase.words)) continue;
          if (!SearchIndex.#maybe(segment, i, phrase.bits)) {
            pending = false;
            break;
          }
          (pending ||= []).push(phrase);
        }
        if (pending === false) continue;
        hits.push({ segment, doc, pending, i });
      }
    }

    hits.sort((a, b) => b.doc.t - a.doc.t || (a.doc.id < b.doc.id ? -1 : 1));

    // Confirm phrases against bodies: all of them when there are few enough
    // to read, else in date order until the page is full, the rest counted
    // by their pair signatures ('about' so many, and said to be so).
    const pendingCount = hits.reduce((n, h) => n + (h.pending ? 1 : 0), 0);
    let exact = true;
    let out = hits;
    if (pendingCount) {
      const confirmed = [];
      let read = 0;
      let held = 0;
      let unread = 0;
      for (const h of hits) {
        if (!h.pending) {
          confirmed.push(h);
          continue;
        }
        if (pendingCount > verifyCap && confirmed.length >= offset + limit) {
          unread++;
          continue;
        }
        read++;
        if (h.pending.every((p) => this.#contains(h.segment, h.i, p))) {
          confirmed.push(h);
          held++;
        }
      }
      exact = unread === 0;
      out = confirmed;
      // Those not read are counted at the share of those read that held.
      if (!exact) out.unconfirmed = Math.round((unread * held) / Math.max(1, read));
    }
    const total = out.length + (out.unconfirmed || 0);
    const rows = out.slice(offset, offset + limit).map(({ segment, doc }) => ({
      ...doc,
      date: doc.t ? new Date(doc.t).toISOString() : null,
      folder: segment.folder.path,
      folderName: segment.folder.name || segment.folder.path,
      accountId: segment.source.id,
      accountName: segment.source.name,
      source: segment.source.id,
      sourceName: segment.source.name,
      sourceKind: segment.source.kind,
    }));
    return { rows, total, exact, searched };
  }

  /** Could the message hold the phrase, by its pair signature. */
  static #maybe(segment, i, bits) {
    if (!segment.sigs) return true;
    const base = i * SIGNATURE;
    for (const bit of bits) if (!(segment.sigs[base + (bit >> 3)] & (1 << (bit & 7)))) return false;
    return true;
  }

  /** Does the message hold the phrase, its words one after another: its headline, then its body. */
  #contains(segment, i, phrase) {
    const doc = segment.docs[i];
    if (inSequence(sequenceOf(headlineText(doc)), phrase.words)) return true;
    if (phrase.words.length > 1 && !SearchIndex.#maybe(segment, i, phrase.bits || [])) return false;
    let text = null;
    try {
      text = segment.textOf ? segment.textOf(doc) : null;
    } catch {
      text = null;
    }
    return text ? inSequence(sequenceOf(text), phrase.words) : false;
  }
}

export { words };
