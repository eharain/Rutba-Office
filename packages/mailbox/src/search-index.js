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
export function tokenise(text, out = new Set()) {
  for (const raw of String(text || '').toLowerCase().split(/[^\p{L}\p{N}@._+-]+/u)) {
    const token = raw.replace(/^[._-]+|[._-]+$/g, '');
    if (token.length < 2 || token.length > 64) continue;
    out.add(token);
    // An address is also its parts, so `ali@example.com` is found by `ali` and
    // by `example`.
    if (token.includes('@') || token.includes('.')) for (const part of token.split(/[@.]/)) if (part.length > 1) out.add(part);
  }
  return out;
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
  const slim = (list) => (list || []).slice(0, 50).map((r) => ({ name: r.name || null, address: r.address || '' }));
  const names = (m.attachments || []).map((a) => (typeof a === 'string' ? a : a.filename)).filter(Boolean);
  return {
    id: String(id),
    subject: m.subject || '',
    from: m.from ? (Array.isArray(m.from) ? slim(m.from)[0] || null : { name: m.from.name || null, address: m.from.address || '' }) : null,
    to: slim(m.to),
    cc: slim(m.cc),
    bcc: slim(m.bcc),
    date: t ? new Date(t).toISOString() : null,
    t,
    size: Number(m.size || 0),
    unread: Boolean(m.unread),
    flagged: Boolean(m.flagged),
    pinned: Boolean(m.pinned),
    hasAttachments: Boolean(m.hasAttachments || names.length),
    attachmentCount: names.length || Number(m.attachmentCount || 0),
    attachments: names,
    messageId: m.messageId ? String(m.messageId).replace(/^<|>$/g, '') : null,
    inReplyTo: m.inReplyTo ? String(m.inReplyTo).replace(/^<|>$/g, '') : null,
    references: (m.references || []).slice(-6),
    preview: String(m.preview || '').replace(/\s+/g, ' ').slice(0, 200),
  };
}

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
      id = this.#tokens.length;
      this.#vocab.set(token, id);
      this.#tokens.push(token);
    }
    return id;
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
    return {
      add(doc, tokens) {
        const ids = tokens instanceof Int32Array ? tokens : index.ids(tokens);
        docs.push(doc);
        lists.push(ids);
        return ids;
      },
      get size() {
        return docs.length;
      },
      finish() {
        const segment = index.#compile({ key, stamp, source, folder, textOf }, docs, lists);
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
  search(query, { limit = 200, offset = 0, sources = null, keys = null, verifyCap = 4000 } = {}) {
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
          return { ...atom, wordIds: ids.some((x) => x === undefined) ? null : ids.map((x) => Int32Array.of(x)) };
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
            holds = this.#contains(segment, doc, clause.any[0].value);
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
          const phrase = clause.any[0].value;
          if (norm(headlineText(doc)).includes(phrase)) continue;
          (pending ||= []).push(phrase);
        }
        hits.push({ segment, doc, pending });
      }
    }

    hits.sort((a, b) => b.doc.t - a.doc.t || (a.doc.id < b.doc.id ? -1 : 1));

    // Confirm phrases against bodies: all of them when there are few enough
    // to read, else in date order until the page is full.
    const pendingCount = hits.reduce((n, h) => n + (h.pending ? 1 : 0), 0);
    let exact = true;
    let out = hits;
    if (pendingCount) {
      const confirmed = [];
      let read = 0;
      let unread = 0;
      for (const h of hits) {
        if (!h.pending) {
          confirmed.push(h);
          continue;
        }
        if (read >= verifyCap && confirmed.length >= offset + limit) {
          unread++;
          continue;
        }
        read++;
        if (h.pending.every((p) => this.#contains(h.segment, h.doc, p))) confirmed.push(h);
      }
      exact = unread === 0;
      out = confirmed;
      if (!exact) out.unconfirmed = unread;
    }
    const total = out.length + (out.unconfirmed || 0);
    const rows = out.slice(offset, offset + limit).map(({ segment, doc }) => ({
      ...doc,
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

  /** Does the message hold the phrase - its headline, then its body. */
  #contains(segment, doc, phrase) {
    if (norm(headlineText(doc)).includes(phrase)) return true;
    let text = null;
    try {
      text = segment.textOf ? segment.textOf(doc) : null;
    } catch {
      text = null;
    }
    return text ? norm(text).includes(phrase) : false;
  }
}

export { words };
