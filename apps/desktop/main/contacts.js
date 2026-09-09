// The address book.
//
// One JSON file in the profile, contacts.json, holding every card as the
// vCard reader gives it — plus an id and the two times. Nothing here reaches
// a server. What comes in: a .vcf or .csv chosen or dropped, a card read out
// of a message, a sender the person asked to keep. What goes out: a .vcf of
// some or all of them. Compose asks for suggestions and gets these cards and
// the people mail has seen, together and once each.

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { readVCards, writeVCards, readContactsCsv, displayName } from '@rutba/contacts';

const readJson = (file, fallback) => {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch {
    return fallback;
  }
};
const writeJson = (file, value) => {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = `${file}.${process.pid}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(value));
  fs.renameSync(tmp, file);
};

const norm = (s) => String(s || '').trim().toLowerCase();

/** A card as the window shows it: the photo as a URL it can put in an <img>. */
function present(c) {
  const photo = c.photo?.data ? `data:${c.photo.mediaType || 'image/jpeg'};base64,${c.photo.data}` : c.photo?.uri || null;
  return { ...c, display: displayName(c), photoUrl: photo, photo: undefined };
}

/** A card as the file keeps it: photo bytes as base64 text. */
function storable(c) {
  const photo = c.photo?.data
    ? { data: Buffer.isBuffer(c.photo.data) ? c.photo.data.toString('base64') : String(c.photo.data), mediaType: c.photo.mediaType || 'image/jpeg', uri: null }
    : c.photo?.uri ? { data: null, mediaType: c.photo.mediaType || null, uri: c.photo.uri } : null;
  return { ...c, photo };
}

/** The same card, as the vCard writer wants it: photo bytes as bytes. */
function forWriting(c) {
  return { ...c, photo: c.photo?.data ? { ...c.photo, data: Buffer.from(c.photo.data, 'base64') } : c.photo };
}

export function createContactsService({ stores, broadcast, people = null }) {
  const file = path.join(stores.dir, 'contacts.json');
  const state = readJson(file, { contacts: [] });
  if (!Array.isArray(state.contacts)) state.contacts = [];

  const save = () => {
    writeJson(file, state);
    broadcast?.('contacts:changed', { count: state.contacts.length });
  };
  const byId = (id) => state.contacts.find((c) => c.id === id) || null;
  const emailsOf = (c) => (c.emails || []).map((e) => norm(e.value)).filter(Boolean);

  /** Add a card, or fold it into the one that already has its address. */
  function merge(incoming) {
    const emails = emailsOf(incoming);
    const existing = emails.length ? state.contacts.find((c) => emailsOf(c).some((e) => emails.includes(e))) : null;
    const now = Date.now();
    if (!existing) {
      state.contacts.push({ ...storable(incoming), id: crypto.randomUUID(), createdAt: now, updatedAt: now });
      return 'added';
    }
    // Fill what the kept card lacks; never overwrite what somebody typed.
    let changed = false;
    const fill = (key) => {
      if ((existing[key] == null || existing[key] === '') && incoming[key]) {
        existing[key] = incoming[key];
        changed = true;
      }
    };
    for (const key of ['nickname', 'org', 'department', 'title', 'role', 'birthday', 'anniversary', 'note']) fill(key);
    if (!existing.name?.full && incoming.name?.full) {
      existing.name = incoming.name;
      changed = true;
    }
    for (const list of ['emails', 'phones', 'urls']) {
      for (const item of incoming[list] || []) {
        if (!(existing[list] || []).some((x) => norm(x.value) === norm(item.value))) {
          existing[list] = [...(existing[list] || []), item];
          changed = true;
        }
      }
    }
    for (const a of incoming.addresses || []) {
      if (!(existing.addresses || []).some((x) => norm(x.street) === norm(a.street) && norm(x.postcode) === norm(a.postcode))) {
        existing.addresses = [...(existing.addresses || []), a];
        changed = true;
      }
    }
    if (!existing.photo && incoming.photo) {
      existing.photo = storable(incoming).photo;
      changed = true;
    }
    for (const cat of incoming.categories || []) {
      if (!(existing.categories || []).includes(cat)) {
        existing.categories = [...(existing.categories || []), cat];
        changed = true;
      }
    }
    if (changed) existing.updatedAt = now;
    return changed ? 'updated' : 'same';
  }

  function readFile(target) {
    const ext = path.extname(target).toLowerCase();
    const text = fs.readFileSync(target, 'utf8');
    if (ext === '.csv' || ext === '.tsv') return readContactsCsv(text);
    return readVCards(text);
  }

  const sorted = () => [...state.contacts].sort((a, b) => displayName(a).localeCompare(displayName(b), undefined, { sensitivity: 'base' }));

  return {
    /** Every card, sorted by name; with `query`, the ones that match it anywhere that matters. */
    list: ({ query = '' } = {}) => {
      const q = norm(query);
      const all = sorted();
      const hit = q
        ? all.filter((c) => [displayName(c), c.org, c.title, c.nickname, ...(c.emails || []).map((e) => e.value), ...(c.phones || []).map((p) => p.value), ...(c.categories || [])].some((v) => norm(v).includes(q)))
        : all;
      return hit.map(present);
    },

    get: ({ id }) => {
      const c = byId(id);
      return c ? present(c) : null;
    },

    /** Keep a card: a new one, or a changed one by id. */
    save: ({ contact }) => {
      const now = Date.now();
      if (contact.id && byId(contact.id)) {
        const kept = byId(contact.id);
        const photo = contact.photoUrl && !contact.photo ? kept.photo : storable(contact).photo;
        Object.assign(kept, { ...contact, photo, photoUrl: undefined, display: undefined, updatedAt: now });
        save();
        return present(kept);
      }
      const fresh = { ...storable(contact), id: crypto.randomUUID(), createdAt: now, updatedAt: now, photoUrl: undefined, display: undefined };
      state.contacts.push(fresh);
      save();
      return present(fresh);
    },

    remove: ({ ids = [], id = null }) => {
      const gone = new Set([...ids, ...(id ? [id] : [])]);
      const before = state.contacts.length;
      state.contacts = state.contacts.filter((c) => !gone.has(c.id));
      if (state.contacts.length !== before) save();
      return { removed: before - state.contacts.length };
    },

    /**
     * Suggestions for an address line: the cards and the people mail has
     * seen, matched on name and address, once each, the cards first.
     */
    suggest: ({ query = '', limit = 8 } = {}) => {
      const q = norm(query);
      if (!q) return [];
      const out = [];
      const seen = new Set();
      for (const c of sorted()) {
        const name = displayName(c);
        for (const e of c.emails || []) {
          const email = norm(e.value);
          if (!email || seen.has(email)) continue;
          if (norm(name).includes(q) || email.includes(q)) {
            seen.add(email);
            out.push({ name, email: e.value, source: 'contacts', id: c.id });
          }
        }
      }
      if (out.length < limit && typeof people === 'function') {
        for (const p of people()) {
          const email = norm(p.email);
          if (!email || seen.has(email)) continue;
          if (norm(p.name).includes(q) || email.includes(q)) {
            seen.add(email);
            out.push({ name: p.name || '', email: p.email, source: 'mail' });
          }
          if (out.length >= limit) break;
        }
      }
      return out.slice(0, limit);
    },

    /** What a file holds, without keeping any of it — for the window to show and offer. */
    peek: ({ path: target }) => {
      const cards = readFile(target);
      return { path: target, name: path.basename(target), contacts: cards.map((c, i) => present({ ...storable(c), id: `file:${i}` })) };
    },

    /** Bring a file's cards in, folding each into the card that already has its address. */
    importFile: ({ path: target }) => {
      const cards = readFile(target);
      const result = { added: 0, updated: 0, same: 0, total: cards.length };
      for (const c of cards) result[merge(c)] += 1;
      if (result.added || result.updated) save();
      return result;
    },

    /** Cards handed over as text — a .vcf attachment, a pasted card. */
    importText: ({ text }) => {
      const cards = readVCards(text);
      const result = { added: 0, updated: 0, same: 0, total: cards.length };
      for (const c of cards) result[merge(c)] += 1;
      if (result.added || result.updated) save();
      return result;
    },

    /** Some or all of the cards, as a .vcf. */
    exportFile: ({ path: target, ids = null }) => {
      const chosen = ids?.length ? state.contacts.filter((c) => ids.includes(c.id)) : sorted();
      fs.writeFileSync(target, writeVCards(chosen.map(forWriting), { version: '3.0' }), 'utf8');
      return { path: target, count: chosen.length };
    },

    /** A sender, kept. Answers with the card, new or already there. */
    fromMail: ({ name, email }) => {
      const address = String(email || '').trim();
      if (!address) return null;
      const existing = state.contacts.find((c) => emailsOf(c).includes(norm(address)));
      if (existing) return present(existing);
      const [given = '', ...rest] = String(name || '').trim().split(/\s+/);
      const contact = {
        uid: null, version: '3.0',
        name: { full: String(name || '').trim(), given, family: rest.join(' '), middle: '', prefix: '', suffix: '' },
        nickname: null, org: null, department: null, title: null, role: null,
        emails: [{ value: address, type: 'other', pref: true, label: null }],
        phones: [], addresses: [], birthday: null, anniversary: null, note: null, urls: [], photo: null, categories: [], rev: null, kind: 'individual', extra: [],
      };
      merge(contact);
      save();
      return present(state.contacts.find((c) => emailsOf(c).includes(norm(address))));
    },

    count: () => state.contacts.length,
  };
}
