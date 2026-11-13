// The languages a first start adds or offers (L5).
//
// On Windows the installer's Languages page asks, and writes the packs ticked
// to HKCU\Software\Rutba Office LanguagesChosen: the first start takes that
// choice and adds them, each with what it needs, and its font where the
// computer has none (L8). With no installer's choice (macOS, Linux, a
// portable copy, a silent install), the first start matches the computer's
// own languages (app.getPreferredSystemLanguages()) and Home offers them,
// once. What cannot be added for want of a network waits and is tried at
// each start; Home says so, and that a language can be added from a file.
// Nothing is added without the person's choice on the installer's page or
// on Home.
//
// After a release that leaves a built-in language out of the core (L7), Home
// says so once to a person who uses it: that it now comes as a pack and was
// added with the update, or, where it was not (an update taken anyway, a
// portable copy), where to add it.
import { matchLanguages, withPacks, notBuiltIn } from '@rutba/office-shell/electron/language-choice';
import { LEFT_CORE, knownLanguages } from '@rutba/office-shell/electron/language-packs';

/** The version and the way the first start chose: set once, never again. */
export const FIRST_KEY = 'languages.firstStart';
/** Packs chosen and not yet added: tried at each start. */
export const PENDING_KEY = 'languages.pending';
/** Packs Home offers, until taken or declined. */
export const OFFER_KEY = 'languages.offer';
/** The languages that left the core whose line Home has shown (L7). */
export const MOVED_KEY = 'languages.leftCoreSaid';

const ID = /^lang-[a-z0-9-]{1,28}$/;
const ids = (v) => (Array.isArray(v) ? [...new Set(v.filter((x) => typeof x === 'string' && ID.test(x) && x !== 'lang-en'))] : []);

/**
 * @param {object} o
 * @param {{ settings: { get, set } }} o.stores
 * @param {object} o.addons          the add-ons service: list, feed, add, installed
 * @param {string} o.office
 * @param {() => Promise<string[]|null>} [o.chosen]   the installer's choice, taken once; null when there is none
 * @param {() => string[]} [o.preferred]              the computer's languages, most wanted first
 * @param {() => string} [o.windowLanguage]       the language the windows are in: the setting, else the system's
 * @param {string[]} [o.leftCore]                 the languages earlier releases carried and this one does not
 * @param {(line: string) => void} [o.log]
 */
export function createFirstLanguages({ stores, addons, office, chosen = async () => null, preferred = () => [], windowLanguage = () => '', leftCore = LEFT_CORE, log = () => {} }) {
  const get = (k, d) => stores.settings.get(k, d);
  const here = () => {
    try {
      const l = addons.list();
      return new Set([...(addons.installed?.() || []).map((p) => p.id), ...(l.staged || []), ...(l.jobs || []).filter((j) => j.state === 'adding' || j.state === 'staged').map((j) => j.id)]);
    } catch {
      return new Set();
    }
  };

  /** The chosen packs not yet here, added when the feed offers them; the rest kept for the next start. */
  async function tryPending() {
    const pending = ids(get(PENDING_KEY, []));
    if (!pending.length) return [];
    const list = await addons.feed({ refresh: false }).catch(() => null);
    if (!list?.feed || list.feed.from !== 'network') return [];
    const offered = new Set((list.offered || []).filter((p) => !p.withdrawn).map((p) => p.id));
    const have = here();
    const added = [];
    const keep = [];
    for (const id of pending) {
      if (have.has(id)) continue;
      // Not offered for this version yet: kept, for a later list.
      if (!offered.has(id)) { keep.push(id); continue; }
      try {
        await addons.add({ id });
        added.push(id);
      } catch (err) {
        log(`${id} not added: ${err?.message || err}`);
        keep.push(id);
      }
    }
    stores.settings.set(PENDING_KEY, keep);
    return added;
  }

  /**
   * At every start: an installer's choice whenever one was written (taken,
   * and cleared by `chosen`, so a later installer's counts too); else, at
   * the first start only, the computer's languages offered; then what waits.
   */
  async function atStart() {
    let fromInstaller = null;
    try { fromInstaller = await chosen(); } catch { fromInstaller = null; }
    if (Array.isArray(fromInstaller)) {
      stores.settings.set(PENDING_KEY, ids([...ids(get(PENDING_KEY, [])), ...ids(fromInstaller)]));
      stores.settings.set(OFFER_KEY, []);
      if (!get(FIRST_KEY, null)) stores.settings.set(FIRST_KEY, { office, from: 'installer' });
    } else if (!get(FIRST_KEY, null)) {
      const have = here();
      let system = [];
      try { system = preferred() || []; } catch { system = []; }
      // Only languages with a pack: the list's, and any the feed already offers.
      const list = await addons.feed({ refresh: false }).catch(() => null);
      // Not one the windows already have built in (L3).
      const known = notBuiltIn(withPacks(knownLanguages((list?.offered || []).filter((p) => !p.withdrawn))));
      stores.settings.set(OFFER_KEY, matchLanguages(system, { list: known }).filter((id) => !have.has(id)));
      stores.settings.set(FIRST_KEY, { office, from: 'system' });
    }
    return tryPending();
  }

  /**
   * What Home shows: the languages offered, those waiting for a network,
   * whether this computer can fetch them now, and which of them the feed
   * offers as drafts (marked so by their status alone).
   * @returns {Promise<{ offer: string[], waiting: string[], online: boolean, drafts: string[] }>}
   */
  async function offer() {
    const have = here();
    const offered = ids(get(OFFER_KEY, [])).filter((id) => !have.has(id));
    const waiting = ids(get(PENDING_KEY, [])).filter((id) => !have.has(id));
    const left = moved(have);
    if (!offered.length && !waiting.length) return { offer: [], waiting: [], online: true, moved: left, drafts: [] };
    const list = await addons.feed({ refresh: false }).catch(() => null);
    const named = new Set([...offered, ...waiting]);
    // A pack's own status (the list's packStatus: its `status` is whether it is added); with no
    // feed, as offline, the languages the list knows still say which are drafts.
    const offeredNow = new Map((list?.offered || []).filter((p) => p?.id).map((p) => [p.id, p]));
    const known = new Map((list?.languages || []).filter((r) => r?.tag).map((r) => [r.pack || `lang-${String(r.tag).toLowerCase().replace(/_/g, '-')}`, r]));
    const draft = (id) => (offeredNow.has(id) ? offeredNow.get(id).packStatus : known.get(id)?.status) === 'draft';
    const drafts = [...named].filter(draft);
    return { offer: offered, waiting, online: list?.feed?.from === 'network', moved: left, drafts };
  }

  /** The window's language, when it has left the core and Home has not said so: { tag, added }, or null. */
  function moved(have = here()) {
    let tag = '';
    try { tag = String(windowLanguage() || ''); } catch { tag = ''; }
    const base = tag.toLowerCase().replace(/_/g, '-').split('-')[0];
    if (!base || !leftCore.includes(base) || (get(MOVED_KEY, []) || []).includes(base)) return null;
    return { tag: base, added: have.has(`lang-${base}`) };
  }

  /** The person took the offer, all of it or the packs named: added now, or when the computer is online. */
  async function takeOffer({ ids: wanted = null } = {}) {
    const offered = ids(get(OFFER_KEY, []));
    const take = Array.isArray(wanted) ? offered.filter((id) => wanted.includes(id)) : offered;
    stores.settings.set(PENDING_KEY, ids([...ids(get(PENDING_KEY, [])), ...take]));
    stores.settings.set(OFFER_KEY, []);
    const added = await tryPending();
    return { added, ...(await offer()) };
  }

  /** Not now: the offer goes, and Voices and languages still has every language. With `moved`, the line on a language that left the core goes instead. */
  function dismissOffer({ moved: said = false } = {}) {
    if (said) {
      const left = moved();
      if (left) stores.settings.set(MOVED_KEY, [...(get(MOVED_KEY, []) || []), left.tag]);
      return offer();
    }
    stores.settings.set(OFFER_KEY, []);
    return { offer: [], waiting: ids(get(PENDING_KEY, [])), online: true, moved: moved(), drafts: [] };
  }

  return { atStart, tryPending, offer, takeOffer, dismissOffer };
}
