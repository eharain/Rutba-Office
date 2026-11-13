// The languages to offer from the computer's own (L5).
//
// At install (the Windows installer's Languages page) and at a first start
// with no installer's choice (macOS, Linux, a portable copy), the computer's
// languages are matched to the packs in the list: the exact tag first, then
// shorter forms of it, then its base, so pt-PT is lang-pt-pt, es-MX is
// lang-es-419, fr-CA is lang-fr-ca, en-US is lang-en-us and de-AT is lang-de.
// English itself is never offered: it is always there. A regional pack brings
// its base with it (requires), so the base is not offered beside it.
//
// The rule is one ordered list of keys, the longest first, each matching a tag
// that is the key or starts with the key and a hyphen. The installer is made
// from the same list (tools/make-installer-languages.mjs), so both choose alike.
import { LANGUAGE_LIST } from './language-packs.js';

/** Spanish of the Americas, written as Latin American Spanish (es-419). */
const AMERICAS = ['ar', 'bo', 'cl', 'co', 'cr', 'cu', 'do', 'ec', 'gt', 'hn', 'mx', 'ni', 'pa', 'pe', 'pr', 'py', 'sv', 'us', 'uy', 've'];
/** Portuguese outside Brazil, written as in Portugal. */
const PORTUGUESE = ['ao', 'cv', 'gw', 'mz', 'st', 'tl', 'ch', 'lu', 'gq', 'mo'];
/** Chinese by region. */
const CHINESE = { hans: ['cn', 'sg', 'my'], hant: ['tw', 'hk', 'mo'] };

const lower = (s) => String(s || '').toLowerCase();

/**
 * The keys, longest first: [key, pack]. A tag matches the first key it is,
 * or starts with followed by a hyphen.
 * @param {{ tag: string, pack: string, variantOf?: string }[]} [list]
 */
export function choiceKeys(list = LANGUAGE_LIST) {
  const rows = (list || []).filter((r) => r?.tag && r?.pack && r.pack !== 'lang-en');
  const byTag = new Map(rows.map((r) => [lower(r.tag), r.pack]));
  const keys = new Map();
  const put = (key, pack) => { if (pack && !keys.has(key)) keys.set(key, pack); };
  for (const r of rows) put(lower(r.tag), r.pack);
  // The regions that write a listed regional form.
  for (const c of AMERICAS) put(`es-${c}`, byTag.get('es-419'));
  for (const c of PORTUGUESE) put(`pt-${c}`, byTag.get('pt-pt'));
  for (const [script, regions] of Object.entries(CHINESE)) for (const c of regions) put(`zh-${c}`, byTag.get(`zh-${script}`));
  // A base, by the first listed row of it that is not a regional form of another.
  for (const r of rows) {
    if (r.variantOf) continue;
    const base = lower(r.tag).split('-')[0];
    if (base !== 'en') put(base, r.pack);
  }
  return [...keys.entries()].sort((a, b) => b[0].length - a[0].length || (a[0] < b[0] ? -1 : 1));
}

/** The pack for one tag, or null. Underscores (en_US, as Linux writes it) count as hyphens. */
export function packForTag(tag, keys = choiceKeys()) {
  const t = lower(tag).replace(/_/g, '-').replace(/\..*$/, '').replace(/@.*$/, '');
  for (const [key, pack] of keys) if (t === key || t.startsWith(`${key}-`)) return pack;
  return null;
}

/**
 * The packs to offer for the computer's languages, in its order: each once,
 * and a regional pack's base left out when the regional pack is there.
 * @param {string[]} tags   the computer's languages, most wanted first
 */
export function matchLanguages(tags, { list = LANGUAGE_LIST } = {}) {
  const keys = choiceKeys(list);
  const out = [];
  for (const tag of tags || []) {
    const pack = packForTag(tag, keys);
    if (pack && !out.includes(pack)) out.push(pack);
  }
  const bases = new Set((list || []).filter((r) => r.variantOf && out.includes(r.pack)).map((r) => r.variantOf));
  return out.filter((p) => !bases.has(p));
}

/**
 * The rows there is a pack for: those with a status (source, draft or
 * reviewed, the pack's own). A row with none names a language whose pack is
 * not yet released, so neither the installer nor a first start offers it; a
 * computer in fr-CA with no lang-fr-ca yet is offered lang-fr.
 */
export const withPacks = (list = LANGUAGE_LIST) => (list || []).filter((r) => typeof r?.status === 'string' && r.status);

/** A regional pack's base, by pack: what the installer leaves out beside it. English, never offered, has nothing to leave out. */
export function regionalBases(list = LANGUAGE_LIST) {
  return (list || []).filter((r) => r.variantOf && r.variantOf !== 'lang-en').map((r) => [r.pack, r.variantOf]);
}
