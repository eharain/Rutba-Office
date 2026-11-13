// The table a window registers for its language when an added pack words it
// (L3). For a language Office carries built in (ur, ar, hi), the pack is laid
// over Office's own catalogue rather than put in its place: a label the pack
// does not word (one new in this Office, after "Update anyway" or a pack
// rolled back to an older revision, or every label when all its entries are
// stale) keeps Office's own translation instead of falling back to English.
// For any other language the pack's table is the whole of it.

/**
 * @param {string} tag            the pack's language tag
 * @param {object|null} table     the pack's table, keyed by the English
 * @param {Record<string, object>} builtIn  Office's own catalogues, by tag
 * @returns {object|null}
 */
export function packTable(tag, table, builtIn = {}) {
  if (!table || typeof table !== 'object') return null;
  const own = builtIn?.[String(tag || '').toLowerCase()];
  return own && typeof own === 'object' ? { ...own, ...table } : table;
}
