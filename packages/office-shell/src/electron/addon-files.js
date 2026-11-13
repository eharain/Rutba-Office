// What a window may load from an add-on: a font pack's fonts, at
// rutba://app/addons/<id>/<file>, and nothing else. An add-on is not an app
// pack and carries no window code, so its folder is never a pack root (whose
// chunks and files a window imports as scripts), and only a font file is
// answered from it: a script, a page or anything else in it is not served.

import path from 'node:path';

/** The files a window may load from an add-on. */
export const ADDON_FILE_TYPES = new Set(['.woff2', '.woff', '.ttf', '.otf', '.ttc']);

/** The address a window loads an add-on's file from. */
export const addonFileUrl = (id, rel) => `rutba://app/addons/${encodeURIComponent(id)}/${String(rel).split('/').map(encodeURIComponent).join('/')}`;

/** Whether a font's address is an add-on's, as addonFileUrl() makes it. */
export const isAddonFileUrl = (url) => typeof url === 'string' && url.startsWith('rutba://app/addons/');

/**
 * The file `rel` (the path after rutba://app/, decoded) names in an add-on,
 * or null when it is not one a window may have: an add-on not served, a
 * type not listed, or a path stepping out of its folder.
 * @param {string} rel
 * @param {Record<string, string>} roots each served add-on's folder, by its id
 */
export function addonFile(rel, roots) {
  const [top, id, ...rest] = String(rel).split('/');
  if (top !== 'addons' || !id || !rest.length || !Object.hasOwn(roots || {}, id)) return null;
  if (!ADDON_FILE_TYPES.has(path.extname(rest[rest.length - 1]).toLowerCase())) return null;
  if (rest.some((s) => s === '' || s === '.' || s === '..' || s.includes('\\') || s.includes(':'))) return null;
  const base = path.resolve(roots[id]);
  const full = path.resolve(base, ...rest);
  return full.startsWith(base + path.sep) ? full : null;
}
