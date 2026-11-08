// The macOS menu bar's list of apps: the sets, in the registry's order, each
// with only the apps installed. A set with none installed is left out, so no
// line is drawn for it.

/**
 * @param {{ apps: string[] }[]} groups  the sets of apps
 * @param {string[]|null} installed  the apps on this computer; null for all of them
 * @returns {string[][]}
 */
export function menuAppSets(groups, installed = null) {
  return groups
    .map((g) => g.apps.filter((key) => !installed || installed.includes(key)))
    .filter((set) => set.length);
}
