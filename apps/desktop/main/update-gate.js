// What the update service does with a version it found, once the add-ons
// have answered (C11; private design: office-pack-compatibility-design.md 5.2).
// No Electron here, so it can be tried on its own.
//
//   'wait'      some add-on is not ready for it, and "Update anyway" was not chosen for it
//   'notice'    a portable copy: told, and nothing downloaded
//   'download'  stage the ready add-ons, then download

/**
 * @param {object} o
 * @param {{ offer: boolean }} o.answer   the add-ons' readiness for the version
 * @param {string} o.version
 * @param {string|null} o.anyway          the version "Update anyway" was chosen for
 * @param {boolean} o.portable
 * @returns {'wait'|'notice'|'download'}
 */
/**
 * The languages an update's own feed (latest.yml, electron-updater's
 * UpdateInfo) says it leaves out of the core (build/write-left-core.js):
 * language tags only, or null when it says nothing (a release before 2.26.0).
 */
export function releaseLeftCore(next) {
  if (!Array.isArray(next?.leftCore)) return null;
  return [...new Set(next.leftCore.filter((t) => typeof t === 'string' && /^[a-z]{2,3}$/.test(t)))].sort();
}

export function gateStep({ answer, version, anyway = null, portable = false }) {
  if (!answer?.offer && anyway !== version) return 'wait';
  return portable ? 'notice' : 'download';
}
