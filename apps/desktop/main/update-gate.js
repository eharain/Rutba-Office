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
export function gateStep({ answer, version, anyway = null, portable = false }) {
  if (!answer?.offer && anyway !== version) return 'wait';
  return portable ? 'notice' : 'download';
}
