// "Not now", kept honest.
//
// A person who is told a release is waiting and says "not now" should not
// be asked again a minute later — and should not be left unasked for ever,
// because a copy that is never quit never installs on quit. So a snooze is
// one version for one day: the prompt stays away for that version until
// the day is up, and a NEWER version is never covered by an older snooze.
// Pure, so it can be tested without Electron.

export const SNOOZE_MS = 24 * 60 * 60 * 1000;

/** What to store when a version is snoozed now. */
export function snoozeFor(version, now = Date.now()) {
  if (!version) return null;
  return { version: String(version), until: now + SNOOZE_MS };
}

/** Is `version` still snoozed by `entry` at `now`? */
export function isSnoozed(entry, version, now = Date.now()) {
  if (!entry || !version) return false;
  if (String(entry.version) !== String(version)) return false;
  return Number(entry.until) > now;
}
