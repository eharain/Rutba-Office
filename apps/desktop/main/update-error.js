/**
 * What a failed update check comes to, in words a person can act on.
 *
 * electron-updater fails with its whole HTTP exchange in the message: the
 * response headers, the request id and a stack, which the About box showed as
 * they came (the owner's report of 2026-10-12, a 404 for latest.yml while a
 * release's installers were still being built). The kind picks the sentence
 * the window says, in its own language; the detail is the first line alone,
 * cut short, for somebody who wants to report it.
 */

/** @returns {{ kind: 'publishing'|'offline'|'refused'|'other', detail: string }} */
export function updateErrorKind(message) {
  const text = String(message || '');
  const first = text.split(/\r?\n/)[0].replace(/\s+/g, ' ').trim();
  const detail = first.length > 160 ? `${first.slice(0, 157)}...` : first;
  // A release whose tag is out before its files: the newest one is still
  // being published, and the copy already installed is fine as it is.
  if (/latest(-\w+)?\.yml/i.test(text) && /\b404\b|Cannot find/i.test(text)) return { kind: 'publishing', detail };
  if (/ENOTFOUND|EAI_AGAIN|ENETUNREACH|ECONNREFUSED|ECONNRESET|ETIMEDOUT|net::ERR_(INTERNET_DISCONNECTED|NAME_NOT_RESOLVED|NETWORK_CHANGED|CONNECTION|TIMED_OUT|PROXY)/i.test(text)) return { kind: 'offline', detail };
  if (/\b(403|429)\b|rate limit/i.test(text)) return { kind: 'refused', detail };
  return { kind: 'other', detail };
}
