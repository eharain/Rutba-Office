// One-click unsubscribe (RFC 8058), sent from Mail itself.
//
// Before 2.19 the button said one request would be sent and then opened the
// address in the browser, which sends whatever the browser sends (cookies, a
// page load, the scripts on it) and leaves the person to find a button there.
// Now Mail makes the one request the dialog names: a POST of
// "List-Unsubscribe=One-Click" to the sender's https address, with no
// cookies, no sign-in, no other headers of the person's, and no redirect
// followed (RFC 8058: the sender answers the POST itself, and must not
// redirect it).

import https from 'node:https';

/** How long the sender has to answer. */
const TIMEOUT_MS = 15_000;
const BODY = 'List-Unsubscribe=One-Click';

/**
 * The one request. Answers { done: true, status } when the sender took it
 * (a 2xx), else { done: false, reason, status? }: 'refused' (any other
 * answer, a redirect among them), 'unreachable' (no answer), 'not-offered'
 * (not an https address).
 *
 * @param {string} url the https address from List-Unsubscribe
 * @param {{ agent?: import('node:https').Agent }} [options] a test's own agent, trusting its own certificate
 */
export function oneClickUnsubscribe(url, { agent } = {}) {
  let target;
  try {
    target = new URL(url);
  } catch {
    return Promise.resolve({ done: false, reason: 'not-offered' });
  }
  if (target.protocol !== 'https:') return Promise.resolve({ done: false, reason: 'not-offered' });
  return new Promise((resolve) => {
    const req = https.request(
      target,
      {
        method: 'POST',
        agent,
        timeout: TIMEOUT_MS,
        headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(BODY) },
      },
      (res) => {
        res.resume();
        const status = res.statusCode || 0;
        resolve(status >= 200 && status < 300 ? { done: true, status } : { done: false, reason: 'refused', status });
      }
    );
    req.on('timeout', () => req.destroy(new Error('No answer in time.')));
    req.on('error', () => resolve({ done: false, reason: 'unreachable' }));
    req.end(BODY);
  });
}
