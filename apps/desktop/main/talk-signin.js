// Talk's Rutba sign-in: Office as a native app at auth.rutba.io (RFC 8252),
// as workers/introducer/docs/SIGN-IN.md sets out and its scripts/sign-in.mjs
// does. The system browser opens at auth.rutba.io, where the person signs in
// as they do everywhere else; the code comes back to a listener Office holds
// on 127.0.0.1 at a port the system picks, is checked (its state, and the
// issuer that sent it) and exchanged with the PKCE verifier for an access
// token for the introducer and a refresh token. Nothing is shown inside
// Office: the person sees the real address bar and their password manager.

import http from 'node:http';
import crypto from 'node:crypto';

export const ISSUER = 'https://auth.rutba.io';
export const CLIENT_ID = 'rutba-office';
export const RESOURCE = 'https://introducer.rutba.io';
const WAIT = 5 * 60 * 1000;

/** The provider's endpoints, read from its discovery document, its issuer checked. */
export async function discover(issuer = ISSUER, { fetchImpl = fetch } = {}) {
  const at = issuer.replace(/\/$/, '');
  const res = await fetchImpl(`${at}/.well-known/openid-configuration`);
  if (!res.ok) throw new Error(`Rutba sign-in could not be reached (${res.status}). Check the connection and try again.`);
  const meta = await res.json();
  if (meta.issuer !== at) throw new Error('Rutba sign-in answered as somebody else, so Talk did not go on.');
  return meta;
}

/** A form POST to the token endpoint: the answer, or an error carrying the provider's code. */
async function tokenRequest(meta, form, fetchImpl) {
  const res = await fetchImpl(meta.token_endpoint, { method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams(form) });
  let body = {};
  try { body = await res.json(); } catch { body = {}; }
  if (!res.ok) {
    const err = new Error(body.error_description || body.error || `Rutba sign-in answered ${res.status}.`);
    err.code = body.error || null;
    err.status = res.status;
    throw err;
  }
  return body;
}

/** When an access token runs out, in ms since 1970, from its own `exp`; null when it cannot be read. */
export function expiresAt(token) {
  try {
    const claims = JSON.parse(Buffer.from(String(token).split('.')[1], 'base64url').toString('utf8'));
    return Number.isFinite(claims.exp) ? claims.exp * 1000 : null;
  } catch {
    return null;
  }
}

/**
 * Sign in in the browser. `open(url)` opens the system browser. Answers
 * { access_token, refresh_token, id_token, expires_in }, or throws why not
 * (cancelled at auth.rutba.io, an answer that was not ours, five minutes
 * with no answer).
 */
export async function signInWithBrowser({ issuer = ISSUER, open, fetchImpl = fetch, wait = WAIT } = {}) {
  const meta = await discover(issuer, { fetchImpl });
  const verifier = crypto.randomBytes(32).toString('base64url');
  const challenge = crypto.createHash('sha256').update(verifier).digest('base64url');
  const state = crypto.randomBytes(16).toString('base64url');
  const page = (res, text) => {
    res.writeHead(200, { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store' });
    res.end(`<!doctype html><meta charset="utf-8"><title>Rutba Office</title><p style="font:16px system-ui;margin:3em">${text}</p>`);
  };
  const answer = await new Promise((resolve, reject) => {
    let timer = null;
    let port = 0;
    const server = http.createServer((req, res) => {
      const url = new URL(req.url, 'http://127.0.0.1');
      if (url.pathname !== '/callback') { res.writeHead(404); res.end(); return; }
      const p = url.searchParams;
      // An answer to a sign-in started somewhere else is turned away, and this one waits on.
      if (p.get('state') !== state) { page(res, 'That sign-in was not started here. Close this window and try again from Rutba Office.'); return; }
      clearTimeout(timer);
      server.close();
      if (p.get('error')) {
        page(res, 'Sign-in did not finish. You can close this window.');
        reject(new Error(p.get('error') === 'access_denied' ? 'Sign-in was cancelled.' : `Sign-in did not finish (${p.get('error')}).`));
        return;
      }
      if (meta.authorization_response_iss_parameter_supported && p.get('iss') !== meta.issuer) {
        page(res, 'That answer came from somewhere else. You can close this window.');
        reject(new Error('The sign-in answer did not come from Rutba sign-in, so Talk did not use it.'));
        return;
      }
      page(res, 'Signed in. You can close this window and go back to Rutba Office.');
      resolve({ code: p.get('code'), redirectUri: `http://127.0.0.1:${port}/callback` });
    });
    server.on('error', (err) => reject(err));
    // Port 0: the system picks a free one, and the provider takes any port on a loopback redirect.
    server.listen(0, '127.0.0.1', () => {
      port = server.address().port;
      const redirectUri = `http://127.0.0.1:${port}/callback`;
      const authorize = `${meta.authorization_endpoint}?${new URLSearchParams({
        client_id: CLIENT_ID,
        response_type: 'code',
        scope: 'openid profile email',
        redirect_uri: redirectUri,
        code_challenge: challenge,
        code_challenge_method: 'S256',
        state,
        resource: RESOURCE,
      })}`;
      Promise.resolve(open(authorize)).catch((err) => { server.close(); reject(err); });
    });
    timer = setTimeout(() => { server.close(); reject(new Error('No sign-in came back within five minutes. Try again.')); }, wait);
    timer.unref?.();
  });
  return tokenRequest(meta, { grant_type: 'authorization_code', code: answer.code, redirect_uri: answer.redirectUri, client_id: CLIENT_ID, code_verifier: verifier }, fetchImpl);
}

/** A fresh access token (and the new refresh token) for the refresh token held. An error's `code` is 'invalid_grant' when the sign-in has ended. */
export async function refreshTokens({ issuer = ISSUER, refreshToken, fetchImpl = fetch }) {
  const meta = await discover(issuer, { fetchImpl });
  return tokenRequest(meta, { grant_type: 'refresh_token', refresh_token: refreshToken, client_id: CLIENT_ID }, fetchImpl);
}

/** The refresh token given back, as Sign out asks: the provider answers 200 whatever happened. */
export async function revoke({ issuer = ISSUER, refreshToken, fetchImpl = fetch }) {
  const meta = await discover(issuer, { fetchImpl });
  if (!meta.revocation_endpoint) return;
  await fetchImpl(meta.revocation_endpoint, { method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams({ token: refreshToken, token_type_hint: 'refresh_token', client_id: CLIENT_ID }) });
}
