// Talk's Rutba sign-in in the browser (workers/introducer docs/SIGN-IN.md),
// checked against a stand-in auth.rutba.io and the stand-in introducer:
// Sign in with Rutba opens the browser at the provider with PKCE and a state;
// the check plays the browser, and the code comes back to Office on
// 127.0.0.1; a token near its end is renewed before Talk connects, and one
// the introducer turns away as run out is renewed and the connection made
// again; a sign-in ended at auth.rutba.io signs Talk out; and Sign out gives
// the refresh token back. Run alone with RUTBA_VERIFY_ONLY=talk-signin.

import http from 'node:http';
import crypto from 'node:crypto';
import { standInIntroducer } from './verify-talk.js';

/** A stand-in auth.rutba.io: discovery, authorize, token and revocation, as the real one answers them. */
function standInAuth() {
  const codes = new Map(); // code -> { challenge, redirect }
  const live = new Set(); // refresh tokens that still work
  const log = { codes: 0, refreshes: 0, revoked: [], expiredFirst: true };
  let base = '';
  const b64 = (o) => Buffer.from(JSON.stringify(o)).toString('base64url');
  const access = (life) => `${b64({ alg: 'none' })}.${b64({ iss: base, aud: 'introducer.rutba.io', sub: 'amira', exp: Math.floor((Date.now() + life) / 1000) })}.x`;
  const tokens = () => {
    const refresh = `r${crypto.randomBytes(9).toString('hex')}`;
    live.add(refresh);
    // The first access token handed out has already run out: Office must renew it before it connects.
    const life = log.expiredFirst ? -60_000 : 3_600_000;
    log.expiredFirst = false;
    return { access_token: access(life), refresh_token: refresh, id_token: access(3_600_000), token_type: 'Bearer', expires_in: Math.round(life / 1000) };
  };
  const server = http.createServer(async (req, res) => {
    const url = new URL(req.url, base);
    const json = (status, body) => { res.writeHead(status, { 'content-type': 'application/json' }); res.end(JSON.stringify(body)); };
    if (url.pathname === '/.well-known/openid-configuration') {
      return json(200, { issuer: base, authorization_endpoint: `${base}/oidc/auth`, token_endpoint: `${base}/oidc/token`, revocation_endpoint: `${base}/oidc/token/revocation`, authorization_response_iss_parameter_supported: true });
    }
    if (url.pathname === '/oidc/auth') {
      const p = url.searchParams;
      const ok = p.get('client_id') === 'rutba-office' && p.get('code_challenge_method') === 'S256' && /^http:\/\/127\.0\.0\.1:\d+\/callback$/.test(p.get('redirect_uri') || '') && p.get('scope') === 'openid profile email';
      if (!ok) return json(400, { error: 'invalid_request' });
      const code = crypto.randomBytes(12).toString('hex');
      codes.set(code, { challenge: p.get('code_challenge'), redirect: p.get('redirect_uri') });
      res.writeHead(302, { location: `${p.get('redirect_uri')}?${new URLSearchParams({ code, state: p.get('state'), iss: base })}` });
      return res.end();
    }
    let body = '';
    for await (const chunk of req) body += chunk;
    const form = new URLSearchParams(body);
    if (url.pathname === '/oidc/token') {
      if (form.get('client_id') !== 'rutba-office') return json(401, { error: 'invalid_client' });
      if (form.get('grant_type') === 'authorization_code') {
        const held = codes.get(form.get('code'));
        codes.delete(form.get('code'));
        const challenge = crypto.createHash('sha256').update(form.get('code_verifier') || '').digest('base64url');
        if (!held || held.challenge !== challenge || held.redirect !== form.get('redirect_uri')) return json(400, { error: 'invalid_grant' });
        log.codes++;
        return json(200, tokens());
      }
      if (form.get('grant_type') === 'refresh_token') {
        if (!live.delete(form.get('refresh_token'))) return json(400, { error: 'invalid_grant', error_description: 'grant request is invalid' });
        log.refreshes++;
        return json(200, tokens());
      }
      return json(400, { error: 'unsupported_grant_type' });
    }
    if (url.pathname === '/oidc/token/revocation') {
      log.revoked.push(form.get('token'));
      live.delete(form.get('token'));
      res.writeHead(200);
      return res.end();
    }
    res.writeHead(404);
    return res.end();
  });
  return {
    log,
    live,
    listen: () => new Promise((resolve) => server.listen(0, '127.0.0.1', () => { base = `http://127.0.0.1:${server.address().port}`; resolve(base); })),
    close: () => new Promise((resolve) => { server.closeAllConnections?.(); server.close(() => resolve()); }),
  };
}

export async function verifyTalkSignIn(h) {
  const { open, check, until, wait } = h;
  const ok = (p) => p.then(() => true, () => false);
  const auth = standInAuth();
  const intro = standInIntroducer({});
  let win = null;
  try {
    const issuer = await auth.listen();
    const port = await intro.listen();
    const server = `ws://127.0.0.1:${port}/ws`;
    win = await open('talk');
    const js = (code) => win.webContents.executeJavaScript(code);
    await js(`window.rutbaOffice.talk.signOut().then(() => window.rutbaOffice.store.set({ key: 'talk.issuer', value: ${JSON.stringify(issuer)} })).then(() => 1)`);
    await until(() => js(`Boolean(document.querySelector('.tk-rutba'))`), 'the Sign in with Rutba button', 8000).catch(() => {});

    // Sign in with Rutba: the browser asked for at the provider, the check playing it.
    globalThis.__rutbaCheckOpened = [];
    const signing = js(`window.rutbaOffice.talk.signInBrowser({ server: ${JSON.stringify(server)} }).then(() => 'signed in', (e) => e.message)`);
    await until(() => (globalThis.__rutbaCheckOpened || []).some((u) => u.startsWith(issuer)) || null, 'the browser opened', 8000).catch(() => {});
    const asked = new URL((globalThis.__rutbaCheckOpened || []).find((u) => u.startsWith(issuer)) || `${issuer}/none`);
    const waiting = await js(`document.querySelector('.tk-rutba')?.textContent || ''`);
    const p = asked.searchParams;
    const shaped = asked.pathname === '/oidc/auth' && p.get('client_id') === 'rutba-office' && p.get('code_challenge_method') === 'S256' && (p.get('code_challenge') || '').length >= 43
      && (p.get('state') || '').length >= 16 && /^http:\/\/127\.0\.0\.1:\d+\/callback$/.test(p.get('redirect_uri') || '') && p.get('resource') === 'https://introducer.rutba.io';
    // A stranger's answer to Office's listener is turned away; the provider's own comes through.
    const forged = await fetch(`${p.get('redirect_uri')}?code=x&state=not-ours&iss=${encodeURIComponent(issuer)}`).then((r) => r.text(), () => '');
    const hop = await fetch(asked, { redirect: 'manual' });
    const back = await fetch(hop.headers.get('location')).then((r) => r.text(), (e) => e.message);
    const result = await signing;
    const online = await ok(until(() => js(`/Signed in as Amira Khan/.test(document.querySelector('.tk-statusbar')?.textContent || '')`), 'signed in', 10000));
    check('talk: Sign in with Rutba opens the browser at auth.rutba.io with PKCE and a state, waits for it, turns away an answer it did not ask for, and signs in with what comes back',
      shaped && /Waiting for the browser/.test(waiting) && /not started here/.test(forged) && /Signed in/.test(back) && result === 'signed in' && online && auth.log.codes === 1,
      JSON.stringify({ shaped, waiting, forged: forged.slice(-80), back: back.slice(-70), result, online, codes: auth.log.codes }));
    check('talk: an access token already at its end is renewed with the refresh token before Talk connects', auth.log.refreshes === 1 && online, JSON.stringify(auth.log));

    // The introducer turns the next connection away as run out: renewed, and connected again at once.
    intro.options.expireNext = true;
    intro.drop('u-amira');
    const renewed = await ok(until(() => auth.log.refreshes >= 2 && intro.online('u-amira'), 'renewed and connected again', 15000));
    check('talk: a sign-in the introducer turns away as run out is renewed and Talk connects again by itself', renewed, JSON.stringify({ refreshes: auth.log.refreshes, online: intro.online('u-amira') }));

    // Signed out at auth.rutba.io: the refresh no longer works, and Talk says it has signed out.
    auth.live.clear();
    intro.options.expireNext = true;
    intro.drop('u-amira');
    const ended = await ok(until(() => js(`Boolean(document.querySelector('.tk-rutba')) && /sign-in has ended/.test(document.querySelector('.tk-error')?.textContent || '')`), 'signed out, saying why', 15000));
    check('talk: a sign-in ended at auth.rutba.io signs Talk out and says so, with Sign in with Rutba to sign in again', ended, await js(`document.querySelector('.tk-error')?.textContent || ''`));

    // Signed in again, then Sign out: the refresh token given back.
    globalThis.__rutbaCheckOpened = [];
    const again = js(`window.rutbaOffice.talk.signInBrowser({ server: ${JSON.stringify(server)} }).then(() => 'ok', (e) => e.message)`);
    await until(() => (globalThis.__rutbaCheckOpened || []).length || null, 'the browser opened again', 8000).catch(() => {});
    const hop2 = await fetch(globalThis.__rutbaCheckOpened[0], { redirect: 'manual' });
    await fetch(hop2.headers.get('location')).catch(() => {});
    await again;
    await until(() => intro.online('u-amira') || null, 'connected again', 8000).catch(() => {});
    const held = [...auth.live];
    await js(`window.rutbaOffice.talk.signOut().then(() => 1)`);
    const given = await ok(until(() => held.length === 1 && auth.log.revoked.includes(held[0]) || null, 'the refresh token given back', 5000));
    check('talk: Sign out gives the refresh token back to auth.rutba.io and leaves Talk signed out', given && !auth.live.size, JSON.stringify({ held: held.length, revoked: auth.log.revoked.length }));
    await wait(100);
  } catch (err) {
    check('talk: the Rutba sign-in checks ran', false, err.stack || err.message);
  } finally {
    try { await win?.webContents.executeJavaScript(`window.rutbaOffice.store.set({ key: 'talk.issuer', value: null }).then(() => 1)`); } catch { /* gone */ }
    await intro.close().catch(() => {});
    await auth.close().catch(() => {});
  }
}
