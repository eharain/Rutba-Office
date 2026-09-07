// Signing in to Gmail and Outlook.com.
//
// Both providers stopped accepting a password for IMAP on ordinary accounts —
// Google in 2022, Microsoft in 2024 — so a mail client without this cannot add
// the two mailboxes most people have. It is the single largest thing standing
// between this suite and the person downloading it.
//
// The flow is the one RFC 8252 prescribes for a native application, and every
// part of that choice matters:
//
//   The system browser, not an embedded window. Google refuses to authenticate
//   inside an embedded webview, and it is right to: a window the application
//   draws is a window the application can read the password out of. Sending
//   people to their own browser means their password is typed into Google's
//   page, in a browser with their password manager and their security key, and
//   this process never sees it.
//
//   A loopback redirect on an ephemeral port. The browser comes back to
//   127.0.0.1, which only this machine can reach, on a port chosen at the
//   moment of asking. The listener answers exactly one request and closes.
//
//   PKCE, always. A public client cannot keep a secret — anything compiled into
//   a downloadable binary is published — so the authorisation code is bound to
//   a verifier this process generated and never sent, which is what makes the
//   code useless to anyone who intercepts it.
//
// What is stored afterwards is a refresh token, in the operating system's
// keystore alongside the passwords. Access tokens live an hour and are never
// written down.

import crypto from 'node:crypto';
import http from 'node:http';

// Loaded when it is needed rather than at import time: everything above the
// sign-in itself — the provider table, the address matching, the PKCE pair —
// is plain Node, and being able to test it without an Electron process is
// worth one dynamic import.
const openInBrowser = async (url) => (await import('electron')).shell.openExternal(url);

/**
 * The two providers worth building in.
 *
 * `clientId` is deliberately empty. A registered client id is not a secret —
 * Thunderbird publishes its own — but it belongs to whoever ships the build,
 * and inventing one here would produce an application that fails at the last
 * step with a message from Google rather than one from us. See
 * docs/OAUTH.md for how to register one and where to put it.
 */
export const PROVIDERS = {
  google: {
    id: 'google',
    label: 'Google',
    matches: /@(gmail\.com|googlemail\.com)$/i,
    authorize: 'https://accounts.google.com/o/oauth2/v2/auth',
    token: 'https://oauth2.googleapis.com/token',
    scope: 'https://mail.google.com/',
    imap: { host: 'imap.gmail.com', port: 993, secure: true },
    smtp: { host: 'smtp.gmail.com', port: 465, secure: true },
    // Google will not return a refresh token a second time unless asked.
    extra: { access_type: 'offline', prompt: 'consent' },
  },
  microsoft: {
    id: 'microsoft',
    label: 'Microsoft',
    matches: /@(outlook\.com|hotmail\.com|live\.com|msn\.com)$/i,
    authorize: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
    token: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
    scope: 'offline_access https://outlook.office.com/IMAP.AccessAsUser.All https://outlook.office.com/SMTP.Send',
    imap: { host: 'outlook.office365.com', port: 993, secure: true },
    smtp: { host: 'smtp-mail.outlook.com', port: 587, secure: false },
    extra: {},
  },
};

/** Which provider an address belongs to, or none. */
export function providerFor(email) {
  const address = String(email || '');
  for (const p of Object.values(PROVIDERS)) if (p.matches.test(address)) return p;
  return null;
}

/* ── PKCE ────────────────────────────────────────────────────────────────── */

const base64url = (buf) => buf.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

function pkce() {
  const verifier = base64url(crypto.randomBytes(64));
  const challenge = base64url(crypto.createHash('sha256').update(verifier).digest());
  return { verifier, challenge };
}

/* ── the service ─────────────────────────────────────────────────────────── */

export function createOAuthService({ stores, broadcast }) {
  const clientIdFor = (provider) => stores.settings.get(`oauth.${provider}.clientId`, '') || '';
  const key = (email, provider) => `oauth:${provider}:${String(email).toLowerCase()}`;

  /**
   * Trade a code or a refresh token for an access token.
   * The exchange is a plain form post, and PKCE means there is no secret in it.
   */
  async function exchange(provider, body) {
    const response = await fetch(provider.token, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams(body),
    });
    const text = await response.text();
    let payload = {};
    try {
      payload = JSON.parse(text);
    } catch {
      throw new Error(`${provider.label} answered with something that is not a token: ${text.slice(0, 120)}`);
    }
    if (!response.ok) {
      throw new Error(payload.error_description || payload.error || `${provider.label} refused the sign-in.`);
    }
    return payload;
  }

  return {
    /** Which provider an address belongs to, and whether we can sign in to it. */
    provider: ({ email }) => {
      const provider = providerFor(email);
      if (!provider) return null;
      return {
        id: provider.id,
        label: provider.label,
        configured: Boolean(clientIdFor(provider.id)),
        imap: provider.imap,
        smtp: provider.smtp,
      };
    },

    /** Where a client id goes, and what it is for. Never a secret. */
    setClientId: ({ provider, clientId }) => {
      if (!PROVIDERS[provider]) throw new Error(`unknown provider: ${provider}`);
      stores.settings.set(`oauth.${provider}.clientId`, String(clientId || '').trim());
      return { provider, configured: Boolean(String(clientId || '').trim()) };
    },

    clientIds: () =>
      Object.fromEntries(Object.keys(PROVIDERS).map((id) => [id, clientIdFor(id) ? 'set' : ''])),

    /**
     * Sign in, in the person's own browser.
     * @returns {{ email, provider, imap, smtp }} enough to add the account
     */
    signIn: async ({ email, provider: providerId }) => {
      const provider = PROVIDERS[providerId] || providerFor(email);
      if (!provider) throw new Error('That address is not one this can sign in to.');

      const clientId = clientIdFor(provider.id);
      if (!clientId) {
        throw new Error(
          `No ${provider.label} client id is set up in this build. See docs/OAUTH.md — it takes about five minutes and is free.`
        );
      }

      const { verifier, challenge } = pkce();
      const state = base64url(crypto.randomBytes(24));

      // The listener first, because the redirect URI has to name its port.
      const { port, wait } = await listen({ state });
      const redirectUri = `http://127.0.0.1:${port}/`;

      const url = new URL(provider.authorize);
      url.searchParams.set('client_id', clientId);
      url.searchParams.set('redirect_uri', redirectUri);
      url.searchParams.set('response_type', 'code');
      url.searchParams.set('scope', provider.scope);
      url.searchParams.set('state', state);
      url.searchParams.set('code_challenge', challenge);
      url.searchParams.set('code_challenge_method', 'S256');
      if (email) url.searchParams.set('login_hint', email);
      for (const [k, v] of Object.entries(provider.extra)) url.searchParams.set(k, v);

      broadcast?.('mail:oauth', { phase: 'browser', provider: provider.id });
      await openInBrowser(url.href);

      const code = await wait;
      const tokens = await exchange(provider, {
        client_id: clientId,
        code,
        code_verifier: verifier,
        grant_type: 'authorization_code',
        redirect_uri: redirectUri,
      });

      if (!tokens.refresh_token) {
        throw new Error(
          `${provider.label} did not return a refresh token, which means this would stop working within the hour. Remove Rutba Office from your account's connected apps and try again.`
        );
      }

      // Whose mailbox this actually is, from the token rather than from what was
      // typed: signing in as somebody else is a mistake worth catching here.
      const identity = readIdToken(tokens.id_token) || {};
      const address = identity.email || email;

      stores.secrets.set(key(address, provider.id), tokens.refresh_token);
      broadcast?.('mail:oauth', { phase: 'done', provider: provider.id, email: address });

      return {
        email: address,
        name: identity.name || null,
        provider: provider.id,
        imap: { ...provider.imap, user: address, auth: 'oauth' },
        smtp: { ...provider.smtp, user: address, auth: 'oauth' },
      };
    },

    /**
     * A live access token for a stored account.
     *
     * Called by the mail service on every connection rather than cached here: an
     * access token lasts an hour, connections are rarer than that, and a token
     * held in a variable is a token that outlives the account being removed.
     */
    accessToken: async ({ email, provider: providerId }) => {
      const provider = PROVIDERS[providerId] || providerFor(email);
      if (!provider) throw new Error('That account does not use a signed-in provider.');
      const refresh = stores.secrets.get(key(email, provider.id));
      if (!refresh) throw new Error(`${email} is not signed in. Add the account again to sign in.`);

      const tokens = await exchange(provider, {
        client_id: clientIdFor(provider.id),
        refresh_token: refresh,
        grant_type: 'refresh_token',
      });

      // Providers may rotate the refresh token; keeping the old one would sign
      // the account out at an unpredictable point in the future.
      if (tokens.refresh_token && tokens.refresh_token !== refresh) {
        stores.secrets.set(key(email, provider.id), tokens.refresh_token);
      }
      return { accessToken: tokens.access_token, expiresIn: tokens.expires_in ?? 3600 };
    },

    /** Forget the account's token. The provider still lists the app; say so. */
    signOut: ({ email, provider: providerId }) => {
      const provider = PROVIDERS[providerId] || providerFor(email);
      if (provider) stores.secrets.delete(key(email, provider.id));
      return {
        removed: true,
        note: provider
          ? `The token is gone from this computer. To revoke it at ${provider.label} as well, remove Rutba Office from your account's connected apps.`
          : 'Removed.',
      };
    },
  };
}

/**
 * Start the loopback listener and hand back its port with a promise for the
 * code. Split from the redirect handling so the caller can build a redirect URI
 * naming the port before the browser is opened — there is no way to know the
 * port before listening, and no way to authorise before naming it.
 */
function listen({ state }) {
  return new Promise((resolveOuter, rejectOuter) => {
    let settle;
    const wait = new Promise((resolve, reject) => {
      settle = { resolve, reject };
    });

    const server = http.createServer((req, res) => {
      const url = new URL(req.url, 'http://127.0.0.1');
      const page = (title, body, ok) => {
        res.writeHead(ok ? 200 : 400, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(
          `<!doctype html><meta charset="utf-8"><title>${title}</title>` +
            '<style>body{font:15px/1.6 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;' +
            'display:grid;place-items:center;min-height:90vh;margin:0;color:#1a1c20;text-align:center;' +
            'background:#fbfbfd}div{max-width:30rem;padding:2rem}h1{font-size:1.3rem;margin:0 0 .6rem}' +
            'p{margin:0;color:#5b626d}</style>' +
            `<div><h1>${title}</h1><p>${body}</p></div>`
        );
      };

      if (!url.searchParams.has('code') && !url.searchParams.has('error')) {
        res.writeHead(404).end();
        return;
      }

      const done = (error, code) => {
        clearTimeout(timer);
        setTimeout(() => server.close(), 300);
        if (error) settle.reject(error);
        else settle.resolve(code);
      };

      if (url.searchParams.get('state') !== state) {
        page('That did not come from us', 'The sign-in could not be verified. Close this tab and try again.', false);
        done(new Error('The sign-in came back with the wrong state, so it was refused.'));
        return;
      }
      const error = url.searchParams.get('error');
      if (error) {
        page('Sign-in cancelled', 'Nothing has been changed. You can close this tab.', false);
        done(new Error(url.searchParams.get('error_description') || error));
        return;
      }
      page('Signed in', 'You can close this tab and go back to Rutba Office.', true);
      done(null, url.searchParams.get('code'));
    });

    const timer = setTimeout(() => {
      server.close();
      settle.reject(new Error('The sign-in was not completed in five minutes. Nothing has been changed.'));
    }, 5 * 60_000);

    server.on('error', rejectOuter);
    server.listen(0, '127.0.0.1', () => resolveOuter({ port: server.address().port, wait }));
  });
}

/**
 * The email address out of an id token.
 *
 * Read, not verified: this token came back over TLS from the provider's own
 * token endpoint in direct response to our request, so there is no third party
 * between us to lie about it. It is used only to label the account, never to
 * decide anything.
 */
function readIdToken(jwt) {
  try {
    const [, payload] = String(jwt || '').split('.');
    if (!payload) return null;
    return JSON.parse(Buffer.from(payload.replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf8'));
  } catch {
    return null;
  }
}
