# Signing in to Gmail and Outlook.com

Google stopped accepting passwords for IMAP on ordinary accounts in 2022, and
Microsoft did the same in 2024. A mail client without OAuth cannot add the two
mailboxes most people have — so Rutba Office implements it, and this is what has
to be registered before it works in a build you ship.

**A client id is not a secret.** It is published in every open-source mail client
there is; Thunderbird's is in its source tree. What makes the flow safe is PKCE
and the loopback redirect, not the id. Do not treat it as a credential, and do
not put a *client secret* in this application — a downloadable binary cannot keep
one, which is exactly why the flow is designed not to need one.

---

## Google

1. Go to the [Google Cloud Console](https://console.cloud.google.com/) and create
   a project (or use one you have).
2. **APIs & Services → OAuth consent screen.** External. Fill in the application
   name, the support email and the developer email. Add the scope
   `https://mail.google.com/`.
3. **APIs & Services → Credentials → Create credentials → OAuth client ID.**
   Application type: **Desktop app**.
4. Copy the client id. It looks like `1234567890-abcdef.apps.googleusercontent.com`.

Google's Desktop app type allows a loopback redirect on any port, which is what
this uses — there is no redirect URI to register.

While the consent screen is in **Testing**, only the addresses you add as test
users can sign in, and their refresh tokens expire after seven days. Publishing
the app requires Google's review because `https://mail.google.com/` is a
restricted scope; expect that to take a few weeks and to need a demonstration
video and a privacy policy.

## Microsoft

1. Go to the [Azure portal](https://portal.azure.com/) → **Microsoft Entra ID** →
   **App registrations** → **New registration**.
2. Supported account types: **Accounts in any organizational directory and
   personal Microsoft accounts**.
3. **Redirect URI**: platform **Mobile and desktop applications**, and add
   `http://localhost` — Microsoft matches loopback redirects regardless of port,
   but the platform has to be registered as a public client.
4. **API permissions → Add → APIs my organization uses → Office 365 Exchange
   Online → Delegated**: `IMAP.AccessAsUser.All` and `SMTP.Send`. Add
   `offline_access` from Microsoft Graph.
5. **Authentication → Allow public client flows: Yes.**
6. Copy the Application (client) ID.

---

## Putting them in

Either from the running application, once:

```js
// In Mail → Tools, or through the shell from a window's console:
await shell.oauth.setClientId({ provider: 'google', clientId: '…apps.googleusercontent.com' });
await shell.oauth.setClientId({ provider: 'microsoft', clientId: '…' });
```

Or, for a build, in `settings.json` inside the user data directory:

```json
{
  "oauth.google.clientId": "1234567890-abcdef.apps.googleusercontent.com",
  "oauth.microsoft.clientId": "00000000-0000-0000-0000-000000000000"
}
```

With neither set, the account dialog still recognises a Gmail or Outlook address
and says plainly that this build cannot sign in — rather than offering a password
box that would fail at the server.

---

## What the flow does

`apps/desktop/main/oauth.js`, in order:

1. Generates a PKCE verifier and its SHA-256 challenge, and a random `state`.
2. Starts an HTTP listener on `127.0.0.1` on a port the operating system picks.
3. Opens the provider's authorisation page **in the person's own browser** —
   never in a window this application draws. Google refuses embedded webviews,
   and is right to.
4. Waits for one request back. A mismatched `state` is refused; anything without
   a `code` or an `error` gets a 404; nothing is accepted after five minutes.
5. Exchanges the code, with the verifier, for an access token and a refresh
   token.
6. Reads the address out of the id token, so an account signed in as somebody
   else is caught here rather than three folders later.
7. Stores **only the refresh token**, in the operating system keystore.

Access tokens are fetched per connection and never written down. If a provider
rotates the refresh token, the new one replaces the old immediately — keeping the
old one signs the account out at an unpredictable point weeks later.

Signing out deletes the token from this computer. It does not revoke it at the
provider, and the message says so: that is done by removing the application from
the account's connected apps.
