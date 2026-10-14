// Setting up a mail account from an address and a password, trying hard.
//
// Discovery (mail-discover.js) says where the servers probably are. That is
// a guess until the password is tried, and the first guess is not always
// right: a company's mail answers on mail.example.com and on imap.example.com
// but only one takes the login; a server wants the whole address as the user
// name, another only the part before the @. So the setup tries the found
// incoming servers in the order discovery ranked them, each with the address
// and then the name before the @, until one signs in - and the outgoing
// servers the same way - and says, when none does, which kind of failure it
// was: a password refused, a provider that wants an app password or its own
// browser sign-in, or a server that never answered. What was tried is kept,
// without the password, for the window to show.
//
// Nothing is looked up for text that is not an address, nor before there is
// a password (the provider table alone says whether it signs in with a
// browser page instead). A server discovery marked foreign - not under the
// address's domain, not in the table - is not sent the password until the
// person has agreed to it by name: the answer is then 'confirm', with the
// servers to show. And no server is tried without encryption: TLS, or a
// STARTTLS upgrade that must happen before the password is sent.
//
// The tries and the discovery are handed in, so the tests run it with no
// network and a check run never reaches one.

import { isAddress } from '@rutba/mailbox/account-rows';

/** What a failed try's message means. */
export function failureOf(message) {
  const m = String(message || '');
  if (/web ?login required|application-specific password|app password|\[ALERT\].*(web|browser)|534-5\.7\.9|please log in via your web browser/i.test(m)) return 'app-password';
  if (/auth(entication)?[ -]?(failed|failure|error)|invalid credentials|login failed|incorrect|bad (user|password)|\bNO \[AUTH|535|AUTHENTICATIONFAILED|username and password not accepted|password/i.test(m)) return 'password';
  if (/ENOTFOUND|EAI_AGAIN|getaddrinfo/i.test(m)) return 'no-server';
  if (/ECONNREFUSED|ETIMEDOUT|timed? ?out|ECONNRESET|EHOSTUNREACH|socket|closed/i.test(m)) return 'unreachable';
  if (/certificate|self[- ]signed|CERT_|TLS/i.test(m)) return 'certificate';
  return 'other';
}

/**
 * The server software a greeting names - "* OK [CAPABILITY ...] Dovecot ready."
 * - for the window to say what was found.
 */
export function serverSoftware(banner) {
  const b = String(banner || '');
  const known = [
    [/gimap|gsmtp|google/i, 'Gmail'],
    [/microsoft exchange|outlook\.com|office365|exchange online|\.outlook\.|protection\.outlook/i, 'Microsoft Exchange'],
    [/dovecot/i, 'Dovecot'],
    [/courier/i, 'Courier'],
    [/cyrus/i, 'Cyrus'],
    [/zimbra/i, 'Zimbra'],
    [/yahoo/i, 'Yahoo'],
    [/icloud|me\.com|apple/i, 'iCloud'],
    [/zoho/i, 'Zoho'],
    [/fastmail|messagingengine/i, 'Fastmail'],
    [/hmailserver/i, 'hMailServer'],
    [/mdaemon/i, 'MDaemon'],
    [/kerio/i, 'Kerio Connect'],
    [/icewarp|merak/i, 'IceWarp'],
    [/smartermail/i, 'SmarterMail'],
    [/axigen/i, 'Axigen'],
    [/postfix/i, 'Postfix'],
    [/exim/i, 'Exim'],
    [/qmail/i, 'qmail'],
    [/sendmail/i, 'Sendmail'],
  ];
  for (const [re, name] of known) if (re.test(b)) return name;
  return null;
}

/** The user names to try with a server: one it was given, the address, the part before the @. */
function userNames(email, given) {
  const local = email.split('@')[0];
  return [...new Set([given, email, local].filter(Boolean))];
}

/**
 * The candidates of one kind, best first, the one discovery chose first of
 * all; a foreign one only when the person agreed to its host. Each is TLS or
 * an upgrade the connection insists on, never plain.
 */
function ranked(found, kind, agreed) {
  const chosen = found?.[kind] ? [found[kind]] : [];
  const rest = (found?.candidates || []).filter((c) => c.kind === kind).sort((a, b) => Number(Boolean(b.verified)) - Number(Boolean(a.verified)) || (b.confidence || 0) - (a.confidence || 0));
  const seen = new Set();
  const out = [];
  for (const c of [...chosen, ...rest]) {
    const key = `${c.host}|${c.port}|${c.secure}`;
    if (seen.has(key) || (c.foreign && !agreed.has(String(c.host).toLowerCase()))) continue;
    seen.add(key);
    out.push({ host: c.host, port: c.port, secure: Boolean(c.secure), starttls: !c.secure, banner: c.banner || null, user: c.user || null });
  }
  return out;
}

/** A server the person typed (a row's own columns): TLS on its usual port, otherwise an upgrade, never plain. */
const given = (s, tlsPort) => {
  const secure = s.secure ?? s.port === tlsPort;
  return { ...s, secure, starttls: !secure };
};

const WORST = ['password', 'app-password', 'certificate', 'unreachable', 'no-server', 'other'];

/**
 * Set an account up from an address and a password.
 * @param {{ email: string, password: string, name?: string, user?: string, imap?: object, smtp?: object,
 *   discover: (email) => Promise<object>, tryImap: (server, user, password) => Promise<{ ok, message?, folders?, banner? }>,
 *   trySmtp: (server, user, password) => Promise<{ ok, message?, banner? }>, most?: number }} o
 *   `imap`/`smtp` given (from a row's own columns) are tried before anything found;
 *   `confirmed` the foreign hosts the person has agreed to send the password to
 * @returns {Promise<{ outcome: 'ready'|'sign-in'|'app-password'|'password'|'not-found'|'unreachable'|'certificate'|'outgoing'|'confirm'|'not-an-address', account?: object, provider?: object, server?: string, foreign?: object[], tried: object[], found?: object }>}
 *   'confirm': `foreign` lists the servers ({ kind, host, port }) to show and agree to; nothing was tried
 */
export async function setupAccount({ email, password, name, user, imap = null, smtp = null, confirmed = [], discover, tryImap, trySmtp, most = 6 }) {
  const address = String(email || '').trim();
  const tried = [];
  if (!isAddress(address)) return { outcome: 'not-an-address', provider: null, tried };
  // No password: the table alone says whether the provider signs in with its own page. Nothing is looked up.
  if (!password) {
    const known = await discover(address, { offline: true });
    return { outcome: known?.oauth ? 'sign-in' : 'password', provider: known?.provider || null, found: known, tried };
  }
  const found = await discover(address);
  const provider = found?.provider || null;
  const agreed = new Set(confirmed.map((h) => String(h).toLowerCase()));
  const ask = [found?.imap, found?.smtp].filter((c) => c?.foreign && !agreed.has(String(c.host).toLowerCase()));
  if (ask.length) {
    const kinds = [['imap', found.imap], ['smtp', found.smtp]].filter(([, c]) => ask.includes(c));
    return { outcome: 'confirm', provider, found, foreign: kinds.map(([kind, c]) => ({ kind, host: c.host, port: c.port })), tried };
  }
  const incoming = [...(imap?.host ? [given(imap, 993)] : []), ...ranked(found, 'imap', agreed)].slice(0, most);
  const outgoing = [...(smtp?.host ? [given(smtp, 465)] : []), ...ranked(found, 'smtp', agreed)].slice(0, most);
  if (!incoming.length) return { outcome: 'not-found', provider, found, tried };

  let inServer = null;
  let inUser = null;
  let inAnswer = null;
  const failures = [];
  for (const server of incoming) {
    for (const u of userNames(address, user || server.user)) {
      const r = await tryImap(server, u, password);
      const failure = r.ok ? null : failureOf(r.message);
      tried.push({ kind: 'imap', host: server.host, port: server.port, user: u, ok: Boolean(r.ok), failure });
      if (r.ok) {
        inServer = server;
        inUser = u;
        inAnswer = r;
        break;
      }
      failures.push(failure);
      // A server that never answered will not answer to another name either.
      if (failure === 'unreachable' || failure === 'no-server' || failure === 'certificate') break;
    }
    if (inServer) break;
  }
  if (!inServer) {
    const worst = WORST.find((w) => failures.includes(w)) || 'unreachable';
    // A known provider that refuses the password wants its app password, or its browser sign-in.
    const outcome = worst === 'password' && found?.appPassword ? 'app-password' : worst === 'no-server' || worst === 'other' ? 'unreachable' : worst;
    return { outcome, provider, found, tried };
  }

  let outServer = null;
  let outUser = null;
  for (const server of outgoing) {
    for (const u of [...new Set([inUser, ...userNames(address, user || server.user)])]) {
      const r = await trySmtp(server, u, password);
      tried.push({ kind: 'smtp', host: server.host, port: server.port, user: u, ok: Boolean(r.ok), failure: r.ok ? null : failureOf(r.message) });
      if (r.ok) {
        outServer = server;
        outUser = u;
        break;
      }
      if (['unreachable', 'no-server', 'certificate'].includes(failureOf(r.message))) break;
    }
    if (outServer) break;
  }
  const shape = (s, u) => (s ? { host: s.host, port: s.port, secure: Boolean(s.secure), starttls: !s.secure, user: u && u !== address ? u : undefined } : null);
  const account = { email: address, name: name || address, imap: shape(inServer, inUser), smtp: shape(outServer || outgoing[0] || null, outUser || inUser), provider: provider?.id || null };
  const software = serverSoftware(inAnswer?.banner || inServer.banner);
  return { outcome: outServer || !outgoing.length ? 'ready' : 'outgoing', account, provider, server: software, folders: inAnswer?.folders ?? null, found, tried };
}
