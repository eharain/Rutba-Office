// Finding somebody's mail server from their address.
//
// "Add account" should be an address and a password. Everything else — the
// IMAP host, the SMTP host, the ports, whether the connection is encrypted
// from the first byte or upgraded, whether the provider wants a browser
// sign-in instead of a password — is knowable, and every mail program that
// people find easy knows it. This module asks in every way there is:
//
//   1. A table of the providers most people use, by domain.
//   2. The domain's MX records: mail hosted at Google, Microsoft, Yahoo, Zoho,
//      Fastmail or iCloud under a company's own domain is named by them.
//   3. SRV records, the standard way a domain publishes its own servers
//      (RFC 6186 / 8314): _imaps._tcp, _imap._tcp, _submissions._tcp,
//      _submission._tcp.
//   4. Autoconfig, the format Thunderbird made common: the domain's own
//      autoconfig host, its .well-known path, then Mozilla's public database.
//   5. Microsoft's autodiscover: the JSON endpoint answers for any mailbox
//      hosted on Microsoft 365, whatever the domain.
//   6. The conventional names — imap., mail., smtp., the bare domain — and a
//      knock on each port that matters: does something answer, and does it
//      say IMAP or SMTP? A name that answers outranks a name that is guessed.
//
// The steps run at once, each with its own short timeout, and the whole
// search is over in a few seconds. What is returned is a settings pair, the
// candidates that were considered, and a step-by-step account the dialog can
// show, so that when it is wrong the person can see what was tried and put
// the right thing in the advanced fields.
//
// Nothing here is called during a check run: `offline` restricts the search
// to the table, and the tests inject every lookup.

import dns from 'node:dns';
import net from 'node:net';
import tls from 'node:tls';

/** The providers most people use. `oauth` names the browser sign-in that replaces a password. */
export const KNOWN = [
  { id: 'google', label: 'Google', match: /^(gmail\.com|googlemail\.com)$/i, oauth: 'google',
    imap: { host: 'imap.gmail.com', port: 993, secure: true }, smtp: { host: 'smtp.gmail.com', port: 465, secure: true } },
  { id: 'microsoft', label: 'Microsoft', match: /^(outlook\.(com|co\.uk|de|fr|es|it|jp|com\.au)|hotmail\.(com|co\.uk|de|fr|es|it)|live\.(com|co\.uk|de|fr|it|nl)|msn\.com)$/i, oauth: 'microsoft',
    imap: { host: 'outlook.office365.com', port: 993, secure: true }, smtp: { host: 'smtp-mail.outlook.com', port: 587, secure: false, starttls: true } },
  { id: 'microsoft365', label: 'Microsoft 365', match: /^onmicrosoft\.com$/i, oauth: 'microsoft',
    imap: { host: 'outlook.office365.com', port: 993, secure: true }, smtp: { host: 'smtp.office365.com', port: 587, secure: false, starttls: true } },
  { id: 'yahoo', label: 'Yahoo', match: /^(yahoo\.[a-z.]+|ymail\.com|rocketmail\.com)$/i,
    imap: { host: 'imap.mail.yahoo.com', port: 993, secure: true }, smtp: { host: 'smtp.mail.yahoo.com', port: 465, secure: true },
    note: 'Yahoo wants an app password, made in the account\'s security settings.' },
  { id: 'icloud', label: 'iCloud', match: /^(icloud\.com|me\.com|mac\.com)$/i,
    imap: { host: 'imap.mail.me.com', port: 993, secure: true }, smtp: { host: 'smtp.mail.me.com', port: 587, secure: false, starttls: true },
    note: 'iCloud wants an app-specific password, made at appleid.apple.com.' },
  { id: 'aol', label: 'AOL', match: /^(aol\.com|aim\.com|verizon\.net)$/i,
    imap: { host: 'imap.aol.com', port: 993, secure: true }, smtp: { host: 'smtp.aol.com', port: 465, secure: true } },
  { id: 'zoho', label: 'Zoho', match: /^zoho(mail)?\.(com|eu|in|com\.au)$/i,
    imap: { host: 'imap.zoho.com', port: 993, secure: true }, smtp: { host: 'smtp.zoho.com', port: 465, secure: true } },
  { id: 'fastmail', label: 'Fastmail', match: /^(fastmail\.(com|fm|net|org|us|co\.uk)|sent\.com|messagingengine\.com)$/i,
    imap: { host: 'imap.fastmail.com', port: 993, secure: true }, smtp: { host: 'smtp.fastmail.com', port: 465, secure: true },
    note: 'Fastmail wants an app password, made in Settings → Privacy & Security.' },
  { id: 'gmx', label: 'GMX', match: /^gmx\.(com|net|de|at|ch|co\.uk|fr|es|it)$/i,
    imap: { host: 'imap.gmx.com', port: 993, secure: true }, smtp: { host: 'mail.gmx.com', port: 587, secure: false, starttls: true } },
  { id: 'webde', label: 'WEB.DE', match: /^web\.de$/i,
    imap: { host: 'imap.web.de', port: 993, secure: true }, smtp: { host: 'smtp.web.de', port: 587, secure: false, starttls: true } },
  { id: 'yandex', label: 'Yandex', match: /^(yandex\.(com|ru|ua|kz|by)|ya\.ru)$/i,
    imap: { host: 'imap.yandex.com', port: 993, secure: true }, smtp: { host: 'smtp.yandex.com', port: 465, secure: true } },
  { id: 'mailcom', label: 'mail.com', match: /^(mail\.com|email\.com|usa\.com|consultant\.com|engineer\.com|post\.com)$/i,
    imap: { host: 'imap.mail.com', port: 993, secure: true }, smtp: { host: 'smtp.mail.com', port: 587, secure: false, starttls: true } },
  { id: 'proton', label: 'Proton', match: /^(proton\.me|protonmail\.(com|ch)|pm\.me)$/i,
    imap: { host: '127.0.0.1', port: 1143, secure: false, starttls: true }, smtp: { host: '127.0.0.1', port: 1025, secure: false, starttls: true },
    note: 'Proton Mail reaches a mail program through Proton Mail Bridge, which runs on this computer; the servers are the bridge\'s.' },
  { id: 'bt', label: 'BT', match: /^(btinternet\.com|btopenworld\.com)$/i,
    imap: { host: 'mail.btinternet.com', port: 993, secure: true }, smtp: { host: 'mail.btinternet.com', port: 465, secure: true } },
  { id: 'sky', label: 'Sky', match: /^sky\.com$/i,
    imap: { host: 'imap.tools.sky.com', port: 993, secure: true }, smtp: { host: 'smtp.tools.sky.com', port: 465, secure: true } },
  { id: 'virgin', label: 'Virgin Media', match: /^(virginmedia\.com|ntlworld\.com|blueyonder\.co\.uk)$/i,
    imap: { host: 'imap.virginmedia.com', port: 993, secure: true }, smtp: { host: 'smtp.virginmedia.com', port: 465, secure: true } },
  { id: 'comcast', label: 'Xfinity', match: /^comcast\.net$/i,
    imap: { host: 'imap.comcast.net', port: 993, secure: true }, smtp: { host: 'smtp.comcast.net', port: 465, secure: true } },
  { id: 'att', label: 'AT&T', match: /^(att\.net|sbcglobal\.net|bellsouth\.net)$/i,
    imap: { host: 'imap.mail.att.net', port: 993, secure: true }, smtp: { host: 'smtp.mail.att.net', port: 465, secure: true } },
  { id: 'tonline', label: 'T-Online', match: /^t-online\.de$/i,
    imap: { host: 'secureimap.t-online.de', port: 993, secure: true }, smtp: { host: 'securesmtp.t-online.de', port: 465, secure: true } },
  { id: 'orange', label: 'Orange', match: /^(orange\.fr|wanadoo\.fr)$/i,
    imap: { host: 'imap.orange.fr', port: 993, secure: true }, smtp: { host: 'smtp.orange.fr', port: 465, secure: true } },
  { id: 'free', label: 'Free', match: /^free\.fr$/i,
    imap: { host: 'imap.free.fr', port: 993, secure: true }, smtp: { host: 'smtp.free.fr', port: 465, secure: true } },
];

/** What the MX host says about who hosts the mail. */
const MX_HINTS = [
  [/(^|\.)(google|googlemail)\.com$/i, 'google'],
  [/(^|\.)(mail\.protection\.outlook\.com|outlook\.com|office365\.com|eo\.outlook\.com)$/i, 'microsoft365'],
  [/(^|\.)yahoodns\.net$/i, 'yahoo'],
  [/(^|\.)zoho(mail)?\.(com|eu|in)$/i, 'zoho'],
  [/(^|\.)icloud\.com$/i, 'icloud'],
  [/(^|\.)messagingengine\.com$/i, 'fastmail'],
  [/(^|\.)yandex\.(net|ru)$/i, 'yandex'],
  [/(^|\.)(gmx|web)\.(net|de)$/i, 'gmx'],
];

const STEP_TIMEOUT = 3500;
const TOTAL_TIMEOUT = 9000;

const withTimeout = (promise, ms, what) =>
  new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`${what} took longer than ${ms} ms`)), ms);
    promise.then(
      (v) => {
        clearTimeout(timer);
        resolve(v);
      },
      (e) => {
        clearTimeout(timer);
        reject(e);
      }
    );
  });

/* ── the lookups, each replaceable ────────────────────────────────────── */

const defaultDns = {
  resolveMx: (domain) => dns.promises.resolveMx(domain),
  resolveSrv: (name) => dns.promises.resolveSrv(name),
  resolveCname: (name) => dns.promises.resolveCname(name),
};

async function defaultFetchText(url) {
  const res = await fetch(url, { signal: AbortSignal.timeout(STEP_TIMEOUT), redirect: 'follow', headers: { 'user-agent': 'Rutba Office' } });
  if (!res.ok) throw new Error(`${res.status} from ${new URL(url).host}`);
  return res.text();
}

/**
 * Knock on a port. Connects (with TLS when `secure`), reads the greeting,
 * and answers with what the server said it was. A closed port, a refused
 * connection, a certificate for another name and silence all come back as
 * `null`; nothing is sent but, on a plain port, the one line that asks for
 * the upgrade, so no server ever sees a password from here.
 */
function defaultProbe(host, port, kind, secure) {
  return new Promise((resolve) => {
    let settled = false;
    let banner = '';
    const done = (value) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      try {
        socket.destroy();
      } catch {
        // already gone
      }
      resolve(value);
    };
    const timer = setTimeout(() => done(null), STEP_TIMEOUT);
    const socket = secure
      ? tls.connect({ host, port, servername: host, rejectUnauthorized: true })
      : net.connect({ host, port });
    socket.setEncoding('utf8');
    socket.on('error', () => done(null));
    socket.on('data', (chunk) => {
      banner += chunk;
      const line = banner.split(/\r?\n/)[0] || '';
      if (kind === 'imap' && /^\* (OK|PREAUTH)/i.test(line)) done({ banner: line.slice(0, 120), starttls: !secure });
      else if (kind === 'smtp' && /^220[ -]/.test(line)) done({ banner: line.slice(0, 120), starttls: !secure });
      else if (banner.length > 512) done(null);
    });
    socket.on('close', () => done(null));
  });
}

/* ── the sources ──────────────────────────────────────────────────────── */

const candidate = (kind, host, port, secure, source, confidence, extra = {}) => ({
  kind, host: String(host).toLowerCase().replace(/\.$/, ''), port: Number(port), secure: Boolean(secure), starttls: !secure, source, confidence, ...extra,
});

function fromKnown(entry, source, confidence) {
  return [
    candidate('imap', entry.imap.host, entry.imap.port, entry.imap.secure, source, confidence, { provider: entry.id }),
    candidate('smtp', entry.smtp.host, entry.smtp.port, entry.smtp.secure, source, confidence, { provider: entry.id }),
  ];
}

export function knownProvider(domain) {
  return KNOWN.find((k) => k.match.test(domain)) || null;
}

async function lookupMx(domain, api) {
  const records = await withTimeout(api.dns.resolveMx(domain), STEP_TIMEOUT, 'the MX lookup');
  const hosts = records.sort((a, b) => a.priority - b.priority).map((r) => r.exchange.toLowerCase().replace(/\.$/, ''));
  if (!hosts.length) return { detail: 'no MX record', candidates: [] };
  for (const host of hosts) {
    const hint = MX_HINTS.find(([re]) => re.test(host));
    if (hint) {
      const entry = KNOWN.find((k) => k.id === hint[1]);
      if (entry) return { detail: `${hosts[0]} — mail is hosted at ${entry.label}`, candidates: fromKnown(entry, 'mx', 2), provider: entry };
    }
  }
  // The MX host itself often runs IMAP too; a knock will tell.
  return { detail: hosts[0], candidates: [candidate('imap', hosts[0], 993, true, 'mx', 1), candidate('smtp', hosts[0], 587, false, 'mx', 1)], mxHosts: hosts };
}

async function lookupSrv(domain, api) {
  const ask = async (name, kind, secure) => {
    try {
      const records = await withTimeout(api.dns.resolveSrv(`${name}.${domain}`), STEP_TIMEOUT, 'the SRV lookup');
      return records
        .filter((r) => r.name && r.name !== '.')
        .sort((a, b) => a.priority - b.priority || b.weight - a.weight)
        .map((r) => candidate(kind, r.name, r.port, secure, 'srv', 3));
    } catch {
      return [];
    }
  };
  const [imaps, imap, submissions, submission] = await Promise.all([
    ask('_imaps._tcp', 'imap', true),
    ask('_imap._tcp', 'imap', false),
    ask('_submissions._tcp', 'smtp', true),
    ask('_submission._tcp', 'smtp', false),
  ]);
  const candidates = [...imaps, ...imap, ...submissions, ...submission];
  return { detail: candidates.length ? candidates.map((c) => `${c.host}:${c.port}`).join(', ') : 'none published', candidates };
}

/** Thunderbird's autoconfig XML, read without a parser: the schema is small and fixed. */
export function parseAutoconfig(xml, email) {
  const servers = [];
  const re = /<(incomingServer|outgoingServer)\s+type="([^"]+)"[^>]*>([\s\S]*?)<\/\1>/g;
  let m;
  while ((m = re.exec(xml))) {
    const [, tag, type, body] = m;
    const pick = (name) => (new RegExp(`<${name}>([^<]*)</${name}>`).exec(body) || [])[1]?.trim();
    const kind = tag === 'incomingServer' ? (type === 'imap' ? 'imap' : null) : type === 'smtp' ? 'smtp' : null;
    if (!kind) continue;
    const socket = (pick('socketType') || '').toUpperCase();
    const username = (pick('username') || '').replace('%EMAILADDRESS%', email).replace('%EMAILLOCALPART%', email.split('@')[0]).replace('%EMAILDOMAIN%', email.split('@')[1] || '');
    servers.push({ kind, host: pick('hostname'), port: Number(pick('port')), secure: socket === 'SSL', starttls: socket === 'STARTTLS', user: username || null, auth: pick('authentication') || null });
  }
  return servers.filter((s) => s.host && s.port);
}

async function lookupAutoconfig(domain, email, api) {
  const urls = [
    `https://autoconfig.${domain}/mail/config-v1.1.xml?emailaddress=${encodeURIComponent(email)}`,
    `https://${domain}/.well-known/autoconfig/mail/config-v1.1.xml`,
    `https://autoconfig.thunderbird.net/v1.1/${encodeURIComponent(domain)}`,
  ];
  const labels = ['the domain\'s own autoconfig', 'its .well-known path', 'Mozilla\'s database'];
  const tried = [];
  for (let i = 0; i < urls.length; i++) {
    try {
      const xml = await withTimeout(api.fetchText(urls[i]), STEP_TIMEOUT, labels[i]);
      const servers = parseAutoconfig(xml, email);
      if (servers.length) {
        return {
          detail: `${labels[i]} names ${servers.map((s) => s.host).filter((v, j, a) => a.indexOf(v) === j).join(' and ')}`,
          candidates: servers.map((s) => candidate(s.kind, s.host, s.port, s.secure, 'autoconfig', 3, { user: s.user, starttls: s.starttls })),
        };
      }
      tried.push(`${labels[i]}: nothing for this domain`);
    } catch (err) {
      tried.push(`${labels[i]}: ${String(err.message).slice(0, 60)}`);
    }
  }
  return { detail: tried.join('; '), candidates: [] };
}

/** Microsoft's JSON autodiscover: answers for any mailbox on Microsoft 365. */
async function lookupAutodiscover(email, api) {
  const url = `https://outlook.office365.com/autodiscover/autodiscover.json?Email=${encodeURIComponent(email)}&Protocol=IMAP`;
  try {
    const text = await withTimeout(api.fetchText(url), STEP_TIMEOUT, 'Microsoft autodiscover');
    const json = JSON.parse(text);
    if (json && json.Url && /office365|outlook/i.test(json.Url)) {
      const entry = KNOWN.find((k) => k.id === 'microsoft365');
      return { detail: `Microsoft 365 hosts this mailbox (${json.Url})`, candidates: fromKnown(entry, 'autodiscover', 3), provider: entry };
    }
    return { detail: json?.ErrorCode ? `not a Microsoft 365 mailbox (${json.ErrorCode})` : 'not a Microsoft 365 mailbox', candidates: [] };
  } catch (err) {
    return { detail: `no answer (${String(err.message).slice(0, 60)})`, candidates: [] };
  }
}

function guesses(domain, mxHosts = []) {
  const imapHosts = [`imap.${domain}`, `mail.${domain}`, domain, `imap4.${domain}`, `mx.${domain}`, ...mxHosts];
  const smtpHosts = [`smtp.${domain}`, `mail.${domain}`, domain, `mx.${domain}`, ...mxHosts];
  const uniq = (a) => a.filter((v, i) => a.indexOf(v) === i);
  const out = [];
  for (const h of uniq(imapHosts)) {
    out.push(candidate('imap', h, 993, true, 'guess', 1));
    out.push(candidate('imap', h, 143, false, 'guess', 1));
  }
  for (const h of uniq(smtpHosts)) {
    out.push(candidate('smtp', h, 465, true, 'guess', 1));
    out.push(candidate('smtp', h, 587, false, 'guess', 1));
  }
  return out;
}

/* ── the search ───────────────────────────────────────────────────────── */

/**
 * @param {string} email
 * @param {object} [options]
 * @param {boolean} [options.offline]  the table only — a check run, or no network
 * @param {object}  [options.seed]     settings that came from another client on this computer; they are kept and checked
 * @returns {Promise<{email, domain, imap, smtp, oauth, provider, note, source, steps, candidates}>}
 */
export async function discoverMailServers(email, options = {}) {
  const address = String(email || '').trim();
  const domain = (address.split('@')[1] || '').toLowerCase();
  const api = {
    dns: options.dns || defaultDns,
    fetchText: options.fetchText || defaultFetchText,
    probe: options.probe || defaultProbe,
  };
  const steps = [];
  const step = (name, status, detail) => steps.push({ name, status, detail: detail || '' });
  const result = { email: address, domain, imap: null, smtp: null, oauth: null, provider: null, note: null, source: 'none', steps, candidates: [] };
  if (!domain) {
    step('address', 'failed', 'not an email address');
    return result;
  }

  let candidates = [];
  let provider = knownProvider(domain);
  if (provider) {
    candidates.push(...fromKnown(provider, 'known', 3));
    step('known providers', 'ok', `${domain} is ${provider.label}`);
  } else step('known providers', 'none', `${domain} is not in the table`);

  // What came in from another client is kept, and knocked on like the rest.
  if (options.seed?.imap?.host) candidates.push(candidate('imap', options.seed.imap.host, options.seed.imap.port || 993, options.seed.imap.secure !== false, 'import', 3));
  if (options.seed?.smtp?.host) candidates.push(candidate('smtp', options.seed.smtp.host, options.seed.smtp.port || 465, options.seed.smtp.secure !== false, 'import', 3));

  if (!options.offline) {
    const started = Date.now();
    const remaining = () => Math.max(500, TOTAL_TIMEOUT - (Date.now() - started));
    const safe = (name, p) =>
      withTimeout(p, remaining(), name).then(
        (v) => ({ ok: true, ...v }),
        (err) => ({ ok: false, detail: String(err.message).slice(0, 80), candidates: [] })
      );
    const [mx, srv, autoconfig, autodiscover] = await Promise.all([
      safe('MX', lookupMx(domain, api)),
      safe('SRV', lookupSrv(domain, api)),
      safe('autoconfig', lookupAutoconfig(domain, address, api)),
      provider?.oauth === 'google' ? Promise.resolve({ ok: true, detail: 'not needed', candidates: [] }) : safe('autodiscover', lookupAutodiscover(address, api)),
    ]);
    step('MX records', mx.ok ? (mx.candidates.length ? 'ok' : 'none') : 'failed', mx.detail);
    step('SRV records', srv.candidates.length ? 'ok' : srv.ok ? 'none' : 'failed', srv.detail);
    step('autoconfig', autoconfig.candidates.length ? 'ok' : autoconfig.ok ? 'none' : 'failed', autoconfig.detail);
    step('Microsoft autodiscover', autodiscover.candidates.length ? 'ok' : autodiscover.ok ? 'none' : 'failed', autodiscover.detail);
    for (const r of [mx, srv, autoconfig, autodiscover]) candidates.push(...(r.candidates || []));
    provider = provider || autodiscover.provider || mx.provider || null;

    // Nothing authoritative: the conventional names, and a knock on each.
    const authoritative = candidates.some((c) => c.confidence >= 3);
    if (!authoritative) candidates.push(...guesses(domain, mx.mxHosts || []));

    // The knock. Authoritative answers are checked too — a published server
    // that does not answer is worth knowing about — but the search never
    // waits on more than a handful.
    const toProbe = dedupe(candidates).sort((a, b) => b.confidence - a.confidence).slice(0, 14);
    const knocks = await Promise.all(
      toProbe.map(async (c) => {
        try {
          const answer = await withTimeout(api.probe(c.host, c.port, c.kind, c.secure), Math.min(STEP_TIMEOUT, remaining()), 'the knock');
          c.verified = Boolean(answer);
          if (answer?.banner) c.banner = answer.banner;
          if (answer) c.confidence += 1;
          return c;
        } catch {
          c.verified = false;
          return c;
        }
      })
    );
    const answered = knocks.filter((c) => c.verified);
    step('servers answering', answered.length ? 'ok' : 'none', answered.length ? answered.map((c) => `${c.host}:${c.port}`).join(', ') : 'nothing answered on the usual ports');
  } else {
    step('network', 'none', 'not asked — offline');
  }

  const ranked = dedupe(candidates).sort((a, b) => b.confidence - a.confidence || Number(Boolean(b.verified)) - Number(Boolean(a.verified)) || Number(b.secure) - Number(a.secure));
  // A verified server beats an unverified one whatever its source; among the
  // unverified, the source decides.
  const best = (kind) => ranked.find((c) => c.kind === kind && c.verified) || ranked.find((c) => c.kind === kind) || null;
  result.imap = pick(best('imap'));
  result.smtp = pick(best('smtp'));
  result.candidates = ranked;
  result.provider = provider ? { id: provider.id, label: provider.label } : null;
  result.oauth = provider?.oauth || null;
  result.source = result.imap?.source || 'none';
  result.note = provider?.note || (result.imap ? noteFor(result.imap, result.smtp) : `Nothing at ${domain} said where its mail server is. The advanced settings take the names from your provider.`);
  return result;
}

function dedupe(list) {
  const seen = new Map();
  for (const c of list) {
    const key = `${c.kind}|${c.host}|${c.port}|${c.secure}`;
    const had = seen.get(key);
    if (!had || c.confidence > had.confidence) seen.set(key, had ? { ...had, ...c, confidence: Math.max(had.confidence, c.confidence) } : c);
  }
  return [...seen.values()];
}

function pick(c) {
  if (!c) return null;
  return { host: c.host, port: c.port, secure: c.secure, starttls: !c.secure, user: c.user || null, source: c.source, verified: Boolean(c.verified) };
}

function noteFor(imap, smtp) {
  const how = { known: 'from the provider table', mx: 'from the domain\'s MX record', srv: 'from the SRV records the domain publishes', autoconfig: 'from the domain\'s autoconfig', autodiscover: 'from Microsoft autodiscover', import: 'from the client on this computer', guess: 'by convention' };
  const verified = imap.verified ? ' and it answered' : '';
  return `Found ${how[imap.source] || imap.source}${verified}${smtp && !smtp.verified && imap.verified ? '; the outgoing server has not answered yet' : ''}.`;
}

export default { discoverMailServers, parseAutoconfig, knownProvider, KNOWN };
