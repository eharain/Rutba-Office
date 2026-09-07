// What a message is actually doing to you.
//
// Every mail client blocks remote images. None of the free ones tell you what
// was blocked, who it belonged to, or how to make it stop — you get a grey bar
// saying "images blocked" and no way to act on it. That is the gap this fills.
//
// Three questions, answered from the message that is already in memory, with no
// network call of any kind:
//
//   Who is watching?   Remote images grouped by host, with the known tracking
//                      networks named. A 1×1 pixel from a marketing platform is
//                      not the same as a logo from the sender's own domain, and
//                      the difference is visible here.
//   Can I leave?       RFC 8058 List-Unsubscribe, both the one-click POST form
//                      and the mailto fallback, so leaving a list is a button
//                      rather than a hunt through the footer.
//   Is it really them?  Authentication-Results: SPF, DKIM and DMARC as the
//                      receiving server judged them. A bank mail that fails
//                      DMARC is worth seeing before you read it, not after.
//
// All of it is inference over headers and markup we already hold. Nothing here
// contacts anybody, and the unsubscribe is only ever *offered* — the person
// presses it, we never leave a list on their behalf.

/**
 * Tracking networks worth naming.
 *
 * The list is deliberately short and made of things that are only ever
 * tracking: an open pixel, a click wrapper, an analytics beacon. It is not a
 * blocklist — nothing is blocked because of it, since every remote image is
 * blocked already. It exists so the report can say "Mailchimp" instead of
 * "list-manage.com", which is the difference between a warning and information.
 */
const TRACKERS = [
  [/(^|\.)list-manage\.com$/i, 'Mailchimp'],
  [/(^|\.)mailchimp\.com$/i, 'Mailchimp'],
  [/(^|\.)sendgrid\.net$/i, 'SendGrid'],
  [/(^|\.)sparkpostmail\.com$/i, 'SparkPost'],
  [/(^|\.)mailgun\.(org|net)$/i, 'Mailgun'],
  [/(^|\.)hubspot(links|email|usercontent)?\.(com|net)$/i, 'HubSpot'],
  [/(^|\.)hs-analytics\.net$/i, 'HubSpot'],
  [/(^|\.)marketo\.com$/i, 'Marketo'],
  [/(^|\.)mktoresp\.com$/i, 'Marketo'],
  [/(^|\.)pardot\.com$/i, 'Pardot'],
  [/(^|\.)exacttarget\.com$/i, 'Salesforce Marketing Cloud'],
  [/(^|\.)mcsv\.net$/i, 'Mailchimp'],
  [/(^|\.)klaviyo(mail)?\.com$/i, 'Klaviyo'],
  [/(^|\.)braze\.(com|eu)$/i, 'Braze'],
  [/(^|\.)iterable(email)?\.com$/i, 'Iterable'],
  [/(^|\.)customer\.io$/i, 'Customer.io'],
  [/(^|\.)intercom(assets|cdn|-mail)?\.(com|io)$/i, 'Intercom'],
  [/(^|\.)sendinblue\.com$/i, 'Brevo'],
  [/(^|\.)brevo\.com$/i, 'Brevo'],
  [/(^|\.)constantcontact\.com$/i, 'Constant Contact'],
  [/(^|\.)campaign-archive\.com$/i, 'Mailchimp'],
  [/(^|\.)cmail\d*\.com$/i, 'Campaign Monitor'],
  [/(^|\.)createsend\d*\.com$/i, 'Campaign Monitor'],
  [/(^|\.)mailtrack\.io$/i, 'Mailtrack'],
  [/(^|\.)bananatag\.com$/i, 'Bananatag'],
  [/(^|\.)yesware\.com$/i, 'Yesware'],
  [/(^|\.)streak-app\.com$/i, 'Streak'],
  [/(^|\.)mixmax\.com$/i, 'Mixmax'],
  [/(^|\.)hubapi\.com$/i, 'HubSpot'],
  [/(^|\.)google-analytics\.com$/i, 'Google Analytics'],
  [/(^|\.)googletagmanager\.com$/i, 'Google Tag Manager'],
  [/(^|\.)doubleclick\.net$/i, 'Google Ads'],
  [/(^|\.)facebook\.com$/i, 'Meta'],
  [/(^|\.)omtrdc\.net$/i, 'Adobe Analytics'],
  [/(^|\.)everesttech\.net$/i, 'Adobe'],
  [/(^|\.)dc\.services\.adobe\.com$/i, 'Adobe'],
  [/(^|\.)sailthru\.com$/i, 'Sailthru'],
  [/(^|\.)responsys\.net$/i, 'Oracle Responsys'],
  [/(^|\.)eloqua\.com$/i, 'Oracle Eloqua'],
  [/(^|\.)en25\.com$/i, 'Oracle Eloqua'],
  [/(^|\.)cheetahmail\.com$/i, 'Cheetah Digital'],
  [/(^|\.)emltrk\.com$/i, 'Email Tracker'],
  [/(^|\.)mailerlite\.com$/i, 'MailerLite'],
  [/(^|\.)postmarkapp\.com$/i, 'Postmark'],
  [/(^|\.)amazonses\.com$/i, 'Amazon SES'],
  [/(^|\.)awstrack\.me$/i, 'Amazon SES'],
];

const trackerName = (host) => TRACKERS.find(([re]) => re.test(host))?.[1] ?? null;

/** The registrable-ish part of a host, good enough to compare against a sender. */
function rootDomain(host) {
  const parts = String(host || '').toLowerCase().split('.').filter(Boolean);
  if (parts.length <= 2) return parts.join('.');
  // Two-level public suffixes we actually meet in mail.
  const two = /^(co|com|org|net|ac|gov|edu)\.[a-z]{2}$/.test(parts.slice(-2).join('.'));
  return parts.slice(two ? -3 : -2).join('.');
}

/**
 * Every remote resource the body would have fetched.
 *
 * Both `src` and `background`, plus `url(...)` in style attributes, because a
 * pixel hidden in CSS counts exactly as much as one in an `<img>`.
 */
function remoteHosts(html) {
  const found = [];
  const add = (raw, how, around) => {
    let url;
    try {
      url = new URL(raw);
    } catch {
      return;
    }
    if (!/^https?:$/.test(url.protocol)) return;

    // A 1×1 or zero-sized image is not decoration. The dimension has to be
    // read as a number, not matched as a prefix: `width="180"` starts with a 1
    // and a pattern that stops there calls every logo a tracking pixel.
    const dimension = (name) => {
      const m = new RegExp(`\\b${name}\\s*=\\s*["']?(\\d+)`, 'i').exec(around || '');
      return m ? Number(m[1]) : null;
    };
    const width = dimension('width');
    const height = dimension('height');
    const tiny =
      (width !== null && width <= 1) ||
      (height !== null && height <= 1) ||
      /\b(?:width|height)\s*:\s*[01]px/i.test(around || '');

    found.push({ host: url.hostname, url: url.href, how, pixel: tiny });
  };

  for (const m of String(html || '').matchAll(/<img\b([^>]*)>/gi)) {
    const tag = m[1];
    const src = /\bsrc\s*=\s*["']([^"']+)["']/i.exec(tag)?.[1];
    if (src) add(src, 'image', tag);
  }
  for (const m of String(html || '').matchAll(/\bbackground(?:-image)?\s*[:=]\s*["']?(?:url\()?["']?(https?:\/\/[^"')\s>]+)/gi)) {
    add(m[1], 'background', '');
  }
  return found;
}

/**
 * Group them the way a person thinks about them: whose, and what for.
 * A host that matches the sender's own domain is "the sender"; a named network
 * is that network; anything else is simply third party.
 */
export function trackerReport(message) {
  const senderDomain = rootDomain(String(message?.from?.[0]?.address || '').split('@')[1] || '');
  const hits = remoteHosts(message?.html);
  const byHost = new Map();

  for (const hit of hits) {
    const root = rootDomain(hit.host);
    const entry = byHost.get(hit.host) || {
      host: hit.host,
      count: 0,
      pixels: 0,
      network: trackerName(hit.host),
      sameAsSender: Boolean(senderDomain) && root === senderDomain,
      sample: hit.url,
    };
    entry.count++;
    if (hit.pixel) entry.pixels++;
    byHost.set(hit.host, entry);
  }

  const hosts = [...byHost.values()].sort((a, b) => b.pixels - a.pixels || b.count - a.count);
  const networks = [...new Set(hosts.map((h) => h.network).filter(Boolean))];

  return {
    total: hits.length,
    hosts,
    networks,
    // The honest headline. A pixel from a named network is a read receipt;
    // an image from the sender's own domain usually is not.
    watching: hosts.filter((h) => h.pixels > 0 || (h.network && !h.sameAsSender)).length,
  };
}

/**
 * RFC 2369 List-Unsubscribe, with the RFC 8058 one-click form when the sender
 * offers it.
 *
 * One-click means a POST with a fixed body and no confirmation page, and it is
 * only valid when `List-Unsubscribe-Post` is present — a GET to the same URL is
 * how mailing lists used to get people unsubscribed by link-scanning virus
 * checkers, so the distinction matters.
 */
export function unsubscribeOffer(message) {
  const header = (name) =>
    (message?.headers || []).find((h) => (h.key || h.name || '').toLowerCase() === name)?.value ?? null;

  const raw = header('list-unsubscribe');
  if (!raw) return null;

  const targets = [...String(raw).matchAll(/<([^>]+)>/g)].map((m) => m[1].trim());
  const http = targets.find((t) => /^https?:/i.test(t)) || null;
  const mailto = targets.find((t) => /^mailto:/i.test(t)) || null;
  const oneClick = /one-click/i.test(header('list-unsubscribe-post') || '');

  if (!http && !mailto) return null;

  let address = null;
  let subject = 'unsubscribe';
  if (mailto) {
    try {
      const url = new URL(mailto);
      address = decodeURIComponent(url.pathname);
      subject = url.searchParams.get('subject') || subject;
    } catch {
      address = mailto.replace(/^mailto:/i, '').split('?')[0];
    }
  }

  return {
    http,
    mailto,
    address,
    subject,
    oneClick: oneClick && Boolean(http),
    list: header('list-id') || message?.listId || null,
    // What pressing the button will do, in words, before it is pressed.
    method: oneClick && http ? 'one-click' : http ? 'web' : 'email',
  };
}

/**
 * SPF, DKIM and DMARC as the *receiving* server recorded them.
 *
 * This is a report, not a verification: the trustworthy copy is the one written
 * by the server that accepted the message, and re-checking it here months later
 * against rotated keys would say nothing useful. Where no Authentication-Results
 * header exists — every imported archive, most old mail — the answer is
 * "unknown", which is stated rather than shaded into "pass".
 */
export function authenticationReport(message) {
  const rows = (message?.headers || []).filter((h) => (h.key || h.name || '').toLowerCase() === 'authentication-results');
  if (!rows.length) return { known: false, spf: null, dkim: null, dmarc: null, summary: 'Not recorded' };

  const text = rows.map((r) => r.value).join('; ');
  const verdict = (what) => {
    const m = new RegExp(`\\b${what}\\s*=\\s*([a-z]+)`, 'i').exec(text);
    return m ? m[1].toLowerCase() : null;
  };

  const spf = verdict('spf');
  const dkim = verdict('dkim');
  const dmarc = verdict('dmarc');
  const bad = [spf, dkim, dmarc].filter((v) => v && v !== 'pass' && v !== 'none');

  return {
    known: true,
    spf,
    dkim,
    dmarc,
    ok: !bad.length && [spf, dkim, dmarc].some((v) => v === 'pass'),
    summary: bad.length
      ? `Failed ${[spf !== 'pass' && spf ? 'SPF' : null, dkim !== 'pass' && dkim ? 'DKIM' : null, dmarc !== 'pass' && dmarc ? 'DMARC' : null].filter(Boolean).join(', ')}`
      : 'Sender checks passed',
  };
}

/**
 * Everything the reader should know about a message before trusting it,
 * assembled in one pass so the window makes one call.
 */
export function insightFor(message) {
  const trackers = trackerReport(message);
  const unsubscribe = unsubscribeOffer(message);
  const auth = authenticationReport(message);

  // A bulk message is worth marking, because the actions that suit it —
  // unsubscribe, block, sweep the sender — are not the actions that suit a
  // message a person wrote to you.
  const header = (name) => (message?.headers || []).find((h) => (h.key || h.name || '').toLowerCase() === name)?.value ?? null;
  const bulk = Boolean(
    unsubscribe ||
      header('list-id') ||
      /\b(bulk|list|auto-generated|auto-replied)\b/i.test(header('precedence') || header('auto-submitted') || '')
  );

  return { trackers, unsubscribe, auth, bulk };
}
