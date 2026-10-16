// A link followed in a message: what its words say it goes to, and why it
// should be warned about before the browser opens it (owner, 2026-10-16: "the
// link in email do not open it leads to a blank page . it should open browser
// instead while spam should have a protection on links").

import { t } from '@rutba/office-ui/messages';

const decode = (s) => s.replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>');
const textOf = (html) => decode(String(html || '').replace(/<[^>]*>/g, ' ')).replace(/\s+/g, ' ').trim();

/** An address as the browser writes it (a slash after a bare host, the host lower case), so a link and the address it asked for compare. */
export function normalUrl(url) {
  try {
    return new URL(url).href;
  } catch {
    return String(url || '');
  }
}

/** The words of each <a> in a message, by the address it goes to, written as the browser writes it. */
export function linkWords(html) {
  const out = new Map();
  const re = /<a\b[^<>]*?\bhref\s*=\s*(["'])(.*?)\1[^<>]*>([\s\S]*?)<\/a>/gi;
  const src = String(html || '');
  for (let m = re.exec(src); m; m = re.exec(src)) {
    const href = normalUrl(decode(m[2]).trim());
    const words = textOf(m[3]);
    if (!out.has(href)) out.set(href, []);
    if (words) out.get(href).push(words);
  }
  return out;
}

/** A link's host, lower case and without www. */
export function hostOf(url) {
  try {
    return new URL(url).hostname.toLowerCase().replace(/^www\./, '');
  } catch {
    return '';
  }
}

/** The site a link's words name when they read as an address (paypal.com, https://bank.example/login), or null. */
export function namedHost(words) {
  const m = /^(?:https?:\/\/)?((?:[a-z0-9-]+\.)+[a-z]{2,})(?::\d+)?(?:[/?#]\S*)?$/i.exec(String(words || '').trim());
  return m ? m[1].toLowerCase().replace(/^www\./, '') : null;
}

const sameSite = (a, b) => a === b || a.endsWith(`.${b}`) || b.endsWith(`.${a}`);

/**
 * Where a message's mailto: links point while it is shown: an address on a
 * name that cannot exist (RFC 2606's .invalid). Chromium hands a mailto: link
 * to the system rather than asking for a window, so it is written as a web
 * address the window's handler is given, and fromFrameUrl turns it back.
 */
export const MAILTO_HOST = 'https://mailto.rutba-office.invalid/';

/** A link as the message wrote it, from the address the frame asked for. */
export function fromFrameUrl(url) {
  const u = String(url || '');
  return u.startsWith(MAILTO_HOST) ? `mailto:${u.slice(MAILTO_HOST.length)}` : u;
}

/** What the window does with a followed link: open it in the browser, write to it, or refuse it. */
export function linkKind(url) {
  if (/^https?:\/\//i.test(url)) return 'web';
  if (/^mailto:/i.test(url)) return 'mail';
  return 'refused';
}

/**
 * Why following `url` from a message calls for a warning first, as
 * sentences, empty for an ordinary link:
 * - the message is in Junk;
 * - it failed the sender's checks;
 * - its words name a different site from the one it goes to;
 * - the address itself is a trick: a bare number, a lookalike punycode
 *   name, or a name before an @.
 */
export function linkRisks(url, { html = '', inJunk = false, authFailed = false } = {}) {
  const reasons = [];
  if (inJunk) reasons.push(t('This message is in Junk.'));
  if (authFailed) reasons.push(t('The sender’s address did not pass its checks (SPF, DKIM, DMARC).'));
  if (linkKind(url) !== 'web') return reasons;
  const host = hostOf(url);
  let parsed = null;
  try {
    parsed = new URL(url);
  } catch {
    parsed = null;
  }
  if (parsed?.username) reasons.push(t('The address hides where it goes behind a name and an @: it really goes to {host}.', { host }));
  if (/^\d{1,3}(\.\d{1,3}){3}$/.test(host) || host.startsWith('[')) reasons.push(t('The link goes to a bare number ({host}), not a named site.', { host }));
  if (host.split('.').some((part) => part.startsWith('xn--'))) reasons.push(t('The site’s name is written in letters that can look like others ({host}).', { host }));
  for (const words of linkWords(html).get(normalUrl(url)) || []) {
    const named = namedHost(words);
    if (named && host && !sameSite(named, host)) {
      reasons.push(t('The link reads {named} but goes to {host}.', { named, host }));
      break;
    }
  }
  return reasons;
}

/** The address a mailto: link writes to, and its subject and words when it carries them. */
export function mailtoFields(url) {
  try {
    const u = new URL(url);
    const p = u.searchParams;
    return { to: decodeURIComponent(u.pathname || ''), subject: p.get('subject') || '', text: p.get('body') || '' };
  } catch {
    return { to: '', subject: '', text: '' };
  }
}
