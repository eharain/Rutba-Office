// CalDAV and CardDAV: the protocol, spoken to one server.
//
// A calendar server and an address-book server are WebDAV servers that keep
// one .ics or one .vcf per event or card, each at its own address with an
// ETag that changes whenever it does. Everything here is that much and no
// more: finding the person's calendars and address books (RFC 6764's
// well-known address, then the principal, then its home), listing what a
// collection holds with each item's ETag, fetching the items that changed,
// and writing or removing one item at a time with If-Match, so an item
// changed on the server meanwhile is never overwritten unseen.
//
// Requests go only to the server the person typed, over HTTPS; plain HTTP
// is accepted for this computer alone (a server on localhost).

import { parse, kids, textOf } from '@rutba/office-formats/xml';

const local = (name) => String(name || '').replace(/^.*:/, '').toLowerCase();

/** Children of an XML node by local name, whatever prefix the server used. */
const childrenNamed = (node, name) => kids(node).filter((c) => local(c.name) === name);
const childNamed = (node, name) => childrenNamed(node, name)[0] || null;
function deep(node, name) {
  if (!node) return null;
  for (const c of kids(node)) {
    if (local(c.name) === name) return c;
    const d = deep(c, name);
    if (d) return d;
  }
  return null;
}

/**
 * A WebDAV multistatus response as a list of `{ href, props }`, `props` the
 * properties a 200 propstat carried, by local name, each its XML node.
 */
export function readMultistatus(xml) {
  const root = parse(String(xml || ''));
  const out = [];
  const status = deep(root, 'multistatus');
  for (const response of childrenNamed(status, 'response')) {
    const href = decodeURI(textOf(childNamed(response, 'href')).trim());
    const props = {};
    let missing = false;
    const own = textOf(childNamed(response, 'status'));
    if (/\s404\s/.test(own)) missing = true;
    for (const propstat of childrenNamed(response, 'propstat')) {
      if (!/\s2\d\d\s/.test(textOf(childNamed(propstat, 'status')))) continue;
      for (const p of kids(childNamed(propstat, 'prop'))) props[local(p.name)] = p;
    }
    out.push({ href, props, missing });
  }
  return out;
}

const NS = 'xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav" xmlns:card="urn:ietf:params:xml:ns:carddav" xmlns:cs="http://calendarserver.org/ns/" xmlns:ic="http://apple.com/ns/ical/"';

const escapeXml = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** An ETag as servers send it, quotes and all, or null. */
const etagOf = (node) => (node ? textOf(node).trim() || null : null);

/**
 * A client for one account. `url` is what the person typed — the server's
 * address, or their principal's, or a calendar's own.
 */
export function createDavClient({ url, user = '', password = '', token = null, fetch: fetchImpl = globalThis.fetch, timeoutMs = 20000 }) {
  let base;
  try {
    base = new URL(/^https?:\/\//i.test(String(url)) ? String(url) : `https://${url}`);
  } catch {
    throw new Error(`"${url}" is not a server address`);
  }
  const loopback = /^(localhost|127\.\d+\.\d+\.\d+|\[::1\])$/i.test(base.hostname);
  if (base.protocol !== 'https:' && !loopback) throw new Error('A calendar or contacts server is reached over HTTPS — give its address with https://');
  const auth = token ? `Bearer ${token}` : user || password ? `Basic ${Buffer.from(`${user}:${password}`).toString('base64')}` : null;
  const resolve = (href) => new URL(href, base).toString();

  async function request(method, href, { depth = null, body = null, type = 'application/xml; charset=utf-8', ifMatch = null, ifNoneMatch = null } = {}) {
    const headers = { 'User-Agent': 'Rutba Office' };
    if (auth) headers.Authorization = auth;
    if (depth !== null) headers.Depth = String(depth);
    if (body !== null) headers['Content-Type'] = type;
    if (ifMatch) headers['If-Match'] = ifMatch;
    if (ifNoneMatch) headers['If-None-Match'] = ifNoneMatch;
    let target = resolve(href);
    // Redirects are followed by hand, a few deep, keeping the method: fetch would turn a PROPFIND into a GET.
    for (let hop = 0; hop < 5; hop++) {
      const res = await fetchImpl(target, { method, headers, body, redirect: 'manual', signal: AbortSignal.timeout(timeoutMs) });
      if ([301, 302, 307, 308].includes(res.status) && res.headers.get('location')) {
        target = new URL(res.headers.get('location'), target).toString();
        if (hop === 0 && method === 'PROPFIND') base = new URL(target);
        continue;
      }
      const text = await res.text();
      if (res.status === 401 || res.status === 403) {
        const err = new Error('The server did not accept the user name and password.');
        err.status = res.status;
        throw err;
      }
      return { status: res.status, headers: res.headers, text, url: target };
    }
    throw new Error('The server sent the request round in circles.');
  }

  async function propfind(href, props, depth = 0) {
    const body = `<?xml version="1.0" encoding="utf-8"?><d:propfind ${NS}><d:prop>${props}</d:prop></d:propfind>`;
    const res = await request('PROPFIND', href, { depth, body });
    if (res.status !== 207) {
      const err = new Error(`The server answered ${res.status} where a list was asked for.`);
      err.status = res.status;
      throw err;
    }
    return readMultistatus(res.text);
  }

  /** The href inside a property such as current-user-principal or calendar-home-set. */
  const hrefIn = (node) => {
    const h = node ? deep(node, 'href') : null;
    return h ? textOf(h).trim() : null;
  };

  /**
   * The person's calendars ('caldav') or address books ('carddav'): each
   * collection's href, name, colour and change tag. Empty when the server
   * keeps none of that kind.
   */
  async function discover(kind) {
    const card = kind === 'carddav';
    const homeProp = card ? '<card:addressbook-home-set/>' : '<c:calendar-home-set/>';
    // RFC 6764: the well-known address first, when what was typed is the server's root.
    let start = base.pathname && base.pathname !== '/' ? base.pathname : `/.well-known/${card ? 'carddav' : 'caldav'}`;
    let principal = null;
    let home = null;
    for (const attempt of [start, '/']) {
      try {
        const [here] = await propfind(attempt, `<d:current-user-principal/>${homeProp}<d:resourcetype/>`, 0);
        if (!here) continue;
        home = hrefIn(here.props[card ? 'addressbook-home-set' : 'calendar-home-set']);
        principal = hrefIn(here.props['current-user-principal']);
        // The address typed may be a collection itself.
        const types = here.props.resourcetype ? kids(here.props.resourcetype).map((c) => local(c.name)) : [];
        if (!home && types.includes(card ? 'addressbook' : 'calendar')) return [await collectionAt(here.href || attempt, kind)];
        if (home || principal) { start = attempt; break; }
      } catch (err) {
        if (err.status === 401 || err.status === 403) throw err;
      }
    }
    if (!home && principal) {
      const [p] = await propfind(principal, homeProp, 0);
      home = hrefIn(p?.props[card ? 'addressbook-home-set' : 'calendar-home-set']);
    }
    if (!home) return [];
    const listed = await propfind(home, '<d:resourcetype/><d:displayname/><cs:getctag/><d:sync-token/><ic:calendar-color/><c:supported-calendar-component-set/>', 1);
    const out = [];
    for (const r of listed) {
      const types = r.props.resourcetype ? kids(r.props.resourcetype).map((c) => local(c.name)) : [];
      if (!types.includes(card ? 'addressbook' : 'calendar')) continue;
      if (!card && r.props['supported-calendar-component-set']) {
        const comps = kids(r.props['supported-calendar-component-set']).map((c) => String(c.attrs.name || '').toUpperCase());
        if (comps.length && !comps.includes('VEVENT')) continue;
      }
      out.push(describe(r, kind));
    }
    return out;
  }

  function describe(r, kind) {
    const colour = r.props['calendar-color'] ? textOf(r.props['calendar-color']).trim().slice(0, 7) : null;
    return {
      href: r.href,
      kind,
      name: (r.props.displayname ? textOf(r.props.displayname).trim() : '') || decodeURIComponent(r.href.replace(/\/$/, '').split('/').pop() || '') || (kind === 'carddav' ? 'Contacts' : 'Calendar'),
      colour: /^#[0-9a-f]{6}$/i.test(colour || '') ? colour : null,
      ctag: r.props.getctag ? textOf(r.props.getctag).trim() : r.props['sync-token'] ? textOf(r.props['sync-token']).trim() : null,
    };
  }

  async function collectionAt(href, kind) {
    const [r] = await propfind(href, '<d:resourcetype/><d:displayname/><cs:getctag/><d:sync-token/><ic:calendar-color/>', 0);
    return describe({ ...r, href: r.href || href }, kind);
  }

  /** The collection's change tag now — what says whether anything in it changed. */
  async function ctag(href) {
    const [r] = await propfind(href, '<cs:getctag/><d:sync-token/>', 0);
    if (!r) return null;
    return r.props.getctag ? textOf(r.props.getctag).trim() : r.props['sync-token'] ? textOf(r.props['sync-token']).trim() : null;
  }

  /** Every item in a collection: href to ETag. */
  async function etags(href, kind) {
    const card = kind === 'carddav';
    const body = card
      ? `<?xml version="1.0" encoding="utf-8"?><card:addressbook-query ${NS}><d:prop><d:getetag/></d:prop></card:addressbook-query>`
      : `<?xml version="1.0" encoding="utf-8"?><c:calendar-query ${NS}><d:prop><d:getetag/></d:prop><c:filter><c:comp-filter name="VCALENDAR"><c:comp-filter name="VEVENT"/></c:comp-filter></c:filter></c:calendar-query>`;
    let listed;
    const res = await request('REPORT', href, { depth: 1, body });
    if (res.status === 207) listed = readMultistatus(res.text);
    else listed = await propfind(href, '<d:getetag/><d:resourcetype/>', 1);
    const out = new Map();
    const self = new URL(href, base).pathname.replace(/\/$/, '');
    for (const r of listed) {
      const path = new URL(r.href, base).pathname;
      if (path.replace(/\/$/, '') === self || r.missing) continue;
      if (r.props.resourcetype && kids(r.props.resourcetype).length) continue;
      out.set(path, etagOf(r.props.getetag));
    }
    return out;
  }

  /** Some items' text and ETags, fetched in one request. */
  async function multiget(href, hrefs, kind) {
    if (!hrefs.length) return [];
    const card = kind === 'carddav';
    const list = hrefs.map((h) => `<d:href>${escapeXml(encodeURI(h))}</d:href>`).join('');
    const body = card
      ? `<?xml version="1.0" encoding="utf-8"?><card:addressbook-multiget ${NS}><d:prop><d:getetag/><card:address-data/></d:prop>${list}</card:addressbook-multiget>`
      : `<?xml version="1.0" encoding="utf-8"?><c:calendar-multiget ${NS}><d:prop><d:getetag/><c:calendar-data/></d:prop>${list}</c:calendar-multiget>`;
    const res = await request('REPORT', href, { depth: 1, body });
    if (res.status !== 207) {
      // A server without multiget: one GET each.
      const out = [];
      for (const h of hrefs) {
        const one = await request('GET', h);
        if (one.status === 200) out.push({ href: h, etag: one.headers.get('etag'), data: one.text });
      }
      return out;
    }
    return readMultistatus(res.text)
      .filter((r) => !r.missing)
      .map((r) => ({ href: new URL(r.href, base).pathname, etag: etagOf(r.props.getetag), data: textOf(r.props[card ? 'address-data' : 'calendar-data']) }))
      .filter((r) => r.data);
  }

  /**
   * Write one item. A new one (no `etag`) is refused if something is already
   * there; a changed one if the server's has changed since. Answers the new
   * ETag, or null when the server did not say.
   */
  async function put(href, data, { etag = null, kind = 'caldav' } = {}) {
    const type = kind === 'carddav' ? 'text/vcard; charset=utf-8' : 'text/calendar; charset=utf-8';
    const res = await request('PUT', href, { body: data, type, ifMatch: etag, ifNoneMatch: etag ? null : '*' });
    if (res.status === 412) {
      const err = new Error('changed on the server');
      err.conflict = true;
      throw err;
    }
    if (res.status < 200 || res.status >= 300) throw new Error(`The server refused to keep ${href}: ${res.status}`);
    return res.headers.get('etag');
  }

  /** Remove one item, unless it changed on the server since `etag`. Gone already is fine. */
  async function remove(href, etag = null) {
    const res = await request('DELETE', href, { ifMatch: etag });
    if (res.status === 412) {
      const err = new Error('changed on the server');
      err.conflict = true;
      throw err;
    }
    if (res.status !== 404 && (res.status < 200 || res.status >= 300)) throw new Error(`The server refused to remove ${href}: ${res.status}`);
  }

  return { discover, ctag, etags, multiget, put, remove, collectionAt, origin: () => base.origin, request };
}
