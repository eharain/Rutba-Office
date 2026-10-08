/**
 * A CalDAV and CardDAV server small enough to read, for the sync's tests and
 * window checks: one user, calendars and address books held in memory, and
 * the requests a client makes answered as a real server answers them — the
 * well-known redirect, the principal and its homes, PROPFIND at depth 0 and
 * 1, the query and multiget REPORTs, GET, and PUT and DELETE with If-Match
 * and If-None-Match. Every change bumps its collection's change tag.
 *
 *   const server = await startDavServer({ user: 'ann', password: 'secret' });
 *   server.url; server.calendar('/dav/calendars/ann/work/').items; await server.close();
 */
import http from 'node:http';
import crypto from 'node:crypto';

const xml = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

export async function startDavServer({ user = 'ann', password = 'secret', calendars = [{ name: 'Work', slug: 'work', colour: '#1a9f7a' }], books = [{ name: 'Friends', slug: 'friends' }] } = {}) {
  const collections = new Map();
  const home = { cal: `/dav/calendars/${user}/`, card: `/dav/addressbooks/${user}/` };
  const principal = `/dav/principals/${user}/`;
  for (const c of calendars) collections.set(`${home.cal}${c.slug}/`, { kind: 'cal', name: c.name, colour: c.colour || null, ctag: 1, items: new Map() });
  for (const b of books) collections.set(`${home.card}${b.slug}/`, { kind: 'card', name: b.name, ctag: 1, items: new Map() });
  const log = [];

  const etag = () => `"${crypto.randomUUID().slice(0, 8)}"`;
  const collectionOf = (path) => collections.get(path.replace(/[^/]*$/, ''));
  const response = (href, props, status = 'HTTP/1.1 200 OK') => `<d:response><d:href>${xml(encodeURI(href))}</d:href><d:propstat><d:prop>${props}</d:prop><d:status>${status}</d:status></d:propstat></d:response>`;
  const multistatus = (inner) => `<?xml version="1.0" encoding="utf-8"?><d:multistatus xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav" xmlns:card="urn:ietf:params:xml:ns:carddav" xmlns:cs="http://calendarserver.org/ns/" xmlns:ic="http://apple.com/ns/ical/">${inner}</d:multistatus>`;
  const describe = (path, c) => `<d:resourcetype><d:collection/>${c.kind === 'cal' ? '<c:calendar/>' : '<card:addressbook/>'}</d:resourcetype><d:displayname>${xml(c.name)}</d:displayname><cs:getctag>${c.ctag}</cs:getctag>${c.colour ? `<ic:calendar-color>${c.colour}FF</ic:calendar-color>` : ''}${c.kind === 'cal' ? '<c:supported-calendar-component-set><c:comp name="VEVENT"/></c:supported-calendar-component-set>' : ''}`;

  const server = http.createServer(async (req, res) => {
    const body = await new Promise((resolve) => { let s = ''; req.on('data', (d) => { s += d; }); req.on('end', () => resolve(s)); });
    const path = decodeURI(new URL(req.url, 'http://x').pathname);
    log.push({ method: req.method, path });
    const send = (status, text = '', headers = {}) => { res.writeHead(status, headers); res.end(text); };
    const expected = `Basic ${Buffer.from(`${user}:${password}`).toString('base64')}`;
    if (req.headers.authorization !== expected) return send(401, 'no', { 'WWW-Authenticate': 'Basic realm="dav"' });
    if (path === '/.well-known/caldav' || path === '/.well-known/carddav') return send(301, '', { Location: '/dav/' });
    const depth = req.headers.depth;
    if (req.method === 'PROPFIND') {
      if (path === '/dav/' || path === principal) {
        return send(207, multistatus(response(path, `<d:current-user-principal><d:href>${principal}</d:href></d:current-user-principal><c:calendar-home-set><d:href>${home.cal}</d:href></c:calendar-home-set><card:addressbook-home-set><d:href>${home.card}</d:href></card:addressbook-home-set><d:resourcetype><d:collection/></d:resourcetype>`)), { 'Content-Type': 'application/xml' });
      }
      if (path === home.cal || path === home.card) {
        let inner = response(path, '<d:resourcetype><d:collection/></d:resourcetype>');
        if (depth === '1') for (const [p, c] of collections) if (p.startsWith(path)) inner += response(p, describe(p, c));
        return send(207, multistatus(inner), { 'Content-Type': 'application/xml' });
      }
      const c = collections.get(path);
      if (c) {
        let inner = response(path, describe(path, c) + `<d:getetag>"c${c.ctag}"</d:getetag>`);
        if (depth === '1') for (const [p, item] of c.items) inner += response(p, `<d:getetag>${item.etag}</d:getetag><d:resourcetype/>`);
        return send(207, multistatus(inner), { 'Content-Type': 'application/xml' });
      }
      return send(404);
    }
    if (req.method === 'REPORT') {
      const c = collections.get(path);
      if (!c) return send(404);
      if (/multiget/.test(body)) {
        const hrefs = [...body.matchAll(/<d:href>([^<]+)<\/d:href>/g)].map((m) => decodeURI(m[1]));
        const tag = c.kind === 'cal' ? 'c:calendar-data' : 'card:address-data';
        const inner = hrefs.map((h) => { const item = c.items.get(h); return item ? response(h, `<d:getetag>${item.etag}</d:getetag><${tag}>${xml(item.data)}</${tag}>`) : `<d:response><d:href>${xml(h)}</d:href><d:status>HTTP/1.1 404 Not Found</d:status></d:response>`; }).join('');
        return send(207, multistatus(inner), { 'Content-Type': 'application/xml' });
      }
      const inner = [...c.items].map(([p, item]) => response(p, `<d:getetag>${item.etag}</d:getetag>`)).join('');
      return send(207, multistatus(inner), { 'Content-Type': 'application/xml' });
    }
    if (req.method === 'GET') {
      const item = collectionOf(path)?.items.get(path);
      return item ? send(200, item.data, { ETag: item.etag }) : send(404);
    }
    if (req.method === 'PUT') {
      const c = collectionOf(path);
      if (!c) return send(409);
      const item = c.items.get(path);
      if (req.headers['if-none-match'] === '*' && item) return send(412);
      if (req.headers['if-match'] && (!item || item.etag !== req.headers['if-match'])) return send(412);
      const tag = etag();
      c.items.set(path, { data: body, etag: tag });
      c.ctag += 1;
      return send(item ? 204 : 201, '', { ETag: tag });
    }
    if (req.method === 'DELETE') {
      const c = collectionOf(path);
      const item = c?.items.get(path);
      if (!item) return send(404);
      if (req.headers['if-match'] && item.etag !== req.headers['if-match']) return send(412);
      c.items.delete(path);
      c.ctag += 1;
      return send(204);
    }
    return send(405);
  });
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  const { port } = server.address();
  return {
    url: `http://127.0.0.1:${port}`,
    log,
    collections,
    calendarPath: (slug) => `${home.cal}${slug}/`,
    bookPath: (slug) => `${home.card}${slug}/`,
    /** Put an item straight in, as another device would. */
    put(path, data) {
      const c = collectionOf(path);
      const tag = etag();
      c.items.set(path, { data, etag: tag });
      c.ctag += 1;
      return tag;
    },
    remove(path) { const c = collectionOf(path); c.items.delete(path); c.ctag += 1; },
    close: () => new Promise((resolve) => server.close(resolve)),
  };
}
