// An ONVIF camera for the tests, on this computer's loopback: it answers a
// WS-Discovery probe sent straight to its own UDP port (nothing is sent to the
// discovery group and nothing leaves the computer), and serves the device and
// media services Office asks for a stream address. Not a test itself.

import crypto from 'node:crypto';
import dgram from 'node:dgram';
import http from 'node:http';

const md5 = (text) => crypto.createHash('md5').update(text).digest('hex');
const esc = (text) => String(text).replace(/[<>&"]/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;' })[c]);
const between = (text, name) => {
  const m = new RegExp(`<(?:[A-Za-z0-9_]+:)?${name}(?:\\s[^>]*)?>([^<]*)<`).exec(text);
  return m ? m[1].trim() : '';
};

/**
 * Starts the camera. Options:
 *   name, hardware   what it calls itself in its answer to a probe
 *   user, password   its sign-in; with none, everything is answered unsigned
 *   signIn           'token' (the WS-Security user name token, default), 'http-digest', 'http-basic'
 *   clockSkewMs      how far its clock is from this computer's (a token stamped more than 5 s off its clock is refused)
 *   noTime           true: it will not tell its time
 *   noCapabilities   true: GetCapabilities is not known, GetServices is
 *   profiles         [{ token, name, encoding, width, height }]
 *   streamHost       the host its stream address names (default 127.0.0.1); cameras name all sorts
 *   rtspPort         the port in its stream address (554)
 *   mediaHost        the host its media address names
 *   extraXAddrs      more service addresses its probe answer names, before its own
 *   relatesTo        false: its probe answer does not say which probe it answers
 *   types            the Types of its probe answer
 *   silent           true: it answers no probe
 */
export async function startOnvifCamera(options = {}) {
  const opt = {
    name: 'Front door',
    hardware: 'IPC-T1',
    user: '',
    password: '',
    signIn: 'token',
    clockSkewMs: 0,
    noTime: false,
    noCapabilities: false,
    profiles: [
      { token: 'main', name: 'mainStream', encoding: 'H264', width: 2560, height: 1440 },
      { token: 'sub', name: 'subStream', encoding: 'H264', width: 1280, height: 720 },
    ],
    streamHost: '127.0.0.1',
    rtspPort: 554,
    mediaHost: '127.0.0.1',
    extraXAddrs: [],
    relatesTo: true,
    types: 'dn:NetworkVideoTransmitter',
    silent: false,
    ...options,
  };
  const asked = [];
  const probes = [];
  const nonce = crypto.randomBytes(8).toString('hex');
  const urn = `urn:uuid:${crypto.randomUUID()}`;
  let httpPort = 0;

  const soap = (body) => `<?xml version="1.0" encoding="UTF-8"?><SOAP-ENV:Envelope xmlns:SOAP-ENV="http://www.w3.org/2003/05/soap-envelope" xmlns:tds="http://www.onvif.org/ver10/device/wsdl" xmlns:trt="http://www.onvif.org/ver10/media/wsdl" xmlns:tt="http://www.onvif.org/ver10/schema"><SOAP-ENV:Body>${body}</SOAP-ENV:Body></SOAP-ENV:Envelope>`;
  const fault = (subcode, reason) => soap(`<SOAP-ENV:Fault><SOAP-ENV:Code><SOAP-ENV:Value>SOAP-ENV:Sender</SOAP-ENV:Value><SOAP-ENV:Subcode><SOAP-ENV:Value>${subcode}</SOAP-ENV:Value></SOAP-ENV:Subcode></SOAP-ENV:Code><SOAP-ENV:Reason><SOAP-ENV:Text xml:lang="en">${reason}</SOAP-ENV:Text></SOAP-ENV:Reason></SOAP-ENV:Fault>`);

  /** Whether a request carries a good user name token, by the camera's own clock. */
  const tokenGood = (text) => {
    const user = between(text, 'Username');
    const digest = between(text, 'Password');
    const n = between(text, 'Nonce');
    const created = between(text, 'Created');
    if (!user || !digest || !n || !created) return false;
    if (Math.abs(Date.parse(created) - (Date.now() + opt.clockSkewMs)) > 5000) return false;
    const want = crypto.createHash('sha1').update(Buffer.concat([Buffer.from(n, 'base64'), Buffer.from(created, 'utf8'), Buffer.from(opt.password, 'utf8')])).digest('base64');
    return user === opt.user && digest === want;
  };

  const httpGood = (req) => {
    const line = req.headers.authorization || '';
    if (opt.signIn === 'http-basic') return line === `Basic ${Buffer.from(`${opt.user}:${opt.password}`).toString('base64')}`;
    if (!/^digest /i.test(line)) return false;
    const p = {};
    const re = /([a-z0-9_-]+)\s*=\s*(?:"([^"]*)"|([^\s,]*))/gi;
    for (let m = re.exec(line); m; m = re.exec(line)) p[m[1].toLowerCase()] = m[2] ?? m[3];
    const ha1 = md5(`${opt.user}:onvif:${opt.password}`);
    const ha2 = md5(`POST:${p.uri}`);
    return p.username === opt.user && p.nonce === nonce && p.uri === req.url && p.response === md5(`${ha1}:${nonce}:${p.nc}:${p.cnonce}:auth:${ha2}`);
  };

  const server = http.createServer((req, res) => {
    const parts = [];
    req.on('data', (c) => parts.push(c));
    req.on('end', () => {
      const text = Buffer.concat(parts).toString('utf8');
      const action = /<(?:[A-Za-z0-9_]+:)?Body[^>]*>\s*<(?:[A-Za-z0-9_]+:)?([A-Za-z]+)/.exec(text)?.[1] || '';
      asked.push({ path: req.url, action, signed: /UsernameToken/.test(text), authorization: req.headers.authorization || '' });
      const answer = (status, body, headers = {}) => {
        res.writeHead(status, { 'Content-Type': 'application/soap+xml; charset=utf-8', ...headers });
        res.end(body);
      };
      if (action === 'GetSystemDateAndTime') {
        if (opt.noTime) { answer(400, fault('ter:ActionNotSupported', 'Not supported')); return; }
        const t = new Date(Date.now() + opt.clockSkewMs);
        answer(200, soap(`<tds:GetSystemDateAndTimeResponse><tds:SystemDateAndTime><tt:DateTimeType>NTP</tt:DateTimeType><tt:UTCDateTime><tt:Time><tt:Hour>${t.getUTCHours()}</tt:Hour><tt:Minute>${t.getUTCMinutes()}</tt:Minute><tt:Second>${t.getUTCSeconds()}</tt:Second></tt:Time><tt:Date><tt:Year>${t.getUTCFullYear()}</tt:Year><tt:Month>${t.getUTCMonth() + 1}</tt:Month><tt:Day>${t.getUTCDate()}</tt:Day></tt:Date></tt:UTCDateTime></tds:SystemDateAndTime></tds:GetSystemDateAndTimeResponse>`));
        return;
      }
      if (opt.user) {
        if (opt.signIn === 'token' && !tokenGood(text)) { answer(400, fault('ter:NotAuthorized', 'Sender not Authorized')); return; }
        if (opt.signIn !== 'token' && !httpGood(req)) {
          answer(401, '', { 'WWW-Authenticate': opt.signIn === 'http-basic' ? 'Basic realm="onvif"' : `Digest realm="onvif", qop="auth", nonce="${nonce}"` });
          return;
        }
      }
      const mediaAddress = `http://${opt.mediaHost}:${httpPort}/onvif/media_service`;
      if (action === 'GetCapabilities') {
        if (opt.noCapabilities) { answer(400, fault('ter:ActionNotSupported', 'Not supported')); return; }
        answer(200, soap(`<tds:GetCapabilitiesResponse><tds:Capabilities><tt:Media><tt:XAddr>${mediaAddress}</tt:XAddr><tt:StreamingCapabilities><tt:RTP_RTSP_TCP>true</tt:RTP_RTSP_TCP></tt:StreamingCapabilities></tt:Media></tds:Capabilities></tds:GetCapabilitiesResponse>`));
        return;
      }
      if (action === 'GetServices') {
        answer(200, soap(`<tds:GetServicesResponse><tds:Service><tds:Namespace>http://www.onvif.org/ver10/device/wsdl</tds:Namespace><tds:XAddr>http://127.0.0.1:${httpPort}/onvif/device_service</tds:XAddr></tds:Service><tds:Service><tds:Namespace>http://www.onvif.org/ver10/media/wsdl</tds:Namespace><tds:XAddr>${mediaAddress}</tds:XAddr></tds:Service></tds:GetServicesResponse>`));
        return;
      }
      if (action === 'GetProfiles' && req.url === '/onvif/media_service') {
        answer(200, soap(`<trt:GetProfilesResponse>${opt.profiles.map((p) => `<trt:Profiles token="${esc(p.token)}" fixed="true"><tt:Name>${esc(p.name)}</tt:Name><tt:VideoSourceConfiguration token="vs"><tt:Name>source</tt:Name><tt:Bounds x="0" y="0" width="${p.width}" height="${p.height}"/></tt:VideoSourceConfiguration>${p.encoding ? `<tt:VideoEncoderConfiguration token="ve-${esc(p.token)}"><tt:Name>${esc(p.name)}</tt:Name><tt:Encoding>${p.encoding}</tt:Encoding><tt:Resolution><tt:Width>${p.width}</tt:Width><tt:Height>${p.height}</tt:Height></tt:Resolution></tt:VideoEncoderConfiguration>` : ''}</trt:Profiles>`).join('')}</trt:GetProfilesResponse>`));
        return;
      }
      if (action === 'GetStreamUri' && req.url === '/onvif/media_service') {
        const token = between(text, 'ProfileToken');
        if (!opt.profiles.some((p) => p.token === token)) { answer(400, fault('ter:NoProfile', 'No such profile')); return; }
        answer(200, soap(`<trt:GetStreamUriResponse><trt:MediaUri><tt:Uri>rtsp://${opt.streamHost}:${opt.rtspPort}/Streaming/Channels/${esc(token)}?transportmode=unicast&amp;profile=${esc(token)}</tt:Uri><tt:InvalidAfterConnect>false</tt:InvalidAfterConnect><tt:Timeout>PT60S</tt:Timeout></trt:MediaUri></trt:GetStreamUriResponse>`));
        return;
      }
      answer(400, fault('ter:ActionNotSupported', 'Not supported'));
    });
  });
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  httpPort = server.address().port;

  const udp = dgram.createSocket({ type: 'udp4' });
  udp.on('message', (bytes, rinfo) => {
    const text = bytes.toString('utf8');
    probes.push(text);
    if (opt.silent || !/Probe/.test(text)) return;
    const id = between(text, 'MessageID');
    const xaddrs = [...opt.extraXAddrs, `http://127.0.0.1:${httpPort}/onvif/device_service`].join(' ');
    const reply = `<?xml version="1.0" encoding="UTF-8"?><SOAP-ENV:Envelope xmlns:SOAP-ENV="http://www.w3.org/2003/05/soap-envelope" xmlns:wsa="http://schemas.xmlsoap.org/ws/2004/08/addressing" xmlns:d="http://schemas.xmlsoap.org/ws/2005/04/discovery" xmlns:dn="http://www.onvif.org/ver10/network/wsdl"><SOAP-ENV:Header><wsa:MessageID>urn:uuid:${crypto.randomUUID()}</wsa:MessageID>${opt.relatesTo ? `<wsa:RelatesTo>${id}</wsa:RelatesTo>` : ''}<wsa:To SOAP-ENV:mustUnderstand="true">http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</wsa:To><wsa:Action SOAP-ENV:mustUnderstand="true">http://schemas.xmlsoap.org/ws/2005/04/discovery/ProbeMatches</wsa:Action></SOAP-ENV:Header><SOAP-ENV:Body><d:ProbeMatches><d:ProbeMatch><wsa:EndpointReference><wsa:Address>${urn}</wsa:Address></wsa:EndpointReference><d:Types>${opt.types}</d:Types><d:Scopes>onvif://www.onvif.org/type/video_encoder onvif://www.onvif.org/Profile/Streaming onvif://www.onvif.org/hardware/${encodeURIComponent(opt.hardware)} onvif://www.onvif.org/name/${encodeURIComponent(opt.name)} onvif://www.onvif.org/location/</d:Scopes><d:XAddrs>${xaddrs}</d:XAddrs><d:MetadataVersion>1</d:MetadataVersion></d:ProbeMatch></d:ProbeMatches></SOAP-ENV:Body></SOAP-ENV:Envelope>`;
    udp.send(Buffer.from(reply, 'utf8'), rinfo.port, rinfo.address);
  });
  await new Promise((resolve) => udp.bind(0, '127.0.0.1', resolve));

  return {
    options: opt,
    urn,
    asked,
    probes,
    /** Where the tests send their probe: this camera alone, on the loopback. */
    target: { address: '127.0.0.1', port: udp.address().port },
    service: `http://127.0.0.1:${httpPort}/onvif/device_service`,
    httpPort,
    async close() {
      await new Promise((resolve) => udp.close(resolve));
      server.closeAllConnections?.();
      await new Promise((resolve) => server.close(resolve));
    },
  };
}
