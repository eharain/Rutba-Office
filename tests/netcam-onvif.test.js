// Finding cameras on the network and asking one for its stream address
// (packages/netcam/src/onvif.js and xml.js). The camera is the tests' own, on
// this computer's loopback (tests/onvif-camera.js): the probe is sent straight
// to it and never to the discovery group, so nothing leaves the computer and
// no real camera is asked anything.

import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import { parseXml, find, findAll, textOf, child, escapeXml, unescapeXml } from '../packages/netcam/src/xml.js';
import { DISCOVERY, chooseProfile, discover, isLocalAddress, parseProbeMatches, probeMessage, readCameraService, usernameToken } from '../packages/netcam/src/onvif.js';
import { startOnvifCamera } from './onvif-camera.js';

const code = async (promise) => {
  try {
    await promise;
    return 'no error';
  } catch (e) {
    return e.code || `not a camera error: ${e.message}`;
  }
};

test('the XML reader reads elements by local name and expands nothing', () => {
  const doc = parseXml('<?xml version="1.0"?><!-- a note --><s:Envelope xmlns:s="x" xmlns:t="y"><s:Body><t:Reply token="a&amp;b" t:kind=\'one\'><t:Name>Front &lt;door&gt; &#65;&#x42;</t:Name><t:Empty/><t:Raw><![CDATA[<kept & as is>]]></t:Raw></t:Reply><t:Reply token="2"/></s:Body></s:Envelope>');
  assert.equal(doc.name, 'Envelope');
  const reply = find(doc, 'Reply');
  assert.deepEqual(reply.attrs, { token: 'a&b', kind: 'one' });
  assert.equal(textOf(doc, 'Name'), 'Front <door> AB');
  assert.equal(textOf(doc, 'Raw'), '<kept & as is>');
  assert.equal(textOf(doc, 'Empty'), '');
  assert.equal(textOf(doc, 'Nowhere'), '');
  assert.equal(findAll(doc, 'Reply').length, 2);
  assert.equal(child(doc, 'Body', 'Reply', 'Name').text, 'Front <door> AB');
  assert.equal(child(doc, 'Body', 'Name'), null);
  assert.equal(find(null, 'x'), null);

  // A document type is never read: an answer that carries one is refused whole.
  assert.equal(parseXml('<!DOCTYPE a [<!ENTITY x SYSTEM "file:///c:/windows/win.ini">]><a>&x;</a>'), null);
  assert.equal(parseXml('<a>&unknown;</a>').text, '&unknown;');
  for (const bad of ['', 'not xml', '<a>', '<a></b>', '<a/><b/>', '<a><b></a></b>', '<a><!-- open', `<a>${'x'.repeat(1024 * 1024)}</a>`, '<'.repeat(10)]) assert.equal(parseXml(bad), null, bad.slice(0, 20));
  assert.equal(parseXml(`${'<a>'.repeat(80)}${'</a>'.repeat(80)}`), null, 'too deep');
  assert.equal(parseXml(Buffer.from('<a>bytes</a>')).text, 'bytes');

  assert.equal(escapeXml('a<b>&"c\''), 'a&lt;b&gt;&amp;&quot;c&apos;');
  assert.equal(unescapeXml('&#xD800;&#0;&#1114112;x'), 'x');
});

test('a probe answer names a camera only at the address it came from', () => {
  const probe = probeMessage('uuid:1234');
  const sent = parseXml(probe);
  assert.equal(textOf(sent, 'MessageID'), 'uuid:1234');
  assert.equal(textOf(sent, 'Types'), 'dn:NetworkVideoTransmitter');
  assert.match(textOf(sent, 'Action'), /discovery\/Probe$/);
  assert.deepEqual(DISCOVERY, { address: '239.255.255.250', port: 3702 });

  const answer = (xaddrs, { relates = 'uuid:1234', scopes = 'onvif://www.onvif.org/name/Back%20yard onvif://www.onvif.org/hardware/DS-2CD2', types = 'dn:NetworkVideoTransmitter tds:Device' } = {}) => `<e:Envelope xmlns:e="s" xmlns:w="a" xmlns:d="d"><e:Header><w:RelatesTo>${relates}</w:RelatesTo></e:Header><e:Body><d:ProbeMatches><d:ProbeMatch><w:EndpointReference><w:Address>urn:uuid:cam-1</w:Address></w:EndpointReference><d:Types>${types}</d:Types><d:Scopes>${scopes}</d:Scopes><d:XAddrs>${xaddrs}</d:XAddrs></d:ProbeMatch></d:ProbeMatches></e:Body></e:Envelope>`;
  const from = '192.168.1.64';
  assert.deepEqual(parseProbeMatches(answer('http://192.168.1.64/onvif/device_service'), { from, messageId: 'uuid:1234' }), [
    { key: 'urn:uuid:cam-1', name: 'Back yard', hardware: 'DS-2CD2', host: from, service: 'http://192.168.1.64/onvif/device_service' },
  ]);
  // Of several addresses, the one on the camera itself; the plain one before the secure one.
  const several = parseProbeMatches(answer('http://[fe80::1]/onvif/device_service https://192.168.1.64/onvif/device_service http://192.168.1.64:8000/onvif/device_service http://10.0.0.9/steal'), { from, messageId: 'uuid:1234' });
  assert.equal(several[0].service, 'http://192.168.1.64:8000/onvif/device_service');
  // An answer that points only somewhere else is not a camera: a sign-in must never be sent where it says.
  assert.deepEqual(parseProbeMatches(answer('http://10.0.0.9/onvif/device_service http://evil.example/onvif'), { from, messageId: 'uuid:1234' }), []);
  assert.deepEqual(parseProbeMatches(answer('http://admin:x@192.168.1.64/onvif/device_service ftp://192.168.1.64/x not-an-address'), { from, messageId: 'uuid:1234' }), []);
  // An answer to some other probe, or to none, is not heard.
  assert.deepEqual(parseProbeMatches(answer('http://192.168.1.64/onvif/device_service', { relates: 'uuid:other' }), { from, messageId: 'uuid:1234' }), []);
  assert.deepEqual(parseProbeMatches(answer('http://192.168.1.64/onvif/device_service', { types: 'wsdp:Printer' }), { from, messageId: 'uuid:1234' }), []);
  assert.deepEqual(parseProbeMatches('rubbish', { from, messageId: 'uuid:1234' }), []);
  // A name is tidied: no line breaks, no more than a line's worth; a broken escape is kept as written.
  const odd = parseProbeMatches(answer('http://192.168.1.64/onvif/device_service', { scopes: `onvif://www.onvif.org/name/${encodeURIComponent('Gate\n\tcamera ')}${'x'.repeat(200)} onvif://www.onvif.org/hardware/%E0%A4%A` }), { from, messageId: 'uuid:1234' });
  assert.equal(odd[0].name.length, 80);
  assert.ok(odd[0].name.startsWith('Gate camera x'));
  assert.equal(odd[0].hardware, '%E0%A4%A');

  for (const a of ['10.1.2.3', '172.16.0.1', '172.31.255.254', '192.168.0.10', '169.254.3.4']) assert.equal(isLocalAddress(a), true, a);
  for (const a of ['8.8.8.8', '172.32.0.1', '192.169.0.1', '127.0.0.1', '999.1.1.1', 'fe80::1', '', null, '10.0.0']) assert.equal(isLocalAddress(a), false, String(a));
  assert.equal(isLocalAddress('127.0.0.1', { loopback: true }), true);
});

test('the user name token is the digest ONVIF lays down, never the password', () => {
  // The worked example of the ONVIF application programmer's guide.
  const nonce = Buffer.from('LKqI6G/AikKCQrN0zqZFlg==', 'base64');
  const token = usernameToken({ user: 'admin', password: 'userpassword', created: new Date('2010-09-16T07:50:45Z'), nonce });
  const doc = parseXml(`<e:Envelope xmlns:e="s">${token}</e:Envelope>`);
  assert.equal(textOf(doc, 'Username'), 'admin');
  assert.equal(textOf(doc, 'Password'), 'tuOSpGlFlIXsozq4HFNeeGeFLEI=');
  assert.equal(textOf(doc, 'Nonce'), 'LKqI6G/AikKCQrN0zqZFlg==');
  assert.equal(textOf(doc, 'Created'), '2010-09-16T07:50:45Z');
  assert.match(find(doc, 'Password').attrs.Type, /#PasswordDigest$/);
  assert.equal(token.includes('userpassword'), false);
  const odd = usernameToken({ user: 'a<b>&', password: 'p' });
  assert.equal(textOf(parseXml(`<e:Envelope xmlns:e="s">${odd}</e:Envelope>`), 'Username'), 'a<b>&');
});

test('of the streams a camera offers, the largest H.264 picture up to full HD is read', () => {
  const p = (token, encoding, width, height) => ({ token, encoding, width, height });
  assert.equal(chooseProfile([p('4k', 'H264', 3840, 2160), p('hd', 'H264', 1920, 1080), p('sd', 'H264', 640, 360)]).token, 'hd');
  assert.equal(chooseProfile([p('4k', 'H264', 3840, 2160), p('qhd', 'H264', 2560, 1440)]).token, 'qhd');
  assert.equal(chooseProfile([p('hevc', 'H265', 1920, 1080), p('sd', 'H264', 640, 360), p('jpeg', 'JPEG', 1280, 720)]).token, 'sd');
  assert.equal(chooseProfile([p('hevc', 'H265', 1920, 1080), p('none', '', 0, 0)]), null);
  assert.equal(chooseProfile([]), null);
  assert.equal(chooseProfile(null), null);
});

test('cameras that answer the probe are found, by name', async () => {
  const front = await startOnvifCamera({ name: 'Front door', hardware: 'IPC-T1' });
  const yard = await startOnvifCamera({ name: 'Back yard', hardware: 'DS-2CD2', extraXAddrs: ['http://10.9.9.9/onvif/device_service'] });
  const quiet = await startOnvifCamera({ silent: true });
  const stray = await startOnvifCamera({ name: 'Stray', relatesTo: false });
  const printer = await startOnvifCamera({ name: 'Printer', types: 'wsdp:Device pr:Printer' });
  try {
    const targets = [front, yard, quiet, stray, printer].map((c) => c.target);
    const found = await discover({ ms: 500, targets, interfaces: ['127.0.0.1'], loopback: true });
    assert.deepEqual(found.map((c) => c.name).sort(), ['Back yard', 'Front door']);
    const got = found.find((c) => c.name === 'Back yard');
    assert.deepEqual(got, { key: yard.urn, name: 'Back yard', hardware: 'DS-2CD2', host: '127.0.0.1', service: yard.service });
    // The probe is sent twice, and a camera that answers both is listed once.
    assert.equal(front.probes.length, 2);
    assert.match(front.probes[0], /NetworkVideoTransmitter/);

    // An answer from an address that is not the local network's own is not heard: here, the loopback when it is not allowed.
    assert.deepEqual(await discover({ ms: 300, targets: [front.target], interfaces: ['127.0.0.1'] }), []);
    // A computer on no network finds none, at once.
    const began = Date.now();
    assert.deepEqual(await discover({ ms: 3000, targets: [front.target], interfaces: [] }), []);
    assert.ok(Date.now() - began < 1000);
    // A network that cannot be sent from is left out and the rest still asked.
    const mixed = await discover({ ms: 400, targets: [front.target], interfaces: ['203.0.113.250', '127.0.0.1'], loopback: true });
    assert.deepEqual(mixed.map((c) => c.name), ['Front door']);
  } finally {
    for (const c of [front, yard, quiet, stray, printer]) await c.close();
  }
});

test('a found camera gives its stream address to the right sign-in', async () => {
  const cam = await startOnvifCamera({ user: 'admin', password: 'p@ss <word>&', streamHost: '192.168.77.5', rtspPort: 554 });
  try {
    const got = await readCameraService({ service: cam.service, user: 'admin', password: 'p@ss <word>&' });
    // The 1440p stream is larger than full HD, so the 720p one is read; the address is on the camera that answered, whatever host it named.
    assert.equal(got.profile.token, 'sub');
    assert.deepEqual([got.profile.width, got.profile.height, got.profile.encoding, got.profile.name], [1280, 720, 'H264', 'subStream']);
    assert.equal(got.url, 'rtsp://127.0.0.1/Streaming/Channels/sub?transportmode=unicast&profile=sub');
    assert.equal(got.profiles.length, 2);
    assert.deepEqual(cam.asked.map((a) => a.action), ['GetSystemDateAndTime', 'GetCapabilities', 'GetProfiles', 'GetStreamUri']);
    assert.deepEqual(cam.asked.map((a) => a.signed), [false, true, true, true], 'the time is asked without the sign-in');
    assert.deepEqual(cam.asked.map((a) => a.path), ['/onvif/device_service', '/onvif/device_service', '/onvif/media_service', '/onvif/media_service']);

    assert.equal(await code(readCameraService({ service: cam.service, user: 'admin', password: 'wrong' })), 'sign-in-refused');
    assert.equal(await code(readCameraService({ service: cam.service })), 'sign-in-needed');
  } finally {
    await cam.close();
  }
});

test('a camera whose clock is wrong is signed in to by its own time', async () => {
  const cam = await startOnvifCamera({ user: 'admin', password: 'secret', clockSkewMs: -3 * 3600 * 1000 - 17000 });
  const noTime = await startOnvifCamera({ user: 'admin', password: 'secret', noTime: true });
  try {
    const got = await readCameraService({ service: cam.service, user: 'admin', password: 'secret' });
    assert.equal(got.profile.token, 'sub');
    // One that will not tell its time is asked the rest by this computer's.
    assert.equal((await readCameraService({ service: noTime.service, user: 'admin', password: 'secret' })).profile.token, 'sub');
  } finally {
    await cam.close();
    await noTime.close();
  }
});

test('cameras that ask in other ways, and what they offer', async () => {
  const digest = await startOnvifCamera({ user: 'admin', password: 'secret', signIn: 'http-digest', rtspPort: 8554, profiles: [{ token: 'only', name: 'Stream 1', encoding: 'H264', width: 1920, height: 1080 }] });
  const basic = await startOnvifCamera({ user: 'admin', password: 'secret', signIn: 'http-basic' });
  const open = await startOnvifCamera({ noCapabilities: true, mediaHost: '10.20.30.40' });
  const hevc = await startOnvifCamera({ profiles: [{ token: 'h', name: 'main', encoding: 'H265', width: 3840, height: 2160 }] });
  const jpeg = await startOnvifCamera({ profiles: [{ token: 'j', name: 'main', encoding: 'JPEG', width: 640, height: 480 }, { token: 'n', name: 'no video', encoding: '', width: 0, height: 0 }] });
  try {
    // HTTP Digest: asked once, answered, and kept for the questions after.
    const got = await readCameraService({ service: digest.service, user: 'admin', password: 'secret' });
    assert.equal(got.url, 'rtsp://127.0.0.1:8554/Streaming/Channels/only?transportmode=unicast&profile=only');
    assert.deepEqual(digest.asked.map((a) => a.action), ['GetSystemDateAndTime', 'GetCapabilities', 'GetCapabilities', 'GetProfiles', 'GetStreamUri']);
    assert.deepEqual(digest.asked.map((a) => Boolean(a.authorization)), [false, false, true, true, true]);
    assert.equal(await code(readCameraService({ service: digest.service, user: 'admin', password: 'no' })), 'sign-in-refused');
    assert.equal(await code(readCameraService({ service: digest.service })), 'sign-in-needed');

    // The password itself is never sent over a plain connection, though the camera asks for just that.
    assert.equal(await code(readCameraService({ service: basic.service, user: 'admin', password: 'secret' })), 'sign-in-refused');
    assert.equal(basic.asked.some((a) => /^basic/i.test(a.authorization)), false);

    // No sign-in at all, GetServices in place of GetCapabilities, and a media address naming another host: still asked on the camera itself.
    const plain = await readCameraService({ service: open.service });
    assert.equal(plain.profile.token, 'sub');
    assert.deepEqual(open.asked.map((a) => a.action), ['GetSystemDateAndTime', 'GetCapabilities', 'GetServices', 'GetProfiles', 'GetStreamUri']);
    assert.equal(open.asked.some((a) => a.signed), false);

    assert.equal(await code(readCameraService({ service: hevc.service })), 'h265');
    assert.equal(await code(readCameraService({ service: jpeg.service })), 'no-video');
  } finally {
    for (const c of [digest, basic, open, hevc, jpeg]) await c.close();
  }
});

test('what is not a camera service is said so, and nothing hangs', async () => {
  const cam = await startOnvifCamera();
  const port = cam.httpPort;
  await cam.close();
  assert.equal(await code(readCameraService({ service: `http://127.0.0.1:${port}/onvif/device_service`, timeoutMs: 1500 })), 'unreachable');
  assert.equal(await code(readCameraService({ service: 'not an address' })), 'address');

  // Stand-ins for the camera's answers: not XML, a stream address that is not one, and a camera that never answers.
  const reply = (body) => ({ status: 200, headers: {}, text: `<e:Envelope xmlns:e="s"><e:Body>${body}</e:Body></e:Envelope>` });
  const script = (answers) => {
    const asked = [];
    return { asked, post: async ({ url, body }) => { asked.push(url); const action = /<e:Body><\w+:(\w+)/.exec(body)[1]; return answers[action] ?? { status: 200, headers: {}, text: 'hello' }; } };
  };
  const service = 'http://192.168.1.64/onvif/device_service';
  assert.equal(await code(readCameraService({ service, post: script({}).post })), 'protocol');
  const media = reply('<r><Capabilities><Media><XAddr>http://192.168.1.64/onvif/Media</XAddr></Media></Capabilities></r>');
  const profiles = reply('<r><Profiles token="p1"><Name>one</Name><VideoEncoderConfiguration><Encoding>H264</Encoding><Resolution><Width>1280</Width><Height>720</Height></Resolution></VideoEncoderConfiguration></Profiles></r>');
  assert.equal(await code(readCameraService({ service, post: script({ GetCapabilities: media, GetProfiles: profiles, GetStreamUri: reply('<r><MediaUri><Uri>http://192.168.1.64/video.mjpg</Uri></MediaUri></r>') }).post })), 'protocol');
  assert.equal(await code(readCameraService({ service, post: script({ GetCapabilities: media, GetProfiles: profiles }).post })), 'protocol');
  assert.equal(await code(readCameraService({ service, post: script({ GetCapabilities: reply('<r><Capabilities><Media><XAddr>ftp://x/</XAddr></Media></Capabilities></r>'), GetServices: reply('<r/>') }).post })), 'protocol');

  // A stream address with a sign-in in it, on another host and the usual port: read as on the camera, with no sign-in kept.
  const s = script({ GetCapabilities: reply('<r><Capabilities><Media><XAddr>http://10.0.0.1:8899/onvif/Media</XAddr></Media></Capabilities></r>'), GetProfiles: profiles, GetStreamUri: reply('<r><MediaUri><Uri>rtsp://admin:hunter2@10.0.0.1:554/live/ch0</Uri></MediaUri></r>') });
  const got = await readCameraService({ service, post: s.post });
  assert.equal(got.url, 'rtsp://192.168.1.64/live/ch0');
  assert.deepEqual(s.asked.slice(-2), ['http://192.168.1.64:8899/onvif/Media', 'http://192.168.1.64:8899/onvif/Media']);

  // A token is stamped with the camera's time: here a camera an hour ahead, and a made-up now.
  const stamps = [];
  const timed = { GetSystemDateAndTime: reply('<r><UTCDateTime><Time><Hour>13</Hour><Minute>0</Minute><Second>5</Second></Time><Date><Year>2026</Year><Month>11</Month><Day>22</Day></Date></UTCDateTime></r>'), GetCapabilities: media, GetProfiles: profiles, GetStreamUri: reply('<r><MediaUri><Uri>rtsp://192.168.1.64/1</Uri></MediaUri></r>') };
  await readCameraService({ service, user: 'u', password: 'p', now: () => Date.UTC(2026, 10, 22, 12, 0, 0), post: async (o) => { const m = /Created>([^<]+)</.exec(o.body); if (m) stamps.push(m[1]); return script(timed).post(o); } });
  assert.deepEqual(stamps, ['2026-11-22T13:00:05Z', '2026-11-22T13:00:05Z', '2026-11-22T13:00:05Z']);
  assert.ok(crypto.randomUUID());
});
