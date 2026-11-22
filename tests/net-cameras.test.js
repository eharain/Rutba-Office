// Network cameras, the main process's half (apps/desktop/main/net-cameras.js):
// what is kept of a camera and where (the password in the keystore only, never
// handed back, never beside the address), and the reading of one: only while a
// window holds it open, shared by the windows that show it, tried again when
// it drops and not when the camera says no. The camera is the tests' own, on
// this computer's loopback (tests/netcam-server.js); there is no real one.

import test from 'node:test';
import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import { createNetCameras } from '../apps/desktop/main/net-cameras.js';
import { splitAnnexB } from '../packages/netcam/src/index.js';
import { startCamera, makeSps, makePps, makeSlice } from './netcam-server.js';

const SPS = makeSps({ width: 1280, height: 720 });
const PPS = makePps();

function fakeStores({ keystore = true } = {}) {
  const kept = new Map();
  const secret = new Map();
  return {
    kept,
    secret,
    settings: { get: (k, d) => (kept.has(k) ? structuredClone(kept.get(k)) : d), set: (k, v) => kept.set(k, structuredClone(v)) },
    secrets: { available: () => keystore, get: (k) => secret.get(k) ?? null, set: (k, v) => secret.set(k, v), delete: (k) => secret.delete(k) },
  };
}

/** A window as the service sees one: it can close, reload, and be asked whether it is gone. */
function fakeWindow() {
  const win = new EventEmitter();
  win.gone = false;
  win.isDestroyed = () => win.gone;
  win.webContents = new EventEmitter();
  win.close = () => {
    win.gone = true;
    win.emit('closed');
  };
  return win;
}

function service(options = {}) {
  const stores = fakeStores(options);
  const got = [];
  const cameras = createNetCameras({ stores, send: (win, event, payload) => got.push({ win, event, payload }), retrySeconds: [0.05], ...options.service });
  const of = (win, event) => got.filter((g) => g.win === win && g.event === event).map((g) => g.payload);
  return { stores, cameras, got, of };
}

async function until(check, ms = 4000) {
  const end = Date.now() + ms;
  while (!check()) {
    if (Date.now() > end) throw new Error('waited too long');
    await new Promise((r) => setTimeout(r, 10));
  }
}

test('a camera is kept by its address and user name, its password in the keystore only', () => {
  const { stores, cameras } = service();
  const made = cameras.cameraSave({ url: 'rtsp://admin:s3cret@192.168.1.64:554/Streaming/Channels/101' });
  assert.equal(made.ok, true);
  const { id } = made.camera;
  assert.deepEqual(made.camera, { id, name: '192.168.1.64', url: 'rtsp://192.168.1.64/Streaming/Channels/101', user: 'admin', hasPassword: true });
  assert.deepEqual(cameras.cameraList(), [made.camera]);
  assert.ok(!JSON.stringify([...stores.kept]).includes('s3cret'), 'no password in the settings, in the address or anywhere');
  assert.ok(!JSON.stringify(made).includes('s3cret'), 'no password handed back');
  assert.equal(stores.secret.get(`cameras.pass.${id}`), 's3cret');

  // The fields beside the address win over what is typed inside it; a password not given is kept.
  const named = cameras.cameraSave({ id, name: 'Porch', user: 'viewer' });
  assert.deepEqual(named.camera, { id, name: 'Porch', url: 'rtsp://192.168.1.64/Streaming/Channels/101', user: 'viewer', hasPassword: true });
  assert.equal(stores.secret.get(`cameras.pass.${id}`), 's3cret');
  cameras.cameraSave({ id, password: 'new one' });
  assert.equal(stores.secret.get(`cameras.pass.${id}`), 'new one');
  assert.equal(cameras.cameraList().length, 1, 'saved over itself');

  // A camera left with no user name keeps no password.
  const open = cameras.cameraSave({ id, user: '' });
  assert.deepEqual([open.camera.user, open.camera.hasPassword], ['', false]);
  assert.equal(stores.secret.has(`cameras.pass.${id}`), false);

  // A setup from another computer names its camera by id: the address is given here, under that id.
  const brought = cameras.cameraSave({ id: 'netcam-from-elsewhere', url: '10.0.0.9/stream1', user: 'u', password: 'p' });
  assert.equal(brought.camera.id, 'netcam-from-elsewhere');
  assert.equal(brought.camera.url, 'rtsp://10.0.0.9/stream1');
  assert.equal(cameras.cameraList().length, 2);
  assert.notEqual(cameras.cameraSave({ id: '../../x', url: '10.0.0.9/a' }).camera.id, '../../x', 'an id that is not one is not used');

  assert.deepEqual(cameras.cameraRemove({ id: 'netcam-from-elsewhere' }).map((c) => c.id).includes('netcam-from-elsewhere'), false);
  assert.equal(stores.secret.has('cameras.pass.netcam-from-elsewhere'), false, 'forgotten with its password');
});

test('what is not a camera address is refused, and a password is not kept where there is no keystore', () => {
  const { cameras } = service();
  for (const url of ['', 'https://example.org/cam', 'rtsp://', `rtsp://cam.example/${'x'.repeat(500)}`, undefined]) {
    assert.deepEqual(cameras.cameraSave({ url }), { ok: false, reason: 'address' });
  }
  assert.deepEqual(cameras.cameraList(), []);

  const bare = service({ keystore: false });
  assert.deepEqual(bare.cameras.cameraSave({ url: 'rtsp://cam.example/1', user: 'a', password: 'b' }), { ok: false, reason: 'no-keystore' });
  assert.deepEqual(bare.cameras.cameraSave({ url: 'rtsp://a:b@cam.example/1' }), { ok: false, reason: 'no-keystore' });
  assert.deepEqual(bare.cameras.cameraList(), [], 'nothing kept');
  assert.equal(bare.cameras.cameraSave({ url: 'rtsp://cam.example/1' }).ok, true, 'a camera with no sign-in needs no keystore');
  assert.deepEqual([...bare.stores.secret], []);
});

test('a camera is read only while a window holds it open, one reading shared by the windows that show it', async () => {
  const camera = await startCamera({ auth: { scheme: 'digest', algorithm: 'MD5', user: 'admin', password: 'right' } });
  const { cameras, of } = service();
  try {
    const { id } = cameras.cameraSave({ url: camera.url, user: 'admin', password: 'right' }).camera;
    assert.equal(camera.requests.length, 0, 'saving a camera does not read it');
    assert.deepEqual(cameras.camerasReading(), []);

    const first = fakeWindow();
    assert.deepEqual(cameras.cameraOpen({ id }, first), { id, status: 'connecting', code: '', width: 0, height: 0, sound: false });
    await until(() => camera.players.size === 1);
    const idr = makeSlice({ key: true, size: 3000 });
    const inter = makeSlice({ size: 500 });
    camera.send([makeSlice({ size: 200 })], { timestamp: 0 });
    camera.send([SPS, PPS, idr], { timestamp: 3000 });
    camera.send([inter], { timestamp: 6000 });
    await until(() => of(first, 'camera:frame').length === 2);
    const [key, delta] = of(first, 'camera:frame');
    assert.deepEqual([key.id, key.key, key.codec, key.width, key.height], [id, true, 'avc1.42c01f', 1280, 720], 'a key frame names its codec');
    assert.ok(key.bytes instanceof Uint8Array && key.bytes.byteOffset === 0 && key.bytes.buffer.byteLength === key.bytes.length, 'a copy of its own, sharing no memory');
    assert.deepEqual(splitAnnexB(Buffer.from(key.bytes)), [SPS, PPS, idr]);
    assert.deepEqual([delta.key, delta.codec, delta.timestamp - key.timestamp], [false, undefined, 33333]);
    assert.deepEqual(of(first, 'camera:state').map((s) => s.status), ['connecting', 'live'], 'live from its first key frame, not before');
    assert.deepEqual(of(first, 'camera:state').at(-1), { id, status: 'live', code: '', width: 1280, height: 720, sound: false });

    // A second window showing the same camera joins the one reading.
    const second = fakeWindow();
    assert.equal(cameras.cameraOpen({ id }, second).status, 'live');
    camera.send([SPS, PPS, idr], { timestamp: 9000 });
    await until(() => of(second, 'camera:frame').length === 1);
    assert.equal(of(first, 'camera:frame').length, 3);
    assert.equal(camera.players.size, 1);
    assert.equal(camera.requests.filter((r) => r.method === 'DESCRIBE' && r.headers.authorization).length, 1);
    assert.deepEqual(cameras.camerasReading().map((r) => [r.status, r.owners]), [['live', 2]]);

    // The first lets go: the reading goes on for the second. Its window closing ends it.
    assert.deepEqual(cameras.cameraClose({ id }, first), { closed: true });
    assert.deepEqual(cameras.cameraClose({ id }, first), { closed: false });
    camera.send([inter], { timestamp: 12000 });
    await until(() => of(second, 'camera:frame').length === 2);
    assert.equal(of(first, 'camera:frame').length, 3, 'nothing more to a window that let go');
    second.close();
    await until(() => camera.requests.some((r) => r.method === 'TEARDOWN'));
    await until(() => camera.players.size === 0);
    assert.deepEqual(cameras.camerasReading(), []);

    // A window that reloads lets go as one that closes does.
    const third = fakeWindow();
    cameras.cameraOpen({ id }, third);
    await until(() => camera.players.size === 1);
    third.webContents.emit('did-navigate');
    await until(() => camera.players.size === 0);
    assert.throws(() => cameras.cameraOpen({ id }), /window/);
    assert.throws(() => cameras.cameraOpen({ id: '../x' }, third), /id/);
  } finally {
    cameras.closeCameras();
    await camera.close();
  }
});

test('a camera that says no is not asked again until its details are saved again', async () => {
  const camera = await startCamera({ auth: { scheme: 'digest', algorithm: 'MD5', user: 'admin', password: 'right' }, sprop: [SPS, PPS] });
  const { cameras, of } = service();
  try {
    const { id } = cameras.cameraSave({ url: camera.url, user: 'admin', password: 'wrong' }).camera;
    const win = fakeWindow();
    cameras.cameraOpen({ id }, win);
    await until(() => of(win, 'camera:state').at(-1)?.status === 'failed');
    assert.equal(of(win, 'camera:state').at(-1).code, 'sign-in-refused');
    const asked = camera.requests.length;
    await new Promise((r) => setTimeout(r, 250));
    assert.equal(camera.requests.length, asked, 'a refused password is not sent again and again (a camera may lock its account)');

    cameras.cameraSave({ id, password: 'right' });
    await until(() => camera.players.size === 1);
    camera.send([makeSlice({ key: true })], { timestamp: 0 });
    await until(() => of(win, 'camera:state').at(-1)?.status === 'live');
    assert.equal(of(win, 'camera:frame').length, 1);

    // Forgotten while it is shown: the reading stops and the window is told.
    cameras.cameraRemove({ id });
    assert.deepEqual([of(win, 'camera:state').at(-1).status, of(win, 'camera:state').at(-1).code], ['failed', 'missing']);
    await until(() => camera.players.size === 0);

    const other = fakeWindow();
    assert.deepEqual([cameras.cameraOpen({ id: 'never-added' }, other).status, cameras.cameraOpen({ id: 'never-added' }, other).code], ['failed', 'missing']);
  } finally {
    cameras.closeCameras();
    await camera.close();
  }
});

test('a camera that drops is tried again for as long as it is wanted, and no longer', async () => {
  const camera = await startCamera({ sprop: [SPS, PPS] });
  const { cameras, of } = service();
  try {
    const { id } = cameras.cameraSave({ url: camera.url }).camera;
    const win = fakeWindow();
    cameras.cameraOpen({ id }, win);
    await until(() => camera.players.size === 1);
    camera.send([makeSlice({ key: true })], { timestamp: 0 });
    await until(() => of(win, 'camera:state').at(-1)?.status === 'live');

    camera.cut();
    await until(() => of(win, 'camera:state').some((s) => s.status === 'waiting'));
    assert.equal(of(win, 'camera:state').find((s) => s.status === 'waiting').code, 'closed');
    await until(() => camera.players.size === 1);
    camera.send([makeSlice({ key: true })], { timestamp: 0 });
    await until(() => of(win, 'camera:state').at(-1)?.status === 'live');
    assert.deepEqual(of(win, 'camera:state').map((s) => s.status), ['connecting', 'live', 'waiting', 'connecting', 'live']);

    // Dropped again, and let go while waiting: it is not tried once nobody wants it.
    camera.cut();
    await until(() => of(win, 'camera:state').at(-1)?.status === 'waiting');
    cameras.cameraClose({ id }, win);
    const asked = camera.requests.length;
    await new Promise((r) => setTimeout(r, 250));
    assert.equal(camera.requests.length, asked);
  } finally {
    cameras.closeCameras();
    await camera.close();
  }
});
