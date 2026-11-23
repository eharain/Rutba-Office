// The window's half of a network camera source (renderer/studio/net-camera.js):
// which cameras a scene holds open, how frames reach the decoder, and what the
// source list is told. The decoder, the canvas and the shell are stand-ins; the
// browser's own decoder is proven by the picture check, not here.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { NetCameras } from '../apps/desktop/renderer/studio/net-camera.js';
import { addItem, emptyStack, normaliseStack, PICTURE_KINDS } from '../packages/studio/src/stack.js';
import { makeTallLayers } from '../packages/studio/src/tall.js';
import { rematch } from '../packages/studio/src/setup.js';

const tick = () => new Promise((r) => setImmediate(r));

function fakeShell({ open = () => ({ status: 'connecting', code: '' }) } = {}) {
  const handlers = new Map();
  const calls = [];
  return {
    calls,
    on(name, fn) {
      handlers.set(name, fn);
      return () => handlers.delete(name);
    },
    emit(name, value) { handlers.get(name)?.(value); },
    listening: () => [...handlers.keys()].sort(),
    capture: {
      cameraOpen: async ({ id }) => {
        calls.push(`open ${id}`);
        return open(id);
      },
      cameraClose: async ({ id }) => {
        calls.push(`close ${id}`);
        return { closed: true };
      },
    },
  };
}

/** A decoder that gives one frame back for each chunk it is handed, when told to. */
function fakeDecoders({ size = [640, 360], refuse = false } = {}) {
  const made = [];
  class Decoder {
    constructor(init) {
      if (refuse) throw new Error('no decoder');
      this.init = init;
      this.state = 'unconfigured';
      this.decodeQueueSize = 0;
      this.chunks = [];
      this.config = null;
      made.push(this);
    }
    configure(config) {
      this.config = config;
      this.state = 'configured';
    }
    decode(chunk) { this.chunks.push(chunk); }
    close() { this.state = 'closed'; }
    /** Hands the picture of the chunks so far to whoever made the decoder. */
    show() {
      const frame = { displayWidth: size[0], displayHeight: size[1], closed: false, close() { this.closed = true; } };
      this.init.output(frame);
      return frame;
    }
  }
  class Chunk {
    constructor(init) { Object.assign(this, init); }
  }
  return { Decoder, Chunk, made };
}

function fakeCanvas() {
  const drawn = [];
  return {
    width: 0,
    height: 0,
    dataUrl: 'old',
    drawn,
    getContext: () => ({ drawImage: (frame, x, y, w, h) => drawn.push([x, y, w, h]) }),
  };
}

const source = (ref, live = true, id = `i-${ref}`) => ({ id, kind: 'netcam', ref, live, name: ref });
const bytes = (n = 4) => new Uint8Array(n).fill(7);

function rig(options = {}) {
  const shell = fakeShell(options);
  const decoders = fakeDecoders(options);
  const canvases = [];
  let clock = 1000;
  let changes = 0;
  const cams = new NetCameras({
    shell,
    onChange: () => { changes += 1; },
    makeCanvas: () => {
      const c = fakeCanvas();
      canvases.push(c);
      return c;
    },
    Decoder: options.noDecoder ? undefined : decoders.Decoder,
    Chunk: options.noDecoder ? undefined : decoders.Chunk,
    now: () => clock,
  });
  return { shell, decoders, canvases, cams, changes: () => changes, wait: (ms) => { clock += ms; } };
}

test('a scene holds open only the cameras its started sources show', async () => {
  const { shell, cams, changes } = rig();
  const a = source('cam-a');
  const b = source('cam-b', false);
  cams.sync([a, b, { id: 'x', kind: 'camera', ref: 'usb', live: true }, { id: 'y', kind: 'netcam', ref: '', live: true }]);
  await tick();
  assert.deepEqual(shell.calls, ['open cam-a']);
  assert.equal(cams.rowState(a), 'opening');
  assert.equal(cams.rowState(b), 'off');
  assert.equal(cams.statusOf(b), 'off');
  assert.equal(cams.codeOf(b), '');
  assert.equal(cams.drawable(a), null);

  // The same sources again open nothing twice and say nothing.
  const before = changes();
  cams.sync([a, b]);
  await tick();
  assert.deepEqual(shell.calls, ['open cam-a']);
  assert.equal(changes(), before);

  // Two sources of one camera are one reading; starting the other camera opens it.
  cams.sync([a, source('cam-a', true, 'second'), { ...b, live: true }]);
  await tick();
  assert.deepEqual(shell.calls, ['open cam-a', 'open cam-b']);

  // Stopped or taken out of the scene, a camera is let go.
  cams.sync([{ ...a, live: false }, { ...b, live: true }]);
  await tick();
  assert.deepEqual(shell.calls, ['open cam-a', 'open cam-b', 'close cam-a']);
  assert.equal(cams.rowState(a), 'off');

  cams.dispose();
  await tick();
  assert.deepEqual(shell.calls, ['open cam-a', 'open cam-b', 'close cam-a', 'close cam-b']);
  assert.deepEqual(shell.listening(), []);
});

test('the picture starts at a key frame and is drawn from the camera\'s canvas', async () => {
  const { shell, decoders, canvases, cams, changes, wait } = rig();
  const a = source('cam-a');
  cams.sync([a]);
  await tick();

  // A frame of a camera this window does not show, and one before any key frame, go nowhere.
  shell.emit('camera:frame', { id: 'other', key: true, codec: 'avc1.640028', timestamp: 0, bytes: bytes() });
  shell.emit('camera:frame', { id: 'cam-a', key: false, timestamp: 0, bytes: bytes() });
  assert.equal(decoders.made.length, 0);

  shell.emit('camera:state', { id: 'cam-a', status: 'live', code: '' });
  assert.equal(cams.rowState(a), 'opening', 'live is said once there is a picture to show');
  shell.emit('camera:frame', { id: 'cam-a', key: true, codec: 'avc1.640028', timestamp: 1000, bytes: bytes(9) });
  assert.equal(decoders.made.length, 1);
  const d = decoders.made[0];
  assert.deepEqual(d.config, { codec: 'avc1.640028', optimizeForLatency: true });
  assert.deepEqual(d.chunks.map((c) => [c.type, c.timestamp, c.data.length]), [['key', 1000, 9]]);
  shell.emit('camera:frame', { id: 'cam-a', key: false, timestamp: 2000, bytes: bytes(3) });
  shell.emit('camera:frame', { id: 'cam-a', key: false, timestamp: 3000, bytes: new Uint8Array(0) });
  assert.deepEqual(d.chunks.map((c) => c.type), ['key', 'delta']);

  const before = changes();
  const frame = d.show();
  assert.equal(frame.closed, true);
  assert.equal(canvases.length, 1);
  assert.deepEqual([canvases[0].width, canvases[0].height], [640, 360]);
  assert.deepEqual(canvases[0].drawn, [[0, 0, 640, 360]]);
  assert.equal(cams.drawable(a), canvases[0]);
  assert.equal(cams.rowState(a), 'live');
  assert.equal(canvases[0].dataUrl, null, 'the list makes its small picture again');
  assert.equal(changes(), before + 1);

  // More frames paint the same canvas without stirring the window, until the small picture is due again.
  canvases[0].dataUrl = 'made';
  d.show();
  wait(1000);
  d.show();
  assert.equal(changes(), before + 1);
  assert.equal(canvases[0].dataUrl, 'made');
  wait(5000);
  d.show();
  assert.equal(changes(), before + 2);
  assert.equal(canvases[0].dataUrl, null);
  assert.equal(canvases[0].drawn.length, 4);

  // A key frame of the same kind keeps the decoder; another kind starts a new one.
  shell.emit('camera:frame', { id: 'cam-a', key: true, codec: 'avc1.640028', timestamp: 9000, bytes: bytes() });
  assert.equal(decoders.made.length, 1);
  shell.emit('camera:frame', { id: 'cam-a', key: true, codec: 'avc1.4d401f', timestamp: 10000, bytes: bytes() });
  assert.equal(decoders.made.length, 2);
  assert.equal(d.state, 'closed');
  assert.deepEqual(decoders.made[1].chunks.map((c) => c.type), ['key']);

  // A frame the old decoder gives back late is closed and not painted.
  const late = d.show();
  assert.equal(late.closed, true);
  assert.equal(canvases[0].drawn.length, 4);
});

test('a window that falls behind starts again at the next key frame', async () => {
  const { shell, decoders, cams } = rig();
  cams.sync([source('cam-a')]);
  await tick();
  shell.emit('camera:frame', { id: 'cam-a', key: true, codec: 'avc1.640028', timestamp: 0, bytes: bytes() });
  const d = decoders.made[0];
  d.decodeQueueSize = 31;
  shell.emit('camera:frame', { id: 'cam-a', key: false, timestamp: 1, bytes: bytes() });
  assert.equal(d.state, 'closed');
  shell.emit('camera:frame', { id: 'cam-a', key: false, timestamp: 2, bytes: bytes() });
  assert.equal(decoders.made.length, 1, 'nothing is decoded from a frame that leans on one let go');
  shell.emit('camera:frame', { id: 'cam-a', key: true, codec: 'avc1.640028', timestamp: 3, bytes: bytes() });
  assert.equal(decoders.made.length, 2);
  assert.deepEqual(decoders.made[1].chunks.map((c) => c.type), ['key']);
});

test('what the camera says reaches the source list by its code', async () => {
  const { shell, decoders, cams } = rig({ open: () => ({ status: 'failed', code: 'missing' }) });
  const a = source('cam-a');
  cams.sync([a]);
  await tick();
  assert.equal(cams.rowState(a), 'error');
  assert.equal(cams.codeOf(a), 'missing');

  shell.emit('camera:state', { id: 'cam-a', status: 'connecting', code: '' });
  assert.equal(cams.rowState(a), 'opening');
  assert.equal(cams.codeOf(a), '');
  shell.emit('camera:state', { id: 'cam-a', status: 'waiting', code: 'unreachable' });
  assert.equal(cams.rowState(a), 'error');
  assert.equal(cams.statusOf(a), 'waiting');
  assert.equal(cams.codeOf(a), 'unreachable');
  shell.emit('camera:state', { id: 'cam-a', status: 'failed' });
  assert.equal(cams.codeOf(a), 'closed');
  shell.emit('camera:state', { id: 'nobody', status: 'live' });

  // A camera that drops keeps its last picture on show while it is tried again.
  shell.emit('camera:state', { id: 'cam-a', status: 'live', code: '' });
  shell.emit('camera:frame', { id: 'cam-a', key: true, codec: 'avc1.640028', timestamp: 0, bytes: bytes() });
  decoders.made[0].show();
  assert.equal(cams.rowState(a), 'live');
  shell.emit('camera:state', { id: 'cam-a', status: 'waiting', code: 'stalled' });
  assert.equal(cams.rowState(a), 'error');
  assert.ok(cams.drawable(a));
});

test('what opening a camera answers never overwrites what the camera has said since', async () => {
  let answer;
  const { shell, cams } = rig({ open: () => new Promise((r) => { answer = r; }) });
  const a = source('cam-a');
  cams.sync([a]);
  await tick();
  shell.emit('camera:state', { id: 'cam-a', status: 'failed', code: 'sign-in-refused' });
  answer({ status: 'connecting', code: '' });
  await tick();
  assert.equal(cams.codeOf(a), 'sign-in-refused');
});

test('a computer that cannot decode the picture says so, once, and without stopping the rest', async () => {
  for (const options of [{ noDecoder: true }, { refuse: true }]) {
    const { shell, cams, changes } = rig(options);
    const a = source('cam-a');
    cams.sync([a]);
    await tick();
    shell.emit('camera:state', { id: 'cam-a', status: 'live', code: '' });
    const before = changes();
    shell.emit('camera:frame', { id: 'cam-a', key: true, codec: 'avc1.640028', timestamp: 0, bytes: bytes() });
    shell.emit('camera:frame', { id: 'cam-a', key: true, codec: 'avc1.640028', timestamp: 1, bytes: bytes() });
    assert.equal(cams.rowState(a), 'error');
    assert.equal(cams.codeOf(a), 'decode');
    assert.equal(changes(), before + 1);
    // Read afresh, the camera is given another go.
    shell.emit('camera:state', { id: 'cam-a', status: 'connecting', code: '' });
    assert.equal(cams.codeOf(a), '');
  }

  // A decoder that gives up in the middle is started again at the next key frame, and the picture stays.
  const { shell, decoders, cams } = rig();
  const a = source('cam-a');
  cams.sync([a]);
  await tick();
  shell.emit('camera:state', { id: 'cam-a', status: 'live', code: '' });
  shell.emit('camera:frame', { id: 'cam-a', key: true, codec: 'avc1.640028', timestamp: 0, bytes: bytes() });
  decoders.made[0].show();
  decoders.made[0].init.error(new Error('bad data'));
  assert.equal(decoders.made[0].state, 'closed');
  assert.equal(cams.rowState(a), 'live', 'a camera that has shown a picture is not called broken for one bad frame');
  shell.emit('camera:frame', { id: 'cam-a', key: true, codec: 'avc1.640028', timestamp: 5, bytes: bytes() });
  assert.equal(decoders.made.length, 2);
});

test('a network camera is a picture source like any other in a scene', () => {
  assert.ok(PICTURE_KINDS.includes('netcam'));
  const { stack, id } = addItem(emptyStack(), { kind: 'netcam', ref: 'cam-a', name: 'Front door' });
  const it = stack.items.find((x) => x.id === id);
  assert.equal(it.kind, 'netcam');
  assert.equal(it.ref, 'cam-a');
  assert.equal(it.name, 'Front door');
  // Saved and opened again, it is the same source, with no address in it.
  const again = normaliseStack(JSON.parse(JSON.stringify(stack)));
  assert.equal(again.items[0].kind, 'netcam');
  assert.equal(again.items[0].ref, 'cam-a');
  assert.equal(JSON.stringify(stack).includes('rtsp'), false);
  // In the tall picture it shares the frame as a camera does: a screen whole across the width, the camera filling the rest.
  const items = [{ id: 's', kind: 'screen', on: true }, { id: 'n', kind: 'netcam', on: true }];
  const layers = makeTallLayers(items);
  const wide = makeTallLayers([{ id: 's', kind: 'screen', on: true }, { id: 'n', kind: 'camera', on: true }]);
  assert.deepEqual(layers, wide);
  assert.equal(layers.length, 2);

  // A setup opened on a computer is not told its network camera is missing: that is said where the camera is shown.
  const setup = { stack: { ...stack, sounds: [] } };
  const r = rematch(setup, [], { cameras: [], mics: [] });
  assert.deepEqual(r.missing.filter((m) => m.kind === 'netcam'), []);
});
