// The Recorder's Stop when the disk stopped taking the recording part way:
// what reached it is kept and said, nothing at all is an error, and each
// take's own count is its own (second review, 10 Nov 2026). The window's
// program module is bundled here with the UI kit's t() stood in for, as the
// kit brings React with it.
import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import esbuild from 'esbuild';

const here = path.dirname(fileURLToPath(import.meta.url));

async function loadProgram() {
  const out = await esbuild.build({
    entryPoints: [path.join(here, '../apps/desktop/renderer/studio/program.js')],
    bundle: true,
    format: 'esm',
    platform: 'node',
    write: false,
    logLevel: 'silent',
    plugins: [{
      name: 'ui-kit',
      setup(b) {
        b.onResolve({ filter: /^@rutba\/office-ui$/ }, () => ({ path: 'ui-kit', namespace: 'stub' }));
        b.onLoad({ filter: /.*/, namespace: 'stub' }, () => ({ contents: 'export const t = (s, v = {}) => s.replace(/\\{(\\w+)\\}/g, (_, k) => String(v[k] ?? ""));', loader: 'js' }));
      },
    }],
  });
  const code = out.outputFiles[0].text;
  return import(`data:text/javascript;base64,${Buffer.from(code).toString('base64')}`);
}

/** A recorder already stopped, and its writer's answer. */
const partOf = (done, take = null) => ({
  rec: { state: 'inactive', stop() {} },
  stopped: Promise.resolve(),
  writer: { done: async () => done, written: done.written },
  take,
});
const takeOf = (role) => ({ role, path: `D:/films/${role}.webm`, contentType: 'video/webm', hasAudio: false, width: 1280, height: 720 });

function stopped(Program, { main = null, side = [] }) {
  const p = Object.create(Program.prototype);
  const parts = [...(main ? [main] : []), ...side];
  p.recording = { parts, main, side, clock: { elapsed: () => 12_000, paused: false }, path: main ? 'D:/films/program.webm' : null, heard: true, type: 'video/webm;codecs=vp9', soundOnly: false };
  return p.stop();
}

test('a disk that stopped part way: Stop keeps what was written and says why it ends there', async () => {
  const { Program } = await loadProgram();
  const r = await stopped(Program, { main: partOf({ written: 5_000_000, failed: new Error('There is not enough space on the disk.') }) });
  assert.equal(r.path, 'D:/films/program.webm');
  assert.equal(r.written, 5_000_000);
  assert.equal(r.seconds, 12);
  assert.equal(r.failed, 'There is not enough space on the disk.');
});

test('nothing written at all is an error, not an empty recording', async () => {
  const { Program } = await loadProgram();
  await assert.rejects(stopped(Program, { main: partOf({ written: 0, failed: new Error('Access is denied.') }) }), /could not be kept: Access is denied/);
});

test('a clean Stop says nothing failed', async () => {
  const { Program } = await loadProgram();
  const r = await stopped(Program, { main: partOf({ written: 900, failed: null }) });
  assert.equal(r.failed, null);
  assert.equal(r.written, 900);
});

test('takes kept on their own: each keeps its own count, one that failed and one that did not', async () => {
  const { Program } = await loadProgram();
  const r = await stopped(Program, { side: [partOf({ written: 0, failed: new Error('gone') }, takeOf('camera')), partOf({ written: 4200, failed: null }, takeOf('screen'))] });
  assert.deepEqual(r.takes.map((k) => [k.role, k.written, k.duration]), [['camera', 0, 12], ['screen', 4200, 12]]);
  assert.equal(r.failed, 'gone', 'and the failure is said');
});
