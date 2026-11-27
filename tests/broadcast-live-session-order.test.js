// Live Broadcast's streams turned on and off while a broadcast runs (B6):
// each change is made after the last has been made, from where it left the
// streams, so a second stream turned on while the first change is still under
// way never stops the first, and a stream left on is never stopped. The window
// module is bundled with its encoders stubbed, as no encoder runs in Node.
import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import esbuild from 'esbuild';

const here = path.dirname(fileURLToPath(import.meta.url));

async function loadSession() {
  const out = await esbuild.build({
    entryPoints: [path.join(here, '../apps/desktop/renderer/apps/broadcast/live-session.js')],
    bundle: true,
    format: 'esm',
    platform: 'node',
    write: false,
    logLevel: 'silent',
    plugins: [{
      name: 'stubs',
      setup(b) {
        b.onResolve({ filter: /\/live\.js$/ }, () => ({ path: 'live', namespace: 'stub' }));
        b.onLoad({ filter: /^live$/, namespace: 'stub' }, () => ({
          contents: `export const AUDIO_KBPS = 128;
            export class LiveEncoder { static async check(v) { return { video: v }; } constructor(o) { this.o = o; } async start() {} async stop() {} setBitrate() {} askKey() {} }
            export class LiveSound { start() {} async stop() {} }`,
          loader: 'js',
        }));
      },
    }],
  });
  return import(`data:text/javascript;base64,${Buffer.from(out.outputFiles[0].text).toString('base64')}`);
}

/** A main process that answers each call after `ms`, keeping the order the calls arrive in. */
function fakeShell(ms = 20) {
  const calls = [];
  const later = (v) => new Promise((r) => setTimeout(() => r(v), ms));
  return {
    calls,
    live: {
      start: async ({ ids }) => { calls.push(['start', ...ids]); return later({ started: ids, skipped: [] }); },
      startOne: async ({ id }) => { calls.push(['startOne', id]); return later({}); },
      stopOne: async ({ id }) => { calls.push(['stopOne', id]); return later({}); },
      stop: async () => { calls.push(['stop']); return later({}); },
    },
  };
}

const STREAMS = [{ id: 'a', shape: 'wide', hasKey: true }, { id: 'b', shape: 'wide', hasKey: true }, { id: 'c', shape: 'wide', hasKey: true }];
const program = { canvas: { width: 1280, height: 720 } };

test('a second stream turned on while the first change is still being made leaves the first going', async () => {
  const { LiveSession } = await loadSession();
  const shell = fakeShell();
  const s = new LiveSession({ shell, media: {}, program, fps: 30, kbps: 4500 });
  await s.open({ ids: ['a'], streams: STREAMS });
  shell.calls.length = 0;
  // B on, then C on before B's change is through: each from where the last left it.
  const one = s.setOn(['a', 'b'], STREAMS);
  const two = s.setOn(['a', 'b', 'c'], STREAMS);
  await Promise.all([one, two]);
  assert.deepEqual(shell.calls, [['startOne', 'b'], ['startOne', 'c']]);
  assert.ok(!shell.calls.some((c) => c[0] === 'stopOne'), 'nothing that stays on is stopped');
  assert.deepEqual(s.on, ['a', 'b', 'c']);
});

test('off and on again quickly: stopped, then started, in that order', async () => {
  const { LiveSession } = await loadSession();
  const shell = fakeShell();
  const s = new LiveSession({ shell, media: {}, program, fps: 30, kbps: 4500 });
  await s.open({ ids: ['a', 'b'], streams: STREAMS });
  shell.calls.length = 0;
  await Promise.all([s.setOn(['b'], STREAMS), s.setOn(['a', 'b'], STREAMS)]);
  assert.deepEqual(shell.calls, [['stopOne', 'a'], ['startOne', 'a']]);
});

test('a change that fails does not hold up the next', async () => {
  const { LiveSession } = await loadSession();
  const shell = fakeShell();
  const s = new LiveSession({ shell, media: {}, program, fps: 30, kbps: 4500 });
  await s.open({ ids: ['a'], streams: STREAMS });
  shell.live.startOne = async ({ id }) => { if (id === 'b') throw new Error('That stream has no key yet.'); shell.calls.push(['startOne', id]); };
  shell.calls.length = 0;
  const one = s.setOn(['a', 'b'], STREAMS);
  const two = s.setOn(['a', 'b', 'c'], STREAMS);
  await assert.rejects(one, /no key/);
  await two;
  assert.deepEqual(shell.calls, [['startOne', 'c']]);
});

test('End broadcast while a change is under way, or while it starts: nothing is left running after it', async () => {
  const { LiveSession } = await loadSession();
  const shell = fakeShell();
  const s = new LiveSession({ shell, media: {}, program, fps: 30, kbps: 4500 });
  await s.open({ ids: ['a'], streams: STREAMS });
  const change = s.setOn(['a', 'b'], STREAMS);
  await s.end();
  await change.catch(() => {});
  assert.equal(s.encoders.size, 0, 'no encoder left');
  assert.deepEqual(shell.calls.at(-1), ['stop'], 'the streams are closed last');

  const t = new LiveSession({ shell: fakeShell(), media: {}, program, fps: 30, kbps: 4500 });
  const opening = t.open({ ids: ['a'], streams: STREAMS });
  await t.end();
  await opening;
  assert.equal(t.encoders.size, 0, 'the window closed while it started: no encoder');
  assert.equal(t.sound ?? null, null, 'and no sound');

  // Ended while the main process is still starting the streams.
  const shell3 = fakeShell(30);
  const u = new LiveSession({ shell: shell3, media: {}, program, fps: 30, kbps: 4500 });
  const opening3 = u.open({ ids: ['a'], streams: STREAMS });
  while (!shell3.calls.some((c) => c[0] === 'start')) await new Promise((r) => setTimeout(r, 1));
  await Promise.all([u.end(), opening3]);
  assert.equal(u.encoders.size, 0, 'no encoder');
  assert.equal(u.sound ?? null, null, 'no sound');
  assert.deepEqual(shell3.calls.at(-1), ['stop'], 'and the streams it started are closed');
});
