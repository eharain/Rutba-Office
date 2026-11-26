// The Recorder's frame when a source throws while it is drawn (B17): the
// typed words being said are still drawn across the foot, the canvas says so,
// and the drawing state is put back, so the next frame starts clean. The
// window's program module is bundled here as studio-program-stop does.
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
        b.onLoad({ filter: /.*/, namespace: 'stub' }, () => ({ contents: "export { t } from './messages.js';", resolveDir: path.join(here, '../packages/office-ui/src'), loader: 'js' }));
      },
    }],
  });
  const code = out.outputFiles[0].text;
  return import(`data:text/javascript;base64,${Buffer.from(code).toString('base64')}`);
}

/** A 2D context that keeps what was asked of it and how deep its saved states go. */
function fakeContext() {
  const calls = [];
  let depth = 0;
  const ctx = {
    calls,
    get depth() { return depth; },
    save() { depth += 1; calls.push('save'); },
    restore() { depth -= 1; calls.push('restore'); },
    measureText: (s) => ({ width: String(s).length * 10 }),
    fillText: (s) => calls.push(`text:${s}`),
    drawImage: () => { throw new Error('The image source is detached.'); },
  };
  return new Proxy(ctx, {
    get: (o, k) => (k in o ? o[k] : () => {}),
    set: (o, k, v) => { o[k] = v; return true; },
  });
}

function programWith(Program, { media, stack }) {
  const p = Object.create(Program.prototype);
  p.ctx = fakeContext();
  p.canvas = { width: 1920, height: 1080, dataset: {} };
  p.media = media;
  p.stack = stack;
  p.accent = '#e03b4c';
  p.frames = 0;
  p.listeners = new Set();
  p.fade = null;
  return p;
}

const TYPED = 'Welcome to the new starters';

test('a source that throws as it is drawn leaves the typed words on the frame', async () => {
  const { Program } = await loadProgram();
  const { TYPED_LINE } = await import('../packages/studio/src/spoken.js');
  const picture = { width: 1280, height: 720 };
  const media = {
    drawable: () => picture,
    spokenLine: (id) => (id === TYPED_LINE ? TYPED : null),
  };
  const stack = { video: 'on', items: [{ id: 'screen-1', kind: 'screen', on: true, box: { x: 0, y: 0, w: 1, h: 1 }, shape: 'rect' }] };
  const p = programWith(Program, { media, stack });
  p.draw();
  assert.equal(p.canvas.dataset.typed, '1');
  assert.ok(p.ctx.calls.some((c) => c === `text:${TYPED}`), 'the typed line is written on the frame');
  assert.equal(p.ctx.depth, 0, 'every saved drawing state is put back');
  assert.equal(p.canvas.dataset.layers, '', 'the item that threw is not counted as drawn');
  assert.equal(p.canvas.dataset.frames, '1');
});

test('a tick that throws does not stop the frame', async () => {
  const { Program } = await loadProgram();
  const { TYPED_LINE } = await import('../packages/studio/src/spoken.js');
  const media = { drawable: () => null, spokenLine: (id) => (id === TYPED_LINE ? TYPED : null) };
  const p = programWith(Program, { media, stack: { video: 'on', items: [] } });
  p.onTick = () => { throw new Error('a camera could not be opened again'); };
  p.draw();
  assert.equal(p.canvas.dataset.typed, '1');
  assert.equal(p.canvas.dataset.frames, '1');
});
