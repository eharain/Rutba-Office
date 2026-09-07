// The renderer build.
//
// esbuild, one config, no framework. The renderer is a single bundle loaded
// from rutba://app/, so there is no dev server to run, no route table to keep
// in step, and nothing to go wrong between a developer's machine and an
// installed copy: the file the installer ships is the file the developer ran.
//
//   node build/bundle.js            build once
//   node build/bundle.js --watch    rebuild on save

import esbuild from 'esbuild';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const app = path.resolve(here, '..');
const out = path.join(app, 'build', 'out');
const watch = process.argv.includes('--watch');
const dev = watch || process.argv.includes('--dev');

fs.mkdirSync(out, { recursive: true });

/** The page. Everything else the renderer needs is in the bundle. */
const html = `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Rutba Office</title>
<link rel="stylesheet" href="app.css">
</head>
<body>
<div id="root"></div>
<script src="app.js"></script>
</body>
</html>
`;

const shared = {
  bundle: true,
  format: 'iife',
  platform: 'browser',
  target: ['chrome120'],
  // JSX lives in .js files here, so the loader is told to expect it. The engine
  // packages are Node-side and are never pulled into this bundle.
  loader: { '.js': 'jsx', '.png': 'dataurl', '.svg': 'dataurl', '.woff2': 'dataurl' },
  jsx: 'automatic',
  logLevel: 'info',
  legalComments: 'none',
  define: { 'process.env.NODE_ENV': dev ? '"development"' : '"production"' },
};

const rendererOptions = {
  ...shared,
  entryPoints: [path.join(app, 'renderer', 'index.js')],
  outfile: path.join(out, 'app.js'),
  minify: !dev,
  sourcemap: dev ? 'inline' : false,
};

/**
 * The preload is built from the same contract the renderer's client reads, so
 * a capability cannot exist on one side and be missing on the other. It has to
 * be CommonJS: a sandboxed preload has no module loader.
 */
const preloadOptions = {
  ...shared,
  entryPoints: [path.resolve(app, '../../packages/office-shell/src/electron/preload.js')],
  outfile: path.join(app, 'build', 'out', 'preload.cjs'),
  format: 'cjs',
  platform: 'node',
  external: ['electron'],
  minify: false,
  sourcemap: false,
};

async function run() {
  fs.writeFileSync(path.join(out, 'index.html'), html);

  if (watch) {
    const a = await esbuild.context(rendererOptions);
    const b = await esbuild.context(preloadOptions);
    await Promise.all([a.watch(), b.watch()]);
    console.log('watching renderer and preload…');
    return;
  }

  const started = Date.now();
  await Promise.all([esbuild.build(rendererOptions), esbuild.build(preloadOptions)]);
  const size = fs.statSync(path.join(out, 'app.js')).size;
  console.log(`built in ${Date.now() - started} ms — app.js ${(size / 1024).toFixed(0)} KB`);
}

run().catch((err) => {
  console.error(err);
  process.exit(1);
});
