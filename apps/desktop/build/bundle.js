// The renderer build.
//
// esbuild, one config, no framework. The renderer is loaded from rutba://app/,
// so there is no dev server to run, no route table to keep in step, and
// nothing to go wrong between a developer's machine and an installed copy: the
// files the installer ships are the files the developer ran. app.js is what
// every window shares; each app is a bundle of its own in chunks/, loaded when
// a window of that app opens (owner, 2026-10-20, the build split by app).
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
// What went into each output file, for the app packs: a chunk only one app's
// code is in goes into that app's pack (make-app-packs.js, check-packs.mjs).
// Beside the output, never in it: nothing ships it.
export const RENDERER_META = path.join(app, 'build', 'meta', 'renderer.json');
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
<script type="module" src="app.js"></script>
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
  // Modules, split: what two apps share is a chunk of its own, loaded once.
  format: 'esm',
  splitting: true,
  entryPoints: { app: path.join(app, 'renderer', 'index.js') },
  outdir: out,
  chunkNames: 'chunks/[name]-[hash]',
  metafile: true,
  minify: !dev,
  sourcemap: dev ? 'inline' : false,
};

/**
 * Automate → scripts: the worker a script runs in, a bundle of its own —
 * protocol.js serves it under a policy of its own, which reaches nothing.
 */
const workerOptions = {
  ...shared,
  entryPoints: [path.join(app, 'renderer', 'apps', 'sheets', 'scripts-worker.js')],
  outfile: path.join(out, 'scripts-worker.js'),
  minify: !dev,
  sourcemap: false,
};

/**
 * Image: the worker that works a picture's pixels (tone, detail, a batch's
 * conversions) off the window's thread. Served from the app like app.js.
 */
const imageWorkerOptions = {
  ...shared,
  entryPoints: [path.join(app, 'renderer', 'apps', 'image', 'worker.js')],
  outfile: path.join(out, 'image-worker.js'),
  minify: !dev,
  sourcemap: false,
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

/**
 * What changed in this version, for the window that opens after an update.
 *
 * The notes are the release's own — docs/releases/v<version>.md, the file
 * the GitHub release is published from — so the suite tells the same story
 * offline that the release page tells online. A version with no notes yet
 * ships `null`, and the dialog points at the releases page instead.
 */
function whatsNew() {
  const { version } = JSON.parse(fs.readFileSync(path.join(app, 'package.json'), 'utf8'));
  const dir = path.resolve(app, '..', '..', 'docs', 'releases');
  let file = path.join(dir, `v${version}.md`);
  // A development build (2.13.0-dev) shows the notes of the release it leads to once
  // they are drafted, and until then the last release's.
  const dev = /^(\d+)\.(\d+)\.(\d+)-/.exec(version);
  if (dev && !fs.existsSync(file)) {
    const want = path.join(dir, `v${dev[1]}.${dev[2]}.${dev[3]}.md`);
    const num = (v) => v.split('.').map(Number);
    const older = (a, b) => { const [x, y] = [num(a), num(b)]; for (let i = 0; i < 3; i++) if (x[i] !== y[i]) return x[i] < y[i]; return false; };
    const base = `${dev[1]}.${dev[2]}.${dev[3]}`;
    const last = fs.readdirSync(dir).map((n) => /^v(\d+\.\d+\.\d+)\.md$/.exec(n)?.[1]).filter(Boolean).filter((v) => older(v, base)).sort((a, b) => (older(a, b) ? 1 : -1))[0];
    file = fs.existsSync(want) ? want : last ? path.join(dir, `v${last}.md`) : file;
  }
  const notes = fs.existsSync(file) ? fs.readFileSync(file, 'utf8') : null;
  fs.writeFileSync(path.join(out, 'whatsnew.json'), JSON.stringify({ version, notes }));
}

/**
 * pdf.js, which the PDF app draws pages with: vendored whole in
 * vendor/pdfjs (tools/vendor-pdfjs.mjs says what and why) and served from
 * the app beside the bundle, loaded by the window when it first needs it.
 */
function pdfjs() {
  const from = path.join(app, 'vendor', 'pdfjs');
  const to = path.join(out, 'pdfjs');
  if (!fs.existsSync(from)) return;
  fs.rmSync(to, { recursive: true, force: true });
  fs.cpSync(from, to, { recursive: true });
}

async function run() {
  // The last build's chunks go: their names change with their contents.
  fs.rmSync(path.join(out, 'chunks'), { recursive: true, force: true });
  fs.writeFileSync(path.join(out, 'index.html'), html);
  whatsNew();
  pdfjs();

  if (watch) {
    const a = await esbuild.context(rendererOptions);
    const b = await esbuild.context(preloadOptions);
    const c = await esbuild.context(workerOptions);
    const d = await esbuild.context(imageWorkerOptions);
    await Promise.all([a.watch(), b.watch(), c.watch(), d.watch()]);
    console.log('watching renderer and preload…');
    return;
  }

  const started = Date.now();
  const [renderer] = await Promise.all([esbuild.build(rendererOptions), esbuild.build(preloadOptions), esbuild.build(workerOptions), esbuild.build(imageWorkerOptions)]);
  writeMeta(renderer.metafile);
  const size = fs.statSync(path.join(out, 'app.js')).size;
  const chunks = fs.readdirSync(path.join(out, 'chunks')).filter((f) => f.endsWith('.js'));
  const chunked = chunks.reduce((n, f) => n + fs.statSync(path.join(out, 'chunks', f)).size, 0);
  console.log(`built in ${Date.now() - started} ms: app.js ${(size / 1024).toFixed(0)} KB, ${chunks.length} chunks ${(chunked / 1024).toFixed(0)} KB`);
}

/** Each output file of the renderer and its inputs, by their paths from the repository's root. */
function writeMeta(metafile) {
  const root = path.resolve(app, '..', '..');
  const posix = (p) => p.split(path.sep).join('/');
  const outputs = {};
  for (const [file, o] of Object.entries(metafile.outputs)) {
    if (!file.endsWith('.js')) continue;
    const rel = posix(path.relative(out, path.resolve(file)));
    outputs[rel] = Object.keys(o.inputs).map((i) => posix(path.relative(root, path.resolve(i))));
  }
  fs.mkdirSync(path.dirname(RENDERER_META), { recursive: true });
  fs.writeFileSync(RENDERER_META, `${JSON.stringify({ outputs }, null, 1)}\n`);
}

// Run when called, not when another build script imports the file's names.
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) run().catch((err) => {
  console.error(err);
  process.exit(1);
});
