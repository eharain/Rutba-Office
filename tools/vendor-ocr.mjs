// Copy the text recognition engine and its six language models into
// packages/ocr/vendor, from an installed tesseract.js and a folder of
// tessdata_fast models:
//
//   node tools/vendor-ocr.mjs <path to node_modules/tesseract.js> <tessdata_fast folder>
//
// The engine is Tesseract 5 compiled to WebAssembly as tesseract.js 7 ships
// it (tesseract.js-core 7, Apache-2.0); the models are the "fast" LSTM
// models of tesseract-ocr/tessdata_fast (Apache-2.0). Both are vendored
// rather than installed, as pdf.js is: the suite reads them from its own
// files, so recognition never fetches anything, and the installer carries
// exactly these files and none of tesseract.js's own loader, which can
// fetch models from the network when it is not told otherwise. Our worker
// (packages/ocr/src/engine.js) drives the engine directly.
//
// What is copied:
//   core/tesseract-core-simd-lstm.cjs, .wasm  the engine: SIMD, LSTM only (its
//                                             loader is CommonJS, and the package is ESM)
//   core/LICENSE                              tesseract.js-core's licence
//   tessdata/<model>.traineddata              one for each row of packages/ocr/src/languages.js
//   tessdata/LICENSE                          tessdata_fast's licence
// and VERSION files naming the versions and each file's SHA-256.

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { LANGUAGES as TABLE } from '../packages/ocr/src/languages.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const from = path.resolve(process.argv[2] || '');
const models = path.resolve(process.argv[3] || '');
const to = path.resolve(here, '..', 'packages', 'ocr', 'vendor');

// The models: one a language of the table, which is where a language is added.
const LANGS = [...new Set(TABLE.map((l) => l.model))];
const BUILD = 'tesseract-core-simd-lstm';

const pkg = JSON.parse(fs.readFileSync(path.join(from, 'package.json'), 'utf8'));
if (pkg.name !== 'tesseract.js' || !/^7\./.test(pkg.version)) throw new Error(`${from} is not tesseract.js 7`);
// The engine tesseract.js resolves: its own nested copy first, then a sibling.
const core = [path.join(from, 'node_modules', 'tesseract.js-core'), path.join(from, '..', 'tesseract.js-core')]
  .find((dir) => fs.existsSync(path.join(dir, `${BUILD}.wasm`)));
if (!core) throw new Error(`no tesseract.js-core with ${BUILD}.wasm beside ${from}`);
const corePkg = JSON.parse(fs.readFileSync(path.join(core, 'package.json'), 'utf8'));

const sha = (file) => crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');

fs.rmSync(to, { recursive: true, force: true });
fs.mkdirSync(path.join(to, 'core'), { recursive: true });
fs.mkdirSync(path.join(to, 'tessdata'), { recursive: true });

const lines = [`tesseract.js ${pkg.version}`, `tesseract.js-core ${corePkg.version}`, ''];
for (const [name, as] of [[`${BUILD}.js`, `${BUILD}.cjs`], [`${BUILD}.wasm`, `${BUILD}.wasm`], ['LICENSE', 'LICENSE']]) {
  fs.copyFileSync(path.join(core, name), path.join(to, 'core', as));
  lines.push(`${sha(path.join(to, 'core', as))}  ${as}`);
}
fs.writeFileSync(path.join(to, 'core', 'VERSION'), `${lines.join('\n')}\n`);

const data = ['tessdata_fast (https://github.com/tesseract-ocr/tessdata_fast)', ''];
for (const lang of LANGS) {
  const name = `${lang}.traineddata`;
  if (!fs.existsSync(path.join(models, name))) throw new Error(`${name} is not in ${models}: every language in packages/ocr/src/languages.js needs its model`);
  fs.copyFileSync(path.join(models, name), path.join(to, 'tessdata', name));
  data.push(`${sha(path.join(to, 'tessdata', name))}  ${name}`);
}
fs.copyFileSync(path.join(models, 'LICENSE'), path.join(to, 'tessdata', 'LICENSE'));
fs.writeFileSync(path.join(to, 'tessdata', 'VERSION'), `${data.join('\n')}\n`);

const size = (dir) => fs.readdirSync(dir).reduce((n, f) => n + fs.statSync(path.join(dir, f)).size, 0);
console.log(`tesseract.js ${pkg.version} (core ${corePkg.version}) and ${LANGS.length} models into ${path.relative(process.cwd(), to)}: ${((size(path.join(to, 'core')) + size(path.join(to, 'tessdata'))) / 1048576).toFixed(1)} MB`);
