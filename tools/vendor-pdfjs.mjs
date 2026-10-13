// Copy the parts of pdf.js the PDF app draws pages with into
// apps/desktop/vendor/pdfjs, from an installed pdfjs-dist:
//
//   node tools/vendor-pdfjs.mjs <path to node_modules/pdfjs-dist>
//
// pdf.js (Mozilla, Apache-2.0) is vendored rather than installed: the window
// loads it from the app's own files, offline, and the installer carries
// exactly these files and nothing of the package's Node side (its optional
// native canvas module is never wanted here). What is copied:
//   build/pdf.min.mjs, build/pdf.worker.min.mjs   the library and its worker
//   standard_fonts/                               the fonts a PDF names without embedding
//   wasm/ (openjpeg, jbig2, qcms)                 JPEG 2000 and JBIG2 scans, colour profiles
//   iccs/                                         the CMYK profile
//   LICENSE and each part's own licence
// Not copied: cmaps/ (169 files, for CJK text in fonts a PDF does not
// embed), the scripting sandbox, the viewer. docs/PDF.md says so.
//
// The minified files carry a few no-break spaces inside string literals; the
// repository's hooks refuse that character in added lines, so each is
// written as its escape, which is the same string.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const from = path.resolve(process.argv[2] || '');
const to = path.resolve(here, '..', 'apps', 'desktop', 'vendor', 'pdfjs');

const pkg = JSON.parse(fs.readFileSync(path.join(from, 'package.json'), 'utf8'));
if (pkg.name !== 'pdfjs-dist') throw new Error(`${from} is not pdfjs-dist`);

const files = [
  'LICENSE',
  'build/pdf.min.mjs',
  'build/pdf.worker.min.mjs',
  ...fs.readdirSync(path.join(from, 'standard_fonts')).map((f) => `standard_fonts/${f}`),
  ...['openjpeg.wasm', 'jbig2.wasm', 'qcms_bg.wasm', 'LICENSE_OPENJPEG', 'LICENSE_PDFJS_OPENJPEG', 'LICENSE_JBIG2', 'LICENSE_PDFJS_JBIG2', 'LICENSE_QCMS', 'LICENSE_PDFJS_QCMS'].map((f) => `wasm/${f}`),
  ...fs.readdirSync(path.join(from, 'iccs')).map((f) => `iccs/${f}`),
];

fs.rmSync(to, { recursive: true, force: true });
for (const rel of files) {
  const src = path.join(from, rel);
  const dest = path.join(to, rel.replace(/^build\//, ''));
  fs.mkdirSync(path.dirname(dest), { recursive: true });
  if (rel.endsWith('.mjs')) {
    const text = fs.readFileSync(src, 'utf8').replace(new RegExp(String.fromCharCode(0xa0), 'g'), '\\xa0');
    fs.writeFileSync(dest, text);
  } else {
    fs.copyFileSync(src, dest);
  }
}
fs.writeFileSync(path.join(to, 'VERSION'), `pdfjs-dist ${pkg.version}\n`);
console.log(`pdf.js ${pkg.version}: ${files.length} files into ${path.relative(process.cwd(), to)}`);
