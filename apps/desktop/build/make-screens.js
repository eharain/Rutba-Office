// Screenshots of the real application, for the website.
//
//   node build/make-screens.js [outDir]
//   npm run screens
//
// `npm run smoke` already boots every window and photographs it — that is what
// this borrows. What it adds is the two things that stop a smoke capture being
// usable as a picture of a product:
//
//   1. **Documents worth photographing.** A smoke run opens each app empty, so
//      the spreadsheet screenshot is an empty grid and the document screenshot
//      is a blank page. That proves the window paints and sells nothing. This
//      builds a small, realistic file for each of the three document apps
//      first, and hands it to that app's capture.
//
//   2. **A profile of its own.** The smoke run writes into the real user
//      profile, so its seeded mail account accumulates: after eleven runs the
//      mail screenshot showed eleven identical accounts in the sidebar. Each
//      capture here runs against a fresh `--user-data-dir` in a temp
//      directory, which is also the only honest way to photograph a mail
//      client — nobody's actual inbox belongs in a marketing asset.
//
// The content is invented and says so where it can. The numbers are made up,
// the correspondents are fictional, and the one thing that is real is the
// software drawing them.
import { spawn } from 'node:child_process';
import zlib from 'node:zlib';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import electron from 'electron';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation';

const appDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const outDir = path.resolve(process.argv[2] || path.join(appDir, 'build', 'screens'));

const work = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-screens-'));
// The status bar shows this path, so it is named rather than random: a
// screenshot whose footer reads `rutba-screens-JAsyZI\docs` reads as a
// test artefact, which is what it is and not what it is for.
const docsDir = path.join(os.tmpdir(), 'Rutba Office Samples');
fs.rmSync(docsDir, { recursive: true, force: true });
fs.mkdirSync(docsDir, { recursive: true });
fs.mkdirSync(outDir, { recursive: true });

const at = (name) => path.join(docsDir, name);

/**
 * A workbook with the two things a spreadsheet screenshot has to show: money
 * formatted as money, and a total that is a formula rather than a typed
 * number. Everything else is restraint — a screenshot of forty columns reads
 * as noise at the size a web page renders it.
 */
fs.writeFileSync(
  at('regional-sales.xlsx'),
  buildXlsx({
    sheets: [
      {
        name: 'Sales',
        rows: [
          ['Region', 'Q1', 'Q2', 'Q3', 'Q4', 'Year'],
          ['North', 142000, 161000, 158500, 174200, '=SUM(B2:E2)'],
          ['South', 86000, 91000, 94750, 98300, '=SUM(B3:E3)'],
          ['East', 51200, 57400, 60100, 66800, '=SUM(B4:E4)'],
          ['West', 73900, 71100, 78650, 82400, '=SUM(B5:E5)'],
          ['Total', '=SUM(B2:B5)', '=SUM(C2:C5)', '=SUM(D2:D5)', '=SUM(E2:E5)', '=SUM(F2:F5)'],
          [],
          ['Growth, Q4 on Q1', '=E6/B6-1'],
        ],
        styles: {
          'A1:F1': { bold: true, fill: '#1F3A5F', colour: '#FFFFFF', border: true },
          'A2:A5': { bold: true },
          'B2:F5': { numFmt: '#,##0' },
          // Declared after the number format and repeated for the numeric
          // half, because the last matching declaration wins: a total row
          // styled as a block silently dropped the format from its own
          // figures and printed 353100 beside 142,000.
          'A6:F6': { bold: true, fill: '#EEF2F7', border: true },
          'B6:F6': { bold: true, fill: '#EEF2F7', border: true, numFmt: '#,##0' },
          'A8:A8': { bold: true },
          'B8:B8': { numFmt: '0.0%', bold: true, colour: '#0E7C66' },
        },
      },
    ],
  }),
);

/** A document with a heading hierarchy, because a screenshot of one paragraph
 * of Lorem Ipsum tells a reader nothing about a word processor. */
fs.writeFileSync(
  at('quarterly-review.docx'),
  buildDocx({
    styles: true,
    paragraphs: [
      { text: 'Quarterly Review', style: 'Title', bold: true, size: 30 },
      { text: 'Prepared for the board — figures are illustrative', size: 20 },
      { text: 'Summary', style: 'Heading1', bold: true, size: 26 },
      {
        text:
          'The northern region grew fastest in absolute terms, adding £32,200 between the first and fourth quarters. Growth was steadier in the east, which finished the year 30% ahead of where it started without a single down quarter.',
      },
      {
        text:
          'Churn fell to 2.1% across the year, from 3.4%. Most of the improvement came in the second half and is attributable to the change in onboarding rather than to pricing.',
      },
      { text: 'What we are watching', style: 'Heading1', bold: true, size: 26 },
      {
        text:
          'Two things. The west was flat in the second quarter and has not fully recovered its trend, and the cost of servicing the largest three accounts has risen faster than the revenue they bring.',
      },
      {
        text:
          'Neither is urgent. Both would become urgent if the pattern repeats in the first quarter of next year, which is when we should look again.',
      },
    ],
  }),
);

/** Three slides, so the thumbnail rail in the screenshot is not one item. */
fs.writeFileSync(
  at('product-review.pptx'),
  buildPptx({
    title: 'Quarterly Review',
    slides: [
      { layout: 'title', title: 'Quarterly Review', body: 'Illustrative figures, prepared for the board' },
      { layout: 'obj', title: 'Where growth came from', body: ['North, in absolute terms', 'East, in consistency', 'West is flat and worth watching'] },
      { layout: 'obj', title: 'What changed', body: ['Churn 3.4% to 2.1%', 'Onboarding rewritten in Q3', 'Pricing unchanged'] },
    ],
  }),
);

/**
 * Sample pictures, generated here.
 *
 * **This exists because of what happened without it.** Given no file, the
 * Pictures app opens the operating system's own pictures folder — so the
 * first run of this script photographed the machine owner's family
 * photographs and portraits, filenames included, and those frames were one
 * commit away from a public website. Nothing was published; the capture was
 * deleted the moment it was looked at.
 *
 * The fix is not to remember to check. It is that these three apps are never
 * again pointed at a real folder: they open a directory this script created,
 * containing images this script drew. A generated gradient is a poor
 * photograph and an excellent guarantee.
 */
const crcTable = (() => {
  const t = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c;
  }
  return t;
})();

function crc32(buf) {
  let c = -1;
  for (let i = 0; i < buf.length; i++) c = crcTable[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
}

function chunk(type, data) {
  const head = Buffer.alloc(8);
  head.writeUInt32BE(data.length, 0);
  head.write(type, 4, 'ascii');
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(Buffer.concat([Buffer.from(type, 'ascii'), data])), 0);
  return Buffer.concat([head, data, crc]);
}

/** A diagonal two-colour gradient as a real, valid PNG. */
function gradientPng(width, height, from, to) {
  const raw = Buffer.alloc((width * 3 + 1) * height);
  let o = 0;
  for (let y = 0; y < height; y++) {
    raw[o++] = 0; // filter: none
    for (let x = 0; x < width; x++) {
      const t = (x / width + y / height) / 2;
      for (let c = 0; c < 3; c++) raw[o++] = Math.round(from[c] + (to[c] - from[c]) * t);
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 2; // colour type: truecolour
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', zlib.deflateSync(raw, { level: 9 })),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

const picturesDir = path.join(docsDir, 'Sample Pictures');
fs.mkdirSync(picturesDir, { recursive: true });
const SAMPLES = [
  ['gradient-teal.png', [16, 74, 92], [126, 214, 214]],
  ['gradient-navy.png', [14, 26, 56], [86, 110, 178]],
  ['gradient-sand.png', [122, 92, 52], [232, 214, 178]],
  ['gradient-rose.png', [104, 32, 62], [226, 154, 176]],
];
/**
 * Real photographs, when there are any to use.
 *
 * `RUTBA_SCREEN_PICTURES=<dir>` copies images from that directory into the
 * sample folder instead of drawing gradients — so a set of properly licensed
 * stock photographs can be dropped in without touching this file. The bar for
 * what goes in that directory is the same bar as any published marketing
 * asset: a licence that permits it, and a model release for anybody
 * recognisable in it. That is why the default is a gradient rather than
 * whatever happens to be on the machine.
 */
const supplied = process.env.RUTBA_SCREEN_PICTURES;
let usedSupplied = 0;
if (supplied && fs.existsSync(supplied)) {
  for (const name of fs.readdirSync(supplied)) {
    if (!/.(jpe?g|png|webp|avif|gif)$/i.test(name)) continue;
    fs.copyFileSync(path.join(supplied, name), path.join(picturesDir, name));
    usedSupplied += 1;
  }
}

if (!usedSupplied) {
  for (const [name, from, to] of SAMPLES) {
    fs.writeFileSync(path.join(picturesDir, name), gradientPng(1200, 800, from, to));
  }
}

const firstPicture = () => path.join(picturesDir, fs.readdirSync(picturesDir).sort()[0]);

/** app -> the file it should open, or null for its empty state. */
const PLAN = [
  ['home', null],
  ['word', at('quarterly-review.docx')],
  ['sheets', at('regional-sales.xlsx')],
  ['slides', at('product-review.pptx')],
  ['mail', null],
  // NEVER null. Null means the app falls back to the operating system's
  // own folders, which is how somebody's family photographs ended up in a
  // build artefact. See the note above `gradientPng`.
  ['pictures', firstPicture()],
  ['image', firstPicture()],
  // Video has no generated sample — this suite writes no video — so it is
  // captured empty rather than pointed anywhere real.
  ['video', null],
];

function capture(app, file, profile) {
  return new Promise((resolve) => {
    const env = {
      ...process.env,
      RUTBA_OFFICE_SMOKE: '1',
      RUTBA_SMOKE_OUT: outDir,
      ELECTRON_DISABLE_SECURITY_WARNINGS: '1',
    };
    if (file) env.RUTBA_SMOKE_FILE = file;
    else delete env.RUTBA_SMOKE_FILE;
    // Mail is the one app whose empty state is not worth a picture, so it gets
    // the smoke seeder's invented correspondents.
    if (app === 'mail') env.RUTBA_SMOKE_SEED = '1';
    else delete env.RUTBA_SMOKE_SEED;

    const child = spawn(electron, [appDir, `--user-data-dir=${profile}`, app], {
      stdio: 'ignore',
      env,
    });
    child.on('exit', (code) => resolve(code === 0));
  });
}

const results = [];
for (const [app, file] of PLAN) {
  // A profile per app, not per run: the mail seeder is idempotent-ish but the
  // recent-files list is not, and a "recently opened" panel full of temp paths
  // is exactly the detail that makes a screenshot look like a test artefact.
  const profile = path.join(work, `profile-${app}`);
  fs.mkdirSync(profile, { recursive: true });
  const ok = await capture(app, file, profile);
  const png = path.join(outDir, `${app}.png`);
  const size = fs.existsSync(png) ? fs.statSync(png).size : 0;
  results.push({ app, ok: ok && size > 5000, size });
  process.stdout.write(
    `${ok && size > 5000 ? 'ok  ' : 'FAIL'} ${app.padEnd(9)} ${String(Math.round(size / 1024)).padStart(5)} KB\n`,
  );
}

fs.rmSync(work, { recursive: true, force: true });

const failed = results.filter((r) => !r.ok);
console.log(`\n${results.length - failed.length}/${results.length} captured into ${outDir}`);
process.exit(failed.length ? 1 : 0);
