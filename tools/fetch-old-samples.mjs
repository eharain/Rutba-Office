// Old spreadsheets and databases from public test collections, for checking
// that Worksheets opens them and Data Files imports them (owner, 2026-10-21:
// "if we can find few publicly available"). Fetched on demand into
// .corpus/old/<source>/, which git ignores; nothing here is committed but this
// list. Each source keeps its own licence, written beside its files.
//
//   node tools/fetch-old-samples.mjs
//
// LibreOffice's are its import tests, several of them files made to break a
// reader (fuzzing and security cases): they are here to prove a bad file is
// refused, not to be read.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT = path.join(ROOT, '.corpus', 'old');

const SOURCES = [
  {
    id: 'libreoffice',
    licence: 'MPL-2.0',
    home: 'https://github.com/LibreOffice/core/tree/master/sc/qa/unit/data',
    note: 'LibreOffice Calc import tests. Files named ofz*, CVE* and sf_* are fuzzing and security cases.',
    dirs: ['123', 'dif', 'wks/pass', 'slk/pass', 'dbf/pass'].map((d) => `LibreOffice/core/contents/sc/qa/unit/data/${d}`),
  },
  {
    id: 'infused-dbf',
    licence: 'MIT',
    home: 'https://github.com/infused/dbf/tree/master/spec/fixtures',
    note: 'dBASE and FoxPro tables with their memo files; each *_summary.txt lists the fields a table should have.',
    dirs: ['infused/dbf/contents/spec/fixtures'],
    keep: /\.(dbf|dbt|fpt|txt)$/i,
  },
  {
    id: 'jackcess',
    licence: 'Apache-2.0',
    home: 'https://github.com/jahlborn/jackcess/tree/master/src/test/data',
    note: 'Access databases from the Jackcess project, Access 97 to 2019.',
    files: [
      'V1997/testV1997.mdb', 'V2000/testV2000.mdb', 'V2003/testV2003.mdb', 'V2003/testUnicodeCompV2003.mdb',
      'V2007/testV2007.accdb', 'V2007/oldDatesV2007.accdb', 'V2010/testV2010.accdb', 'V2010/complexDataTestV2010.accdb',
      'V2010/calcFieldTestV2010.accdb', 'V2019/extDateTestV2019.accdb',
    ].map((f) => `https://raw.githubusercontent.com/jahlborn/jackcess/master/src/test/data/${f}`),
  },
];

async function listed(api, keep) {
  const r = await fetch(`https://api.github.com/repos/${api}`, { headers: { accept: 'application/vnd.github+json' } });
  if (!r.ok) throw new Error(`${api}: ${r.status}`);
  return (await r.json()).filter((x) => x.type === 'file' && !x.name.startsWith('.') && (!keep || keep.test(x.name))).map((x) => x.download_url);
}

let total = 0;
for (const src of SOURCES) {
  const dir = path.join(OUT, src.id);
  fs.mkdirSync(dir, { recursive: true });
  const urls = [...(src.files || [])];
  for (const api of src.dirs || []) urls.push(...(await listed(api, src.keep)));
  for (const url of urls) {
    const name = decodeURIComponent(url.split('/').pop());
    const to = path.join(dir, name);
    if (fs.existsSync(to)) continue;
    const r = await fetch(url);
    if (!r.ok) { console.log(`  skipped ${name}: ${r.status}`); continue; }
    const bytes = Buffer.from(await r.arrayBuffer());
    fs.writeFileSync(to, bytes);
    total += bytes.length;
  }
  fs.writeFileSync(path.join(dir, 'SOURCE.md'), `# ${src.id}\n\nFrom ${src.home}\n\nLicence: ${src.licence}. ${src.note}\n`);
  console.log(`${src.id}: ${fs.readdirSync(dir).length - 1} files (${src.licence})`);
}
console.log(`fetched ${(total / 1048576).toFixed(1)} MB into ${OUT}`);
