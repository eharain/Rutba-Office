// Files written to do harm, which damage alone does not make.
//
//   node tools/fuzz-grammar.js            600 of them
//   node tools/fuzz-grammar.js 3000 4242  more, from a seed
//
// fuzz-open.js damages good files byte by byte, and almost none of what that
// makes gets past the zip's checksum to the XML, or past the first line of a
// text format. These are written to be well formed and hostile: a Word part
// whose numbers are as large as a number goes, a sheet with a formula nested
// fifty thousand deep, a message with parts inside parts inside parts, an
// invitation that recurs on no day that exists. Each goes to the parser a
// person's double-click reaches, in a worker with a time and a memory limit,
// and the answer must be a result or a refusal in a sentence.
//
// A last round writes each text format at two sizes. Four times the bytes is
// allowed about four times the time; sixteen times is a loop inside a loop,
// which a person would meet as a window that stops.

import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Worker, isMainThread, parentPort } from 'node:worker_threads';

const FAULT_NAMES = new Set(['RangeError', 'TypeError', 'ReferenceError', 'EvalError', 'InternalError', 'AggregateError', 'SyntaxError', 'AssertionError']);
const FAULT_WORDS = /Maximum call stack|Invalid array length|Array buffer allocation|out of memory|outside the bounds|outside buffer bounds|out of range|Offset is outside|is not a function|is not iterable|Cannot read propert|Cannot set propert|undefined is not|null is not|assertion failed/i;
const RAW_WORDS = /invalid distance|invalid stored block|invalid block type|invalid code lengths|incorrect header check|incorrect data check|unexpected end of|Unexpected token|at parser\.js|\.js line \d+/i;

function judge(err) {
  const links = [];
  for (let e = err, n = 0; e && n < 8; e = e.cause, n++) links.push({ name: e?.name || typeof e, message: String(e?.message ?? e) });
  const fault = links.find((l) => FAULT_NAMES.has(l.name) || FAULT_WORDS.test(l.message));
  const top = links[0]?.message || '';
  const raw = RAW_WORDS.test(top) ? links[0] : null;
  return fault || raw || !/[a-z]{3}.*[ .]/.test(top)
    ? { outcome: 'FAULT', message: `${(fault || raw || links[0]).name}: ${(fault || raw || links[0]).message}`.slice(0, 220) }
    : { outcome: 'refused', message: top.slice(0, 200) };
}

if (!isMainThread) {
  const { parseMessage } = await import('@rutba/mailbox/mime');
  const { readCalendar } = await import('../packages/calendar/src/ical.js');
  const { occurrences, exceptionsOf } = await import('../packages/calendar/src/recur.js');
  const { parseCsv } = await import('../packages/contacts/src/csv.js');
  const vcard = await import('../packages/contacts/src/vcard.js');
  const { createDocumentService } = await import('../apps/desktop/main/documents.js');
  const service = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
  const from = Date.UTC(2020, 0, 1);
  const to = Date.UTC(2030, 0, 1);
  const TARGETS = {
    eml: (bytes) => parseMessage(Buffer.from(bytes)),
    ics: (bytes) => {
      const cal = readCalendar(Buffer.from(bytes).toString('utf8'));
      const siblingsOf = exceptionsOf(cal.events);
      for (const e of cal.events || []) if (!e.recurrenceId) occurrences(e, { from, to, siblings: siblingsOf(e), limit: 500 });
    },
    vcf: (bytes) => vcard.readVCards(Buffer.from(bytes).toString('utf8')),
    'contacts-csv': (bytes) => parseCsv(Buffer.from(bytes).toString('utf8')),
  };
  parentPort.on('message', ({ target, file, bytes }) => {
    const t = performance.now();
    try {
      if (file) {
        const opened = service.open({ path: file });
        if (!opened.locked) {
          service.model({ id: opened.id });
          service.close({ id: opened.id });
        }
      } else TARGETS[target](bytes);
      parentPort.postMessage({ outcome: 'read', ms: performance.now() - t });
    } catch (err) {
      parentPort.postMessage({ ...judge(err), ms: performance.now() - t });
    }
  });
} else if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  await main();
}

async function main() {
  const { readZip, writeZip } = await import('@rutba/ooxml/zip');
  const { buildDocx, buildXlsx } = await import('@rutba/ooxml/build');
  const { buildPptx } = await import('@rutba/presentation');
  const total = Number(process.argv[2] || 600);
  let state = (Number(process.argv[3] || Date.now() % 100000) >>> 0) || 1;
  const seed = state;
  const rand = () => { state ^= state << 13; state ^= state >>> 17; state ^= state << 5; return ((state >>> 0) % 1000000) / 1000000; };
  const pick = (n) => Math.floor(rand() * n);
  const one = (list) => list[pick(list.length)];
  const LIMIT_MS = 5000;

  const NUMBERS = ['-1', '0', '1', '65535', '65536', '2147483647', '2147483648', '4294967295', '4294967296', '99999999999999999999', '1e309', '-1e309', 'NaN', 'Infinity', '0.0000001', ''];
  const WORDS = ['\u0000', '&#0;', '&#x10FFFF;', '&amp;&amp;', ']]>', '"', "'", '<', '%00', '../../x', '\\\\evil\\share', '\ud800'];

  // XML inside an Office file, changed in ways that keep it XML.
  const XML = {
    'a number as large as it goes': (x) => {
      const spots = [...x.matchAll(/(\s[\w:]+=")(-?\d+(?:\.\d+)?)"/g)];
      for (let i = 1 + pick(6); i > 0 && spots.length; i--) {
        const m = one(spots);
        x = x.slice(0, m.index) + m[1] + one(NUMBERS) + '"' + x.slice(m.index + m[0].length);
        break;
      }
      return x;
    },
    'every number the same bad one': (x) => { const v = one(NUMBERS); return x.replace(/(\s[\w:]+=")-?\d+(?:\.\d+)?"/g, `$1${v}"`); },
    'a cell address that is not one': (x) => x.replace(/\sr="([A-Z]+)(\d+)"/, () => ` r="${one(['XFD1048577', 'AAAA1', 'A0', 'A-1', '@', '', 'A1:A1048576'])}"`),
    'an element repeated thousands of times': (x) => {
      const all = [...x.matchAll(/<[\w:]+(?:\s[^<>]*)?\/>/g)];
      if (!all.length) return x;
      const m = one(all);
      return x.slice(0, m.index) + m[0].repeat(2000 + pick(20000)) + x.slice(m.index);
    },
    'elements nested very deep': (x) => {
      const at = x.lastIndexOf('</');
      if (at < 0) return x;
      const n = 2000 + pick(60000);
      return `${x.slice(0, at)}${'<a:x>'.repeat(n)}${'</a:x>'.repeat(n)}${x.slice(at)}`;
    },
    'an element taken out': (x) => {
      const all = [...x.matchAll(/<([\w:]+)(?:\s[^<>]*)?>[^<]*<\/\1>/g)];
      if (!all.length) return x;
      const m = one(all);
      return x.slice(0, m.index) + x.slice(m.index + m[0].length);
    },
    'an attribute taken out': (x) => {
      const all = [...x.matchAll(/\s[\w:]+="[^"]*"/g)];
      if (!all.length) return x;
      const m = one(all);
      return x.slice(0, m.index) + x.slice(m.index + m[0].length);
    },
    'strange words in the text': (x) => x.replace(/>([^<]{1,40})</, () => `>${one(WORDS).replace(/</g, '&lt;')}<`),
    'a formula made to be hard': (x) => {
      const n = 1000 + pick(60000);
      const f = one([`${'('.repeat(n)}1${')'.repeat(n)}`, `${'SUM('.repeat(n)}1${')'.repeat(n)}`, `A1${'+A1'.repeat(n)}`, 'A1:XFD1048576', 'INDIRECT("A1:XFD1048576")', 'OFFSET(A1,0,0,1048576,16384)', 'SUM(A:XFD)*SUM(1:1048576)', 'REPT("x",2147483647)']);
      return x.includes('<f>') ? x.replace(/<f>[^<]*<\/f>/, `<f>${f}</f>`) : x.replace(/<c r="([A-Z]+\d+)"[^>]*>(?:<v>[^<]*<\/v>)?<\/c>/, `<c r="$1"><f>${f}</f></c>`);
    },
    'an entity expansion': (x) => {
      const defs = Array.from({ length: 10 }, (_, i) => `<!ENTITY e${i} "${`&e${i + 1};`.repeat(10)}">`).join('');
      return x.replace(/^(<\?xml[^>]*\?>)?/, `$1<!DOCTYPE d [${defs}<!ENTITY e10 "x">]>`).replace(/>([^<]{1,40})</, '>&e0;<');
    },
  };
  const xmlWays = Object.keys(XML);

  const OFFICE = {
    docx: buildDocx({ title: 'Agreement', paragraphs: ['The supplier shall deliver the goods.', 'Second.', 'Third.'] }),
    xlsx: buildXlsx({ sheets: [{ name: 'S', rows: [['Region', 'Total'], ['North', 12], ['South', 8], [null, '=SUM(B2:B3)']] }] }),
    pptx: buildPptx({ title: 'Deck', slides: [{ layout: 'title', title: 'Rutba' }, { layout: 'titleAndContent', title: 'Two', bullets: ['a', 'b'] }] }),
  };
  const office = (ext) => () => {
    const { entries } = readZip(Buffer.from(OFFICE[ext]));
    const xml = entries.filter((e) => /\.(xml|rels)$/.test(e.name));
    const ways = [];
    for (let i = 1 + pick(3); i > 0; i--) {
      const e = one(xml);
      const way = one(xmlWays);
      e.data = Buffer.from(XML[way](e.data.toString('utf8')), 'utf8');
      ways.push(`${way} in ${e.name}`);
    }
    return { name: `f.${ext}`, bytes: Buffer.from(writeZip(entries)), what: ways.join('; ') };
  };

  const ODS_NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"';
  const odfZip = async (mime, content, ext) => {
    const { ZipEntry } = await import('@rutba/ooxml/zip');
    const entry = (name, text) => {
      const e = new ZipEntry({ name, method: 0, crc: 0, compressedSize: 0, uncompressedSize: 0, compressed: Buffer.alloc(0), dosTime: 0, dosDate: 0x2821 });
      e.data = Buffer.from(text, 'utf8');
      return e;
    };
    return { name: `f.${ext}`, bytes: Buffer.from(writeZip([entry('mimetype', mime), entry('content.xml', content)])) };
  };
  const ods = async () => {
    const big = () => one(['1', '256', '65536', '1048576', '2147483647', '-1', 'NaN', '1e309']);
    const cell = () => `<table:table-cell table:number-columns-repeated="${big()}" table:number-columns-spanned="${big()}" table:number-rows-spanned="${big()}" office:value-type="${one(['float', 'string', 'date', 'x'])}" office:value="${big()}" table:formula="of:=${one(['SUM([.A1:.XFD1048576])', '1/0', `${'('.repeat(5000)}1${')'.repeat(5000)}`])}"><text:p>x</text:p></table:table-cell>`;
    const row = () => `<table:table-row table:number-rows-repeated="${big()}">${Array.from({ length: 1 + pick(5) }, cell).join('')}</table:table-row>`;
    const out = await odfZip('application/vnd.oasis.opendocument.spreadsheet', `<office:document-content ${ODS_NS}><office:body><office:spreadsheet><table:table table:name="S">${Array.from({ length: 1 + pick(6) }, row).join('')}</table:table></office:spreadsheet></office:body></office:document-content>`, 'ods');
    return { ...out, what: 'repeats and spans in a spreadsheet' };
  };
  const odt = async () => {
    const n = 2000 + pick(60000);
    const out = await odfZip('application/vnd.oasis.opendocument.text', `<office:document-content ${ODS_NS}><office:body><office:text>${'<text:list><text:list-item>'.repeat(n)}<text:p>x</text:p>${'</text:list-item></text:list>'.repeat(n)}</office:text></office:body></office:document-content>`, 'odt');
    return { ...out, what: `lists nested ${n} deep` };
  };

  const eml = () => {
    const kind = pick(6);
    const head = 'From: a@example.com\r\nTo: b@example.com\r\nSubject: x\r\nMIME-Version: 1.0\r\n';
    if (kind === 0) {
      const depth = 200 + pick(5000);
      let body = 'Content-Type: text/plain\r\n\r\nhello\r\n';
      for (let i = 0; i < depth; i++) body = `Content-Type: multipart/mixed; boundary="b${i}"\r\n\r\n--b${i}\r\n${body}--b${i}--\r\n`;
      return { target: 'eml', bytes: Buffer.from(head + body), what: `parts inside parts, ${depth} deep` };
    }
    if (kind === 1) {
      const params = Array.from({ length: 1 + pick(4000) }, (_, i) => `name*${one(['0', String(i), '4294967294', '999999999', '-1', 'x'])}${one(['', '*'])}=${one(['a', "utf-8''%E2%82%AC", '%%%', '""'])}`).join('; ');
      return { target: 'eml', bytes: Buffer.from(`${head}Content-Type: application/x; ${params}\r\n\r\nx`), what: 'parameters with hostile indexes' };
    }
    if (kind === 2) return { target: 'eml', bytes: Buffer.from(`${head}Subject: ${'=?utf-8?B?'.repeat(20000 + pick(40000))}\r\n\r\nx`), what: 'encoded words never closed' };
    if (kind === 3) return { target: 'eml', bytes: Buffer.from(`${head}To: ${'"'.repeat(pick(200000))}<${'a,'.repeat(pick(100000))}\r\n\r\nx`), what: 'addresses that never end' };
    if (kind === 4) return { target: 'eml', bytes: Buffer.from(`${head}Content-Type: multipart/mixed; boundary=""\r\n\r\n${'--\r\n'.repeat(100000 + pick(200000))}`), what: 'an empty boundary' };
    return { target: 'eml', bytes: Buffer.from(`${head}Content-Transfer-Encoding: ${one(['base64', 'quoted-printable', 'uuencode'])}\r\n\r\n${'=\r\n'.repeat(pick(300000))}`), what: 'an encoding of nothing' };
  };

  const ics = () => {
    const rule = () => `FREQ=${one(['DAILY', 'WEEKLY', 'MONTHLY', 'YEARLY', 'SECONDLY', 'x'])};INTERVAL=${one(NUMBERS)};COUNT=${one(NUMBERS)}${one(['', ';BYMONTH=13', ';BYDAY=9XX', ';BYMONTHDAY=31;BYMONTH=2', ';BYSETPOS=-99999', ';BYDAY=1MO,2MO,3MO'])}${one(['', ';UNTIL=99991231T000000Z', ';UNTIL=x'])}`;
    const date = () => one(['20260101T090000Z', '00000000T000000Z', '99999999T999999Z', '20260230T090000', 'x', '-20260101']);
    const event = () => `BEGIN:VEVENT\r\nUID:u${pick(4)}\r\nDTSTART:${date()}\r\nDTEND:${date()}\r\nDURATION:${one(['PT1H', 'P99999999999W', 'x'])}\r\nRRULE:${rule()}\r\nRECURRENCE-ID:${date()}\r\nEXDATE:${Array.from({ length: pick(2000) }, date).join(',')}\r\nEND:VEVENT\r\n`;
    const kind = pick(4);
    if (kind === 0) return { target: 'ics', bytes: Buffer.from(`BEGIN:VCALENDAR\r\n${'BEGIN:VEVENT\r\n'.repeat(5000 + pick(50000))}`), what: 'components never ended' };
    if (kind === 1) return { target: 'ics', bytes: Buffer.from(`BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nSUMMARY:${' x\r\n '.repeat(pick(200000))}\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n`), what: 'a line folded for ever' };
    return { target: 'ics', bytes: Buffer.from(`BEGIN:VCALENDAR\r\n${Array.from({ length: 1 + pick(40) }, event).join('')}END:VCALENDAR\r\n`), what: 'events with hostile rules and dates' };
  };

  const vcf = () => {
    const kind = pick(3);
    if (kind === 0) return { target: 'vcf', bytes: Buffer.from('BEGIN:VCARD\r\n'.repeat(5000 + pick(50000))), what: 'cards never ended' };
    if (kind === 1) return { target: 'vcf', bytes: Buffer.from(`BEGIN:VCARD\r\nVERSION:4.0\r\nPHOTO;ENCODING=b:${'A'.repeat(pick(4000000))}\r\nEND:VCARD\r\n`), what: 'a very long picture' };
    return { target: 'vcf', bytes: Buffer.from(`BEGIN:VCARD\r\n${`N;${'X=;'.repeat(pick(2000))}:${';'.repeat(pick(2000))}\r\n`.repeat(1 + pick(40))}END:VCARD\r\n`), what: 'properties full of separators' };
  };

  const textFile = (ext, make, what) => () => ({ name: `f.${ext}`, bytes: Buffer.from(make()), what });
  const TEXT = {
    csv: [() => `"${'x'.repeat(pick(2000000))}`, () => `${'"",'.repeat(pick(800000))}`, () => `${'a,'.repeat(20000)}\n`.repeat(1 + pick(4)), () => `=${'('.repeat(pick(100000))}1`],
    rtf: [() => `{\\rtf1 ${'{'.repeat(pick(400000))}x`, () => `{\\rtf1 ${'\\u-99999999 '.repeat(pick(100000))}}`, () => `{\\rtf1 \\trowd${'\\cellx99999999'.repeat(pick(100000))}\\row}`, () => `{\\rtf1 {\\pict\\wmetafile8\\picw99999999\\pich99999999 ${'ff'.repeat(pick(100000))}}}`, () => `{\\rtf1 \\bin4294967295 x}`],
    html: [() => '<div>'.repeat(pick(300000)), () => `<table>${'<tr><td colspan=99999999 rowspan=99999999>x'.repeat(pick(20000))}`, () => `<p ${'a=1 '.repeat(pick(200000))}>`, () => '<!--'.repeat(pick(300000)), () => `<img src="${'x'.repeat(pick(3000000))}">`],
    md: [() => '>'.repeat(pick(300000)), () => '['.repeat(pick(300000)), () => '*'.repeat(pick(300000)), () => `${'- '.repeat(pick(100000))}x`, () => `|${'a|'.repeat(pick(100000))}\n|${'-|'.repeat(pick(100000))}\n`],
    txt: [() => '\n'.repeat(pick(4000000)), () => '\u0000'.repeat(pick(2000000)), () => 'x'.repeat(pick(8000000))],
  };
  const textMakers = Object.entries(TEXT).flatMap(([ext, list]) => list.map((m, i) => textFile(ext, m, `${ext} pattern ${i}`)));

  const makers = [office('docx'), office('docx'), office('xlsx'), office('xlsx'), office('pptx'), ods, odt, eml, eml, ics, ics, vcf, ...textMakers];

  let worker = null;
  const start = () => {
    worker = new Worker(fileURLToPath(import.meta.url), { resourceLimits: { maxOldGenerationSizeMb: 1024, maxYoungGenerationSizeMb: 64 } });
    worker.unref();
  };
  const run = (job, limit = LIMIT_MS) => new Promise((resolve) => {
    if (!worker) start();
    const w = worker;
    const done = (m) => { clearTimeout(timer); w.off('message', done); w.off('error', failed); resolve(m); };
    const failed = (err) => { worker = null; done({ outcome: 'FAULT', message: `${err?.name}: ${String(err?.message || err)}`.slice(0, 200), ms: limit }); };
    const grace = w.warm ? 0 : 8000;
    w.warm = true;
    const timer = setTimeout(() => { worker = null; w.terminate(); done({ outcome: 'FAULT', message: `did not finish in ${limit / 1000} seconds`, ms: limit }); }, limit + grace);
    w.on('message', done);
    w.on('error', failed);
    w.postMessage(job);
  });

  console.log(`fuzz-grammar: ${total} files, seed ${seed}, ${LIMIT_MS / 1000} s each`);
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-fuzz-grammar-'));
  const counts = { read: 0, refused: 0, FAULT: 0 };
  const faults = [];
  let slowest = { ms: 0 };
  const save = (made) => {
    const name = `${crypto.randomBytes(5).toString('hex')}-${made.name || `x.${made.target}`}`;
    const file = path.join(dir, name);
    fs.writeFileSync(file, made.bytes);
    return file;
  };
  for (let i = 0; i < total; i++) {
    const made = await one(makers)();
    const file = save(made);
    const r = await run(made.target ? { target: made.target, bytes: made.bytes } : { file });
    if (r.outcome !== 'FAULT' && r.ms > LIMIT_MS) Object.assign(r, { outcome: 'FAULT', message: `took ${r.ms.toFixed(0)} ms` });
    counts[r.outcome] = (counts[r.outcome] || 0) + 1;
    if (r.ms > slowest.ms) slowest = { ms: r.ms, what: made.what };
    if (r.outcome === 'FAULT') {
      faults.push({ what: `${made.name || made.target}: ${made.what}`, message: r.message, file });
      console.log(`  ${made.name || made.target}: ${made.what}: ${r.message}`);
    }
    else fs.rmSync(file, { force: true });
  }

  // The same pattern at two sizes, a quarter apart: the time should grow as the bytes do.
  // A document costs far more for each byte than mail does, so it is measured smaller.
  const fit = (seedText, size) => seedText.repeat(Math.ceil(size / seedText.length)).slice(0, size);
  const SCALE = {
    'eml body lines': ['eml', 'From: a@b.c\r\n\r\n', 'line of body text =\r\n'],
    'eml headers': ['eml', '', 'X-Header: value\r\n'],
    'eml addresses': ['eml', 'To: ', 'a@b.c, '],
    'ics events': ['ics', 'BEGIN:VCALENDAR\r\n', 'BEGIN:VEVENT\r\nUID:a\r\nDTSTART:20260101T090000Z\r\nEND:VEVENT\r\n'],
    'ics exdates': ['ics', 'BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:a\r\nDTSTART:20200101T090000Z\r\nRRULE:FREQ=DAILY\r\nEXDATE:', '20260101T090000Z,'],
    'vcf cards': ['vcf', '', 'BEGIN:VCARD\r\nFN:a\r\nEND:VCARD\r\n'],
    'contacts csv': ['contacts-csv', 'name,email\r\n', 'a,b@c.d\r\n'],
    'csv document': ['f.csv', '', '1,2,3,4,5\n'],
    'rtf document': ['f.rtf', '{\\rtf1 ', '\\par some words '],
    'html document': ['f.html', '', '<p>some words</p>'],
    'markdown document': ['f.md', '', '- some words\n'],
    'text document': ['f.txt', '', 'some words on a line\n'],
  };
  const timed = async (kind, bytes) => {
    if (kind.startsWith('f.')) return run({ file: save({ name: kind, bytes }) }, 60000);
    return run({ target: kind, bytes }, 60000);
  };
  for (const [what, [kind, head, unit]] of Object.entries(SCALE)) {
    const base = kind.startsWith('f.') && kind !== 'f.csv' && kind !== 'f.txt' ? 256 * 1024 : 1024 * 1024;
    const small = await timed(kind, Buffer.from(head + fit(unit, base)));
    const large = await timed(kind, Buffer.from(head + fit(unit, base * 4)));
    counts.read += 2;
    const fault = small.outcome === 'FAULT' ? small : large.outcome === 'FAULT' ? large : null;
    console.log(`  scale ${what}: ${base / 1024} KB ${small.ms.toFixed(0)} ms, four times that ${large.ms.toFixed(0)} ms${fault ? `, ${fault.message}` : ''}`);
    if (fault) faults.push({ what: `scale, ${what}`, message: fault.message, file: '' });
    else if (large.ms > 2000 && large.ms > small.ms * 8) faults.push({ what: `scale, ${what}`, message: `${base / 1024} KB took ${small.ms.toFixed(0)} ms and four times that took ${large.ms.toFixed(0)} ms: a loop inside a loop`, file: '' });
  }
  counts.FAULT = faults.length;

  await worker?.terminate();
  console.log(`${total + Object.keys(SCALE).length * 2} damaged files: read ${counts.read}, refused in a sentence ${counts.refused}, faults ${counts.FAULT}`);
  console.log(`slowest: ${slowest.ms.toFixed(0)} ms (${slowest.what})`);
  if (faults.length) {
    console.log(`faults (files kept in ${dir}):`);
    process.exit(1);
  }
  fs.rmSync(dir, { recursive: true, force: true });
  console.log('every file was read or refused in a sentence, each within its time.');
}
