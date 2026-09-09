// Pack a Flat OPC document — the XML Word hands over as Document.WordOpenXML —
// into a .docx.
//
//   node tools/flat-opc-to-docx.mjs <flat.xml> <out.docx>
//
// Why this exists: on the machine that makes the rich fixtures, Word's
// SaveAs never returns from an automated session, in any variant tried —
// hidden or minimised, add-ins on or off, background save off, SaveAs or
// SaveAs2, any format. WordOpenXML is the same package as a Flat OPC string
// and involves no save at all, so the document is taken that way and packed
// here: every pkg:part becomes an entry, xmlData as XML with its declaration,
// binaryData decoded from base64, and [Content_Types].xml is built from the
// content types the parts declare.
import fs from 'node:fs';
import { ZipEntry, writeZip } from '../packages/ooxml/src/zip.js';

const [, , input, output] = process.argv;
if (!input || !output) {
  console.error('usage: node tools/flat-opc-to-docx.mjs <flat.xml> <out.docx>');
  process.exit(2);
}

const flat = fs.readFileSync(input, 'utf8');
const parts = [];
const partRe = /<pkg:part\b([^>]*)>([\s\S]*?)<\/pkg:part>/g;
for (const m of flat.matchAll(partRe)) {
  const attrs = m[1];
  const name = /pkg:name="([^"]+)"/.exec(attrs)?.[1];
  const contentType = /pkg:contentType="([^"]+)"/.exec(attrs)?.[1];
  if (!name) continue;
  const body = m[2];
  const xml = /<pkg:xmlData\b[^>]*>([\s\S]*?)<\/pkg:xmlData>/.exec(body);
  const bin = /<pkg:binaryData\b[^>]*>([\s\S]*?)<\/pkg:binaryData>/.exec(body);
  let data;
  if (xml) data = Buffer.from('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml[1].trim(), 'utf8');
  else if (bin) data = Buffer.from(bin[1].replace(/\s+/g, ''), 'base64');
  else continue;
  parts.push({ name: name.replace(/^\//, ''), contentType, data });
}
if (!parts.some((p) => p.name === 'word/document.xml')) {
  console.error('no word/document.xml among the parts — not a Flat OPC document');
  process.exit(1);
}

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
const types =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  parts
    .filter((p) => p.contentType && !p.name.endsWith('.rels'))
    .map((p) => `<Override PartName="/${esc(p.name)}" ContentType="${esc(p.contentType)}"/>`)
    .join('') +
  '</Types>';

const entry = (name, data) => {
  const e = new ZipEntry({ name, method: 8, crc: 0, compressedSize: 0, uncompressedSize: 0, compressed: Buffer.alloc(0), dosTime: 0, dosDate: 0, flags: 0, externalAttrs: 0, comment: Buffer.alloc(0) });
  e.data = data;
  return e;
};
const entries = [entry('[Content_Types].xml', Buffer.from(types, 'utf8')), ...parts.map((p) => entry(p.name, p.data))];
fs.writeFileSync(output, writeZip(entries));
console.log(`${output}: ${parts.length} parts, ${fs.statSync(output).size} bytes`);
