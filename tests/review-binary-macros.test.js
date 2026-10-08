// A Word 97-2003 document carrying a macro project: the project is not
// carried into the .docx the suite writes, and nothing in it is run.
//
// The fixture is a real Word file (structure.doc) with the storage Word
// keeps a VBA project in added to its compound file, holding a module whose
// source has a marker no other part of the file has.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { CompoundFile } from '@rutba/office-formats/cfb';
import { writeCompoundFile } from '@rutba/office-formats/cfb-write';
import { readWordDocument } from '@rutba/office-formats/msword';
import { docModelToDocx } from '@rutba/office-formats/msdoc-docx';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const MARKER = 'Sub AutoOpen_MARKER_7f3a()';

/** structure.doc with a Macros storage added: the project record, its dir and one module. */
function withMacros() {
  const file = new CompoundFile(new Uint8Array(fs.readFileSync(path.join(HERE, 'fixtures', 'binary', 'structure.doc'))));
  const streams = [];
  const walk = (entry, trail) => {
    for (const kid of file.childrenOf(entry)) {
      if (kid.name === '') continue;
      if (kid.type === 1) walk(kid, [...trail, kid.name]);
      else streams.push({ path: [...trail, kid.name], data: file.read(kid) });
    }
  };
  walk(file.root, []);
  const text = (s) => new TextEncoder().encode(s);
  streams.push({ path: ['Macros', 'PROJECT'], data: text('ID="{00000000-0000-0000-0000-000000000000}"\r\nModule=Module1\r\n') });
  streams.push({ path: ['Macros', 'VBA', '_VBA_PROJECT'], data: text('\xcc\x61\xff\xff\x00\x00\x00') });
  streams.push({ path: ['Macros', 'VBA', 'Module1'], data: text(MARKER + '\r\nShell "calc.exe"\r\nEnd Sub') });
  return writeCompoundFile(streams);
}

test('a .doc with a macro project opens as its words, and converts to a .docx holding no macro part, no module and no source', () => {
  const bytes = withMacros();
  assert.ok(Buffer.from(bytes).includes(MARKER), 'the module really is in the file');

  const plain = readWordDocument(new Uint8Array(fs.readFileSync(path.join(HERE, 'fixtures', 'binary', 'structure.doc'))));
  const model = readWordDocument(new Uint8Array(bytes));
  const words = (m) => JSON.stringify(m.body);
  assert.equal(words(model), words(plain), 'the document reads the same with or without the project');

  const docx = Buffer.from(docModelToDocx(model));
  const pkg = OoxmlPackage.read(docx);
  const names = pkg.partNames();
  assert.ok(!names.some((n) => /vba|macro|Macros/i.test(n)), 'no macro part: ' + names.join(', '));
  assert.ok(!docx.includes(MARKER) && !docx.includes('calc.exe'), 'the source is nowhere in the package, stored or not');
  for (const n of names.filter((x) => /\.(xml|rels)$/.test(x))) assert.doesNotMatch(pkg.text(n), /macroEnabled|vbaProject|vbaData/, n);
  assert.match(pkg.text('word/document.xml'), /Second page|Region/, 'the document itself came through');
  assert.ok(openDocx(docx).blocks.length > 0);
});
