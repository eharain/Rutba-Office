// Any Word binary document, whichever Word wrote it.
//
// Word 97 to 2003 write the format msdoc.js reads. Before them: Word 6.0
// and 95 (compound files too, with an older FIB and one-byte sprms), Word
// 2.0 and 1.x for Windows (no compound file — the FIB opens the file
// itself), Word for DOS and Windows Write (a 128-byte header, the text,
// then pages of character and paragraph properties). msdoc-old.js reads
// those into the same model, so one converter writes them all as .docx.

import { CompoundFile } from './cfb.js';
import { readDoc, DocError } from './msdoc.js';
import { readOldWord, oldWordKind } from './msdoc-old.js';

/** Which Word wrote these bytes — 'word97', 'word6', 'word2', 'word1', 'worddos', 'write' — or null. */
export function wordKind(bytes) {
  if (CompoundFile.is(bytes)) {
    const cfb = new CompoundFile(bytes);
    const entry = cfb.find(['WordDocument']);
    if (!entry) return null;
    const wd = cfb.read(entry);
    const ident = wd[0] | (wd[1] << 8);
    const nFib = wd[2] | (wd[3] << 8);
    if (ident === 0xa5ec && nFib >= 0xc1) return 'word97';
    if (ident === 0xa5ec || ident === 0xa5dc) return 'word6';
    return null;
  }
  return oldWordKind(bytes);
}

/** Read any Word binary document into the model msdoc-docx.js writes. */
export function readWordDocument(bytes) {
  const kind = wordKind(bytes);
  if (kind === 'word97') return readDoc(bytes);
  if (kind) return readOldWord(bytes, kind);
  throw new DocError('not a Word document this suite knows');
}
