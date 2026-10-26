// The Strict saves Office is asked to reopen (tools/check-strict-office.ps1):
// each rich fixture saved by Rutba Office as Strict Open XML, and each of
// Word's, Excel's and PowerPoint's own Strict fixtures edited here and saved,
// so staying Strict.
//
//   node tools/strict-saves.mjs <out-dir>
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.resolve(process.argv[2] || path.join(root, 'out', 'strict-saves'));
fs.mkdirSync(out, { recursive: true });
const doc = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });

const EDITS = {
  xlsx: [{ op: 'sheet', name: 'Sales' }, { op: 'setCell', row: 30, col: 0, value: 'Edited in Rutba Office' }],
  docx: [{ op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } }, { op: 'insertText', text: 'Edited in Rutba Office. ' }],
  pptx: [{ op: 'setNotes', slide: 0, text: 'Edited in Rutba Office.' }],
};

for (const ext of ['docx', 'xlsx', 'pptx']) {
  const plain = doc.open({ path: path.join(root, 'tests', 'fixtures', 'rich', `showcase.${ext}`) });
  const a = doc.save({ id: plain.id, path: path.join(out, `saved-as-strict.${ext}`), conformance: 'strict' });
  doc.close({ id: plain.id });
  const strict = doc.open({ path: path.join(root, 'tests', 'fixtures', 'strict', `showcase-strict.${ext}`) });
  doc.apply({ id: strict.id, ops: EDITS[ext] });
  const b = doc.save({ id: strict.id, path: path.join(out, `strict-edited.${ext}`) });
  doc.close({ id: strict.id });
  console.log(`${path.basename(a.path)} (${a.conformance}), ${path.basename(b.path)} (${b.conformance})`);
}
