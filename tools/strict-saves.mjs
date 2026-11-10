// The saves Office is asked to reopen (tools/check-strict-office.ps1, the
// gate's office:reopen pass): each rich fixture saved by Rutba Office as it
// is, saved after an edit, and saved as Strict Open XML, and each of Word's,
// Excel's and PowerPoint's own Strict fixtures edited here and saved, so
// staying Strict.
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

/** Open `from`, make `edits`, save to `name` in the out folder; the conformance the save chose. */
function saved(from, name, { edits = null, conformance } = {}) {
  const opened = doc.open({ path: from });
  try {
    if (edits) doc.apply({ id: opened.id, ops: edits });
    const r = doc.save({ id: opened.id, path: path.join(out, name), ...(conformance ? { conformance } : {}) });
    return `${path.basename(r.path)} (${r.conformance})`;
  } finally {
    doc.close({ id: opened.id });
  }
}

for (const ext of ['docx', 'xlsx', 'pptx']) {
  const rich = path.join(root, 'tests', 'fixtures', 'rich', `showcase.${ext}`);
  const strict = path.join(root, 'tests', 'fixtures', 'strict', `showcase-strict.${ext}`);
  console.log([
    saved(rich, `saved.${ext}`),
    saved(rich, `edited.${ext}`, { edits: EDITS[ext] }),
    saved(rich, `saved-as-strict.${ext}`, { conformance: 'strict' }),
    saved(strict, `strict-edited.${ext}`, { edits: EDITS[ext] }),
  ].join(', '));
}
