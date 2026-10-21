// Worksheets: words with looks of their own in a cell (owner, 2026-10-21:
// "the excel formatted contents or styled text does not reload same after
// save, close reopen"). A workbook whose cell has a plain start, an italic
// red middle and a plain end, and another with a superscript, opens with
// those runs drawn as they look; the first cell's words retyped keep each
// run's look round the change; saved and opened again, they are still there.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..', '..');

const CELL = 'B18';
const SUPER = 'B21';

export async function verifySheetRuns(h, { dir }) {
  const { open, check, until, wait, doc } = h;
  const file = path.join(dir, 'runs.xlsx');
  try {
    fs.copyFileSync(path.join(ROOT, 'tests', 'corpus-public', 'xlsx', 'office-showcase.xlsx'), file);
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'sheet' && s.path === file).pop();
    await until(() => js(`[...document.querySelectorAll('.sh-tab')].some((t) => t.textContent.trim() === 'Data types')`), 'the sheet tabs', 8000).catch(() => {});
    await js(`[...document.querySelectorAll('.sh-tab')].find((t) => t.textContent.trim() === 'Data types')?.click(), 1`);
    await until(() => js(`document.querySelectorAll('.sh-cell[data-ref="${CELL}"] .sh-run').length >= 3`), 'the runs drawn', 6000).catch(() => {});
    const drawn = await js(`(() => {
      const runs = [...document.querySelectorAll('.sh-cell[data-ref="${CELL}"] .sh-run')].map((s) => ({ text: s.textContent, italic: s.style.fontStyle === 'italic', colour: s.style.color }));
      const sup = [...document.querySelectorAll('.sh-cell[data-ref="${SUPER}"] .sh-run')].map((s) => ({ text: s.textContent, up: s.style.verticalAlign }));
      return { runs, sup };
    })()`);
    const red = drawn.runs.find((r) => r.italic);
    check("sheets: a cell's words with looks of their own are drawn run by run: the italic red middle, and a superscript",
      drawn.runs.length >= 3 && red && /255, 0, 0/.test(red.colour) && drawn.sup.some((r) => r.text === '2' && r.up === 'super'),
      JSON.stringify(drawn));

    // Retyped: the end changed, the italic red run kept.
    await doc.apply({ id: session.id, ops: [{ op: 'setCell', row: 17, col: 1, value: 'Bold start, red middle, plain end changed' }] });
    const runsNow = () => doc.model({ id: session.id }).cells?.find((c) => c.ref === CELL)?.runs || null;
    const kept = runsNow();
    await doc.save({ id: session.id, path: file });
    const again = await doc.open({ path: file });
    await doc.apply({ id: again.id, ops: [{ op: 'sheet', name: 'Data types' }, { op: 'select', row: 17, col: 1 }] });
    const reopened = doc.model({ id: again.id }).cells?.find((c) => c.ref === CELL)?.runs || null;
    doc.close({ id: again.id });
    const italicRed = (runs) => (runs || []).some((r) => r.italic && r.colour === '#FF0000' && /red middl/.test(r.text));
    check("sheets: retyping a cell's words keeps each run's look round the change, and saved and opened again the runs are there",
      italicRed(kept) && (kept || []).map((r) => r.text).join('') === 'Bold start, red middle, plain end changed' && italicRed(reopened),
      JSON.stringify({ kept: kept?.map((r) => [r.text, r.italic, r.colour]), reopened: reopened?.map((r) => [r.text, r.italic, r.colour]) }));
    await wait(100);
  } catch (err) {
    check('sheets: the runs in cells ran', false, err.message);
  }
}
