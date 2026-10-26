// Strict Open XML from the app menu (2.19.0, C2): Documents, Worksheets and
// Presentations (which gains Save as… itself) each offer Save as Strict Open
// XML beside Save as…, named as Word, Excel and PowerPoint name the form, and
// the file it writes is Strict; a Strict file then offers the usual form
// instead, and Save as… keeps it Strict. Run alone with
// RUTBA_VERIFY_ONLY=strictsave.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { OoxmlPackage } from '@rutba/ooxml';

const FIXTURES = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..', '..', 'tests', 'fixtures', 'rich');
/** The form of a file on disk: 'strict', 'transitional', or 'none' when it cannot be read. */
function formOf(file) {
  try {
    return OoxmlPackage.read(fs.readFileSync(file)).conformance;
  } catch {
    return 'none';
  }
}

const APPS = [
  { app: 'word', ext: 'docx', strict: 'Save as Strict Open XML Document…', usual: 'Save as Word Document…', ready: `Boolean(document.querySelector('.wd-page [data-block="1"]'))` },
  { app: 'sheets', ext: 'xlsx', strict: 'Save as Strict Open XML Spreadsheet…', usual: 'Save as Excel Workbook…', ready: `Boolean(document.querySelector('.rw-appmark'))` },
  { app: 'slides', ext: 'pptx', strict: 'Save as Strict Open XML Presentation…', usual: 'Save as PowerPoint Presentation…', ready: `Boolean(document.querySelector('.rw-appmark'))` },
];

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the files are written
 */
export async function verifyStrictSave({ open, check, until, wait, errorsIn }, { dir }) {
  for (const { app, ext, strict, usual, ready } of APPS) {
    const file = path.join(dir, `usual-form.${ext}`);
    const asStrict = path.join(dir, `strict-form.${ext}`);
    const keptStrict = path.join(dir, `strict-kept.${ext}`);
    const back = path.join(dir, `usual-again.${ext}`);
    for (const f of [asStrict, keptStrict, back]) fs.rmSync(f, { force: true });
    fs.copyFileSync(path.join(FIXTURES, `showcase.${ext}`), file);
    try {
      const win = await open(app, file);
      const js = (code) => win.webContents.executeJavaScript(code);
      const items = () => js(`[...document.querySelectorAll('.rw-menu button')].map((b) => b.textContent.trim())`);
      const menu = async (label) => {
        await js(`document.querySelector('.rw-appmark').click(), 1`);
        return until(async () => (await items()).some((s) => s.startsWith(label)), `the ${label} item`, 5000).then(() => true, () => false);
      };
      const choose = (label, answer) => {
        globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer });
        return js(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith(${JSON.stringify(label)}))?.click(), 1`);
      };
      const written = (f) => until(() => fs.existsSync(f) && fs.statSync(f).size > 0, `${path.basename(f)} written`, 8000).then(() => formOf(f), () => 'none');
      await until(() => js(ready), 'the file', 10000);

      const offered = await menu(strict);
      await choose(strict, asStrict);
      const first = await written(asStrict);
      await wait(400);
      const turned = await menu(usual);
      await choose('Save as…', keptStrict);
      const kept = await written(keptStrict);
      await wait(400);
      await menu(usual);
      await choose(usual, back);
      const again = await written(back);
      await wait(400);
      const offeredAgain = await menu(strict);
      await js(`document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), 1`);
      check(`${app}: Save as Strict Open XML writes Strict, a Strict file offers the usual form and Save as… keeps it Strict, and the usual form goes back to Transitional`,
        offered && first === 'strict' && turned && kept === 'strict' && again === 'transitional' && offeredAgain,
        JSON.stringify({ offered, first, turned, kept, again, offeredAgain }));
      const complaints = await errorsIn(win);
      check(`${app}: the Strict Save As checks report nothing`, complaints.length === 0, complaints.join(' | ') || 'nothing reported');
      // Left open: the last window closing would end the run.
    } catch (err) {
      check(`${app}: the Strict Save As checks ran`, false, err?.message || String(err));
    }
  }
}
