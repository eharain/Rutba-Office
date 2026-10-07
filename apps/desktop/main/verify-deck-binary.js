// Presentations: a PowerPoint 97-2003 presentation, opened in full.
//
// A .ppt PowerPoint itself wrote (tests/fixtures/binary/showcase.ppt) opens
// as its eight slides — the title on its blue background, the bullets, the
// shapes in their colours with their words (the ones PowerPoint saved as
// freeforms drawn from their own points), the pictures, the notes — where
// it used to open as one slide saying it could not be converted. Run alone
// with RUTBA_VERIFY_ONLY=ppt.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const FIXTURE = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..', '..', 'tests', 'fixtures', 'binary', 'showcase.ppt');

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the copy is opened from
 */
export async function verifyDeckBinary(h, { dir }) {
  const { open, check, until, wait, errorsIn, doc, sessionFor } = h;
  const file = path.join(dir, 'showcase.ppt');
  try {
    fs.copyFileSync(FIXTURE, file);
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('deck');
    const capture = async (name) => { if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await win.webContents.capturePage()).toPNG()); };
    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 8`), 'the eight slides', 8000).catch(() => {});
    await wait(400);
    await capture('deck-binary-1.png');
    const model = () => doc.model({ id: session.id });
    const first = await js(`(() => ({ thumbs: document.querySelectorAll('.sl-thumb').length, text: document.querySelector('.sl-stage')?.innerText || '', svg: document.querySelector('.sl-stage svg')?.outerHTML.length || 0 }))()`);
    const slideTexts = JSON.stringify(model().slide?.shapes || []);
    check('slides: a PowerPoint 97-2003 presentation opens as its eight slides, not a "could not be converted" slide',
      session.converted?.from === 'ppt' && !session.converted?.partial && first.thumbs === 8 && /Rutba Office showcase/.test(slideTexts + first.text),
      `from ${session.converted?.from}${session.converted?.partial ? ' (partial)' : ''}; ${first.thumbs} slides; title ${/Rutba Office showcase/.test(slideTexts)}`);

    // Slide 3: the shapes, the freeforms among them.
    await js(`(() => { document.querySelectorAll('.sl-thumb')[2]?.click(); return 1; })()`);
    await until(() => js(`document.querySelectorAll('.sl-thumb')[2]?.classList.contains('active')`), 'the third slide', 4000).catch(() => {});
    await wait(400);
    await capture('deck-binary-3.png');
    // What the stage draws: the slide's SVG, its words among it.
    const all = await js(`document.querySelector('.sl-stage svg')?.textContent || ''`);
    const shapes = await js(`[...document.querySelectorAll('.sl-stage .sl-hit')]`).then(() => js(`document.querySelectorAll('.sl-stage .sl-hit').length`));
    const wanted = ['Rectangle', 'Oval', 'Star', 'Chevron', 'Heart', 'Cloud', 'Diamond'];
    check('slides: the .ppt\'s shapes come with their words, the freeforms among them',
      wanted.every((w) => all.includes(w)), `${shapes} shapes; missing ${wanted.filter((w) => !all.includes(w)).join(', ') || 'none'}`);
    // Slide 4: the table, a table again — five rows of four, each cell its words.
    await js(`(() => { document.querySelectorAll('.sl-thumb')[3]?.click(); return 1; })()`);
    await until(() => js(`document.querySelectorAll('.sl-thumb')[3]?.classList.contains('active')`), 'the fourth slide', 4000).catch(() => {});
    await wait(400);
    await capture('deck-binary-4.png');
    const table = (doc.model({ id: session.id, slide: 3 }).slide?.shapes || []).find((s) => s.kind === 'table');
    const drawn = await js(`document.querySelector('.sl-stage svg')?.textContent || ''`);
    check('slides: the .ppt\'s table opens as a table, its cells with their words',
      table?.table?.rows === 5 && table.table.cols === 4 && ['Region', 'North', '7,590'].every((w) => drawn.includes(w)),
      table ? `${table.table?.rows} rows of ${table.table?.cols}; words drawn ${['Region', 'North', '7,590'].filter((w) => drawn.includes(w)).length}/3` : 'no table on the slide');
    const notesShown = (model().slide?.notes ?? model().notes ?? '') || '';
    await js(`(() => { document.querySelectorAll('.sl-thumb')[0]?.click(); return 1; })()`);
    await until(() => js(`document.querySelectorAll('.sl-thumb')[0]?.classList.contains('active')`), 'the first slide again', 4000).catch(() => {});
    const notes = JSON.stringify(model()).includes('Speaker notes for the title slide');
    check('slides: the .ppt\'s speaker notes came with it', notes, notes ? 'the title slide\'s notes' : `no notes found (${String(notesShown).slice(0, 60)})`);

    const complaints = await errorsIn(win);
    check('slides: opening a .ppt reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('slides: the .ppt checks ran', false, err.message);
  }
}
