// Worksheets: conditional formatting drawn in the cells.
//
// The showcase workbook's Sales sheet: a colour scale over the regions, data
// bars on the totals and traffic lights on the growth — the bars drawn under
// the numbers, each as long as its share of the range (Excel 2010's from
// nothing to the whole cell, as its half of the rule says), the icons at the
// cells' left in the light each value earns; the numbers still read. Run
// alone with RUTBA_VERIFY_ONLY=cfdraw.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const FIXTURE = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..', '..', 'tests', 'fixtures', 'rich', 'showcase.xlsx');

/**
 * @param {object} h the harness: open, check, until, errorsIn
 * @param {{ dir: string }} args where the copy is opened from
 */
export async function verifySheetCf(h, { dir }) {
  const { open, check, until, errorsIn } = h;
  const file = path.join(dir, 'cf-showcase.xlsx');
  try {
    fs.copyFileSync(FIXTURE, file);
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="F13"] .sh-cf-bar'))`), 'the data bars', 8000).catch(() => {});
    if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'sheet-cf.png'), (await win.webContents.capturePage()).toPNG());
    const seen = await js(`(() => {
      const cell = (ref) => document.querySelector('.sh-cell[data-ref="' + ref + '"]');
      const bar = (ref) => { const b = cell(ref)?.querySelector('.sh-cf-bar'); return b ? { width: b.getBoundingClientRect().width, of: cell(ref).getBoundingClientRect().width, look: getComputedStyle(b).backgroundImage } : null; };
      const icon = (ref) => { const i = cell(ref)?.querySelector('.sh-cf-icon'); return i ? { set: i.dataset.set, index: Number(i.dataset.index), fill: i.querySelector('circle')?.getAttribute('fill') } : null; };
      return { f2: bar('F2'), f5: bar('F5'), f13: bar('F13'), f5text: cell('F5')?.innerText.trim(), h3: icon('H3'), h9: icon('H9'), h11: icon('H11'), h3text: cell('H3')?.innerText.trim() };
    })()`);
    check('sheets: data bars are drawn under their numbers, each as long as its share, from nothing to the whole cell',
      seen.f2 && seen.f5 && seen.f13 && seen.f2.width < 2 && seen.f5.width > seen.f2.width && seen.f13.width > seen.f5.width && seen.f13.width > seen.f13.of * 0.9
        && /linear-gradient/.test(seen.f5.look) && seen.f5text === '10,086',
      `F2 ${seen.f2?.width.toFixed(1)}, F5 ${seen.f5?.width.toFixed(1)}, F13 ${seen.f13?.width.toFixed(1)} of ${seen.f13?.of.toFixed(1)} px; ${seen.f5?.look.slice(0, 40)}; F5 reads ${JSON.stringify(seen.f5text)}`);
    check('sheets: an icon set\'s icons are drawn at their cells\' left, each in the light its value earns',
      seen.h3?.set === '3TrafficLights1' && seen.h3.index === 2 && seen.h11?.index === 0 && seen.h9?.index === 1 && seen.h3.fill !== seen.h11.fill && seen.h3text === '11.4%',
      `H3 ${JSON.stringify(seen.h3)}, H9 ${JSON.stringify(seen.h9)}, H11 ${JSON.stringify(seen.h11)}; H3 reads ${JSON.stringify(seen.h3text)}`);
    const complaints = await errorsIn(win);
    check('sheets: drawing conditional formatting reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the conditional formatting checks ran', false, err.message);
  }
}
