// OpenDocument charts, shapes and gradients, in the windows.
//
// showcase.ods and showcase.odp are what Excel and PowerPoint themselves
// saved in ODF (tests/fixtures/rich). The workbook opens with its charts on
// the Charts sheet — drawn, plotting their cells — and its shapes, picture
// and text box below them; the deck opens with its first slide's gradient
// background and its shapes' gradients drawn. Run alone with
// RUTBA_VERIFY_ONLY=odfcharts.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const FIXTURES = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..', '..', 'tests', 'fixtures', 'rich');

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the files are copied
 */
export async function verifyOdfCharts(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const capture = async (win, name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE) return;
    win.webContents.invalidate();
    await wait(400);
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await win.webContents.capturePage()).toPNG());
  };
  try {
    const ods = path.join(dir, 'showcase.ods');
    fs.copyFileSync(path.join(FIXTURES, 'showcase.ods'), ods);
    const sheets = await open('sheets', ods);
    const js = (code) => sheets.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.sh-tab'))`), 'the grid', 15000);
    await js(`(() => { [...document.querySelectorAll('.sh-tab')].find((t) => t.textContent.trim() === 'Charts')?.click(); return 1; })()`);
    const drawn = await until(() => js(`document.querySelectorAll('.sh-drawing[data-kind="chart"] svg').length >= 4`), 'the four charts drawn', 10000).then(() => true, () => false);
    const seen = await js(`(() => {
      const charts = [...document.querySelectorAll('.sh-drawing[data-kind="chart"]')];
      return { charts: charts.length, words: charts.map((c) => c.textContent.replace(/\\s+/g, ' ').trim()).join(' | ').slice(0, 300) };
    })()`);
    check('odf: an .ods from Excel opens with its charts drawn on their sheet, titled and plotting their months',
      drawn && seen.charts === 4 && /Sales by region, by month/.test(seen.words) && /Jan/.test(seen.words), JSON.stringify(seen));
    await capture(sheets, 'odf-charts.png');
    // Down to the shapes, the picture and the text box.
    await js(`(() => { document.querySelector('.sh-grid').scrollTop = 640; return 1; })()`);
    await wait(500);
    const below = await until(() => js(`document.querySelectorAll('.sh-drawing[data-kind="shape"]').length >= 10`), 'the shapes', 8000).then(() => true, () => false);
    const kinds = await js(`[...document.querySelectorAll('.sh-drawing')].map((d) => d.dataset.kind)`);
    check('odf: and its shapes, picture and text box below them', below && kinds.includes('image'), JSON.stringify(kinds.reduce((m, k) => ({ ...m, [k]: (m[k] || 0) + 1 }), {})));
    await capture(sheets, 'odf-shapes.png');

    const odp = path.join(dir, 'showcase.odp');
    fs.copyFileSync(path.join(FIXTURES, 'showcase.odp'), odp);
    const slides = await open('slides', odp);
    const sj = (code) => slides.webContents.executeJavaScript(code);
    await until(() => sj(`Boolean(document.querySelector('.sl-stage svg'))`), 'the slide', 15000);
    await wait(400);
    const gradients = await sj(`(() => { const svg = document.querySelector('.sl-stage svg'); return { linear: svg.querySelectorAll('linearGradient').length, html: svg.innerHTML.includes('#1f4e79') || svg.innerHTML.includes('#1F4E79') }; })()`);
    check('odf: an .odp from PowerPoint opens with its first slide\'s gradient background drawn', gradients.linear >= 1 && gradients.html, JSON.stringify(gradients));
    await capture(slides, 'odf-gradients.png');
    // The slide of shapes: each drawn as its own outline, not a box.
    await sj(`(() => { document.querySelectorAll('.sl-thumb')[2]?.click(); return 1; })()`);
    await wait(600);
    const outlines = await sj(`(() => { const svg = document.querySelector('.sl-stage svg'); return { paths: svg.querySelectorAll('path').length, curves: [...svg.querySelectorAll('path')].filter((p) => /C/.test(p.getAttribute('d') || '')).length }; })()`);
    check('odf: its shapes are drawn as their own outlines — the oval, the heart, the smile curved', outlines.curves >= 4, JSON.stringify(outlines));
    await capture(slides, 'odf-outlines.png');

    const complaints = [...(await errorsIn(sheets)), ...(await errorsIn(slides))];
    check('odf: the charts, shapes and gradients report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('odf: the chart and gradient checks ran', false, err.message);
  }
}
