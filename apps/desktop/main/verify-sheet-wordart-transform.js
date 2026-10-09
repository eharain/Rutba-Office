// Worksheets: Shape Format → Text Effects → Transform — a WordArt shape
// picked on the sheet, Arch Up pressed: its words are drawn along an arc
// (a <textPath>), and the saved workbook carries a:prstTxWarp as Excel
// writes it. Run alone with RUTBA_VERIFY_ONLY=sheetwordart.

import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { SheetView } from '@rutba/sheet-view';

export async function verifySheetWordArtTransform({ open, check, until, wait, errorsIn }, { dir }) {
  const file = path.join(dir, 'sheet-wordart.xlsx');
  try {
    const made = SheetView.open(buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['Sales', 10]] }] }), { viewportWidth: 900, viewportHeight: 500 });
    made.select(3, 1);
    made.insertWordArt({ text: 'Rutba Office', size: 36, style: { color: '#2F5597' } });
    fs.writeFileSync(file, Buffer.from(made.serialize()));
    const win = await open('sheets', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.sh-drawing').length === 1`), 'the WordArt drawn', 8000);
    const at = await js(`(() => { const el = document.querySelector('.sh-drawing'); el.scrollIntoView({ block: 'center', inline: 'center' }); const r = el.getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; })()`);
    wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    const tabShown = await until(() => js(`[...document.querySelectorAll('.rw-tab')].some((t) => t.textContent.trim() === 'Shape Format')`), 'the Shape Format tab', 4000).then(() => true).catch(() => false);
    await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Shape Format')?.click(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sh-warp[data-preset="textArchUp"]'))`), 'the Transform buttons', 3000).catch(() => {});
    await js(`document.querySelector('.sh-warp[data-preset="textArchUp"]')?.click(), 1`);
    const drawn = await until(() => js(`(() => { const t = document.querySelector('.sh-drawing textPath'); return Boolean(t) && t.textContent.includes('Rutba Office'); })()`), 'the words along the arc', 5000).then(() => true).catch(() => false);
    if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'sheet-wordart-arch.png'), (await wc.capturePage()).toPNG());
    check('worksheets: Shape Format → Transform → Arch Up lays a WordArt\'s words along an arc on the sheet', tabShown && drawn, JSON.stringify({ tabShown, drawn }));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const pkg = OoxmlPackage.read(fs.readFileSync(file));
        const drawing = pkg.partNames().find((n) => /xl\/drawings\/drawing\d+\.xml$/.test(n));
        return /<a:prstTxWarp prst="textArchUp">/.test(pkg.text(drawing));
      } catch { return false; }
    }, 'the save', 6000).then(() => true).catch(() => false);
    check('worksheets: the saved workbook carries a:prstTxWarp as Excel writes it', saved, saved ? 'textArchUp' : 'not in the file');
    const complaints = await errorsIn(win);
    check('worksheets: WordArt\'s Transform reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    void wait;
  } catch (err) {
    check('worksheets: the WordArt Transform checks ran', false, err?.message || String(err));
  }
}
