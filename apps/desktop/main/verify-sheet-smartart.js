// Worksheets: Insert → SmartArt, pressed on the ribbon — Hierarchy picked,
// a lead and two items under it typed in its text pane (a Tab in for each):
// one group on the sheet at the selected cell, the lead over the two with a
// line down, across and down to each, every shape's words drawn; and the
// saved workbook keeps it as a group of shapes Excel opens.

import fs from 'node:fs';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait, doc
 */
export async function verifySheetSmartArt(h, { dir }) {
  const { open, check, until, wait } = h;
  const file = path.join(dir, 'smartart.xlsx');
  try {
    fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['Team', 'Lead'], ['North', 'Ana'], ['South', 'Ben']] }] }));
    const win = await open('sheets', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const capture = async (name) => { if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG()); };
    const drawn = () => js(`(() => {
      const d = [...document.querySelectorAll('.sh-drawing[data-kind="group"]')].find((n) => /^Hierarchy \\d+$/.test(n.dataset.name || ''));
      if (!d) return null;
      const r = d.getBoundingClientRect();
      return { name: d.dataset.name, members: d.querySelectorAll('.sh-member').length, words: [...d.querySelectorAll('text')].map((t) => t.textContent).join(','), lines: d.querySelectorAll('.sh-member line').length, w: Math.round(r.width), h: Math.round(r.height) };
    })()`);

    await until(() => js(`Boolean(document.querySelector('.rw-tab'))`), 'the ribbon', 8000);
    const pressed = await js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'SmartArt');
      if (!b || b.disabled) return 'no button'; b.click(); return 'clicked';
    })()`);
    await until(() => js(`Boolean(document.querySelector('.sa-ok'))`), 'the SmartArt box', 4000).catch(() => {});
    await js(`document.querySelector('.sa-item[data-layout="hierarchy"]')?.click(), 1`);
    // Typed as one types it: a line, then Tab at the start of the next to put it under the first.
    await js(`(() => { const ta = document.querySelector('.sa-text'); const set = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set; set.call(ta, 'Lead\\nNorth\\nSouth'); ta.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    for (const line of [1, 2]) {
      await js(`(() => { const ta = document.querySelector('.sa-text'); const at = ta.value.split('\\n').slice(0, ${line}).join('\\n').length + 1; ta.focus(); ta.selectionStart = ta.selectionEnd = at; ta.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true })); return 1; })()`);
      await wait(120);
    }
    const pane = await js(`document.querySelector('.sa-text')?.value ?? null`);
    await capture('sheet-smartart-box.png');
    await js(`document.querySelector('.sa-ok')?.click(), 1`);
    await until(async () => (await drawn())?.members === 7, 'the diagram on the sheet', 8000).catch(() => {});
    const got = await drawn();
    await capture('sheet-smartart.png');
    check('worksheets: Insert → SmartArt → Hierarchy puts the lead over the items tabbed in under it as one group at the cell, a line to each, every word drawn',
      pressed === 'clicked' && pane === 'Lead\n\tNorth\n\tSouth' && got?.members === 7 && got.words === 'Lead,North,South' && got.lines === 4, JSON.stringify({ pressed, pane, got }));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const pkg = OoxmlPackage.read(fs.readFileSync(file));
        const part = pkg.partNames().find((p) => /^xl\/drawings\/drawing\d+\.xml$/.test(p));
        const xml = part ? pkg.text(part) : '';
        return /<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="\d+" name="Hierarchy \d+"\/>/.test(xml) && (xml.match(/<a:prstGeom prst="roundRect">/g) || []).length === 3
          && (xml.match(/<a:prstGeom prst="line">/g) || []).length === 4 && /<a:t>North<\/a:t>/.test(xml);
      } catch { return false; }
    }, 'the saved diagram', 8000).then(() => true).catch(() => false);
    check('worksheets: the saved workbook keeps the diagram as a group of shapes Excel opens', saved, saved ? 'xdr:grpSp of shapes in the drawing' : 'not in the file');
    await wait(50);
  } catch (err) {
    check('worksheets: the SmartArt checks ran', false, err.message);
  }
}
