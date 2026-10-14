// Documents: ink as Word keeps it. A document with ink Word drew (an
// InkML part in a w14 content part, its traces packed as Word packs them,
// a picture its fallback) shows the stroke drawn from its InkML; a pen and
// a highlighter stroke drawn here are saved the same way, a content part
// that requires wpi, the InkML with the pen's colour, width and a
// highlighter's tip, the freeform its fallback, and drawn again when the
// file is opened again.
// Run alone with RUTBA_VERIFY_ONLY=word-ink.

import fs from 'node:fs';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A document holding ink as Word 365 writes it: the content part, its InkML with packed traces, a picture its fallback. */
function wordInkDocx() {
  const pkg = OoxmlPackage.read(buildDocx({ styles: true, paragraphs: [{ text: 'Signed below' }, { text: '' }] }));
  const inkml = '<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"><inkml:definitions><inkml:context xml:id="ctx0"><inkml:inkSource xml:id="inkSrc0"><inkml:traceFormat>'
    + '<inkml:channel name="X" type="integer" max="32767" units="cm"/><inkml:channel name="Y" type="integer" max="32767" units="cm"/><inkml:channel name="T" type="integer" max="2.14748E9" units="dev"/></inkml:traceFormat>'
    + '<inkml:channelProperties><inkml:channelProperty channel="X" name="resolution" value="1000" units="1/cm"/><inkml:channelProperty channel="Y" name="resolution" value="1000" units="1/cm"/><inkml:channelProperty channel="T" name="resolution" value="1" units="1/dev"/></inkml:channelProperties>'
    + '</inkml:inkSource><inkml:timestamp xml:id="ts0" timeString="2026-10-14T09:00:00.000"/></inkml:context><inkml:brush xml:id="br0"><inkml:brushProperty name="width" value="0.05" units="cm"/><inkml:brushProperty name="height" value="0.05" units="cm"/><inkml:brushProperty name="color" value="#0000FF"/></inkml:brush></inkml:definitions>'
    + '<inkml:trace contextRef="#ctx0" brushRef="#br0">5000 9000 0,\'400\'50 16,"0\'-10 15,0 0 16,0\'40\'0 17,\'-300\'200 16,0 0 15,0 0 16</inkml:trace></inkml:ink>';
  pkg.addPart('word/ink/ink1.xml', inkml, 'application/inkml+xml');
  pkg.write_('word/_rels/document.xml.rels', pkg.text('word/_rels/document.xml.rels').replace('</Relationships>', `<Relationship Id="rIdInk" Type="http://schemas.microsoft.com/office/2011/relationships/ink" Target="ink/ink1.xml"/></Relationships>`));
  const anchor = (graphic) => '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
    + '<wp:positionH relativeFrom="page"><wp:posOffset>1800000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>1700000</wp:posOffset></wp:positionV>'
    + '<wp:extent cx="1440000" cy="540000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="11" name="Ink 1"/><wp:cNvGraphicFramePr/>' + graphic + '</wp:anchor>';
  const run = '<w:r><mc:AlternateContent><mc:Choice Requires="wpi"><w:drawing>' + anchor('<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"><w14:contentPart bwMode="auto" r:id="rIdInk"><w14:nvContentPartPr><w14:cNvContentPartPr/></w14:nvContentPartPr><w14:xfrm><a:off x="0" y="0"/><a:ext cx="1440000" cy="540000"/></w14:xfrm></w14:contentPart></a:graphicData></a:graphic>')
    + '</w:drawing></mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r>';
  let xml = pkg.text('word/document.xml');
  xml = xml.replace('<w:document ', `<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:r="${REL}" mc:Ignorable="w14" `);
  xml = xml.replace(/<w:p><w:r><w:t xml:space="preserve">Signed below<\/w:t><\/w:r><\/w:p>/, (p) => p.replace('</w:p>', run + '</w:p>'));
  pkg.write_('word/document.xml', xml);
  return pkg.write();
}

/**
 * @param {object} h the harness: open, check, until, wait, capture, doc
 */
export async function verifyWordInkml(h, { dir }) {
  const { open, check, wait, capture, doc } = h;
  // Waits as the harness's until does, and answers what the condition found once it held.
  const until = async (fn, what, ms) => { let found; await h.until(async () => { found = await fn(); return found; }, what, ms); return found; };
  const own = path.join(dir, 'word-ink');
  const file = path.join(own, 'signed.docx');
  try {
    fs.mkdirSync(own, { recursive: true });
    fs.writeFileSync(file, wordInkDocx());
  } catch (err) {
    return check('word-ink: the fixture was made', false, err.message);
  }
  try {
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'doc' && s.path === file).pop();
    const inks = () => js(`[...document.querySelectorAll('.wd-page .wd-drawing')].filter((n) => /^Ink \\d+$/.test(n.getAttribute('alt') || '')).map((n) => { const r = n.getBoundingClientRect(); return { name: n.getAttribute('alt'), w: Math.round(r.width), h: Math.round(r.height) }; })`);
    const press = (label) => js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!b || b.disabled) return 'no button'; b.click(); return 'clicked'; })()`);
    const frame = () => js(`(() => { const s = document.querySelector('.sl-ink-surface'); if (!s) return null; const r = s.getBoundingClientRect(); return { left: r.left, top: r.top, k: r.width / parseFloat(s.style.width) }; })()`);
    const stroke = async (points) => {
      const f = await frame();
      if (!f) return false;
      const toWin = ([x, y]) => ({ x: Math.round(f.left + x * f.k), y: Math.round(f.top + y * f.k) });
      const first = toWin(points[0]);
      wc.sendInputEvent({ type: 'mouseDown', x: first.x, y: first.y, button: 'left', clickCount: 1 });
      for (const p of points.slice(1)) { const q = toWin(p); wc.sendInputEvent({ type: 'mouseMove', x: q.x, y: q.y, button: 'left', modifiers: ['leftButtonDown'] }); await wait(8); }
      const last = toWin(points[points.length - 1]);
      wc.sendInputEvent({ type: 'mouseUp', x: last.x, y: last.y, button: 'left', clickCount: 1 });
      return true;
    };
    const line = (a, b, n = 20) => Array.from({ length: n + 1 }, (_, i) => [a[0] + ((b[0] - a[0]) * i) / n, a[1] + ((b[1] - a[1]) * i) / n]);

    // 1. Word's own ink, drawn from its InkML.
    const word = await until(async () => { const l = await inks(); return l.length === 1 && l[0].w > 20 ? l : null; }, 'Word\'s ink on the page', 8000).catch(() => null);
    check('word-ink: ink Word drew is drawn from its InkML, the content part read, its packed traces undone, where the fallback is only a placeholder', Array.isArray(word), JSON.stringify(word));

    // 2. A pen stroke and a highlighter stroke, saved as Word saves ink.
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Draw')?.click(), 1`);
    await wait(250);
    await press('Pen');
    await until(() => js(`Boolean(document.querySelector('.sl-ink-surface'))`), 'the pen surface', 3000).catch(() => {});
    await stroke(line([150, 420], [450, 460]));
    await until(async () => (await inks()).length === 2, 'the pen stroke', 6000).catch(() => {});
    await press('Highlighter');
    await stroke(line([120, 520], [520, 520]));
    const three = await until(async () => (await inks()).length === 3, 'the highlighter stroke', 6000).then(() => true).catch(() => false);
    await press('Select');
    wc.invalidate();
    await wait(700);
    await capture(win, 'word-ink-inkml.png');
    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    await until(() => doc.meta({ id: session.id }).dirty === false, 'the save', 15000).catch(() => {});
    const pkg = OoxmlPackage.read(fs.readFileSync(file));
    const body = pkg.text('word/document.xml');
    const parts = pkg.partNames().filter((n) => /^word\/ink\/ink\d+\.xml$/.test(n));
    const inkml = parts.map((n) => pkg.text(n));
    const saved = {
      choices: (body.match(/<mc:Choice Requires="wpi">/g) || []).length === 3,
      contentParts: (body.match(/<w14:contentPart bwMode="auto" r:id="[^"]+">/g) || []).length === 3,
      fallbacks: (body.match(/<mc:Fallback><w:drawing>[\s\S]*?<wps:wsp>/g) || []).length === 2,
      parts: parts.length === 3 && /application\/inkml\+xml/.test(pkg.text('[Content_Types].xml')),
      highlighter: inkml.some((x) => /name="tip" value="rectangle"/.test(x) && /name="rasterOp" value="maskPen"/.test(x)),
      traces: inkml.every((x) => /<inkml:trace\b[^>]*>\s*-?\d+ -?\d+/.test(x)),
    };
    check('word-ink: pen and highlighter strokes are saved as Word saves ink, a content part requiring wpi, its InkML part (the highlighter\'s rectangular tip and mask pen), the freeform their fallback', three && Object.values(saved).every(Boolean), JSON.stringify(saved));

    // 3. Opened again: drawn from the InkML.
    const win2 = await open('word', file);
    const js2 = (code) => win2.webContents.executeJavaScript(code);
    const again = await until(() => js2(`[...document.querySelectorAll('.wd-page .wd-drawing')].filter((n) => /^Ink \\d+$/.test(n.getAttribute('alt') || '')).length`).then((n) => (n === 3 ? n : null)), 'the ink opened again', 8000).catch(() => null);
    check('word-ink: opened again, all three strokes are drawn from their InkML', again === 3, String(again));
  } catch (err) {
    check('word-ink: the ink checks ran', false, err.stack || err.message);
  }
}
