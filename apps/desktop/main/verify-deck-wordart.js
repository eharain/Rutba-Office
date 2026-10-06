// Presentations: Insert → WordArt, pressed on the ribbon — the gallery shows
// each style on an "A"; the white-with-blue-outline-and-glow style puts
// PowerPoint's "Your text here" in the middle of the slide, selected, its
// letters drawn with a stroke and the line with a glow; new words typed in
// keep the style; and the saved deck keeps the outline and the glow in the
// run, as PowerPoint writes WordArt.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, doc, capture
 */
export async function verifyDeckWordArt(h, { dir }) {
  const { open, check, until, wait, doc, capture } = h;
  const file = path.join(dir, 'wordart.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Art', slides: [{ layout: 'blank' }] }));
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'deck' && s.path === file).pop();
    const art = () => (doc.model({ id: session.id }).slide?.shapes || []).find((s) => (s.name || '').startsWith('WordArt'));
    const runOf = () => art()?.text?.paragraphs?.[0]?.runs?.[0] || null;

    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 1`), 'the slide', 8000);
    const opened = await js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'WordArt');
      if (!b || b.disabled) return 'no button';
      b.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 500, clientY: 110 }));
      await new Promise((r) => setTimeout(r, 250));
      return [...document.querySelectorAll('.rw-menu button')].map((n) => ({ label: n.textContent.trim().replace(/^A/, ''), sample: Boolean(n.querySelector('.rw-menu-preview span')) }));
    })()`);
    if (capture) await capture(win, 'deck-wordart-menu.png');
    const gallery = Array.isArray(opened) && opened.length >= 6 && opened.every((i) => i.sample);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.querySelector('span:not(.rw-menu-preview):not(.rw-menu-preview *)')?.textContent.startsWith('Fill: White; Outline: Blue')); if (!b) return 'no item'; b.click(); return 'picked'; })()`);
    const placed = await until(() => runOf()?.text === 'Your text here', 'the WordArt on the slide', 5000).then(() => true).catch(() => false);
    const run = runOf();
    const shape = art()?.geometry;
    const size = doc.model({ id: session.id }).size || { width: 1280, height: 720 };
    const centred = Boolean(shape) && Math.abs(shape.x + shape.w / 2 - size.width / 2) < 4 && Math.abs(shape.y + shape.h / 2 - size.height / 2) < 4;
    check('presentations: Insert → WordArt shows the gallery on an "A" each, and a style puts "Your text here" in the middle of the slide in it',
      gallery && placed && centred && run?.outline?.color === '#4472c4' && run?.textEffects?.glow?.radiusPt === 5 && run?.size === 54,
      JSON.stringify({ gallery, opened: Array.isArray(opened) ? opened.length : opened, placed, centred, run, box: shape && [shape.x, shape.y, shape.w, shape.h], size }));

    // The stage draws it: a stroke round the letters, a filter round the line.
    const drawn = await until(() => js(`(() => {
      const t = [...document.querySelectorAll('.sl-slide svg tspan')].find((n) => n.textContent === 'Your text here');
      if (!t) return false;
      const text = t.closest('text');
      const id = /url\\(#([^)]+)\\)/.exec(text.getAttribute('filter') || '')?.[1];
      return t.getAttribute('stroke') === '#4472c4' && Boolean(id && document.getElementById(id)?.querySelector('feFlood'));
    })()`), 'the WordArt drawn', 5000).then(() => true).catch(() => false);
    if (capture) await capture(win, 'deck-wordart.png');
    check('presentations: the slide draws WordArt — the outline as a stroke round each letter, the glow round the words', drawn, drawn ? 'stroke and glow filter on the stage' : 'not drawn');

    // New words, written the way the window writes a shape's words: the style stays.
    const paragraphs = art().text.paragraphs.map((p) => ({ ...p, runs: p.runs.map((r) => ({ ...r, text: 'Quarterly results' })) }));
    doc.apply({ id: session.id, ops: [{ op: 'setText', slide: 0, shape: art().id, paragraphs }] });
    const kept = runOf()?.text === 'Quarterly results' && runOf()?.outline?.color === '#4472c4' && Boolean(runOf()?.textEffects?.glow);
    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const xml = Deck.open(fs.readFileSync(file)).pkg.text('ppt/slides/slide1.xml');
        return /<a:rPr\b[^>]*><a:ln w="12700"><a:solidFill><a:srgbClr val="4472C4"\/><\/a:solidFill><\/a:ln><a:solidFill><a:srgbClr val="FFFFFF"\/><\/a:solidFill><a:effectLst><a:glow /i.test(xml) && xml.includes('Quarterly results');
      } catch { return false; }
    }, 'the saved WordArt', 8000).then(() => true).catch(() => false);
    check('presentations: new words keep the WordArt style, and the saved deck keeps the outline and glow in the run as PowerPoint writes them', kept && saved, JSON.stringify({ kept, saved, run: runOf() }));
  } catch (err) {
    check('presentations: the WordArt checks ran', false, err.message);
  }
}
