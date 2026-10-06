// Presentations: Insert → SmartArt, pressed on the ribbon — Basic Process
// picked in the gallery, three steps typed in its text pane: one group of
// rounded boxes holding the words, an arrow between each, drawn on the
// slide; Home → Convert to SmartArt on a box of three lines: the box gone,
// a Basic Cycle of its lines where it stood; and the saved deck keeps both
// as groups of shapes any PowerPoint opens.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, doc, sessionFor
 */
export async function verifyDeckSmartArt(h, { dir }) {
  const { open, check, until, wait, doc, sessionFor } = h;
  const file = path.join(dir, 'smartart.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'SmartArt', slides: [{ layout: 'title', title: 'Diagrams' }, { layout: 'obj', title: 'Stages', body: ['Plan', 'Build', 'Ship'] }] }));
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    // A picture of the window, when the run asks for them.
    const capture = async (name) => { if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG()); };
    const model = (slide) => doc.model({ id: sessionFor('deck').id, slide });
    const words = (s) => (s.text?.paragraphs || []).map((p) => p.plain).join('');
    const diagram = (slide, re) => {
      const shapes = model(slide).slide?.shapes || [];
      const group = shapes.find((s) => s.kind === 'group' && re.test(s.name || ''));
      return group ? { group, members: shapes.filter((s) => String(s.groupId) === String(group.id)) } : null;
    };
    const press = (tab, title) => js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(tab)})?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'SmartArt' && (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
      if (!b || b.disabled) return 'no button'; b.click(); return 'clicked';
    })()`);
    const pick = async (layout, text) => {
      await until(() => js(`Boolean(document.querySelector('.sa-ok'))`), 'the SmartArt box', 4000);
      await js(`document.querySelector('.sa-item[data-layout="${layout}"]')?.click(), 1`);
      if (text != null) {
        await js(`(() => { const ta = document.querySelector('.sa-text'); const set = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set; set.call(ta, ${JSON.stringify(text)}); ta.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
      }
      await wait(150);
      await capture(`slides-smartart-box-${layout}.png`);
      const pane = await js(`document.querySelector('.sa-text')?.value ?? null`);
      await js(`document.querySelector('.sa-ok')?.click(), 1`);
      return pane;
    };

    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slides', 8000);
    const pressed = await press('Insert', 'SmartArt');
    await pick('process', 'One\nTwo\nThree');
    const made = await until(() => {
      const d = diagram(0, /^Basic Process \d+$/);
      return d && d.members.filter((s) => /^Rectangle: Rounded Corners \d+$/.test(s.name)).map(words).join(',') === 'One,Two,Three' && d.members.filter((s) => /^Arrow: Right \d+$/.test(s.name)).length === 2;
    }, 'the Basic Process group', 8000).then(() => true).catch(() => false);
    const drawn = await until(() => js(`(() => { const t = document.querySelector('.sl-stage .sl-slide')?.textContent || ''; return ['One', 'Two', 'Three'].every((w) => t.includes(w)); })()`), 'the words on the slide', 4000).then(() => true).catch(() => false);
    const picked = await js(`document.querySelectorAll('.sl-hit.selected').length`);
    await capture('slides-smartart-process.png');
    check('presentations: Insert → SmartArt → Basic Process puts the steps typed in as one group — boxes holding the words, an arrow between each — drawn and selected',
      pressed === 'clicked' && made && drawn && picked >= 1, JSON.stringify({ pressed, made, drawn, picked, shapes: (diagram(0, /^Basic Process/)?.members || []).map((s) => [s.name, words(s)]) }));

    // Convert to SmartArt: the box of three lines on slide 2.
    await js(`(() => { document.querySelectorAll('.sl-thumb')[1]?.click(); return 1; })()`);
    await until(() => js(`document.querySelectorAll('.sl-thumb')[1]?.classList.contains('active')`), 'the second slide', 4000).catch(() => {});
    const box = (model(1).slide?.shapes || []).find((s) => s.placeholder?.type === 'body');
    await until(() => js(`Boolean(document.querySelector('.sl-hit[data-shape="${box?.id}"]'))`), 'the box', 4000).catch(() => {});
    await js(`(() => { const el = document.querySelector('.sl-hit[data-shape="${box?.id}"]'); el?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); el?.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true })); return Boolean(el); })()`);
    await wait(200);
    const converting = await press('Home', 'Convert to SmartArt');
    const pane = await pick('cycle', null);
    const converted = await until(() => {
      const d = diagram(1, /^Basic Cycle \d+$/);
      // The box's id may be taken again by a new shape; the box itself is the placeholder.
      const left = (model(1).slide?.shapes || []).some((s) => s.placeholder?.type === 'body');
      return d && !left && d.members.filter((s) => /^Oval \d+$/.test(s.name) && words(s)).map(words).join(',') === 'Plan,Build,Ship';
    }, 'the Basic Cycle in place of the box', 8000).then(() => true).catch(() => false);
    const where = diagram(1, /^Basic Cycle/)?.group.geometry;
    await wait(300);
    await capture('slides-smartart-cycle.png');
    const inBox = Boolean(where && box?.geometry && where.x >= box.geometry.x - 1 && where.y >= box.geometry.y - 1 && where.x + where.w <= box.geometry.x + box.geometry.w + 1 && where.y + where.h <= box.geometry.y + box.geometry.h + 1);
    check('presentations: Home → Convert to SmartArt makes a box\'s lines a diagram where the box stood, and the box goes',
      converting === 'clicked' && pane === 'Plan\nBuild\nShip' && converted && inBox, JSON.stringify({ converting, pane, converted, inBox, where, box: box?.geometry }));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const deck = Deck.open(fs.readFileSync(file));
        const one = deck.pkg.text('ppt/slides/slide1.xml');
        const two = deck.pkg.text('ppt/slides/slide2.xml');
        return /<p:grpSp><p:nvGrpSpPr><p:cNvPr id="\d+" name="Basic Process \d+"\/>/.test(one) && (one.match(/prst="rightArrow"/g) || []).length === 2
          && /name="Basic Cycle \d+"/.test(two) && !/<p:ph type="body"/.test(two);
      } catch { return false; }
    }, 'the saved diagrams', 8000).then(() => true).catch(() => false);
    check('presentations: the saved deck keeps each diagram as a group of shapes', saved, saved ? 'p:grpSp of preset shapes in the file' : 'not in the file');
    await wait(50);
  } catch (err) {
    check('presentations: the SmartArt checks ran', false, err.message);
  }
}
