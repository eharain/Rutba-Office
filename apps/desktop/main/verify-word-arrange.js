// Word: Layout → Arrange on a document's drawings.
//
// A page with a floating picture and two overlapping text boxes. The
// picture goes Behind Text — the words are drawn over it — then In Front of
// Text, over them. The Selection Pane lists the three top-most first; a row
// selects, Bring to Front puts the lower box over the other where they
// overlap, Align Right sets it against the margin, Shift adds the second
// box and Align Top lines the two up, Rotate turns one, Group makes the two
// one wpg group and Ungroup parts them, and the pane's eye hides and shows
// the picture. The saved file carries the order and the wraps.
// Run alone with RUTBA_VERIFY_ONLY=wordarrange.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { gradientPng } from './sample-picture.js';

const lorem = 'The northern region grew fastest in absolute terms, and the margin it opened in the spring held through the autumn. ';

/** The fixture: words, a floating picture, and two boxes that overlap. */
export function makeArrangeFixture(file) {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Arrange', style: 'Heading1' }, { text: lorem.repeat(7) }, { text: lorem.repeat(4) }] }));
  view.setSelection({ block: 0, offset: 0 });
  view.insertImage({ name: 'Logo', contentType: 'image/png', data: gradientPng(160, 120, [40, 120, 200], [230, 200, 60]), widthPx: 160, heightPx: 120 });
  const pic = view.drawings().find((d) => d.kind === 'picture').id;
  // The picture floats at the left, the words beside it.
  view.updateDrawings({ id: pic, wrap: 'square', h: { rel: 'column', offsetPx: 0 }, v: { rel: 'paragraph', offsetPx: 0 } });
  view.setSelection({ block: 3, offset: 0 });
  view.insertTextBox({ name: 'Box A', widthPx: 200, heightPx: 90, fill: 'DEEBF7', h: { rel: 'column', offsetPx: 330 }, v: { rel: 'paragraph', offsetPx: 0 }, paragraphs: [{ text: 'Box A' }] });
  view.setSelection({ block: 3, offset: 0 });
  view.insertTextBox({ name: 'Box B', widthPx: 200, heightPx: 90, fill: 'FFF2CC', h: { rel: 'column', offsetPx: 390 }, v: { rel: 'paragraph', offsetPx: 40 }, paragraphs: [{ text: 'Box B' }] });
  fs.writeFileSync(file, view.save());
}

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 * @param {{ dir: string }} args where the fixture is written
 */
export async function verifyWordArrange(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, capture, doc, sessionFor } = h;
  const file = path.join(dir, 'arrange.docx');
  try {
    makeArrangeFixture(file);
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = sessionFor('doc');
    const model = () => doc.model({ id: session.id });
    const drawing = (name) => (model().drawings || []).find((d) => d.name === name);
    const clickTab = (name) => js(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)}); if (!t) return 'no tab'; t.click(); return 'tab'; })()`);
    const clickRibbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no button ' + ${JSON.stringify(title)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const menuPick = async (button, label) => {
      const pressed = await clickRibbon(button);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)} && !b.disabled))`), `"${label}" in ${button}`, 3000).catch(() => false);
      const picked = await js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); if (!b || b.disabled) return 'missing'; b.click(); return 'clicked'; })()`);
      return pressed === 'clicked' && picked === 'clicked' ? 'clicked' : `${pressed}/${picked}`;
    };
    const row = (name, shift = false) => js(`(() => {
      const r = [...document.querySelectorAll('.wd-selpane .wd-layer')].find((x) => x.querySelector('.wd-layer-title')?.textContent === ${JSON.stringify(name)});
      if (!r) return 'no row';
      r.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, shiftKey: ${shift} }));
      return 'row';
    })()`);
    const rectOf = (name) => js(`(() => {
      const id = ${JSON.stringify(String(drawing(name)?.id))};
      const el = document.querySelector('.wd-drawing[data-drawing="' + id + '"]');
      if (!el) return null;
      const r = el.getBoundingClientRect();
      return { left: Math.round(r.left), top: Math.round(r.top), right: Math.round(r.right), bottom: Math.round(r.bottom), layer: el.closest('.wd-drawlayer')?.classList.contains('behind') ? 'behind' : el.closest('.wd-drawlayer') ? 'front' : 'flow', transform: el.style.transform || '' };
    })()`);
    const shot = async (name) => {
      wc.invalidate();
      await wait(700);
      await capture(win, name);
    };
    const columnNow = () => js(`(() => { const p = document.querySelector('.wd-page'); const r = p.getBoundingClientRect(); const cs = getComputedStyle(p); return { left: Math.round(r.left + parseFloat(cs.paddingLeft)), right: Math.round(r.right - parseFloat(cs.paddingRight)) }; })()`);

    await until(() => js(`document.querySelectorAll('.wd-drawing[data-drawing]').length === 3`), 'the three drawings', 8000);
    await wait(500);

    // The picture Behind Text: the words run across it and are drawn over it.
    await js(`(() => { document.querySelector('.wd-image[data-drawing]').click(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.wd-frame-sel'))`), 'the picture to be selected', 4000).catch(() => false);
    await clickTab('Layout');
    await wait(200);
    const behindPress = await menuPick('Wrap Text', 'Behind text');
    const behind = await until(async () => (await rectOf('Logo'))?.layer === 'behind', 'the picture behind the words', 5000).catch(() => false);
    const under = await js(`(() => {
      const img = document.querySelector('.wd-drawlayer.behind .wd-image');
      const r = img.getBoundingClientRect();
      const hit = document.elementFromPoint(r.left + r.width / 2, r.top + r.height * 0.6);
      return { onTop: hit ? (hit.closest('.wd-block') ? 'words' : hit.className) : null };
    })()`);
    check('word: Wrap Text → Behind Text lays the picture under the words — they run across it and are drawn over it',
      behindPress === 'clicked' && behind === true && under.onTop === 'words' && drawing('Logo')?.behind === true && drawing('Logo')?.wrap === 'none',
      `${behindPress}; ${JSON.stringify(under)}; ${JSON.stringify({ wrap: drawing('Logo')?.wrap, behind: drawing('Logo')?.behind })}`);
    await shot('word-behind-text.png');
    const frontPress = await menuPick('Wrap Text', 'In front of text');
    const front = await until(async () => (await rectOf('Logo'))?.layer === 'front', 'the picture in front of the words', 5000).catch(() => false);
    const over = await js(`(() => {
      const img = document.querySelector('.wd-drawlayer.front .wd-image');
      const r = img.getBoundingClientRect();
      const hit = document.elementFromPoint(r.left + r.width / 2, r.top + r.height * 0.6);
      return { onTop: hit === img ? 'picture' : hit?.closest('.wd-block') ? 'words' : hit?.className };
    })()`);
    check('word: Wrap Text → In Front of Text lays it over the words',
      frontPress === 'clicked' && front === true && over.onTop === 'picture' && drawing('Logo')?.behind === false, `${frontPress}; ${JSON.stringify(over)}`);

    // The Selection Pane: every drawing, the top-most first.
    const pane = await clickRibbon('Selection Pane');
    await until(() => js(`document.querySelectorAll('.wd-selpane .wd-layer').length === 3`), 'the three rows', 4000).catch(() => false);
    const rows = await js(`[...document.querySelectorAll('.wd-selpane .wd-layer-title')].map((t) => t.textContent)`);
    check('word: Layout → Selection Pane lists every drawing, the one on top first',
      pane === 'clicked' && JSON.stringify(rows) === JSON.stringify(['Box B', 'Box A', 'Logo']), JSON.stringify(rows));

    // Box A is under B where they overlap; Bring to Front puts it over.
    await row('Box A');
    await until(() => js(`Boolean(document.querySelector('.wd-selpane .wd-layer.active'))`), 'the row to be selected', 3000).catch(() => false);
    const overlap = async () => {
      const a = await rectOf('Box A');
      const b = await rectOf('Box B');
      if (!a || !b) return null;
      const x = Math.round((Math.max(a.left, b.left) + Math.min(a.right, b.right)) / 2);
      const y = Math.round((Math.max(a.top, b.top) + Math.min(a.bottom, b.bottom)) / 2);
      return js(`(() => { const hit = document.elementFromPoint(${x}, ${y}); const box = hit?.closest('.wd-textbox'); return box ? box.dataset.drawing : 'at ${x},${y}: ' + (hit ? hit.tagName + '.' + hit.className : 'nothing'); })()`);
    };
    const onTopBefore = await overlap();
    const toFront = await menuPick('Bring Forward', 'Bring to Front');
    await until(() => (drawing('Box A')?.relativeHeight ?? 0) > (drawing('Box B')?.relativeHeight ?? 0), 'A above B in the file', 5000).catch(() => false);
    const onTopAfter = await until(async () => (await overlap()) === String(drawing('Box A').id), 'A drawn over B', 5000).catch(() => false);
    const rowsAfter = await js(`[...document.querySelectorAll('.wd-selpane .wd-layer-title')].map((t) => t.textContent)`);
    check('word: Bring to Front puts the lower box over the other where they overlap, and the pane follows',
      toFront === 'clicked' && onTopBefore === String(drawing('Box B').id) && onTopAfter === true && rowsAfter[0] === 'Box A',
      `${toFront}; before ${onTopBefore}, after ${await overlap()}; rows ${JSON.stringify(rowsAfter)}`);
    await shot('word-arrange-front.png');

    // Align Right, to the margin.
    const alignRight = await menuPick('Align', 'Align Right');
    const column = await columnNow();
    const righted = await until(async () => { const a = await rectOf('Box A'); const c = await columnNow(); return a && Math.abs(a.right - c.right) <= 2; }, 'A against the right margin', 5000).catch(() => false);
    check('word: Align → Align Right sets a drawing against the right margin, as Word writes it',
      alignRight === 'clicked' && righted === true && drawing('Box A')?.hAlign === 'right' && drawing('Box A')?.hRel === 'margin',
      `${alignRight}; ${JSON.stringify(await rectOf('Box A'))}; column ${JSON.stringify(await columnNow())}`);

    // Shift adds B; Align Top lines the two up (Align Selected Objects).
    await row('Box B', true);
    await until(() => js(`document.querySelectorAll('.wd-selpane .wd-layer.active').length === 2`), 'two rows selected', 3000).catch(() => false);
    const alignTop = await menuPick('Align', 'Align Top');
    const topped = await until(async () => { const a = await rectOf('Box A'); const b = await rectOf('Box B'); return a && b && Math.abs(a.top - b.top) <= 2; }, 'the two tops level', 5000).catch(() => false);
    check('word: Shift adds a drawing to the selection, and Align Top lines the selected ones up',
      alignTop === 'clicked' && topped === true, `${alignTop}; A ${JSON.stringify(await rectOf('Box A'))} B ${JSON.stringify(await rectOf('Box B'))}`);

    // Group, then Ungroup.
    const grouped = await menuPick('Group', 'Group');
    const oneGroup = await until(() => (model().drawings || []).some((d) => d.kind === 'group' && d.members?.length === 2), 'one group', 5000).catch(() => false);
    const groupOnPage = await until(() => js(`document.querySelectorAll('.wd-group .wd-textbox').length === 2`), 'the group on the page', 5000).catch(() => false);
    await shot('word-arrange-group.png');
    const ungrouped = await menuPick('Group', 'Ungroup');
    const two = await until(() => (model().drawings || []).filter((d) => d.kind === 'textbox').length === 2 && !(model().drawings || []).some((d) => d.kind === 'group'), 'two boxes again', 5000).catch(() => false);
    check('word: Group makes the two boxes one group, drawn as one, and Ungroup parts them again',
      grouped === 'clicked' && oneGroup === true && groupOnPage === true && ungrouped === 'clicked' && two === true,
      `${grouped} ${oneGroup} ${groupOnPage} ${ungrouped} ${two}`);

    // Rotate Right 90° on the picture.
    await row('Logo');
    await until(() => js(`document.querySelectorAll('.wd-selpane .wd-layer.active').length === 1`), 'the picture selected', 3000).catch(() => false);
    const turned = await menuPick('Rotate', 'Rotate Right 90°');
    const rot = await until(async () => drawing('Logo')?.rot === 90 && /rotate\(90deg\)/.test((await rectOf('Logo'))?.transform || ''), 'the picture turned', 5000).catch(() => false);
    check('word: Rotate → Rotate Right 90° turns the picture on the page and in the file', turned === 'clicked' && rot === true, `${turned}; rot ${drawing('Logo')?.rot}; ${JSON.stringify(await rectOf('Logo'))}`);
    await shot('word-arrange-rotate.png');

    // The pane's eye hides the picture and shows it again.
    const eye = (name) => js(`(() => { const r = [...document.querySelectorAll('.wd-selpane .wd-layer')].find((x) => x.querySelector('.wd-layer-title')?.textContent === ${JSON.stringify(name)}); const b = r?.querySelector('.wd-eye'); if (!b) return 'no eye'; b.click(); return 'eye'; })()`);
    await eye('Logo');
    const hidden = await until(async () => drawing('Logo')?.hidden === true && !(await rectOf('Logo')), 'the picture hidden', 5000).catch(() => false);
    await eye('Logo');
    const back = await until(async () => drawing('Logo')?.hidden === false && Boolean(await rectOf('Logo')), 'the picture shown', 5000).catch(() => false);
    check('word: the Selection Pane\'s eye hides a drawing and shows it again', hidden === true && back === true, `${hidden} ${back}`);

    // Saved as Word writes it: the wraps and the order.
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(wc, 's', { modifiers: ['control'] });
    await wait(1400);
    const saved = openDocx(fs.readFileSync(file)).drawings();
    const byName = Object.fromEntries(saved.map((d) => [d.name, d]));
    check('word: the saved file keeps the order, the wraps and the turn',
      byName['Box A']?.relativeHeight > byName['Box B']?.relativeHeight && byName.Logo?.wrap === 'none' && byName.Logo?.behind === false && byName.Logo?.rot === 90,
      JSON.stringify(saved.map((d) => ({ name: d.name, z: d.relativeHeight, wrap: d.wrap, behind: d.behind, rot: d.rot, hAlign: d.hAlign }))));

    const complaints = await errorsIn(win);
    check('word: arranging drawings reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the arrange checks ran', false, err.message);
  }
}
