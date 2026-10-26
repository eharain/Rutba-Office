// Presentations: New Slide, Shapes, Shape Fill and Outline, Design Ideas and
// Presenter View, as a person meets them. New Slide lists the deck's own
// layouts and makes the slide from the one chosen, its text where that
// layout has it; a shape chosen from Shapes is drawn on the slide by a drag,
// and Esc puts the choice away; Shape Fill and Shape Outline offer colours
// at once; Design Ideas tried one after another do not pile up, and one
// Ctrl+Z gives the slide back; Esc in a show with Presenter View ends it
// for both windows. Run alone with RUTBA_VERIFY_ONLY=decktools.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { BrowserWindow, screen } from 'electron';
import { buildPptx } from '@rutba/presentation';
import { putShapeDown } from './verify-shape-place.js';

const PITCH = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..', '..', 'packages', 'templates', 'files', 'deck-pitch.pptx');

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 */
export async function verifyDeckTools(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, doc } = h;
  const opened = [];
  try {
    /* -- New Slide on the deck's own layouts -------------------------------- */
    const pitchFile = path.join(dir, 'tools-pitch.pptx');
    fs.copyFileSync(PITCH, pitchFile);
    const pitch = await open('slides', pitchFile);
    const pjs = (code) => pitch.webContents.executeJavaScript(code);
    const sessionOf = (file) => doc.sessions().find((s) => s.path && path.resolve(s.path) === path.resolve(file));
    const pModel = (slide) => doc.model({ id: sessionOf(pitchFile).id, slide });
    const tab = (js, label) => js(`(async () => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(label)})?.click(); await new Promise((r) => setTimeout(r, 200)); return 1; })()`);
    const ribbon = (js, label) => js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => !n.disabled && n.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'no button ' + ${JSON.stringify(label)}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
    const items = (js) => js(`[...document.querySelectorAll('.rw-menu button')].map((b) => b.textContent.trim())`);
    const pick = (js, label) => js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'no item ' + ${JSON.stringify(label)}; b.click(); return 'picked'; })()`);
    await until(() => pjs(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the pitch deck', 10000);
    const count = pModel(0).count;
    await tab(pjs, 'Home');
    await ribbon(pjs, 'New Slide');
    await until(async () => (await items(pjs)).includes('Section Header'), 'the deck\'s layouts', 3000).catch(() => {});
    const listed = await items(pjs);
    await pick(pjs, 'Section Header');
    await until(() => pModel(0).count === count + 1, 'the new slide', 5000).catch(() => {});
    // After the slide that was on screen (the first), and chosen.
    await until(() => pjs(`document.querySelectorAll('.sl-thumb')[1]?.classList.contains('active') || document.querySelectorAll('.sl-thumb')[1]?.getAttribute('aria-current') === 'true'`), 'the new slide chosen', 3000).catch(() => {});
    const active = await pjs(`[...document.querySelectorAll('.sl-thumb')].findIndex((t) => t.classList.contains('active') || t.getAttribute('aria-current') === 'true')`);
    const made = (pModel(1).slide?.shapes || []).filter((s) => s.placeholder).map((s) => `${s.placeholder.type || 'obj'}:${s.placeholder.idx ?? ''}`);
    await tab(pjs, 'Design');
    const designLabels = await pjs(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].map((n) => n.textContent.trim())`);
    check('presentations: New Slide lists the deck\'s own layouts, and Section Header gives a slide with that layout\'s title and text placeholder',
      ['Title Slide', 'Title and Content', 'Section Header', 'Title Only', 'Picture with Caption'].every((n) => listed.includes(n))
        && pModel(0).count === count + 1 && made.join(',') === 'title:,body:13' && active === 1,
      JSON.stringify({ listed, made, active }));
    check('presentations: the Design tab says New Slide for a new slide, and Layout is Home\'s alone',
      designLabels.includes('New Slide') && !designLabels.includes('Layout'), JSON.stringify(designLabels.filter((l) => /Slide|Layout/.test(l))));
    opened.push(pitch);

    /* -- Shapes drawn by a drag; Shape Fill and Outline colours ------------- */
    const file = path.join(dir, 'tools.pptx');
    fs.writeFileSync(file, buildPptx({ title: 'Tools', slides: [{ title: 'Quarterly results', body: ['Sales rose', 'Costs fell'] }, { layout: 'blank' }] }));
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const model = (slide) => doc.model({ id: sessionOf(file).id, slide });
    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the deck', 8000);
    await js(`document.querySelectorAll('.sl-thumb')[1].click(), 1`);
    await until(() => js(`[...document.querySelectorAll('.sl-thumb')].findIndex((t) => t.classList.contains('active')) === 1`), 'the blank slide', 4000).catch(() => {});
    const size = model(1).size;
    await tab(js, 'Insert');
    await ribbon(js, 'Shapes');
    await until(async () => (await items(js)).includes('Oval'), 'the Shapes menu', 3000).catch(() => {});
    await pick(js, 'Oval');
    const waiting = await until(() => js(`Boolean(document.querySelector('[data-role="shape-draw"]'))`), 'the drawing surface', 3000).then(() => true).catch(() => false);
    win.focus();
    await press(wc, 'Escape');
    await until(() => js(`!document.querySelector('[data-role="shape-draw"]')`), 'the choice put away', 3000).catch(() => {});
    const putAway = await js(`!document.querySelector('[data-role="shape-draw"]')`) && (model(1).slide?.shapes || []).length === 0;
    await ribbon(js, 'Shapes');
    await until(async () => (await items(js)).includes('Rectangle'), 'the Shapes menu', 3000).catch(() => {});
    await pick(js, 'Rectangle');
    const drew = await putShapeDown(wc, { x: 0.1, y: 0.2, to: { x: 0.4, y: 0.6 } });
    await until(() => (model(1).slide?.shapes || []).length === 1, 'the rectangle drawn', 5000).catch(() => {});
    const g = (model(1).slide?.shapes || [])[0]?.geometry || null;
    const near = (a, b) => Math.abs(a - b) <= Math.max(3, b * 0.02);
    check('presentations: a shape chosen from Shapes waits to be drawn, Esc puts the choice away, and a drag draws it in the box dragged',
      waiting && putAway && drew === 'put' && g && near(g.x, size.width * 0.1) && near(g.y, size.height * 0.2) && near(g.w, size.width * 0.3) && near(g.h, size.height * 0.4),
      JSON.stringify({ waiting, putAway, drew, g, size }));
    if (process.env.RUTBA_VERIFY_CAPTURE) { wc.invalidate(); await wait(600); fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-drawn-shape.png'), (await wc.capturePage()).toPNG()); }

    const id = (model(1).slide?.shapes || [])[0]?.id;
    await tab(js, 'Home');
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Shape Fill' && !n.disabled))`), 'Shape Fill', 3000).catch(() => {});
    await ribbon(js, 'Shape Fill');
    await until(async () => (await items(js)).includes('Red'), 'the fill colours', 3000).catch(() => {});
    const fillMenu = await js(`({ items: [...document.querySelectorAll('.rw-menu button')].map((b) => b.textContent.trim()), swatches: document.querySelectorAll('.rw-menu button span[style*="background"]').length })`);
    await pick(js, 'Red');
    await until(() => /^#?ff0000$/i.test(model(1).slide.shapes.find((s) => s.id === id)?.fill?.color || ''), 'the red fill', 4000).catch(() => {});
    await ribbon(js, 'Shape Outline');
    await until(async () => (await items(js)).includes('Blue'), 'the outline colours', 3000).catch(() => {});
    await pick(js, 'Blue');
    await until(() => /^#?0070c0$/i.test(model(1).slide.shapes.find((s) => s.id === id)?.line?.color || ''), 'the blue outline', 4000).catch(() => {});
    const styled = model(1).slide.shapes.find((s) => s.id === id);
    check('presentations: Shape Fill and Shape Outline offer colours at once, each with its swatch, and More Options for the Format pane',
      fillMenu.items.includes('No Fill') && fillMenu.items.includes('More Fill Options…') && fillMenu.swatches >= 10
        && /^#?ff0000$/i.test(styled?.fill?.color || '') && /^#?0070c0$/i.test(styled?.line?.color || ''),
      JSON.stringify({ fillMenu: { ...fillMenu, items: fillMenu.items.slice(0, 4) }, fill: styled?.fill, line: styled?.line }));

    /* -- Design Ideas, one after another ------------------------------------ */
    await js(`document.querySelectorAll('.sl-thumb')[0].click(), 1`);
    await until(() => js(`[...document.querySelectorAll('.sl-thumb')].findIndex((t) => t.classList.contains('active')) === 0`), 'the first slide', 4000).catch(() => {});
    const bands = () => (model(0).slide?.shapes || []).filter((s) => s.name === 'Design idea').length;
    await tab(js, 'Design');
    await ribbon(js, 'Design Ideas');
    await until(() => js(`Boolean(document.querySelector('.sl-idea[data-idea="title-band"]')) && Boolean(document.querySelector('.sl-idea[data-idea="side-bar"]'))`), 'the ideas', 8000).catch(() => {});
    await js(`document.querySelector('.sl-idea[data-idea="title-band"]').click(), 1`);
    await until(() => bands() === 1, 'the title band', 4000).catch(() => {});
    await wait(300);
    await js(`document.querySelector('.sl-idea[data-idea="side-bar"]').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.sl-idea.on[data-idea="side-bar"]'))`), 'the side bar marked', 4000).catch(() => {});
    await wait(300);
    const after = { bands: bands(), listed: await js(`[...document.querySelectorAll('.sl-idea')].map((n) => n.dataset.idea)`), on: await js(`[...document.querySelectorAll('.sl-idea.on')].map((n) => n.dataset.idea)`) };
    win.focus();
    await js(`document.activeElement?.blur?.(), 1`);
    await press(wc, 'Z', { modifiers: ['control'] });
    await until(() => bands() === 0, 'the slide as it was', 4000).catch(() => {});
    const undone = bands();
    check('presentations: a second Design Idea tried takes the place of the first, the ideas stay those of the slide, and one Ctrl+Z gives the slide back',
      after.bands === 1 && after.listed.includes('title-band') && after.on.join() === 'side-bar' && undone === 0,
      JSON.stringify({ after, undone }));

    /* -- Presenter View: Esc ends the show for both windows ------------------ */
    await tab(js, 'Slide Show');
    const tips = await js(`(() => { const tip = (l) => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === l); return b ? (b.dataset.tip || b.title || '') : null; }; return { presenter: tip('Use Presenter View'), narrations: tip('Play Narrations') }; })()`);
    const before = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
    await ribbon(js, 'Use Presenter View');
    await until(() => BrowserWindow.getAllWindows().some((w) => !before.has(w.id)), 'the presenter window', 5000).catch(() => {});
    const presenter = BrowserWindow.getAllWindows().find((w) => !before.has(w.id));
    if (presenter) opened.push(presenter);
    const presenting = () => js(`Boolean(document.querySelector('.sl-present'))`);
    await until(() => presenting(), 'the show', 5000).catch(() => {});
    const screens = screen.getAllDisplays().length;
    const said = await js(`[...document.querySelectorAll('.rw-toast')].map((n) => n.textContent).join(' | ')`);
    await wait(400);
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' });
    await until(async () => !(await presenting()), 'the show to end', 4000).catch(() => {});
    const startShown = () => presenter ? presenter.webContents.executeJavaScript(`[...document.querySelectorAll('.rw-btn')].some((b) => b.textContent.trim() === 'Start')`) : Promise.resolve(false);
    const presenterStopped = await until(() => startShown(), 'the presenter told the show ended', 4000).then(() => true).catch(() => false);
    // A word from the presenter (its clock reset) does not start the show again.
    if (presenter) await presenter.webContents.executeJavaScript(`[...document.querySelectorAll('.rw-btn')].find((b) => b.textContent.trim() === 'Reset')?.click(), 1`);
    await wait(1200);
    const stillEnded = !(await presenting());
    await until(() => !win.isFullScreen() && !win.rutbaCheckFull, 'the window back', 3000).catch(() => {});
    check('presentations: Use Presenter View says it starts the show, Esc in the show ends it for the presenter too and it stays ended; with one screen, the way between the windows is said',
      /starts the show/.test(tips.presenter || '') && /Record/.test(tips.narrations || '') && Boolean(presenter) && presenterStopped && stillEnded
        && (screens > 1 || /Alt\+Tab/.test(said)),
      JSON.stringify({ tips, presenter: Boolean(presenter), presenterStopped, stillEnded, screens, said: said.slice(0, 120) }));

    const complaints = [...await errorsIn(win), ...await errorsIn(pitch)];
    check('presentations: the slide tools checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('presentations: the slide tools checks ran', false, err?.message || String(err));
  } finally {
    for (const w of opened) if (w && !w.isDestroyed() && w.webContents.getURL().includes('presenter')) w.destroy();
  }
}
