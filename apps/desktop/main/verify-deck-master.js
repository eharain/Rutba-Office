// Presentation: View → Slide Master.
//
// On a deck of its own (three slides on two layouts, written beside the
// shared fixtures): Slide Master view opens on the current slide's layout
// with the master above it in the strip and the layouts indented under it,
// each drawn with its prompts; the master's title placeholder takes a
// colour and a size from the Home tab, which is the master's title style,
// and every slide's title follows; a band drawn on the master sits under
// every slide; Insert Layout, Rename, Title, Footers and Delete work on a
// new layout; Hide Background Graphics takes the band off one layout's
// slides and Undo puts it back; Close Master View returns to the slides,
// redrawn. The saved file carries the master's edits. Run alone with
// RUTBA_VERIFY_ONLY=master.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';
import { consoleMessage } from './console-message.js';

const MASTER = 'ppt/slideMasters/slideMaster1.xml';
const CONTENT = 'ppt/slideLayouts/slideLayout2.xml';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where to write the block's own deck
 */
export async function verifyDeckMaster(h, { dir }) {
  const { open, check, until, wait, doc, sessionFor } = h;
  try {
    const file = path.join(dir, 'master.pptx');
    fs.writeFileSync(file, buildPptx({
      title: 'Master checks',
      slides: [
        { layout: 'title', title: 'Annual report', body: 'Twenty twenty-six' },
        { layout: 'obj', title: 'Highlights', body: ['Revenue grew', 'Costs held'] },
        { layout: 'obj', title: 'Outlook', body: ['Two launches', 'One new market'] },
      ],
    }));
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const consoleErrors = [];
    wc.on('console-message', (...args) => { const m = consoleMessage(args); if (m.level >= 3) consoleErrors.push(m.text.split('\n')[0].slice(0, 160)); });
    const id = () => sessionFor('deck').id;
    const model = (slide = 0, master = null) => doc.model({ id: id(), slide, master });
    const masterList = () => (model(0, MASTER).masterView?.items || []);
    const capture = async (name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      wc.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
    };
    const tab = (label) => js(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(label)}); if (!t) return 'no tab'; t.click(); return 'tab'; })()`);
    const clickRibbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no button ' + ${JSON.stringify(title)};
      b.click();
      return 'clicked';
    })()`);
    const clickMenuItem = async (label) => {
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)} && !b.disabled))`), `the "${label}" menu item`, 3000).catch(() => {});
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'missing'; b.click(); return 'clicked'; })()`);
    };
    const clickIn = (selector) => js(`(() => { const b = document.querySelector(${JSON.stringify(selector)}); if (!b) return 'missing ' + ${JSON.stringify(selector)}; b.click(); return 'clicked'; })()`);
    const stageHas = (needle) => js(`(document.querySelector('.sl-svg')?.innerHTML || '').includes(${JSON.stringify(needle)})`);
    const focusStage = () => js(`(() => { document.querySelector('.sl-stage')?.focus(); return 1; })()`);
    const key = (keyCode, modifiers = []) => {
      wc.sendInputEvent({ type: 'keyDown', keyCode, modifiers });
      wc.sendInputEvent({ type: 'keyUp', keyCode, modifiers });
    };
    const inFile = () => { try { return Deck.open(fs.readFileSync(file)); } catch { return null; } };

    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 3`), 'the slide sorter', 8000);
    await js(`(() => { document.querySelectorAll('.sl-thumb')[1]?.click(); return 1; })()`);
    await until(() => js(`document.querySelectorAll('.sl-thumb')[1]?.classList.contains('active')`), 'slide 2', 4000).catch(() => {});

    // 1. View → Slide Master: the strip, the tab, the current slide's layout on the stage.
    await tab('View');
    const opened = await clickRibbon('Slide Master');
    const strip = await until(() => js(`document.querySelectorAll('.sl-mthumb-master').length === 1 && document.querySelectorAll('.sl-mthumb-layout').length === 2 && document.querySelectorAll('.sl-mthumb svg').length === 3`), 'the master strip', 8000).catch(() => false);
    const onLayout = await js(`document.querySelector('.sl-mthumb.active')?.dataset.part || null`);
    const masterTab = await js(`[...document.querySelectorAll('.rw-tab')].some((t) => t.textContent.trim() === 'Slide Master')`);
    const prompts = await until(() => stageHas('Click to edit Master title style'), 'the prompts on the stage', 6000).catch(() => false);
    const indented = await js(`(() => { const m = document.querySelector('.sl-mthumb-master .sl-thumb-card')?.getBoundingClientRect(); const l = document.querySelector('.sl-mthumb-layout .sl-thumb-card')?.getBoundingClientRect(); return m && l ? Math.round(l.left - m.left) : null; })()`);
    check('slides: View → Slide Master shows the master with its layouts indented under it, each drawn with its prompts, on the current slide\'s layout',
      opened === 'clicked' && strip === true && onLayout === CONTENT && masterTab && prompts === true && indented > 8,
      `${opened}; strip ${strip}; on ${onLayout}; tab ${masterTab}; prompts ${prompts}; indent ${indented}px`);
    await capture('slides-master-view.png');

    // 2. The master: its title placeholder's colour and size are the master's title style.
    await clickIn(`.sl-mthumb[data-part="${MASTER}"]`);
    await until(() => stageHas('Fifth level'), 'the master on the stage', 6000).catch(() => false);
    const titleId = model(0, MASTER).slide.shapes.find((s) => s.placeholder?.type === 'title')?.id;
    await until(() => js(`Boolean(document.querySelector('.sl-hit[data-shape="${titleId}"]'))`), 'the master\'s title on the stage', 5000).catch(() => {});
    await js(`(() => { document.querySelector('.sl-hit[data-shape="${titleId}"]')?.click(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sl-hit.selected[data-shape="${titleId}"]'))`), 'the title selected', 3000).catch(() => {});
    await tab('Home');
    await wait(250);
    const colour = await clickRibbon('Font colour');
    const red = await clickMenuItem('Dark red');
    const styled = await until(() => model(1).slide.shapes.find((s) => s.placeholder?.type === 'title')?.textDefaults?.color === '#c00000', 'every slide\'s title in red', 6000).catch(() => false);
    await wait(300);
    const bigger = await clickRibbon('Increase font size');
    const sized = await until(() => model(1).slide.shapes.find((s) => s.placeholder?.type === 'title')?.textDefaults?.size === 48, 'the title style a step bigger', 6000).catch(() => false);
    const drawnRed = await until(() => stageHas('#c00000'), 'the master drawn red', 6000).catch(() => false);
    check('slides: in Slide Master view, Font colour and Increase font size on the master\'s title write its title style — every slide\'s title follows',
      colour === 'clicked' && red === 'clicked' && styled === true && bigger === 'clicked' && sized === true && drawnRed === true,
      `${colour}/${red}/${bigger}; title ${JSON.stringify(model(1).slide.shapes.find((s) => s.placeholder?.type === 'title')?.textDefaults)}`);

    // 3. A band on the master sits under every slide.
    await js(`(() => { document.querySelector('.sl-stage')?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); return 1; })()`);
    await tab('Insert');
    const shapes = await clickRibbon('Shapes');
    const rect = await clickMenuItem('Rectangle');
    const banded = await until(() => [0, 1, 2].every((i) => (model(i).slide.underlay || []).length === 1), 'the master\'s rectangle under every slide', 6000).catch(() => false);
    check('slides: a shape inserted on the master is drawn under every slide, not added to one', shapes === 'clicked' && rect === 'clicked' && banded === true && model(1).slide.shapes.length === 2,
      `${shapes}/${rect}; underlay ${[0, 1, 2].map((i) => (model(i).slide.underlay || []).length).join(',')}; slide 2 shapes ${model(1).slide.shapes.length}`);
    await capture('slides-master-edited.png');

    // 4. Insert Layout, Rename, Title off, Footers on, Delete.
    await tab('Slide Master');
    const inserted = await clickRibbon('Insert Layout');
    const added = await until(() => js(`document.querySelectorAll('.sl-mthumb-layout').length === 3 && document.querySelector('.sl-mthumb.active')?.dataset.part === 'ppt/slideLayouts/slideLayout3.xml'`), 'the new layout', 6000).catch(() => false);
    const renamed1 = await clickRibbon('Rename');
    await until(() => js(`Boolean(document.querySelector('.sl-partname'))`), 'the rename dialog', 3000).catch(() => {});
    await js(`(() => { const el = document.querySelector('.sl-partname'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, 'Quote slide'); el.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await clickIn('.sl-partname-ok');
    const renamed = await until(() => masterList().some((it) => it.name === 'Quote slide'), 'the new name', 5000).catch(() => false);
    const titleOff = await clickIn('.sl-master-title');
    const noTitle = await until(() => masterList().find((it) => it.part === 'ppt/slideLayouts/slideLayout3.xml')?.hasTitle === false, 'the title off', 5000).catch(() => false);
    await wait(300);
    const footersOn = await js(`(() => { const b = document.querySelector('.sl-master-footers'); if (!b) return 'missing'; if (!b.checked) b.click(); return 'clicked'; })()`);
    const footers = await until(() => masterList().find((it) => it.part === 'ppt/slideLayouts/slideLayout3.xml')?.hasFooters === true, 'the footers', 5000).catch(() => false);
    await capture('slides-master-layout.png');
    await wait(300);
    const deleted = await clickRibbon('Delete');
    const gone = await until(() => js(`document.querySelectorAll('.sl-mthumb-layout').length === 2`), 'the layout deleted', 5000).catch(() => false);
    check('slides: Insert Layout adds a layout under the master; Rename, Title and Footers change it; Delete takes it away again',
      inserted === 'clicked' && added === true && renamed1 === 'clicked' && renamed === true && titleOff === 'clicked' && noTitle === true && footersOn === 'clicked' && footers === true && deleted === 'clicked' && gone === true,
      `${inserted}/${renamed1}/${titleOff}/${footersOn}/${deleted}; added ${added}; renamed ${renamed}; title ${noTitle}; footers ${footers}; gone ${gone}`);

    // 5. Hide Background Graphics on the content layout, then Undo.
    await clickIn(`.sl-mthumb[data-part="${CONTENT}"]`);
    await until(() => js(`document.querySelector('.sl-mthumb.active')?.dataset.part === ${JSON.stringify(CONTENT)}`), 'the content layout', 4000).catch(() => {});
    await wait(300);
    const hide = await clickIn('.sl-master-hidebg');
    const hidden = await until(() => (model(1).slide.underlay || []).length === 0 && (model(0).slide.underlay || []).length === 1, 'the band hidden on the content layout\'s slides', 5000).catch(() => false);
    await wait(300);
    await focusStage();
    key('Z', ['control']);
    const back = await until(() => (model(1).slide.underlay || []).length === 1, 'the band back after Undo', 5000).catch(() => false);
    check('slides: Hide Background Graphics takes the master\'s band off the content layout\'s slides only, and Ctrl+Z puts it back',
      hide === 'clicked' && hidden === true && back === true, `${hide}; hidden ${hidden}; back ${back}`);

    // 6. Close Master View: the slides again, redrawn with the master's edits.
    await tab('Slide Master');
    const closed = await clickRibbon('Close Master View');
    const slidesBack = await until(() => js(`document.querySelectorAll('.sl-mthumb').length === 0 && document.querySelectorAll('.sl-thumb').length === 3`), 'the slides', 5000).catch(() => false);
    const thumbsRed = await until(async () => (await js(`[...document.querySelectorAll('.sl-thumb .sl-thumb-pic')].filter((p) => /#c00000/i.test(p.innerHTML)).length`)) === 3, 'the thumbnails redrawn', 8000).catch(() => false);
    const noTab = await js(`![...document.querySelectorAll('.rw-tab')].some((t) => t.textContent.trim() === 'Slide Master')`);
    check('slides: Close Master View returns to the slides, every thumbnail redrawn with the red title and the master\'s band, and the Slide Master tab goes',
      closed === 'clicked' && slidesBack === true && thumbsRed === true && noTab, `${closed}; slides ${slidesBack}; thumbs ${thumbsRed}; tab gone ${noTab}`);
    await capture('slides-master-closed.png');

    // 7. Saved: the master's title style and band are in the file.
    await focusStage();
    key('S', ['control']);
    const saved = await until(() => /<p:titleStyle><a:lvl1pPr[^>]*><a:defRPr[^>]*sz="4800"[^>]*><a:solidFill><a:srgbClr val="C00000"\/>/.test(inFile()?.pkg.text(MASTER) || ''), 'the saved master', 8000).catch(() => false);
    const reopened = inFile();
    check('slides: saved, the file carries the master\'s title style, its band and its two layouts, and opens again',
      saved === true && reopened?.slide(1).underlay.length === 1 && reopened?.masterList()[0].layouts.length === 2,
      `saved ${saved}; underlay ${reopened?.slide(1).underlay.length}; layouts ${reopened?.masterList()[0].layouts.length}`);
    check('slides: the Slide Master checks raised no errors in the window', consoleErrors.length === 0, consoleErrors.slice(0, 3).join(' | '));
  } catch (err) {
    check('slides: the Slide Master checks ran', false, err.stack || err.message);
  }
}
