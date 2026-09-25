// Presentation: Animations — authored from the ribbon, listed in the
// Animation Pane, numbered on the stage and played in the show.
//
// On the fixture's first slide: Fade is picked for the title and Fly In for
// the subtitle from the Animation gallery, each read back from the engine's
// model; the Animation Pane lists them in order and the stage wears
// PowerPoint's numbered tags beside both shapes. Fly In is moved earlier,
// the Fade set to start With Previous (both tags then read 1), Preview plays
// the sequence over the stage, and the saved file carries PowerPoint's
// timing tree. In the show both shapes start hidden and a click brings them
// in; with nothing left to play, the next click moves to the next slide.
// Both effects are taken out again from the pane and the deck is saved
// exactly as it was found. Run alone with RUTBA_VERIFY_ONLY=animations.

import fs from 'node:fs';
import path from 'node:path';
import { Deck } from '@rutba/presentation';
import { consoleMessage } from './console-message.js';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ file: string }} args the deck fixture (files.pptx)
 */
export async function verifyDeckAnimations(h, { file }) {
  const { open, check, until, wait, errorsIn, doc, sessionFor } = h;
  if (!file) return check('slides: the animation checks have a fixture', false, 'no files.pptx');
  try {
    const original = (() => { try { const d = Deck.open(fs.readFileSync(file)); return d.pkg.text(d.slideParts[0].part); } catch { return null; } })();
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const consoleErrors = [];
    wc.on('console-message', (...args) => { const m = consoleMessage(args); if (m.level >= 3) consoleErrors.push(m.text.split('\n')[0].slice(0, 160)); });
    const id = () => sessionFor('deck').id;
    const model = (slide = 0) => doc.model({ id: id(), slide });
    const anims = () => model(0).slide.animations || [];
    const capture = async (name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      wc.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
    };
    const clickRibbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no button ' + ${JSON.stringify(title)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const clickShape = (shapeId) => js(`(() => {
      const el = document.querySelector('.sl-hit[data-shape="${shapeId}"]');
      if (!el) return 'no shape ${shapeId}';
      el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true }));
      el.click();
      return 'clicked';
    })()`);
    const rows = () => js(`[...document.querySelectorAll('.sl-animrow')].map((r) => r.querySelector('.sl-layer-title')?.textContent + ' / ' + r.querySelector('.sl-layer-words')?.textContent)`);
    const badges = () => js(`[...document.querySelectorAll('.sl-anim-badge')].map((b) => b.dataset.shape + ':' + b.textContent.trim())`);
    const inFile = () => { try { return Deck.open(fs.readFileSync(file)); } catch { return null; } };

    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slide sorter', 8000);
    // The show below leaves the first slide for the second, which an
    // earlier block may have left hidden: shown for these checks (the
    // window told by a trip to that slide and back), hidden again before
    // the last save.
    const wasHidden = model(1).slide?.hidden === true;
    if (wasHidden) await doc.apply({ id: id(), ops: [{ op: 'setSlideHidden', slide: 1, hidden: false }] });
    await js(`(() => { document.querySelectorAll('.sl-thumb')[1]?.click(); return 1; })()`);
    await until(() => js(`document.querySelectorAll('.sl-thumb')[1]?.classList.contains('active') && !document.querySelectorAll('.sl-thumb')[1]?.classList.contains('hidden')`), 'slide 2, shown', 4000).catch(() => {});
    await js(`(() => { document.querySelectorAll('.sl-thumb')[0]?.click(); return 1; })()`);
    await until(() => js(`document.querySelectorAll('.sl-thumb')[0]?.classList.contains('active')`), 'slide 1', 4000).catch(() => {});
    await wait(250);
    const shapes = model(0).slide.shapes.filter((s) => s.geometry && s.groupId == null && !s.hidden);
    const [a, b] = shapes.map((s) => String(s.id));
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Animations')?.click(), 'tab'`);
    await until(() => js(`document.querySelectorAll('.sl-an-pick').length >= 8`), 'the animation gallery', 4000).catch(() => {});

    // Fade on the first shape, Fly In on the second, from the gallery.
    await clickShape(a);
    await until(() => js(`Boolean(document.querySelector('.sl-hit.selected[data-shape="${a}"]'))`), 'the first shape selected', 3000).catch(() => {});
    await wait(200);
    const faded = await clickRibbon('Fade');
    const one = await until(() => anims().length === 1 && anims()[0].effect === 'fade' && String(anims()[0].shapeId) === a, 'the Fade', 5000).catch(() => false);
    await until(() => js(`document.querySelector('.sl-an-pick[data-effect="fade"]')?.getAttribute('aria-pressed') === 'true'`), 'Fade pressed', 3000).catch(() => {});
    await wait(300);
    await clickShape(b);
    await until(() => js(`Boolean(document.querySelector('.sl-hit.selected[data-shape="${b}"]'))`), 'the second shape selected', 3000).catch(() => {});
    await wait(200);
    const flown = await clickRibbon('Fly In');
    const two = await until(() => anims().length === 2 && anims()[1].effect === 'fly' && String(anims()[1].shapeId) === b, 'the Fly In', 5000).catch(() => false);
    const flyPressed = await until(() => js(`document.querySelector('.sl-an-pick[data-effect="fly"]')?.getAttribute('aria-pressed') === 'true'`), 'Fly In pressed', 3000).catch(() => false);
    check('slides: Animations → Fade on one shape and Fly In on another write two entrance effects, one click each, and the gallery marks the selected shape\'s',
      faded === 'clicked' && flown === 'clicked' && one === true && two === true && flyPressed === true && anims().map((e) => `${e.effect}:${e.trigger}:${e.group}`).join() === 'fade:onClick:1,fly:onClick:2',
      `${faded}/${flown}; ${JSON.stringify(anims().map((e) => [e.shapeId, e.effect, e.direction, e.trigger, e.group]))}`);

    // The Animation Pane and the numbers on the stage.
    const paned = await clickRibbon('Animation Pane');
    const listed = await until(async () => (await rows()).length === 2, 'two rows in the pane', 4000).catch(() => false);
    const rowText = await rows();
    const tagged = await until(async () => (await badges()).join() === `${a}:1,${b}:2`, 'the numbers beside the shapes', 4000).catch(() => false);
    check('slides: the Animation Pane lists both effects in order and the stage numbers the shapes 1 and 2',
      paned === 'clicked' && listed === true && /Fade/.test(rowText[0] || '') && /Fly In/.test(rowText[1] || '') && tagged === true,
      `${paned}; rows ${JSON.stringify(rowText)}; tags ${JSON.stringify(await badges())}`);

    // Move Earlier: Fly In (the selected shape's) goes first.
    await wait(300);
    const moved = await clickRibbon('Move Earlier');
    const reordered = await until(() => anims().map((e) => e.effect).join() === 'fly,fade', 'Fly In first', 4000).catch(() => false);
    const rowsAfter = await until(async () => /Fly In/.test((await rows())[0] || ''), 'the pane to follow', 4000).catch(() => false);
    check('slides: Move Earlier puts Fly In first, in the file\'s sequence and in the pane', moved === 'clicked' && reordered === true && rowsAfter === true, `${moved}; ${anims().map((e) => e.effect).join(', ')}; rows ${JSON.stringify(await rows())}`);

    // The Fade, picked in the pane, starts With Previous: one click brings both.
    await js(`(() => { document.querySelector('.sl-animrow[data-anim="1"]')?.click(); return 1; })()`);
    await until(() => js(`document.querySelector('.sl-animrow[data-anim="1"]')?.classList.contains('active')`), 'the Fade row picked', 3000).catch(() => {});
    await until(() => js(`document.querySelector('.sl-an-start')?.disabled === false`), 'the Start box', 3000).catch(() => {});
    await wait(200);
    const started = await js(`(() => {
      const s = document.querySelector('.sl-an-start');
      if (!s) return 'no Start box';
      Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, 'withPrevious');
      s.dispatchEvent(new Event('change', { bubbles: true }));
      return 'set';
    })()`);
    const withPrev = await until(() => anims()[1]?.trigger === 'withPrevious' && anims()[1]?.group === 1, 'With Previous', 4000).catch(() => false);
    const sameTag = await until(async () => { const t = await badges(); return t.length === 2 && t.every((x) => x.endsWith(':1')); }, 'both tags to read 1', 4000).catch(() => false);
    check('slides: Start → With Previous puts the Fade in the first click with Fly In, and both tags read 1',
      started === 'set' && withPrev === true && sameTag === true, `${started}; ${JSON.stringify(anims().map((e) => [e.effect, e.trigger, e.group]))}; tags ${JSON.stringify(await badges())}`);
    await capture('slides-animations.png');

    // Preview plays the sequence over the stage, then goes.
    const previewed = await clickRibbon('Preview');
    const shown = await until(() => js(`Boolean(document.querySelector('.sl-preview-anim'))`), 'the preview', 3000).catch(() => false);
    const moving = await until(() => js(`[...document.querySelectorAll('.sl-preview-anim g[data-shape]')].some((g) => g.getAnimations().length > 0)`), 'a shape moving in the preview', 3000).catch(() => false);
    const ended = await until(() => js(`!document.querySelector('.sl-preview-anim')`), 'the preview to end', 6000).catch(() => false);
    check('slides: Animations → Preview plays the sequence on the stage and ends by itself, reporting nothing',
      previewed === 'clicked' && shown === true && moving === true && ended === true && consoleErrors.length === 0,
      `${previewed}; shown ${shown}; moving ${moving}; ended ${ended}; console ${consoleErrors.join(' | ') || 'clean'}`);

    // Saved: PowerPoint's tree, its preset numbers and build list.
    await clickRibbon('Save');
    const savedOk = await until(() => inFile()?.animations(0)?.length === 2, 'the animations in the file', 8000).catch(() => false);
    const d = inFile();
    const xml0 = d ? d.pkg.text(d.slideParts[0].part) : '';
    const timing = /<p:timing>[\s\S]*<\/p:timing>/.exec(xml0)?.[0] || '';
    check('slides: the saved file carries PowerPoint\'s timing tree — Fly In (preset 2) on the click, Fade (preset 10) with it, a build entry for each',
      savedOk === true && /nodeType="tmRoot"/.test(timing) && /nodeType="mainSeq"/.test(timing) && /presetID="2" presetClass="entr" presetSubtype="4"[^>]*nodeType="clickEffect"/.test(timing) && /presetID="10" presetClass="entr" presetSubtype="0"[^>]*nodeType="withEffect"/.test(timing) && new RegExp(`<p:bldP spid="${a}"`).test(timing) && new RegExp(`<p:bldP spid="${b}"`).test(timing),
      timing ? `${timing.length} bytes of timing; ${(timing.match(/presetID="\d+"/g) || []).join(' ')}` : 'no timing in the file');

    // The show: both shapes hidden until the click, then in; the next click leaves the slide.
    const began = await clickRibbon('Start from the beginning');
    await until(() => js(`document.querySelector('.sl-show-layer')?.dataset.slide === '0'`), 'the show on slide 1', 5000).catch(() => {});
    const visibility = (shapeId) => js(`(() => { const g = document.querySelector('.sl-show-layer:last-child g[data-shape="${shapeId}"]'); return g ? getComputedStyle(g).visibility : 'missing'; })()`);
    const hiddenFirst = await until(async () => (await visibility(a)) === 'hidden' && (await visibility(b)) === 'hidden', 'both shapes hidden', 4000).catch(() => false);
    const before = [await visibility(a), await visibility(b)];
    await wait(400);
    await js(`document.querySelector('.sl-present')?.click(), 'click'`);
    const playing = await until(() => js(`[...document.querySelectorAll('.sl-show-layer g[data-shape]')].some((g) => g.getAnimations().some((x) => x.playState === 'running'))`), 'the shapes animating in', 3000).catch(() => false);
    if (playing === true && process.env.RUTBA_VERIFY_CAPTURE) {
      await wait(150);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-animation-show.png'), (await wc.capturePage()).toPNG());
    }
    const inNow = await until(async () => (await visibility(a)) === 'visible' && (await visibility(b)) === 'visible' && (await js(`[...document.querySelectorAll('.sl-show-layer g[data-shape]')].every((g) => g.getAnimations().every((x) => x.playState !== 'running'))`)), 'both shapes in', 5000).catch(() => false);
    check('slides: in the show both animated shapes start hidden, and a click brings them in together',
      began === 'clicked' && hiddenFirst === true && playing === true && inNow === true, `${began}; before the click ${before.join('/')}; animating ${playing}; after ${await visibility(a)}/${await visibility(b)}`);
    await wait(300);
    await js(`document.querySelector('.sl-present')?.click(), 'click'`);
    const nextSlide = await until(() => js(`document.querySelector('.sl-show-layer:last-child')?.dataset.slide === '1'`), 'the next slide', 5000).catch(() => false);
    check('slides: with nothing left to play, the next click moves the show to the next slide', nextSlide === true, `on slide ${await js(`document.querySelector('.sl-show-layer:last-child')?.dataset.slide`)}`);
    await js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))`);
    await until(() => js(`!document.querySelector('.sl-present')`), 'the show to close', 5000).catch(() => {});

    // Back in the editor: the Fade becomes a Wipe from the gallery, from the
    // left (Effect Options), two seconds long (Duration) — played by Preview.
    await js(`(() => { document.querySelectorAll('.sl-thumb')[0]?.click(); return 1; })()`);
    await until(() => js(`document.querySelectorAll('.sl-animrow').length === 2`), 'the pane again', 5000).catch(() => {});
    await js(`(() => { document.querySelector('.sl-animrow[data-anim="1"]')?.click(); return 1; })()`);
    await until(() => js(`document.querySelector('.sl-animrow[data-anim="1"]')?.classList.contains('active')`), 'the Fade row picked', 3000).catch(() => {});
    await wait(250);
    const wiped = await clickRibbon('Wipe');
    const isWipe = await until(() => anims()[1]?.effect === 'wipe' && anims()[1]?.trigger === 'withPrevious' && String(anims()[1]?.shapeId) === a && anims().length === 2, 'the Wipe in the Fade\'s place', 4000).catch(() => false);
    await until(() => js(`!document.querySelector('.sl-preview-anim')`), 'the pick\'s own preview to end', 4000).catch(() => {});
    await wait(300);
    await clickRibbon('Effect Options');
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'From Left'))`), 'the options', 3000).catch(() => {});
    await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'From Left')?.click(); return 1; })()`);
    const fromLeft = await until(() => anims()[1]?.direction === 'left', 'From Left', 4000).catch(() => false);
    await until(() => js(`!document.querySelector('.sl-preview-anim')`), 'the option\'s own preview to end', 4000).catch(() => {});
    await wait(300);
    const typed = await js(`(() => {
      const el = document.querySelector('.sl-an-duration');
      if (!el || el.disabled) return 'no field';
      el.focus();
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, '2');
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
      return 'set';
    })()`);
    const twoSeconds = await until(() => anims()[1]?.duration === 2, 'two seconds', 4000).catch(() => false);
    await wait(300);
    await clickRibbon('Preview');
    const clipping = await until(() => js(`[...document.querySelectorAll('.sl-preview-anim g[data-shape="${a}"]')].some((g) => g.getAnimations().some((x) => x.playState === 'running' && /inset/.test(getComputedStyle(g).clipPath)))`), 'the wipe clipping the shape', 4000).catch(() => false);
    if (clipping === true && process.env.RUTBA_VERIFY_CAPTURE) {
      await wait(500);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-animation-wipe.png'), (await wc.capturePage()).toPNG());
    }
    await until(() => js(`!document.querySelector('.sl-preview-anim')`), 'the preview to end', 8000).catch(() => {});
    check('slides: another gallery pick replaces the picked effect in place, and Effect Options and Duration change it — a two-second Wipe from the left, played by Preview',
      wiped === 'clicked' && isWipe === true && fromLeft === true && typed === 'set' && twoSeconds === true && clipping === true,
      `${wiped}/${typed}; ${JSON.stringify(anims().map((e) => [e.effect, e.direction, e.duration, e.trigger]))}; clipping ${clipping}`);

    // More Effects: PowerPoint's three headings over their effects.
    await wait(250);
    const more = await clickRibbon('More Effects');
    const heads = await until(async () => (await js(`[...document.querySelectorAll('.rw-menu .rw-menu-head')].map((h) => h.textContent.trim()).join()`)) === 'Entrance,Emphasis,Exit', 'the three headings', 3000).catch(() => false);
    const effectsListed = await js(`[...document.querySelectorAll('.rw-menu button')].map((b) => b.textContent.trim())`);
    await capture('slides-animation-menu.png');
    await js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), 'closed'`);
    await until(() => js(`!document.querySelector('.rw-menu')`), 'the menu to close', 3000).catch(() => {});
    check('slides: More Effects lists every effect under Entrance, Emphasis and Exit',
      more === 'clicked' && heads === true && effectsListed.includes('Spin') && effectsListed.includes('Fly Out') && effectsListed.length === 17,
      `${more}; ${effectsListed.join(', ')}`);

    // Both taken out from the pane: the tags and the strip's star go, the file goes back to how it was.
    for (let k = 0; k < 2; k++) {
      await js(`(() => { document.querySelector('.sl-animrow[data-anim="0"]')?.click(); return 1; })()`);
      await until(() => js(`document.querySelector('.sl-animrow[data-anim="0"]')?.classList.contains('active')`), 'the first row picked', 3000).catch(() => {});
      await wait(200);
      await js(`(() => { const b = document.querySelector('.sl-animpane-remove'); if (b && !b.disabled) b.click(); return 1; })()`);
      await until(() => anims().length === 1 - k, `${1 - k} left`, 4000).catch(() => {});
    }
    const cleared = await until(async () => anims().length === 0 && (await badges()).length === 0 && (await js(`document.querySelectorAll('.sl-animrow').length`)) === 0, 'no animations, no tags, no rows', 4000).catch(() => false);
    const noStar = await until(() => js(`!document.querySelectorAll('.sl-thumb')[0]?.querySelector('.sl-thumb-fx')`), 'the star to go', 4000).catch(() => false);
    check('slides: Remove in the Animation Pane takes both effects out, and the tags and the strip\'s star go with them', cleared === true && noStar === true, `left ${anims().length}; tags ${JSON.stringify(await badges())}; star gone ${noStar}`);

    if (wasHidden) {
      await doc.apply({ id: id(), ops: [{ op: 'setSlideHidden', slide: 1, hidden: true }] });
      await js(`(() => { document.querySelectorAll('.sl-thumb')[1]?.click(); return 1; })()`);
      await until(() => js(`document.querySelectorAll('.sl-thumb')[1]?.classList.contains('hidden')`), 'slide 2 hidden again', 4000).catch(() => {});
      await js(`(() => { document.querySelectorAll('.sl-thumb')[0]?.click(); return 1; })()`);
      await wait(250);
    }
    await clickRibbon('Save');
    const restored = await until(() => { const f = inFile(); return Boolean(f) && f.pkg.text(f.slideParts[0].part) === original && f.isSlideHidden(1) === wasHidden; }, 'the slide as it was found', 8000).catch(() => false);
    check('slides: the animation checks leave the deck exactly as they found it', restored === true, restored ? 'slide 1 byte for byte' : 'slide 1 differs');

    const complaints = await errorsIn(win);
    check('slides: the animation checks report nothing', complaints.length === 0 && consoleErrors.length === 0, [...complaints, ...consoleErrors].join(' | ') || 'nothing reported');
  } catch (err) {
    check('slides: the animation checks ran', false, err.message);
  }
}
