// Presentation: Transitions — authored from the ribbon, played in the show.
//
// On the two-slide fixture: Push is picked from the Transition to This Slide
// gallery for the second slide, its Effect Options and Duration set, and
// each read back from the engine's own model and, once saved, from the
// file's raw XML (an mc:AlternateContent with p14:dur, after clrMapOvr).
// The strip marks the slide with PowerPoint's little star, Preview plays on
// the stage and goes, Apply To All puts the transition on both slides and
// None takes it off again. The first slide is given Advance Slide → After
// one second, so the show started from the beginning moves on by itself —
// and the second slide is seen coming in through Push (the incoming layer
// carries the transition and is animating) before it settles. Esc leaves
// the show; every transition comes off and the deck is saved as found.
// Run alone with RUTBA_VERIFY_ONLY=transitions.

import fs from 'node:fs';
import path from 'node:path';
import { Deck } from '@rutba/presentation';
import { consoleMessage } from './console-message.js';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ file: string }} args the deck fixture (files.pptx)
 */
export async function verifyDeckTransitions(h, { file }) {
  const { open, check, until, wait, errorsIn, doc, sessionFor } = h;
  if (!file) return check('slides: the transition checks have a fixture', false, 'no files.pptx');
  try {
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const consoleErrors = [];
    wc.on('console-message', (...args) => { const m = consoleMessage(args); if (m.level >= 3) consoleErrors.push(m.text.split('\n')[0].slice(0, 160)); });
    const id = () => sessionFor('deck').id;
    const model = (slide) => doc.model({ id: id(), slide });
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
    const clickMenuItem = async (label) => {
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)} && !b.disabled))`), `the "${label}" menu item`, 3000).catch(() => {});
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'missing'; b.click(); return 'clicked'; })()`);
    };
    const pressed = (key) => js(`document.querySelector('.sl-tr-pick[data-transition="${key}"]')?.getAttribute('aria-pressed')`);
    // A number typed into a ribbon field and committed with Enter, the way a person does.
    const setSeconds = (selector, value) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el || el.disabled) return 'no field ' + ${JSON.stringify(selector)};
      el.focus();
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, ${JSON.stringify(String(value))});
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
      return 'set';
    })()`);
    const tick = (selector, on) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return 'no box ' + ${JSON.stringify(selector)};
      if (el.checked !== ${on ? 'true' : 'false'}) el.click();
      return 'ticked';
    })()`);
    const selectSlide = async (i) => {
      await js(`(() => { document.querySelectorAll('.sl-thumb')[${i}]?.click(); return 1; })()`);
      await until(() => js(`document.querySelectorAll('.sl-thumb')[${i}]?.classList.contains('active')`), `slide ${i + 1}`, 4000).catch(() => {});
      await until(async () => (await js(`document.querySelector('.sl-slide') ? 1 : 0`)) === 1, 'the stage', 3000).catch(() => {});
      await wait(250);
    };
    const inFile = () => { try { return Deck.open(fs.readFileSync(file)); } catch { return null; } };

    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slide sorter', 8000);
    // The show below moves from the first slide to the second, which an
    // earlier block may have left hidden (Slide Show → Hide Slide): shown
    // for these checks, and hidden again before the last save.
    const wasHidden = model(1).slide?.hidden === true;
    if (wasHidden) await doc.apply({ id: id(), ops: [{ op: 'setSlideHidden', slide: 1, hidden: false }] });
    await selectSlide(1);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Transitions')?.click(), 'tab'`);
    await until(() => js(`document.querySelectorAll('.sl-tr-pick').length >= 10`), 'the transition gallery', 4000).catch(() => {});
    const gallery = await js(`[...document.querySelectorAll('.sl-tr-pick')].map((b) => b.textContent.trim())`);
    const nonePressed = await pressed('none');

    // Push, from the gallery: PowerPoint's own first pick — From Bottom, one second.
    const picked = await clickRibbon('Push');
    const pushed = await until(() => model(1).slide.transition?.type === 'push', 'the Push transition', 5000).catch(() => false);
    const t1 = model(1).slide.transition;
    const pushPressed = await until(async () => (await pressed('push')) === 'true', 'Push to show pressed', 4000).catch(() => false);
    check('slides: Transitions → Push writes a push transition, From Bottom for one second, and the gallery shows Push pressed (None was, before)',
      picked === 'clicked' && pushed === true && t1?.direction === 'u' && t1?.duration === 1 && pushPressed === true && nonePressed === 'true' && gallery.length === 11,
      `${picked}; ${JSON.stringify(t1)}; pressed ${pushPressed}; none before ${nonePressed}; gallery ${gallery.join('/')}`);

    // Effect Options → From Left, then Duration 1.5 s.
    await wait(300);
    const opened = await clickRibbon('Effect Options');
    const option = await clickMenuItem('From Left');
    const turned = await until(() => model(1).slide.transition?.direction === 'r', 'the direction', 4000).catch(() => false);
    await wait(300);
    const typed = await setSeconds('.sl-tr-duration', 1.5);
    const timed = await until(() => model(1).slide.transition?.duration === 1.5, 'the duration', 4000).catch(() => false);
    const shownDuration = await until(async () => (await js(`document.querySelector('.sl-tr-duration')?.value`)) === '1.50', 'the field to show 1.50', 3000).catch(() => false);
    check('slides: Effect Options → From Left and Duration 1.5 s change the push, and the field shows the new duration',
      opened === 'clicked' && option === 'clicked' && turned === true && typed === 'set' && timed === true && shownDuration === true,
      `${opened}/${option}/${typed}; ${JSON.stringify(model(1).slide.transition)}`);

    // The strip's star: on the second slide, not the first.
    const starred = await until(() => js(`Boolean(document.querySelectorAll('.sl-thumb')[1]?.querySelector('.sl-thumb-fx')) && !document.querySelectorAll('.sl-thumb')[0]?.querySelector('.sl-thumb-fx')`), 'the star in the strip', 4000).catch(() => false);
    check('slides: the strip marks the slide that has a transition with a star under its number', starred === true, `star on slide 2 only: ${starred}`);
    await until(() => js(`!document.querySelector('.sl-preview')`), 'the pick\'s own preview to end', 5000).catch(() => {});
    await capture('slides-transitions.png');

    // Preview plays on the stage, then goes.
    const previewed = await clickRibbon('Preview');
    const appeared = await until(() => js(`Boolean(document.querySelector('.sl-preview[data-transition="push"]'))`), 'the preview', 3000).catch(() => false);
    const gone = await until(() => js(`!document.querySelector('.sl-preview')`), 'the preview to end', 5000).catch(() => false);
    check('slides: Transitions → Preview plays the push over the stage and ends by itself, reporting nothing',
      previewed === 'clicked' && appeared === true && gone === true && consoleErrors.length === 0,
      `${previewed}; shown ${appeared}; ended ${gone}; console ${consoleErrors.join(' | ') || 'clean'}`);

    // Saved: the file carries the transition in PowerPoint's p14 form, after clrMapOvr.
    await clickRibbon('Save');
    const savedOk = await until(() => inFile()?.transition(1)?.duration === 1.5, 'the transition in the file', 8000).catch(() => false);
    const d = inFile();
    const xml1 = d ? d.pkg.text(d.slideParts[1].part) : '';
    const block = /<mc:AlternateContent[\s\S]*?<\/mc:AlternateContent>/.exec(xml1)?.[0] || '';
    check('slides: the saved file carries the push as PowerPoint writes it — mc:AlternateContent, p14:dur="1500" beside a speed-only fallback, after p:clrMapOvr',
      savedOk === true && /<mc:Choice[^>]*Requires="p14"><p:transition spd="slow" p14:dur="1500"><p:push dir="r"\/><\/p:transition><\/mc:Choice><mc:Fallback><p:transition spd="slow"><p:push dir="r"\/><\/p:transition><\/mc:Fallback>/.test(block) && xml1.indexOf('</p:clrMapOvr>') < xml1.indexOf('<mc:AlternateContent'),
      block.slice(0, 260) || 'no transition in the file');

    // Apply To All: both slides push; then None on the first takes its own off again.
    await wait(300);
    const all = await clickRibbon('Apply To All');
    const both = await until(() => model(0).slide.transition?.type === 'push' && model(0).slide.transition?.direction === 'r', 'the push on slide 1', 4000).catch(() => false);
    const bothStars = await until(() => js(`document.querySelectorAll('.sl-thumb .sl-thumb-fx').length === 2`), 'two stars', 4000).catch(() => false);
    await selectSlide(0);
    await clickRibbon('None');
    const noneAgain = await until(() => model(0).slide.transition === null, 'slide 1 back to none', 4000).catch(() => false);
    check('slides: Apply To All puts the transition on every slide, and None takes this slide\'s off again',
      all === 'clicked' && both === true && bothStars === true && noneAgain === true,
      `${all}; slide 1 had ${both}; stars ${bothStars}; none ${noneAgain}`);

    // Advance Slide → After 1 s on the first slide, so the show moves on by itself.
    await wait(300);
    const ticked = await tick('.sl-tr-after-on', true);
    await until(() => model(0).slide.transition?.advanceAfter === 0, 'After ticked', 4000).catch(() => {});
    await wait(300);
    await setSeconds('.sl-tr-after', 1);
    const after = await until(() => model(0).slide.transition?.advanceAfter === 1, 'After 1 s', 4000).catch(() => false);
    check('slides: Advance Slide → After 1 s writes a timing-only transition on the first slide', ticked === 'ticked' && after === true, JSON.stringify(model(0).slide.transition));

    // The show: starts on slide 1, moves on after a second, slide 2 comes in through Push.
    await wait(300);
    const started = await clickRibbon('Start from the beginning');
    const onFirst = await until(() => js(`document.querySelector('.sl-show-layer')?.dataset.slide === '0'`), 'the show on slide 1', 5000).catch(() => false);
    const animating = await until(() => js(`(() => { const l = document.querySelector('.sl-show-layer.in[data-transition="push"]'); return Boolean(l && l.getAnimations().length > 0 && document.querySelector('.sl-show-layer.out')); })()`), 'slide 2 coming in through Push', 6000).catch(() => false);
    if (animating === true && process.env.RUTBA_VERIFY_CAPTURE) {
      await wait(500);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-transition-show.png'), (await wc.capturePage()).toPNG());
    }
    const settled = await until(() => js(`document.querySelectorAll('.sl-show-layer').length === 1 && document.querySelector('.sl-show-layer')?.dataset.slide === '1' && !document.querySelector('.sl-show-layer.in')`), 'slide 2 to settle', 6000).catch(() => false);
    check('slides: the show moves on by itself after a second, and the next slide comes in through its Push — the incoming layer animating over the outgoing one — then settles',
      started === 'clicked' && onFirst === true && animating === true && settled === true,
      `${started}; on slide 1 ${onFirst}; push seen animating ${animating}; settled on slide 2 ${settled}`);

    await js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))`);
    const left = await until(() => js(`!document.querySelector('.sl-present')`), 'the show to close', 5000).catch(() => false);
    check('slides: Esc leaves the show', left === true, `left ${left}`);

    // Every other effect in the gallery, picked on the second slide: each is
    // written as itself and plays its own preview without a complaint.
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Transitions')?.click(), 'tab'`);
    await selectSlide(1);
    const expected = { Cut: 'cut', Fade: 'fade', Wipe: 'wipe', Split: 'split', Uncover: 'pull', Cover: 'cover', 'Random Bars': 'randomBar', Shape: 'circle', Dissolve: 'dissolve' };
    const played = [];
    for (const [label, type] of Object.entries(expected)) {
      await wait(250);
      await clickRibbon(label);
      const written = await until(() => model(1).slide.transition?.type === type, `${label} written`, 4000).catch(() => false);
      // Cut is a tenth of a second: it may be over before a poll sees it.
      const seen = type === 'cut' ? true : await until(() => js(`[...document.querySelectorAll('.sl-preview[data-transition="${type}"] [data-layer]')].some((l) => l.getAnimations().length > 0)`), `${label} previewing`, 3000).catch(() => false);
      if (type === 'dissolve' && seen === true && process.env.RUTBA_VERIFY_CAPTURE) {
        await wait(350);
        fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-transition-dissolve.png'), (await wc.capturePage()).toPNG());
      }
      played.push(`${label} ${written === true && seen === true ? 'ok' : `written ${written}, seen ${seen}`}`);
      await until(() => js(`!document.querySelector('.sl-preview')`), `${label} preview to end`, 5000).catch(() => {});
    }
    check('slides: every effect in the gallery is written as itself and plays its preview, with nothing in the console',
      played.every((p) => p.endsWith(' ok')) && consoleErrors.length === 0,
      `${played.join('; ')}; console ${consoleErrors.join(' | ') || 'clean'}`);

    // Clean up: every transition off, saved, the file as it was found.
    await doc.apply({ id: id(), ops: [{ op: 'setTransition', slide: 0, spec: null }, { op: 'setTransition', slide: 1, spec: null }, ...(wasHidden ? [{ op: 'setSlideHidden', slide: 1, hidden: true }] : [])] });
    await selectSlide(1);
    await selectSlide(0);
    await clickRibbon('Save');
    const clean = await until(() => { const f = inFile(); return Boolean(f) && f.transition(0) === null && f.transition(1) === null && f.isSlideHidden(1) === wasHidden; }, 'the deck without transitions', 8000).catch(() => false);
    const noStars = await until(() => js(`document.querySelectorAll('.sl-thumb .sl-thumb-fx').length === 0`), 'the stars to go', 4000).catch(() => false);
    check('slides: the transition checks leave the deck as they found it', clean === true && noStars === true, `file clean ${clean}; stars gone ${noStars}`);

    const complaints = await errorsIn(win);
    check('slides: the transition checks report nothing', complaints.length === 0 && consoleErrors.length === 0, [...complaints, ...consoleErrors].join(' | ') || 'nothing reported');
  } catch (err) {
    check('slides: the transition checks ran', false, err.message);
  }
}
